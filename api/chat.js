// TRACE 도우미 챗봇 API (Vercel Function)
// GET  /api/chat → 사용할 수 있는 모델 목록
// POST /api/chat → { messages, model, code } 를 받아 OpenRouter 응답을 줄 단위 JSON으로 스트리밍
//   {"d":"텍스트 조각"} … {"error":"메시지"}
// 환경변수: OPENROUTER_API_KEY(필수), OPENROUTER_MODEL(기본 모델), CHAT_ACCESS_CODE(설정 시 접속 코드 필요)
import { readFile } from 'node:fs/promises';
import path from 'node:path';

const OPENROUTER_URL = 'https://openrouter.ai/api/v1/chat/completions';

// 무료(:free) 모델만 허용한다. 클라이언트가 유료 모델을 지정해 키 크레딧을 쓰지 못하게 막는 목록이기도 하다.
const FREE_MODELS = {
  'nvidia/nemotron-3-super-120b-a12b:free': 'Nemotron 3 Super 120B',
  'nvidia/nemotron-3-ultra-550b-a55b:free': 'Nemotron 3 Ultra 550B',
  'nvidia/nemotron-3.5-lightning:free': 'Nemotron 3.5 Lightning',
};
const DEFAULT_MODEL = process.env.OPENROUTER_MODEL || 'nvidia/nemotron-3-super-120b-a12b:free';

const MAX_TURNS = 20;
const MAX_CHARS = 8000;

let contextPromise;
function loadContext() {
  contextPromise ??= Promise.all(
    ['README.md', 'docs/프로젝트정의.md', 'docs/입력양식_변경방법.md'].map(async (f) => {
      try {
        return `### ${f}\n\n${await readFile(path.join(process.cwd(), f), 'utf8')}`;
      } catch {
        return '';
      }
    }),
  ).then((parts) => parts.filter(Boolean).join('\n\n---\n\n'));
  return contextPromise;
}

function systemPrompt(context) {
  return `너는 'TRACE 도우미'다. TRACE는 보험사 예산·사업비 담당자를 위한 사업비 변동 원인분석·예산 점검 도구(Excel 2016 + VBA + HTML)이며, IFRS 17(GMM) CSM·당기손익 영향까지 연결한다.

- 한국어로, 간결하고 정확하게 답한다. 필요하면 단계별 목록이나 표를 쓴다.
- TRACE 사용법·시트·버튼·입력 양식 질문은 아래 문서를 근거로 답하고, 문서에 없는 내용은 추측하지 말고 문서에 없다고 말한다.
- 사업비·예산·IFRS 17 일반 개념 질문에는 일반 지식으로 답하되, 회사 회계정책에 따라 달라질 수 있는 부분은 그렇다고 밝힌다.
- 사용자가 실제 업무 데이터(실제 금액·거래처·전표 번호 등)를 붙여 넣으려 하면, 이 챗봇은 외부망 서비스이므로 가상 데이터만 쓰라고 안내한다.

## TRACE 문서

${context}`;
}

function json(body, status = 200) {
  return Response.json(body, { status, headers: { 'Cache-Control': 'no-store' } });
}

function allowedModels() {
  const models = { ...FREE_MODELS };
  if (!models[DEFAULT_MODEL]) models[DEFAULT_MODEL] = DEFAULT_MODEL;
  return models;
}

function cleanMessages(raw) {
  if (!Array.isArray(raw)) return null;
  const msgs = raw
    .filter((m) => m && (m.role === 'user' || m.role === 'assistant') && typeof m.content === 'string' && m.content.trim())
    .slice(-MAX_TURNS)
    .map((m) => ({ role: m.role, content: m.content.slice(0, MAX_CHARS) }));
  return msgs.length && msgs.at(-1).role === 'user' ? msgs : null;
}

function upstreamError(status, text) {
  let msg = '';
  try {
    msg = JSON.parse(text)?.error?.message || '';
  } catch {}
  if (status === 401) return 'OpenRouter API 키가 올바르지 않습니다. Vercel 환경변수 OPENROUTER_API_KEY를 확인하세요.';
  if (status === 429) return '무료 모델 사용 한도에 걸렸습니다. 잠시 후 다시 시도하거나 다른 모델을 고르세요.';
  return `OpenRouter 오류 (${status})${msg ? `: ${msg}` : ''}`;
}

export function GET() {
  return json({
    models: Object.entries(allowedModels()).map(([id, label]) => ({ id, label })),
    defaultModel: DEFAULT_MODEL,
    needsCode: Boolean(process.env.CHAT_ACCESS_CODE),
  });
}

export async function POST(request) {
  const key = process.env.OPENROUTER_API_KEY;
  if (!key) return json({ error: 'Vercel 환경변수 OPENROUTER_API_KEY가 없습니다.' }, 500);

  let body;
  try {
    body = await request.json();
  } catch {
    return json({ error: '요청 형식이 올바르지 않습니다.' }, 400);
  }

  const accessCode = process.env.CHAT_ACCESS_CODE;
  if (accessCode && body.code !== accessCode) return json({ error: '접속 코드가 올바르지 않습니다.' }, 401);

  const model = body.model || DEFAULT_MODEL;
  if (!allowedModels()[model]) return json({ error: '허용되지 않은 모델입니다.' }, 400);

  const messages = cleanMessages(body.messages);
  if (!messages) return json({ error: '보낼 메시지가 없습니다.' }, 400);

  const upstream = await fetch(OPENROUTER_URL, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${key}`,
      'Content-Type': 'application/json',
      'HTTP-Referer': new URL(request.url).origin,
      'X-Title': 'TRACE Assistant',
    },
    body: JSON.stringify({
      model,
      messages: [{ role: 'system', content: systemPrompt(await loadContext()) }, ...messages],
      stream: true,
      reasoning: { exclude: true },
    }),
    signal: request.signal,
  });

  if (!upstream.ok || !upstream.body) {
    return json({ error: upstreamError(upstream.status, await upstream.text()) }, upstream.status === 429 ? 429 : 502);
  }

  // OpenRouter SSE → {"d": ...} / {"error": ...} 줄 단위 JSON
  const encoder = new TextEncoder();
  const decoder = new TextDecoder();
  const reader = upstream.body.getReader();
  const stream = new ReadableStream({
    async start(controller) {
      const send = (obj) => controller.enqueue(encoder.encode(JSON.stringify(obj) + '\n'));
      let buf = '';
      try {
        for (;;) {
          const { done, value } = await reader.read();
          if (done) break;
          buf += decoder.decode(value, { stream: true });
          let nl;
          while ((nl = buf.indexOf('\n')) >= 0) {
            const line = buf.slice(0, nl).trim();
            buf = buf.slice(nl + 1);
            if (!line.startsWith('data:')) continue; // ": OPENROUTER PROCESSING" 같은 주석 줄은 건너뜀
            const data = line.slice(5).trim();
            if (data === '[DONE]') continue;
            let chunk;
            try {
              chunk = JSON.parse(data);
            } catch {
              continue;
            }
            if (chunk.error) send({ error: chunk.error.message || '모델 응답 중 오류가 났습니다.' });
            const text = chunk.choices?.[0]?.delta?.content;
            if (text) send({ d: text });
          }
        }
      } catch (err) {
        if (!request.signal.aborted) send({ error: `응답을 받는 중 연결이 끊겼습니다: ${err.message}` });
      } finally {
        controller.close();
      }
    },
    cancel() {
      reader.cancel();
    },
  });

  return new Response(stream, {
    headers: { 'Content-Type': 'application/x-ndjson; charset=utf-8', 'Cache-Control': 'no-store' },
  });
}
