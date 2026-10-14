// TRACE 도우미 챗봇 (모든 TRACE 페이지 공용)
// Vercel에 배포했을 때만 보인다: 같은 주소의 /api/chat (OpenRouter 무료 Nemotron)이 응답할 때만 창을 띄운다.
// 내부망에서 파일로 열거나 이 파일이 없으면 아무것도 하지 않으므로 기존 화면은 그대로 동작한다.
// 대화 내용과 창 열림 상태는 브라우저에 저장해 페이지를 옮겨도 이어진다.
(function () {
  "use strict";
  if (!/^https?:$/.test(location.protocol) || document.getElementById("cb")) return;
  var API = "api/chat";
  var style = document.createElement("style");
  style.textContent = "/* 챗봇 전용 토큰: 페이지 토큰을 우선 쓰고, 없는 페이지에서는 TRACE 기본색 */\n.cb, .cb-fab { --cb-bg: var(--bg, #f3f5f4); --cb-surface: var(--surface, #ffffff); --cb-ink: var(--ink, #16231f); --cb-muted: var(--muted, #5b6b66); --cb-line: var(--line, #d5ddd9); --cb-grid: var(--grid, #e6ebe8);\n  --cb-accent: var(--accent, #17705a); --cb-accent-soft: var(--accent-soft, #e1f0ea); --cb-warn: var(--warn, #9a6512); --cb-crit: var(--crit, #b4372f); --cb-crit-soft: var(--crit-soft, #f8e2df); font-family: var(--font, \"Malgun Gothic\", \"Apple SD Gothic Neo\", \"Noto Sans KR\", system-ui, sans-serif); line-height: 1.55; }\n@media (prefers-color-scheme: dark) { :root:not([data-theme=\"light\"]) .cb, :root:not([data-theme=\"light\"]) .cb-fab { --cb-bg: var(--bg, #0f1513); --cb-surface: var(--surface, #17201d); --cb-ink: var(--ink, #e4ece8); --cb-muted: var(--muted, #93a49e); --cb-line: var(--line, #2a3833); --cb-grid: var(--grid, #222d29);\n  --cb-accent: var(--accent, #4fbf9a); --cb-accent-soft: var(--accent-soft, #183128); --cb-warn: var(--warn, #e0a84a); --cb-crit: var(--crit, #ef7a6e); --cb-crit-soft: var(--crit-soft, #3a1f1c); } }\n:root[data-theme=\"dark\"] .cb, :root[data-theme=\"dark\"] .cb-fab { --cb-bg: var(--bg, #0f1513); --cb-surface: var(--surface, #17201d); --cb-ink: var(--ink, #e4ece8); --cb-muted: var(--muted, #93a49e); --cb-line: var(--line, #2a3833); --cb-grid: var(--grid, #222d29);\n  --cb-accent: var(--accent, #4fbf9a); --cb-accent-soft: var(--accent-soft, #183128); --cb-warn: var(--warn, #e0a84a); --cb-crit: var(--crit, #ef7a6e); --cb-crit-soft: var(--crit-soft, #3a1f1c); }\n.cb { color: var(--cb-ink); font-size: 14px; }\n.cb-btn { font: inherit; font-size: 13px; font-weight: 600; color: #fff; background: var(--cb-accent); border: none; border-radius: 8px; padding: 6px 14px; cursor: pointer; white-space: nowrap; }\n.cb-btn:hover { filter: brightness(1.08); } .cb-btn:focus-visible { outline: 2px solid var(--cb-ink); outline-offset: 2px; }\n.cb-btn.ghost { background: transparent; color: var(--cb-ink); border: 1px solid var(--cb-line); font-weight: 500; }\n/* TRACE 도우미 챗봇 (Vercel · OpenRouter 무료 Nemotron) */\n.cb-fab { position: fixed; right: calc(16px + env(safe-area-inset-right, 0px)); bottom: calc(16px + env(safe-area-inset-bottom, 0px)); z-index: 40; font: inherit; font-size: 14px; font-weight: 700; color: #fff; background: var(--cb-accent); border: none; border-radius: 999px; padding: 11px 18px; cursor: pointer; box-shadow: 0 6px 18px rgba(0,0,0,.18); }\n.cb-fab:focus-visible { outline: 2px solid var(--cb-ink); outline-offset: 2px; }\n.cb { position: fixed; right: calc(16px + env(safe-area-inset-right, 0px)); bottom: calc(16px + env(safe-area-inset-bottom, 0px)); z-index: 41; width: min(420px, calc(100vw - 32px)); height: min(620px, calc(100dvh - 96px)); display: flex; flex-direction: column; background: var(--cb-surface); border: 1px solid var(--cb-line); border-radius: 14px; box-shadow: 0 12px 36px rgba(0,0,0,.22); overflow: hidden; }\n.cb-head { display: flex; align-items: center; gap: 8px; padding: 10px 12px; border-bottom: 1px solid var(--cb-line); }\n.cb-head b { font-size: 14.5px; }\n.cb-head .sub { font-size: 11.5px; color: var(--cb-muted); }\n.cb-head .sp { margin-left: auto; display: flex; gap: 6px; align-items: center; }\n.cb-head a.cb-x { text-decoration: none; color: var(--cb-accent); }\n.cb-x { white-space: nowrap; font: inherit; font-size: 12.5px; background: none; border: 1px solid var(--cb-line); border-radius: 7px; padding: 3px 9px; cursor: pointer; color: var(--cb-ink); }\n.cb-log { flex: 1; overflow-y: auto; padding: 12px; display: grid; gap: 10px; align-content: start; background: var(--cb-bg); }\n.cb-m { display: grid; gap: 3px; min-width: 0; }\n.cb-m .who { font-size: 11px; color: var(--cb-muted); font-weight: 700; }\n.cb-m .body { background: var(--cb-surface); border: 1px solid var(--cb-line); border-radius: 10px; padding: 8px 11px; font-size: 13.5px; overflow-wrap: anywhere; min-width: 0; }\n.cb-m.user { justify-items: end; }\n.cb-m.user .body { background: var(--cb-accent-soft); border-color: transparent; white-space: pre-wrap; max-width: 88%; }\n.cb-m.error .body { background: var(--cb-crit-soft); border-color: transparent; color: var(--cb-crit); }\n.cb-m .body > :first-child { margin-top: 0; } .cb-m .body > :last-child { margin-bottom: 0; }\n.cb-m .body p, .cb-m .body ul, .cb-m .body ol { margin: 6px 0; } .cb-m .body ul, .cb-m .body ol { padding-left: 20px; }\n.cb-m .body h4 { margin: 8px 0 4px; font-size: 13.5px; }\n.cb-m .body pre { background: var(--cb-grid); border-radius: 7px; padding: 8px 10px; overflow-x: auto; font-family: var(--mono, Consolas, Menlo, monospace); font-size: 12px; }\n.cb-m .body code { font-family: var(--mono, Consolas, Menlo, monospace); font-size: .92em; background: var(--cb-grid); border-radius: 4px; padding: 0 3px; }\n.cb-m .body pre code { background: none; padding: 0; }\n.cb-m .body table { display: block; overflow-x: auto; font-size: 12px; margin: 6px 0; }\n.cb-m .body a { color: var(--cb-accent); }\n.cb-typing { color: var(--cb-muted); font-size: 12.5px; }\n.cb-chips { display: flex; flex-wrap: wrap; gap: 6px; }\n.cb-chip { font: inherit; font-size: 12px; border: 1px solid var(--cb-line); background: var(--cb-surface); color: var(--cb-ink); border-radius: 999px; padding: 4px 10px; cursor: pointer; }\n.cb-chip:hover { border-color: var(--cb-accent); color: var(--cb-accent); }\n.cb-form { border-top: 1px solid var(--cb-line); padding: 8px 10px; display: grid; gap: 6px; }\n.cb-row { display: flex; gap: 6px; align-items: flex-end; }\n.cb-row textarea { flex: 1; min-width: 0; resize: none; font: inherit; font-size: 13.5px; color: var(--cb-ink); background: var(--cb-bg); border: 1px solid var(--cb-line); border-radius: 9px; padding: 7px 10px; max-height: 140px; line-height: 1.5; }\n.cb-row input { font: inherit; font-size: 13px; color: var(--cb-ink); background: var(--cb-bg); border: 1px solid var(--cb-line); border-radius: 7px; padding: 4px 8px; width: 140px; }\n.cb-note { font-size: 11px; color: var(--cb-muted); }\n.cb-note b { color: var(--cb-warn); }\n.cb-ctx { font-size: 11.5px; color: var(--cb-muted); display: flex; align-items: center; gap: 5px; }\n@media print { .cb, .cb-fab { display: none !important; } }\n";
  document.head.appendChild(style);
  var holder = document.createElement("div");
  holder.innerHTML = "<button type=\"button\" class=\"cb-fab\" id=\"cb-fab\" aria-controls=\"cb\" aria-expanded=\"false\" hidden>TRACE 도우미</button>\n<section class=\"cb\" id=\"cb\" role=\"dialog\" aria-label=\"TRACE 도우미 챗봇\" hidden>\n  <div class=\"cb-head\"><div style=\"display:grid\"><b>TRACE 도우미</b><span class=\"sub\">무료 Nemotron · 사용법 · 사업비 · IFRS 17</span></div>\n    <div class=\"sp\"><button type=\"button\" class=\"cb-x\" id=\"cb-reset\">새 대화</button><button type=\"button\" class=\"cb-x\" id=\"cb-close\" aria-label=\"닫기\">닫기</button></div></div>\n  <div class=\"cb-log\" id=\"cb-log\" aria-live=\"polite\"></div>\n  <form class=\"cb-form\" id=\"cb-form\">\n    <div class=\"cb-row\" id=\"cb-code-row\" hidden><label for=\"cb-code\" class=\"cb-note\">접속 코드</label><input id=\"cb-code\" type=\"password\" autocomplete=\"off\"></div>\n    <label class=\"cb-ctx\"><input type=\"checkbox\" id=\"cb-ctx\" checked> <span id=\"cb-ctx-label\">지금 화면 내용을 같이 보내기</span></label>\n    <div class=\"cb-row\"><textarea id=\"cb-input\" rows=\"1\" placeholder=\"질문 입력 (Enter 보내기 · Shift+Enter 줄바꿈)\"></textarea>\n      <button type=\"submit\" class=\"cb-btn\" id=\"cb-send\">보내기</button><button type=\"button\" class=\"cb-btn ghost\" id=\"cb-stop\" hidden>중지</button></div>\n    <div class=\"cb-note\"><b>실제 업무 데이터는 입력하지 마세요.</b> 외부 무료 모델(OpenRouter · NVIDIA)로 전송됩니다.</div>\n  </form>\n</section>";
  while (holder.firstChild) document.body.appendChild(holder.firstChild);

  var $ = function (id) { return document.getElementById(id); };
  var fab = $("cb-fab"), box = $("cb"), log = $("cb-log"), form = $("cb-form"), input = $("cb-input"),
      send = $("cb-send"), stop = $("cb-stop"), useCtx = $("cb-ctx"), codeRow = $("cb-code-row"), code = $("cb-code");
  var KEY = "trace-dash-chat-vercel-v1";
  var store = {
    get: function () { try { return JSON.parse(localStorage.getItem(KEY)) || {}; } catch (e) { return {}; } },
    set: function (v) { try { localStorage.setItem(KEY, JSON.stringify(v)); } catch (e) { /* 저장이 막힌 환경 */ } }
  };
  var saved = store.get(), messages = Array.isArray(saved.messages) ? saved.messages : [], closed = !!saved.closed, ctrl = null, pending = "";
  code.value = saved.code || "";
  function save() { store.set({ messages: messages, code: code.value, closed: closed }); }

  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  // 작은 마크다운 변환 (외부 라이브러리 없이): 코드블록 · 제목 · 목록 · 표 · 굵게 · 코드
  function inline(t) { return esc(t).replace(/`([^`]+)`/g, "<code>$1</code>").replace(/\*\*([^*]+)\*\*/g, "<b>$1</b>"); }
  function md(src) {
    var out = [], lines = src.replace(/\r/g, "").split("\n"), i = 0;
    while (i < lines.length) {
      var l = lines[i];
      if (/^```/.test(l)) { var buf = []; i++; while (i < lines.length && !/^```/.test(lines[i])) buf.push(lines[i++]); i++; out.push("<pre><code>" + esc(buf.join("\n")) + "</code></pre>"); continue; }
      if (/^#{1,6}\s/.test(l)) { out.push("<h4>" + inline(l.replace(/^#+\s/, "")) + "</h4>"); i++; continue; }
      if (/^\s*\|.*\|\s*$/.test(l) && i + 1 < lines.length && /^\s*\|?\s*:?-{2,}/.test(lines[i + 1])) {
        var cells = function (r) { return r.trim().replace(/^\||\|$/g, "").split("|").map(function (c) { return c.trim(); }); };
        var h = cells(l); i += 2; var body = "";
        while (i < lines.length && /^\s*\|.*\|\s*$/.test(lines[i])) { body += "<tr>" + cells(lines[i++]).map(function (c) { return "<td>" + inline(c) + "</td>"; }).join("") + "</tr>"; }
        out.push("<table><thead><tr>" + h.map(function (c) { return "<th>" + inline(c) + "</th>"; }).join("") + "</tr></thead><tbody>" + body + "</tbody></table>"); continue;
      }
      if (/^\s*([-*]|\d+\.)\s/.test(l)) {
        var ordered = /^\s*\d+\./.test(l), items = [];
        while (i < lines.length && /^\s*([-*]|\d+\.)\s/.test(lines[i])) items.push("<li>" + inline(lines[i++].replace(/^\s*([-*]|\d+\.)\s/, "")) + "</li>");
        out.push((ordered ? "<ol>" : "<ul>") + items.join("") + (ordered ? "</ol>" : "</ul>")); continue;
      }
      if (!l.trim()) { i++; continue; }
      var para = [];
      while (i < lines.length && lines[i].trim() && !/^(```|#{1,6}\s|\s*([-*]|\d+\.)\s|\s*\|)/.test(lines[i])) para.push(inline(lines[i++]));
      if (!para.length) para.push(inline(lines[i++]));
      out.push("<p>" + para.join("<br>") + "</p>");
    }
    return out.join("");
  }

  // 지금 화면에 불러온 대시보드 자료 요약
  function dashContext() {
    var D = window.TRACE_CURRENT;
    if (!D || !D.meta) return "";
    var m = D.meta, L = [];
    L.push("회사 " + m.company + " · 기준 " + m.period + " · 단위 " + m.unit + " · 자가진단 " + m.check + " · 3자 대사 " + m.recon);
    (D.kpi || []).forEach(function (k) { L.push("KPI " + k.label + ": " + k.value + (k.rate === null || k.rate === undefined ? "" : " (" + k.basis + " 대비 " + (k.rate * 100).toFixed(1) + "%)")); });
    if (D.status) L.push("분기 한도 판정: " + D.status.map(function (s) { return s.status + " " + s.count; }).join(", "));
    (D.depts || []).forEach(function (d) { L.push("부서 " + d.name + ": 분기 한도 " + d.avail + ", 분기 실적 " + d.ytd + ", 분기말 예상 " + d.forecast + ", " + d.status); });
    (D.top10 || []).forEach(function (t) {
      var c = t.cause || {}, k = (c.candidates || []).map(function (x) { return x.pattern + (x.strength ? "(" + x.strength + ")" : "") + (x.text ? ": " + x.text : ""); }).join(" / ");
      L.push("주의 " + t.rank + "위 " + t.dept + " · " + t.account + ": 당월 " + t.cur + ", 전월대비 " + (typeof t.mom === "number" ? (t.mom * 100).toFixed(0) + "%" : "-") +
        ", 분기 " + t.status + (t.signal ? ", 신호 " + t.signal : "") + (c.summary ? ", 요약 " + c.summary : "") + (k ? ", 원인 후보 " + k : "") + (t.comment ? ", 확인된 원인 " + t.comment : ""));
    });
    if (D.ifrs17 && D.ifrs17.summary) L.push("IFRS 17: " + D.ifrs17.summary);
    return L.join("\n").slice(0, 12000);
  }


  // 대시보드가 아닌 페이지: 지금 보고 있는 페이지 본문을 짧게 보낸다
  function pageContext() {
    var root = document.querySelector("main") || document.body, copy = root.cloneNode(true);
    copy.querySelectorAll("script,style,.cb,.cb-fab,header.topnav,nav,svg").forEach(function (n) { n.remove(); });
    var text = (copy.textContent || "").replace(/[ \t]+/g, " ").replace(/\n\s*\n+/g, "\n").trim();
    return text ? "페이지: " + document.title + "\n" + text.slice(0, 6000) : "";
  }
  var onDash = function () { return !!(window.TRACE_CURRENT && window.TRACE_CURRENT.meta); };
  $("cb-ctx-label").textContent = onDash() ? "지금 대시보드 요약을 같이 보내기" : "지금 보고 있는 페이지 내용을 같이 보내기";

  function bubble(role, text) {
    var w = document.createElement("div"); w.className = "cb-m " + role;
    var who = document.createElement("div"); who.className = "who"; who.textContent = role === "user" ? "나" : role === "error" ? "알림" : "TRACE 도우미";
    var b = document.createElement("div"); b.className = "body";
    if (role === "assistant") b.innerHTML = md(text); else b.textContent = text;
    w.appendChild(who); w.appendChild(b); log.appendChild(w); return b;
  }
  var SUGGEST = ["주의 대상 1위는 왜 늘었어?", "이번 달 IFRS 17 영향을 요약해줘", "분기 한도 초과 예상인 곳은?", "매월 사용 순서를 알려줘"];
  function redraw() {
    log.innerHTML = "";
    if (!messages.length) {
      var hi = bubble("assistant", onDash() ? "무엇이 궁금하세요? 지금 화면의 대시보드 숫자나 TRACE 사용법을 물어보세요." : "무엇이 궁금하세요? 지금 보고 있는 페이지 내용이나 TRACE 사용법을 물어보세요.");
      var chips = document.createElement("div"); chips.className = "cb-chips";
      SUGGEST.forEach(function (q) { var c = document.createElement("button"); c.type = "button"; c.className = "cb-chip"; c.textContent = q; c.onclick = function () { ask(q); }; chips.appendChild(c); });
      hi.appendChild(chips);
    }
    messages.forEach(function (m) { bubble(m.role, m.content); });
    log.scrollTop = log.scrollHeight;
  }
  function busy(on) { send.hidden = on; stop.hidden = !on; input.disabled = on; }
  function autosize() { input.style.height = "auto"; input.style.height = Math.min(input.scrollHeight, 140) + "px"; }


  var RULES = "너는 'TRACE 도우미'다. TRACE는 보험사 예산·사업비 담당자용 사업비 원인추적기(Excel 2016 + VBA + HTML)로, IFRS 17(GMM·VFA·PAA) CSM·당기손익 영향까지 연결한다.\n" +
    "- 한국어로 간결하게, 결론부터 답한다. 필요하면 목록이나 표를 쓴다. 금액은 단위를 밝힌다.\n" +
    "- 대시보드 숫자 질문은 [대시보드 요약]만 근거로 답하고, 요약에 없는 숫자는 지어내지 말고 없다고 말한다.\n" +
    "- [지금 보고 있는 페이지]가 있으면 그 내용을 우선 근거로 쓴다.\n" +
    "- 원인 후보는 규칙 기반 가설이므로 근거 전표와 담당 부서 확인이 필요하다는 점을 필요할 때 짚는다.\n" +
    "- 화면의 자료는 가상 예시(예시보험㈜)다. 사용자가 실제 업무 데이터를 붙여 넣으려 하면 가상 데이터만 쓰라고 안내한다.";

  function ask(text) {
    text = (text || "").trim();
    if (!text || ctrl) return;
    if (!messages.length) log.innerHTML = "";
    messages.push({ role: "user", content: text });
    bubble("user", text);
    input.value = ""; autosize(); save();
    var b = bubble("assistant", ""); b.innerHTML = '<span class="cb-typing">생각 중…</span>';
    log.scrollTop = log.scrollHeight;
    busy(true); ctrl = new AbortController(); pending = "";
    var err = "", myCtrl = ctrl;
    var ctx = "";
    if (useCtx.checked) ctx = onDash() ? "[대시보드 요약 · 가상 예시 데이터]\n" + dashContext() : (pageContext() ? "[지금 보고 있는 페이지]\n" + pageContext() : "");
    // 서버가 README·docs를 근거로 넣으므로, 여기서는 화면 내용과 지시만 마지막 질문 앞에 붙인다
    var payload = messages.map(function (m, i) {
      if (i !== messages.length - 1) return m;
      return { role: "user", content: RULES + (ctx ? "\n\n" + ctx : "") + "\n\n[질문]\n" + m.content };
    });
    fetch(API, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ messages: payload, code: code.value }), signal: myCtrl.signal })
      .then(function (res) {
        if (!res.ok) return res.json().catch(function () { return {}; }).then(function (d) { throw new Error(d.error || "서버 오류 (" + res.status + ")"); });
        var reader = res.body.getReader(), dec = new TextDecoder(), buf = "";
        function pump() {
          return reader.read().then(function (r) {
            if (r.done) return;
            buf += dec.decode(r.value, { stream: true });
            var lines = buf.split("\n"); buf = lines.pop();
            lines.forEach(function (line) {
              if (!line) return;
              var ev = JSON.parse(line);
              if (ev.error) err = ev.error;
              if (ev.d) { var stick = log.scrollHeight - log.scrollTop - log.clientHeight < 60; pending += ev.d; b.innerHTML = md(pending); if (stick) log.scrollTop = log.scrollHeight; }
            });
            return pump();
          });
        }
        return pump();
      })
      .catch(function (e) {
        if (e.name === "AbortError") return;
        err = e instanceof TypeError ? "챗봇 서버에 연결하지 못했습니다. 잠시 뒤 다시 보내 주세요." : e.message;
      })
      .then(function () {
        if (ctrl !== myCtrl) return;   // 페이지를 떠나며 이미 저장함
        ctrl = null; busy(false);
        if (pending) messages.push({ role: "assistant", content: pending }); else b.parentNode.remove();
        if (!pending && !err && !myCtrl.signal.aborted) err = "모델이 빈 응답을 보냈습니다. 다시 시도하거나 잠시 뒤 보내 주세요.";
        pending = "";
        if (err) bubble("error", err);
        save(); log.scrollTop = log.scrollHeight; input.focus();
      });
  }
  // 답을 받는 도중 다른 메뉴로 옮기면, 받은 데까지 저장해 다음 페이지에서 이어 보이게 한다
  window.addEventListener("pagehide", function () {
    if (!ctrl) return;
    if (pending) messages.push({ role: "assistant", content: pending + "\n\n(페이지를 옮겨 여기서 멈춤)" });
    ctrl = null; save();
  });

  function open(on, quiet) {
    box.hidden = !on; fab.hidden = on; fab.setAttribute("aria-expanded", on ? "true" : "false");
    if (closed === on) { closed = !on; save(); }
    if (on) { redraw(); if (!quiet) input.focus(); }
  }
  fetch(API).then(function (r) { return r.ok ? r.json() : null; }).then(function (cfg) {
    if (!cfg || !cfg.models) return;
    codeRow.hidden = !cfg.needsCode;
    if (closed) fab.hidden = false; else open(true, true);   // 기본은 열림, 닫아 두었으면 다른 페이지에서도 닫힌 채
    if (cfg.keyError) bubble("error", cfg.keyError);
  }).catch(function () { /* 서버 없음: 챗봇을 띄우지 않는다 */ });
  fab.addEventListener("click", function () { open(true); });
  $("cb-close").addEventListener("click", function () { open(false); fab.focus(); });
  code.addEventListener("change", save);
  $("cb-reset").addEventListener("click", function () { if (ctrl) return; messages = []; save(); redraw(); input.focus(); });
  box.addEventListener("keydown", function (e) { if (e.key === "Escape") { open(false); fab.focus(); } });
  form.addEventListener("submit", function (e) { e.preventDefault(); ask(input.value); });
  input.addEventListener("keydown", function (e) { if (e.key === "Enter" && !e.shiftKey && !e.isComposing) { e.preventDefault(); ask(input.value); } });
  input.addEventListener("input", autosize);
  stop.addEventListener("click", function () { if (ctrl) ctrl.abort(); });
})();
