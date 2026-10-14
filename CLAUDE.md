# TRACE 프로젝트 안내 (Claude Code용)

TRACE(사업비 원인추적기)는 보험사 예산·사업비 담당자가 내부망에서 쓰는 엑셀 + HTML 도구다. 이 파일은 다른 컴퓨터·클라우드에서 작업을 이어갈 때 먼저 읽는다. 지금까지의 요구사항과 결정은 `docs/작업이력.md`에 있다.

## 사용자와 환경 (반드시 지킬 제약)
- 사용자: 보험사 예산·사업비·대외기관대응 담당. 한국어로 소통. 답변·문서·UI 문구 모두 한국어.
- **망분리:** 개발은 외부망에서 가상 데이터로, 실행은 내부망. 실제 업무 데이터는 저장소에 올리지 않는다.
- **내부망 실행 환경:** Excel 2016 + VBA(매크로 허용) + HTML(Edge·Chrome). 인터넷 없음.
  - 엑셀 수식은 Excel 2016 함수만: `XLOOKUP`·`FILTER`·`TEXTJOIN`·`IFS`·`MAXIFS`·`LET` 등 금지. `SUMIFS`·`INDEX/MATCH`·`SUMPRODUCT`·`COUNTIFS` 사용. 배열 트릭은 `SUMPRODUCT(MAX(...))`, `INDEX(...,0)`.
  - SUMPRODUCT 배열 안에 `""`(빈 문자열)이 들어가면 `#VALUE!` → 숫자 열은 빈 행에 0을 돌려준다.
  - VBA는 `.bas` 텍스트로 제공하고 **CP949 + CRLF**로 저장 (한국어 Windows에서 가져오기 시 한글 깨짐 방지). 프로시저 이름은 영문.
  - HTML은 외부 CDN·폰트에 의존하지 않는다 (내부망 오프라인). 차트는 직접 그린 SVG. 엑셀 읽기는 `DecompressionStream`으로 직접 ZIP 해제.
- 저장소 `bicebersa/TRACE`는 **공개**. 예시 데이터는 모두 가상(예시보험㈜).

## 구조
- `src/build_trace.py`: **엑셀 본체 생성기(단일 소스)**. openpyxl로 시트·수식·예시 데이터를 만든다. 환경변수 `JE_END`(전표 수식 범위 끝 행, 기본 100000), `SEL`(검증용 원인분석 선택 "부서,G/L,기준"), `EMBED_JSON`(data.js 경로 → 숨김 시트 `대시보드데이터`에 결과 내장).
- `src/build_datajs.py`: 엑셀 수식과 **같은 로직을 파이썬으로 옮긴 것**. 예시 `data.js`를 만들고, 검증 때 기대값(대조 기준)으로 쓴다. 수식 로직을 바꾸면 여기도 같이 바꾼다.
- `src/build_vba.py`: `modTrace.bas` 생성 (시트 레이아웃 상수를 `tool_data.json`에서 읽어 넣음).
- `src/build_sample_fss.py`: 예시 감독원 양식(가로형) 생성.
- `tools/verify.py`: 작은 범위로 엑셀을 만들어 `formulas` 라이브러리로 전 수식을 계산하고, 파이썬 대조값과 비교한다.
- HTML 5개(대시보드·처리 흐름·IFRS 17 연결·사용설명서·내려받기)는 공통 상단 메뉴로 연결되고, 공통 디자인(색 토큰·메뉴·카드·표·탭)은 `trace.css` 한 파일에 있다. 페이지 `<style>`에는 그 페이지 전용 규칙만 둔다. 외부 글꼴 없이 맑은 고딕·D2Coding/Consolas. 내려받기 페이지는 claude.ai에서는 `downloads` 기능(아티팩트 capabilities)으로, 로컬에서는 `a[download]`로 받는다. 아티팩트는 xlsx·zip을 파일로 못 올리므로 `build_bundle.py`가 만든 `dl/*.js`(base64)를 함께 게시한다. 대시보드는 탭 5개, `data.js` 또는 엑셀 업로드로 자료를 읽는다.
- `legacy/`: 초기 산출물(SAP 시산표·자산운용 양식, 예산실적리포트 v0.1)과 생성 스크립트.

## 빌드 순서
```bash
pip install openpyxl formulas
python src/build_trace.py                       # TRACE_v0.5.xlsx (+ src/tool_data.json)
python src/build_datajs.py                      # data.js
EMBED_JSON=data.js python src/build_trace.py    # 결과 내장본으로 다시 생성 (엑셀만 올려도 대시보드가 읽음)
python src/build_vba.py                         # modTrace.bas
python src/build_xlsm.py                        # TRACE_v0.5.xlsm (매크로 내장, vbaProject.bin 직접 생성 · olevba로 구조 확인)
python src/build_sample_fss.py                  # samples/샘플_감독원양식_2026-09.xlsx
# 두 번째 예시: SEED=27 SCALE=1.35 COMPANY=샘플생명㈜ 로 src 사본에서 build_trace → build_datajs → EMBED 빌드 → samples/TRACE_예시2_샘플생명_v0.5.xlsx(.xlsm)
python src/build_forms.py                       # samples/TRACE_입력양식_v0.5.xlsx (본체 헤더설정에서 생성)
python src/build_bundle.py                      # TRACE_v0.5_전체.zip (내려받기 페이지용 전체 묶음)
```

## 예산 고갈 예상 (선형 추세)
- 변동분석 AP~AW열: 월 소진속도(기울기)·절편·추세 연말 누계·고갈 예상(개월)·고갈 예상일·추세 판정·분기 고갈 예상일·정렬키. 월별집계 AB~AM열(당해 누계 1~12월, 기준월 이후 0)을 쓴다.
- 기울기는 설정 C19(`BURN_METHOD` 환경변수로 생성 시 지정): 연내누적 = 점 (0,0)…(기준월, 누계)에 최소제곱, 최근월·최근분기·전년추세(변동분석 AZ열) = (기준월, 누계)를 지나는 직선. 대시보드도 같은 식(burnCalc)으로 다시 계산한다.
- 비교 기준(전월·전분기·전년동기누적): 원인분석 C6, 변동분석 K~M·AX~AY, 주의대상 O~R(율)·AN~AP(증감액), 차트데이터 계정별 C~J. VBA 내보내기는 TOP 10 × 3기준 원인분석.
- 직선 계산: Excel 2016 호환을 위해 SLOPE 대신 SUMPRODUCT 닫힌 식. 날짜는 `YEAR/MONTH/DAY`로 문자열 조립 (TEXT 날짜 서식은 지역 설정에 따라 다름).
- 차트데이터 96행부터 소진 추세 표(전사 1행 + 고갈 빠른 순 15행) → VBA `burn` 배열 → 대시보드 사업비 탭 카드.

## 검증 방법
- 이 개발 환경에는 엑셀·LibreOffice가 없을 수 있다. `python tools/verify.py`로 수식을 계산·대조한다 (전표 범위를 줄인 사본으로 10분 안팎, 메모리 8GB면 **한 번에 하나씩** 실행).
- `formulas` 라이브러리는 `TEXT(x,"+0%")`의 `+` 부호, `TEXT(0,"#,##0.0")`를 잘못 표시한다. 이 차이는 엑셀과 무관하다.
- VBA는 엑셀이 없어 실행 검증을 못 했다. 사용자가 내부망에서 처음 돌릴 때 나온 오류 문장을 받아 고친다.

## 남은 일 (우선순위)
0. **v0.5 수식 검증 마무리 (가장 먼저):** 로컬 Mac(메모리 8GB)에서는 시간이 오래 걸려 중단했다. `python tools/verify.py`(1위 항목 기준)와 원인 패턴 시나리오 4개 `--sel D200,62901000,전월`(재분류·역분개), `--sel D600,63201000,전월`(시점 이동), `--sel D700,63002000,전월`(단가 인상), `--sel D300,62802000,전월`(연간 키워드)를 **하나씩** 돌려 불일치 0건을 확인한다. v0.4까지는 같은 방식으로 수식 오류 0건·대조 불일치 0건을 확인했고, v0.5에서 바뀐 것은 계정대체 규칙·VFA·분기 한도 판정이다.
1. 실제 양식 반영: SAP 예산판·전표 화면, 감독원 업무보고서(BS·PL·보험계약부채 변동표) 실제 양식이 오면 `헤더설정`·`입력옵션` 기본값과 예시 데이터를 맞춘다.
2. 상품·증권번호별 공통비 귀속 결과(관리회계)를 받으면 `계정대체기준` 3번 표(포트폴리오 비율)를 그 결과로 대체.
3. 계리 예정사업비의 실제 단위(상품·포트폴리오·계약집단) 확인.
4. VBA 내부망 실행 결과 반영.
5. 월마감 이력 기반 추세, 예정사업비 대비 사업비차익 보고서.

## 문서 작성 원칙
- 결론부터, 짧은 문장, 표로 정리. 금액 단위 명시(원본 원, 보고 백만원).
- 엑셀 셀 색: 파란 글씨 입력 · 검정 수식 · 초록 다른 시트 참조 · 노란 배경 설정 · 회색 머리글 보조수식.
