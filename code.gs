/*************************************************************************
 * 실시간 찬반 토론 웹앱 – GAS 백엔드 (Code.gs)
 * -----------------------------------------------------------------------
 * ★ 이 스크립트는 '토론주제' 시트에 컨테이너 바인딩된 상태를 전제로 함
 *   (시트 > 확장 프로그램 > Apps Script 로 생성). 그래서 SPREADSHEET_ID 불필요.
 *
 * 역할
 *  1) 토론 주제(어젠다) 목록 제공                         → action=topics
 *  2) 어젠다 + 학생 스탠스 → 시스템 프롬프트 조합          → action=session
 *  3) Gemini Live API 임시 토큰(ephemeral token) 발급 후 함께 반환
 *
 * 프론트엔드(Vercel)는 이 웹앱 URL로 요청 → 토큰/프롬프트/wsUrl 수신 →
 * 브라우저에서 직접 Live API에 WebSocket 연결 (client-to-server).
 *
 * ── 사용 준비 ──────────────────────────────────────────────────────────
 *  1) 시트 상단 메뉴 [🤖 토론앱 설정] > [🔑 Gemini API 키 등록] 으로 키 저장
 *     (스크립트 속성 GEMINI_API_KEY 에 저장됨. 코드에 하드코딩 안 함)
 *  2) [배포] > [새 배포] > 유형 '웹 앱'
 *     - 실행 계정: 나        - 액세스 권한: 모든 사용자
 *     - 코드 수정 시 반드시 '새 배포' 또는 버전 새로 만들기
 *************************************************************************/

const PROPS = PropertiesService.getScriptProperties();

const CONFIG = {
  SHEET_NAME: '토론주제_v2',   // 열 구조: 사용 · 분류 · 어젠다 · AI가 찬성할 때 · AI가 반대할 때 (헤더 이름으로 인식)
  // ▼ Live 모델. 메뉴에서 GEMINI_MODEL 속성으로 덮어쓸 수 있음(속성이 있으면 그 값이 우선!).
  //   모델명은 자주 바뀌므로 AI Studio에서 현재 사용 가능한 이름을 확인해 갱신.
  //   이전 값: 'gemini-3.1-flash-live-preview'
  DEFAULT_MODEL: 'gemini-3.8-live',
};

function getModelName_() {
  return PROPS.getProperty('GEMINI_MODEL') || CONFIG.DEFAULT_MODEL;
}

/*───────────────────────────────────────────────────────────────────────
 * 공통 시스템 프롬프트 (모든 주제 공통. 시트 '공통지침' 시트와 동일 내용)
 *───────────────────────────────────────────────────────────────────────*/
const COMMON_INSTRUCTION =
`너는 대한민국 중·고등학생과 실시간 음성으로 토론하는 AI 토론 상대야. 아래 규칙을 반드시 지켜.

[말투] 음성 대화이므로 자연스러운 구어체로 말하고, 한 번에 2~4문장 정도로 짧게 끊어서 말해. 절대 길게 늘어놓지 마.
[태도] 네 입장은 분명하고 당당하게 주장하되, 학생의 말을 끝까지 듣고 그 주장에 조목조목 반박해. 항상 예의 바르고 존중하는 태도를 유지하고, 비꼬거나 무시하거나 인신공격하지 마.
[눈높이] 어려운 용어나 개념은 학생이 이해할 수 있게 쉬운 말로 풀어서 설명해.
[교육적 역할] 학생이 근거를 제시하면 "왜 그렇게 생각해?", "그 근거가 항상 맞을까?"처럼 되물으며 스스로 더 깊이 생각하도록 이끌어. 정답을 강요하지 말고 사고를 자극해.
[마무리] 네 주장을 말한 뒤에는 학생이 반박할 수 있도록 질문을 던지며 대화를 이어가.
[안전] 욕설, 차별, 혐오, 특정 집단이나 개인을 비하하는 표현은 절대 쓰지 마. 민감하거나 논쟁적인 주제도 감정적으로 흐르지 않게 침착하고 균형 잡힌 자세로 다뤄.
[시작] 토론을 시작할 때는 네 입장을 한두 문장으로 간단히 밝히고, 학생의 생각을 먼저 물어봐.

────────────────────────
아래는 이번 토론에서 네가 맡은 구체적인 입장과 논거야. 이 입장을 일관되게 유지해.`;

/*═══════════════════════════════════════════════════════════════════════
 * 시트 메뉴 – 스크립트 속성(API 키/모델명)을 메뉴에서 등록
 *═══════════════════════════════════════════════════════════════════════*/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🤖 토론앱 설정')
    .addItem('🔑 Gemini API 키 등록', 'showApiKeyDialog')
    .addItem('🧠 모델명 설정(선택)', 'showModelDialog')
    .addSeparator()
    .addItem('✅ 현재 설정 확인', 'showConfigStatus')
    .addToUi();
}

function showApiKeyDialog() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Gemini API 키 등록',
    'Google AI Studio에서 발급받은 API 키를 입력하세요:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const key = r.getResponseText().trim();
  if (!key) { ui.alert('경고', '올바른 키를 입력해주세요.', ui.ButtonSet.OK); return; }
  PROPS.setProperty('GEMINI_API_KEY', key);
  ui.alert('설정 완료', 'API 키가 스크립트 속성에 안전하게 저장되었습니다.', ui.ButtonSet.OK);
}

function showModelDialog() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('모델명 설정',
    '사용할 Live 모델명을 입력하세요.\n비워두고 확인하면 기본값(' + CONFIG.DEFAULT_MODEL + ')을 사용합니다:',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const m = r.getResponseText().trim();
  if (m) { PROPS.setProperty('GEMINI_MODEL', m); ui.alert('설정 완료', '모델명: ' + m, ui.ButtonSet.OK); }
  else   { PROPS.deleteProperty('GEMINI_MODEL'); ui.alert('설정 완료', '기본 모델을 사용합니다: ' + CONFIG.DEFAULT_MODEL, ui.ButtonSet.OK); }
}

function showConfigStatus() {
  const ui = SpreadsheetApp.getUi();
  const key = PROPS.getProperty('GEMINI_API_KEY');
  const masked = key ? (key.substring(0, 4) + '••••' + key.slice(-4)) : '(미등록)';
  let topicCount = 0;
  try { topicCount = getTopics_().length; } catch (e) {}
  ui.alert('현재 설정',
    '• API 키: ' + masked + '\n' +
    '• 모델명: ' + getModelName_() + '\n' +
    '• 시트 탭: ' + CONFIG.SHEET_NAME + '\n' +
    '• 주제 수: ' + topicCount + '개',
    ui.ButtonSet.OK);
}

/*═══════════════════════════════════════════════════════════════════════
 * 웹앱 엔트리 포인트
 *  - GET  : 미리보기/간단 요청 (CORS 프리플라이트 없음)
 *  - POST : 프론트에서 text/plain(JSON 문자열)로 전송해 프리플라이트 회피
 *═══════════════════════════════════════════════════════════════════════*/
function doGet(e)  { return handleRequest_(parseParams_(e)); }
function doPost(e) { return handleRequest_(parseParams_(e)); }

function parseParams_(e) {
  let body = {};
  try { if (e && e.postData && e.postData.contents) body = JSON.parse(e.postData.contents); } catch (err) {}
  const q = (e && e.parameter) ? e.parameter : {};
  return {
    action: body.action || q.action || 'topics',
    agenda: body.agenda || q.agenda || '',
    stance: body.stance || q.stance || '',
  };
}

function handleRequest_(p) {
  try {
    switch (p.action) {
      case 'topics':
        return json_({ ok: true, topics: getTopics_() });
      case 'session':
        if (!p.agenda) throw new Error('agenda(어젠다)가 필요합니다.');
        if (!p.stance) throw new Error("stance가 필요합니다. 'pro'(학생 찬성) 또는 'con'(학생 반대).");
        return json_(getSession_(p.agenda, p.stance));
      default:
        throw new Error('알 수 없는 action: ' + p.action);
    }
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

/*───────────────────────────────────────────────────────────────────────
 * 시트 읽기 (컨테이너 바인딩 → getActive)
 *  시트 구조(1행 헤더, 열 순서는 바뀌어도 되고 '헤더 이름'으로 찾음):
 *   사용(체크박스) · 분류 · 어젠다 · AI가 찬성할 때 프롬프트 · AI가 반대할 때 프롬프트
 *  - '사용'이 체크 해제(FALSE)된 행은 주제 목록에서 숨김
 *  - 학생 '찬성' → AI는 반대 → 'AI가 반대할 때' 열
 *  - 학생 '반대' → AI는 찬성 → 'AI가 찬성할 때' 열
 *───────────────────────────────────────────────────────────────────────*/
const COLS = {
  use:      '사용',
  category: '분류',
  agenda:   '어젠다',
  aiPro:    'AI가 찬성할 때',
  aiCon:    'AI가 반대할 때',
};

/** 헤더에 키워드가 들어 있는 열 번호(0-base). 없으면 오류 */
function colIndex_(headers, keyword, required) {
  const i = headers.findIndex(h => String(h).indexOf(keyword) !== -1);
  if (i < 0 && required) throw new Error("시트 헤더에서 '" + keyword + "' 열을 찾을 수 없습니다.");
  return i;
}

/** [{ use, category, agenda, aiPro, aiCon }] */
function getRows_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) throw new Error("시트 탭 '" + CONFIG.SHEET_NAME + "' 을 찾을 수 없습니다.");
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const ix = {
    use:      colIndex_(h, COLS.use, false),
    category: colIndex_(h, COLS.category, false),
    agenda:   colIndex_(h, COLS.agenda, true),
    aiPro:    colIndex_(h, COLS.aiPro, true),
    aiCon:    colIndex_(h, COLS.aiCon, true),
  };
  return values.slice(1)
    .map(r => ({
      use:      ix.use < 0 ? true : (r[ix.use] === true || String(r[ix.use]).toUpperCase() === 'TRUE'),
      category: ix.category < 0 ? '' : String(r[ix.category] || '').trim(),
      agenda:   String(r[ix.agenda] || '').trim(),
      aiPro:    String(r[ix.aiPro] || '').trim(),
      aiCon:    String(r[ix.aiCon] || '').trim(),
    }))
    .filter(r => r.agenda !== '' && r.use);
}

/** 주제 카드용 목록: [{ agenda, category }] */
function getTopics_() {
  return getRows_().map(r => ({ agenda: r.agenda, category: r.category }));
}

/*───────────────────────────────────────────────────────────────────────
 * 세션 생성: 어젠다 + 스탠스 → 프롬프트 조합 + 임시 토큰
 *───────────────────────────────────────────────────────────────────────*/
function getSession_(agenda, stance) {
  const studentPro = isProStance_(stance);
  const target = String(agenda).trim();

  const row = getRows_().filter(r => r.agenda === target)[0];
  if (!row) throw new Error('해당 어젠다를 시트에서 찾지 못했습니다: ' + target);

  // 학생이 찬성이면 AI는 반대편 프롬프트를 사용
  const stanceInstruction = studentPro ? row.aiCon : row.aiPro;
  if (!stanceInstruction) throw new Error("해당 입장의 프롬프트('" + (studentPro ? COLS.aiCon : COLS.aiPro) + "' 열)가 비어 있습니다.");

  const systemInstruction = COMMON_INSTRUCTION + '\n\n' + stanceInstruction;
  const model = getModelName_();
  const token = createEphemeralToken_(model); // { name, expireTime, newSessionExpireTime }

  return {
    ok: true,
    agenda: target,
    studentStance: studentPro ? 'pro' : 'con',
    aiStance:      studentPro ? 'con' : 'pro',
    model: model,                          // 프론트 setup에는 'models/' + model 형태로 사용
    systemInstruction: systemInstruction,  // 잠그지 않았으므로 프론트가 setup에 주입
    token: token.name,                     // 'authTokens/...' (API 키처럼 사용)
    wsUrl: buildWsUrl_(token.name),        // 잠금 토큰 → Constrained 엔드포인트
    expireTime: token.expireTime,
    newSessionExpireTime: token.newSessionExpireTime,
  };
}

function isProStance_(s) {
  s = String(s).trim().toLowerCase();
  return (s === 'pro' || s === '찬성' || s === 'agree' || s === 'for' || s === 'yes');
}

/*───────────────────────────────────────────────────────────────────────
 * 임시 토큰 발급 (v1alpha REST)  ※ 검증된 회화앱과 동일한 payload 구조
 *  - fieldMask + bidiGenerateContentSetup 으로 model / generationConfig 잠금
 *  - systemInstruction 은 잠그지 않아 프론트에서 주입 가능
 *───────────────────────────────────────────────────────────────────────*/
function createEphemeralToken_(model) {
  const apiKey = PROPS.getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('메뉴 [🤖 토론앱 설정] > [🔑 Gemini API 키 등록]에서 키를 먼저 등록하세요.');

  const now = Date.now();
  const payload = {
    uses: 1,
    expireTime:           new Date(now + 30 * 60 * 1000).toISOString(), // 30분간 메시지 송수신
    newSessionExpireTime: new Date(now +  2 * 60 * 1000).toISOString(), // 2분 내 세션 시작(마이크 권한 여유)
    fieldMask: 'model,generationConfig',
    bidiGenerateContentSetup: {
      model: 'models/' + model,
      generationConfig: { responseModalities: ['AUDIO'] }
    }
  };

  const url = 'https://generativelanguage.googleapis.com/v1alpha/auth_tokens?key=' + encodeURIComponent(apiKey);
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText();
  Logger.log('auth_tokens ' + code + ' : ' + text);

  if (!text) throw new Error('서버가 빈 응답을 반환했습니다 (HTTP ' + code + ').');

  let data;
  try { data = JSON.parse(text); }
  catch (e) { throw new Error('응답 해석 실패 (HTTP ' + code + '): ' + text.substring(0, 300)); }

  if (code !== 200) {
    const msg = (data && data.error && data.error.message) ? data.error.message : ('토큰 발급 실패 (HTTP ' + code + ')');
    throw new Error(msg);
  }
  if (data.error) throw new Error(data.error.message);
  if (!data.name) throw new Error('토큰 응답에 name 필드가 없습니다: ' + text);

  return data; // { name, expireTime, newSessionExpireTime }
}

/** 잠금(constrained) 토큰용 WebSocket URL */
function buildWsUrl_(tokenName) {
  return 'wss://generativelanguage.googleapis.com/ws/'
       + 'google.ai.generativelanguage.v1alpha.GenerativeService.BidiGenerateContentConstrained'
       + '?access_token=' + encodeURIComponent(tokenName);
}

/*───────────────────────────────────────────────────────────────────────
 * 유틸 & 테스트
 *───────────────────────────────────────────────────────────────────────*/
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ping() { return 'pong'; }                      // 통신 점검용
function testTopics()  { Logger.log(JSON.stringify(getTopics_(), null, 2)); }
function testToken()   { Logger.log(JSON.stringify(createEphemeralToken_(getModelName_()), null, 2)); }
function testSession() {
  const agenda = getTopics_()[0].agenda;
  Logger.log('테스트 어젠다: ' + agenda);
  Logger.log(JSON.stringify(getSession_(agenda, 'pro'), null, 2));
}
