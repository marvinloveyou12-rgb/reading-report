/**
 * ============================================================================
 *  독서 진단도구 - Google Apps Script (설문지 응답 스프레드시트에 연결)
 * ============================================================================
 *
 * 【무엇을 하는 스크립트인가】
 *  - 이 스크립트가 붙는 "설문 응답 스프레드시트"의 데이터를 읽어
 *    ① 5개 대영역(독서흥미와 몰입 / 독서주도성과 소통 / 독서자신감 /
 *       독서가치와 목적인식 / 독서노력과 마음가짐)
 *    ② 6개 하위영역(어휘 / 사실 / 추론 / 비판 / 창의 / 주제통합)
 *    두 가지 기준으로 학생 개인 점수와 전체 평균을 계산한다.
 *  - 웹앱(doGet)으로 배포하면, 학년/반/번호를 입력해 자기 결과를
 *    방사형(레이더) 그래프 2개로 즉시 확인할 수 있다. (조회할 때마다
 *    시트를 다시 읽으므로 새 응답이 들어오면 별도 작업 없이 바로 반영됨)
 *  - 응답이 새로 들어올 때마다(onFormSubmit) "진단결과" 시트에 그 학생의
 *    점수를 자동으로 한 줄씩 기록해 둔다(교사용 기록/출력용, 선택 사항).
 *
 * ============================================================================
 *  【설치 방법】
 *  1) 구글 설문지 편집 화면 → 우측 상단 "응답" 탭 → 초록색 스프레드시트
 *     아이콘을 눌러 응답이 쌓이는 스프레드시트를 만든다(이미 만들어져
 *     있다면 그 시트를 연다).
 *  2) 그 스프레드시트에서 상단 메뉴 [확장 프로그램] → [Apps Script] 클릭.
 *  3) 기본으로 열린 Code.gs 내용을 전부 지우고, 이 파일 내용을 통째로
 *     붙여넣는다.
 *  4) 좌측 "+" → HTML → 파일명을 정확히 "WebApp"으로 만들고, 함께 드린
 *     WebApp.html 내용을 붙여넣는다. (파일명이 WebApp 이어야 함, 확장자는
 *     자동으로 .html 이 붙는다 — 실제 파일명은 WebApp.html 이 됨)
 *     같은 방식으로 "AdminApp"이라는 이름의 HTML 파일도 하나 더 만들고
 *     AdminApp.html 내용을 붙여넣는다. (교사용 관리자 대시보드 화면)
 *     아래 ADMIN_PASSCODE에 원하는 암호를 반드시 설정해 두세요.
 *  5) 상단 함수 선택 드롭다운에서 setupTrigger 선택 후 ▶ 실행 → 최초 1회
 *     권한 승인(본인 계정, 검토되지 않은 앱 경고가 뜨면 "고급"→"이동" 클릭).
 *     이 작업으로 onFormSubmit 트리거가 자동 등록된다.
 *  6) 우측 상단 [배포] → [새 배포] → 유형: "웹 앱" 선택
 *       - 실행 계정: 나(본인)
 *       - 액세스 권한: "전체" 또는 "링크가 있는 모든 사용자"
 *     [배포] 클릭 → 생성된 웹 앱 URL을 복사한다.
 *  7) 구글 설문지 편집 화면 → 설정(톱니바퀴) → [프레젠테이션] →
 *     "확인 메시지"란에 위 URL을 붙여넣은 안내문을 적어 두면, 학생이
 *     제출 직후 바로 자기 결과 페이지로 이동할 수 있다.
 *     예) "제출 완료! 결과 확인: https://script.google.com/macros/s/.../exec"
 *
 *  ※ 코드를 수정해서 다시 배포할 때는 [배포]→[배포 관리]→ 연필 아이콘 →
 *    새 버전으로 배포해야 URL이 최신 코드를 반영한다.
 * ============================================================================
 *
 *  【중요: 설문 문항 순서를 반드시 아래와 동일하게 맞춰 주세요】
 *  아래 QUESTION_MAP은 "진단도구_0718" 원본 문항 텍스트를 기준으로
 *  자동 매칭합니다(문항 앞의 번호 "1. " 등은 무시하고 비교). 구글 폼
 *  문항 문구가 원본 엑셀과 동일하면 그대로 동작합니다. 문구를 다소
 *  수정했다면 아래 QUESTION_MAP의 text 값도 같이 고쳐주세요.
 */

// ---------------------------------------------------------------------------
// 설정
// ---------------------------------------------------------------------------

// 응답 스프레드시트에서 응답이 쌓이는 시트 이름. null이면 첫 번째 시트를 사용.
const SHEET_NAME = null;

// 처리 결과를 기록해 둘 시트 이름(자동 생성됨).
const RESULT_SHEET_NAME = '진단결과';

// 척도 범위 (리커트 척도 1~5점이 아니라면 이 값을 수정하세요)
const SCALE_MIN = 1;
const SCALE_MAX = 5;

// 관리자 대시보드(전체 학생 목록) 접근 암호. 반드시 직접 정해서 바꿔주세요.
// 비워두면("") 대시보드가 암호 없이 열리니 꼭 설정할 것을 권장합니다.
const ADMIN_PASSCODE = '여기에_암호를_설정하세요';

// 학년/반/번호 열을 찾기 위한 헤더 힌트(포함되면 매칭)
// 더 구체적인 힌트를 앞에 두었지만, 실제로는 헤더를 앞에서부터 훑다가
// 아래 힌트 중 하나라도 포함된 첫 열을 채택합니다.
const ID_FIELD_HINTS = {
  grade: ['학년'],
  class: ['반'],
  number: ['번호', '학번', '출석번호', '순번', '번'],
};

// 36개 문항 정의: area(대영역) / subFactor(있으면) / category(하위영역) / text(문항 원문)
const QUESTION_MAP = [
  // Ⅰ. 독서흥미와 몰입
  { area: 'Ⅰ', areaLabel: '독서흥미와 몰입', subFactor: null, category: '어휘', text: '책을 읽을 때 새로운 단어를 알아가는 것이 즐겁다.' },
  { area: 'Ⅰ', areaLabel: '독서흥미와 몰입', subFactor: null, category: '사실', text: '책을 읽을 때 내용을 꼼꼼히 읽으려고 노력하는 것이 즐겁다.' },
  { area: 'Ⅰ', areaLabel: '독서흥미와 몰입', subFactor: null, category: '추론', text: '책에서 단서를 찾아 문제를 해결하거나, 깊은 뜻을 상상하며 읽는 것이 즐겁다.' },
  { area: 'Ⅰ', areaLabel: '독서흥미와 몰입', subFactor: null, category: '비판', text: '책의 내용이나 생각이 나와 맞는지 비교할 때 집중할 수 있고 즐겁다.' },
  { area: 'Ⅰ', areaLabel: '독서흥미와 몰입', subFactor: null, category: '창의', text: '책에서 본 내용을 내 생활이나 경험에 연결해 보는 것이 즐겁다.' },
  { area: 'Ⅰ', areaLabel: '독서흥미와 몰입', subFactor: null, category: '주제통합', text: '같은 주제의 여러 책을 비교하며 선택하여 읽는 것이 즐겁다.' },

  // Ⅱ. 독서주도성과 소통 - 주도성
  { area: 'Ⅱ', areaLabel: '독서주도성과 소통', subFactor: '주도성', category: '어휘', text: '책에 모르는 단어가 있을 때 알아내려고 노력한다.' },
  { area: 'Ⅱ', areaLabel: '독서주도성과 소통', subFactor: '주도성', category: '사실', text: '책의 내용을 잘 이해하고 기억하려고 중요한 내용을 스스로 메모하며 읽기도 한다.' },
  { area: 'Ⅱ', areaLabel: '독서주도성과 소통', subFactor: '주도성', category: '추론', text: '책을 읽는 방법이나 순서를 내가 원하는 대로 정해서 읽기도 한다.' },
  { area: 'Ⅱ', areaLabel: '독서주도성과 소통', subFactor: '주도성', category: '비판', text: '책을 읽을 때 글쓴이의 생각이 맞는지 스스로 질문을 던지며 읽기도 한다.' },
  { area: 'Ⅱ', areaLabel: '독서주도성과 소통', subFactor: '주도성', category: '창의', text: '책을 읽은 후 어떤 활동을 할지 스스로 정하여 실천하기도 한다.' },
  { area: 'Ⅱ', areaLabel: '독서주도성과 소통', subFactor: '주도성', category: '주제통합', text: '같은 주제의 여러 책을 스스로 찾아 비교하며 선택하여 읽기도 한다.' },

  // Ⅱ. 독서주도성과 소통 - 소통
  { area: 'Ⅱ', areaLabel: '독서주도성과 소통', subFactor: '소통', category: '어휘', text: '책에 모르는 단어가 있을 때 친구나 어른에게 물어보는 것이 좋다.' },
  { area: 'Ⅱ', areaLabel: '독서주도성과 소통', subFactor: '소통', category: '사실', text: '책에 있는 내용을 잘 이해하거나 기억하려고 친구나 어른들과 이야기하는 것이 좋다.' },
  { area: 'Ⅱ', areaLabel: '독서주도성과 소통', subFactor: '소통', category: '추론', text: '책 속에 숨어있는 깊은 뜻에 대해 친구나 어른들과 이야기하는 것이 좋다.' },
  { area: 'Ⅱ', areaLabel: '독서주도성과 소통', subFactor: '소통', category: '비판', text: '책을 읽고 서로 다른 생각에 대해 비교하며 친구나 어른들과 이야기하는 것이 좋다.' },
  { area: 'Ⅱ', areaLabel: '독서주도성과 소통', subFactor: '소통', category: '창의', text: '책을 읽고 새롭게 생각해 본 것을 친구나 어른들과 이야기하는 것이 좋다.' },
  { area: 'Ⅱ', areaLabel: '독서주도성과 소통', subFactor: '소통', category: '주제통합', text: '같은 주제의 여러 책을 비교하며 선택하여 읽는 것을 친구나 어른들과 이야기하는 것이 좋다.' },

  // Ⅲ. 독서자신감
  { area: 'Ⅲ', areaLabel: '독서자신감', subFactor: null, category: '어휘', text: '책을 읽을 때 어려운 단어가 나와도 잘 읽을 수 있다.' },
  { area: 'Ⅲ', areaLabel: '독서자신감', subFactor: null, category: '사실', text: '책을 읽을 때 책에 있는 핵심 내용을 잘 파악할 수 있다.' },
  { area: 'Ⅲ', areaLabel: '독서자신감', subFactor: null, category: '추론', text: '책에 직접 나오지 않아도 책 속에 숨어있는 뜻을 알아낼 수 있다.' },
  { area: 'Ⅲ', areaLabel: '독서자신감', subFactor: null, category: '비판', text: '책을 읽고 내 생각과 글쓴이의 생각을 비교하여 말할 자신이 있다.' },
  { area: 'Ⅲ', areaLabel: '독서자신감', subFactor: null, category: '창의', text: '책을 읽고 이해하고 생각한 것을 내 생활이나 경험에 연결해 실천할 수 있다.' },
  { area: 'Ⅲ', areaLabel: '독서자신감', subFactor: null, category: '주제통합', text: '같은 주제의 여러 책을 비교하며 선택하여 읽고 공통점과 차이점을 설명할 자신이 있다.' },

  // Ⅳ. 독서가치와 목적인식
  { area: 'Ⅳ', areaLabel: '독서가치와 목적인식', subFactor: null, category: '어휘', text: '책을 읽고 어려운 단어를 더 많이 알아가는 것은 중요하다.' },
  { area: 'Ⅳ', areaLabel: '독서가치와 목적인식', subFactor: null, category: '사실', text: '책 속에 있는 내용을 정확하게 이해하는 것은 중요하다.' },
  { area: 'Ⅳ', areaLabel: '독서가치와 목적인식', subFactor: null, category: '추론', text: '책의 깊은 뜻까지 생각해 보는 것은 중요하다.' },
  { area: 'Ⅳ', areaLabel: '독서가치와 목적인식', subFactor: null, category: '비판', text: '책을 읽을 때 나만의 생각과 기준을 가진 독자가 되는 것은 중요하다.' },
  { area: 'Ⅳ', areaLabel: '독서가치와 목적인식', subFactor: null, category: '창의', text: '책을 읽고 나서 창의적으로 다양한 활동을 해내는 것은 중요하다.' },
  { area: 'Ⅳ', areaLabel: '독서가치와 목적인식', subFactor: null, category: '주제통합', text: '같은 주제의 여러 책을 비교하며 선택하여 읽고 나만의 생각을 정리하는 것은 중요하다.' },

  // Ⅴ. 독서노력과 마음가짐
  { area: 'Ⅴ', areaLabel: '독서노력과 마음가짐', subFactor: null, category: '어휘', text: '책을 열심히 읽으면 어휘력이 점점 좋아진다.' },
  { area: 'Ⅴ', areaLabel: '독서노력과 마음가짐', subFactor: null, category: '사실', text: '책을 열심히 읽으면 내용을 정확히 파악할 수 있는 능력이 좋아진다.' },
  { area: 'Ⅴ', areaLabel: '독서노력과 마음가짐', subFactor: null, category: '추론', text: '책을 열심히 읽으면 책 속에 적혀 있지 않은 내용까지 생각할 수 있는 능력이 좋아진다.' },
  { area: 'Ⅴ', areaLabel: '독서노력과 마음가짐', subFactor: null, category: '비판', text: '책을 열심히 읽으면 다양한 생각을 엿보며 내 생각을 더 깊이 있게 정리하는 능력이 좋아진다.' },
  { area: 'Ⅴ', areaLabel: '독서노력과 마음가짐', subFactor: null, category: '창의', text: '책을 읽고 나서 다양한 활동을 열심히 해내면 창의력이 점점 좋아진다.' },
  { area: 'Ⅴ', areaLabel: '독서노력과 마음가짐', subFactor: null, category: '주제통합', text: '같은 주제의 여러 책을 비교하며 선택하여 읽으면 여러 가지 내용을 종합하는 능력이 좋아진다.' },
];

const AREA_ORDER = ['Ⅰ', 'Ⅱ', 'Ⅲ', 'Ⅳ', 'Ⅴ'];
const CATEGORY_ORDER = ['어휘', '사실', '추론', '비판', '창의', '주제통합'];

// ---------------------------------------------------------------------------
// 헤더 매칭
// ---------------------------------------------------------------------------

function normalizeHeader_(h) {
  return String(h || '')
    .replace(/^\s*\d+\.\s*/, '')  // 앞의 "1. " 같은 번호 제거
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * 시트 헤더 배열을 받아 { itemIndex(QUESTION_MAP의 index) -> columnIndex(0-based) }
 * 와 학년/반/번호 컬럼 인덱스를 찾아 돌려준다.
 */
function buildColumnMap_(headers) {
  const normHeaders = headers.map(normalizeHeader_);
  const itemColByIndex = {};
  const unmatched = [];

  QUESTION_MAP.forEach((q, i) => {
    const target = normalizeHeader_(q.text);
    let col = normHeaders.findIndex(h => h === target);
    if (col === -1) col = normHeaders.findIndex(h => h.indexOf(target) !== -1 || target.indexOf(h) !== -1);
    if (col === -1) {
      unmatched.push(i);
    } else {
      itemColByIndex[i] = col;
    }
  });

  // 텍스트로 못찾은 문항은, 학년/반/번호 이후 순서를 기준으로 위치 매칭 시도(백업)
  if (unmatched.length > 0) {
    const idCols = findIdColumns_(headers);
    const usedCols = new Set(Object.values(itemColByIndex));
    idCols.forEach(c => usedCols.add(c));
    let cursor = 0;
    const freeCols = headers.map((_, i) => i).filter(i => !usedCols.has(i) && i !== 0 /* Timestamp */);
    unmatched.forEach(i => {
      while (cursor < freeCols.length && usedCols.has(freeCols[cursor])) cursor++;
      if (cursor < freeCols.length) {
        itemColByIndex[i] = freeCols[cursor];
        usedCols.add(freeCols[cursor]);
        cursor++;
      }
    });
  }

  return { itemColByIndex, idCols: findIdColumns_(headers), unmatchedCount: 36 - Object.keys(itemColByIndex).length };
}

function findIdColumns_(headers) {
  const norm = headers.map(h => String(h || ''));
  function find(hints) {
    for (let i = 0; i < norm.length; i++) {
      for (const hint of hints) {
        if (norm[i].indexOf(hint) !== -1) return i;
      }
    }
    return -1;
  }
  return {
    grade: find(ID_FIELD_HINTS.grade),
    class: find(ID_FIELD_HINTS.class),
    number: find(ID_FIELD_HINTS.number),
  };
}

// ---------------------------------------------------------------------------
// 점수 계산
// ---------------------------------------------------------------------------

function average_(nums) {
  const valid = nums.filter(n => typeof n === 'number' && !isNaN(n));
  if (valid.length === 0) return null;
  return valid.reduce((a, b) => a + b, 0) / valid.length;
}

/**
 * 한 응답 행(row, 0-based 배열)에서 36문항 값을 뽑아
 * 5개 대영역 점수 + 6개 하위영역 점수를 계산한다.
 */
function computeScoresForRow_(row, itemColByIndex) {
  const values = QUESTION_MAP.map((q, i) => {
    const col = itemColByIndex[i];
    if (col === undefined) return null;
    const v = Number(row[col]);
    return isNaN(v) ? null : v;
  });

  const areaScores = {};
  AREA_ORDER.forEach(area => {
    const idxs = QUESTION_MAP.map((q, i) => (q.area === area ? i : -1)).filter(i => i !== -1);
    areaScores[area] = average_(idxs.map(i => values[i]));
  });

  const categoryScores = {};
  CATEGORY_ORDER.forEach(cat => {
    const idxs = QUESTION_MAP.map((q, i) => (q.category === cat ? i : -1)).filter(i => i !== -1);
    categoryScores[cat] = average_(idxs.map(i => values[i]));
  });

  return { areaScores, categoryScores, rawValues: values };
}

/**
 * 응답 시트 전체를 읽어 모든 학생의 점수 + 전체 평균을 계산한다.
 */
/**
 * 학년/반/번호 값을 비교용 키로 정규화한다.
 * "4", "4학년", " 4 ", "04" 등 표기가 달라도 숫자만 뽑아 "4"로 통일해서
 * 객관식/드롭다운 문항으로 만들었거나 앞뒤 공백·0이 붙은 경우에도 검색이 되게 한다.
 * 숫자가 전혀 없으면(드문 경우) 원래 문자열을 그대로 쓴다.
 */
function normalizeIdValue_(v) {
  const s = String(v == null ? '' : v).trim();
  const digits = s.match(/\d+/);
  return digits ? String(Number(digits[0])) : s;
}

function computeAllScores_() {
  const sheet = getResponseSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const { itemColByIndex, idCols, unmatchedCount } = buildColumnMap_(headers);

  const students = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const grade = idCols.grade !== -1 ? String(row[idCols.grade]).trim() : '';
    const cls = idCols.class !== -1 ? String(row[idCols.class]).trim() : '';
    const number = idCols.number !== -1 ? String(row[idCols.number]).trim() : '';
    if (!grade && !cls && !number) continue; // 빈 행 스킵

    const { areaScores, categoryScores } = computeScoresForRow_(row, itemColByIndex);
    students.push({
      rowIndex: r + 1, // 시트상 실제 행 번호(1-based, 헤더 포함)
      grade, cls, number,
      key: normalizeIdValue_(grade) + '-' + normalizeIdValue_(cls) + '-' + normalizeIdValue_(number),
      areaScores, categoryScores,
      timestamp: row[0],
    });
  }

  // 같은 학년-반-번호가 여러 번 제출되었으면 가장 최근(마지막) 응답만 사용
  const latestByKey = {};
  students.forEach(s => { latestByKey[s.key] = s; });
  const latestStudents = Object.values(latestByKey);

  const overallArea = {};
  AREA_ORDER.forEach(area => {
    overallArea[area] = average_(latestStudents.map(s => s.areaScores[area]).filter(v => v !== null));
  });
  const overallCategory = {};
  CATEGORY_ORDER.forEach(cat => {
    overallCategory[cat] = average_(latestStudents.map(s => s.categoryScores[cat]).filter(v => v !== null));
  });

  return {
    students, latestStudents, overallArea, overallCategory,
    unmatchedCount, count: latestStudents.length,
  };
}

function getResponseSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];
}

// ---------------------------------------------------------------------------
// onFormSubmit 트리거: 제출될 때마다 "진단결과" 시트에 한 줄 기록
// ---------------------------------------------------------------------------

function setupTrigger() {
  // 기존 동일 트리거 중복 등록 방지
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onFormSubmit_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onFormSubmit_')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
  // 주의: SpreadsheetApp.getUi().alert(...)는 스크립트 편집기에서 직접 실행할 때
  // "Cannot call SpreadsheetApp.getUi() from this context" 오류를 일으키므로 사용하지 않음.
  // 대신 실행 로그에 기록 — 편집기 좌측 "실행" 메뉴에서 결과를 확인할 수 있음.
  Logger.log('트리거 등록 완료! 이제 설문 제출 시 자동으로 "진단결과" 시트가 갱신됩니다.');
}

function onFormSubmit_(e) {
  writeResultSheet_();
}

function writeResultSheet_() {
  const all = computeAllScores_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(RESULT_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(RESULT_SHEET_NAME);
  sheet.clear();

  const headers = ['학년', '반', '번호',
    ...AREA_ORDER.map(a => QUESTION_MAP.find(q => q.area === a).areaLabel),
    ...CATEGORY_ORDER,
  ];
  sheet.appendRow(headers);

  all.latestStudents
    .sort((a, b) => (a.key > b.key ? 1 : -1))
    .forEach(s => {
      sheet.appendRow([
        s.grade, s.cls, s.number,
        ...AREA_ORDER.map(a => round1_(s.areaScores[a])),
        ...CATEGORY_ORDER.map(c => round1_(s.categoryScores[c])),
      ]);
    });

  sheet.appendRow(new Array(headers.length).fill('')); // 빈 줄 하나 (appendRow는 완전히 빈 배열을 허용하지 않음)
  sheet.appendRow(['전체 평균', '', '',
    ...AREA_ORDER.map(a => round1_(all.overallArea[a])),
    ...CATEGORY_ORDER.map(c => round1_(all.overallCategory[c])),
  ]);

  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.autoResizeColumns(1, headers.length);
}

function round1_(v) {
  return v === null || v === undefined ? '' : Math.round(v * 10) / 10;
}

// 수동으로 "진단결과" 시트를 다시 생성하고 싶을 때 이 함수를 실행하면 됨
function manualRefreshResultSheet() {
  writeResultSheet_();
}

// ---------------------------------------------------------------------------
// 웹앱: 학년/반/번호 입력 → 방사형 그래프 2개
// ---------------------------------------------------------------------------

function doGet(e) {
  const params = (e && e.parameter) || {};
  const isAdmin = params.view === 'admin';

  const template = HtmlService.createTemplateFromFile(isAdmin ? 'AdminApp' : 'WebApp');
  template.scaleMin = SCALE_MIN;
  template.scaleMax = SCALE_MAX;
  if (!isAdmin) {
    template.prefillGrade = params.grade || '';
    template.prefillClass = params.cls || '';
    template.prefillNumber = params.number || '';
  }
  return template.evaluate()
    .setTitle(isAdmin ? '독서 진단 관리자 대시보드' : '독서 진단 결과')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * AdminApp.html에서 google.script.run으로 호출한다.
 * 암호가 맞으면 전체 학생 명단 + 5대영역/6하위영역 점수 + 전체 평균을 돌려준다.
 * (교사만 사용하는 화면이므로 학생용 조회 링크와는 분리해서 안내할 것)
 */
function getAdminRoster(passcode) {
  // String()으로 양쪽 다 감싸서, ADMIN_PASSCODE를 숫자로 적었어도(예: 1234, 따옴표 없이)
  // 입력값(항상 문자열)과 타입이 달라 비교가 실패하는 일이 없게 한다.
  if (ADMIN_PASSCODE && String(passcode).trim() !== String(ADMIN_PASSCODE).trim()) {
    return { ok: false, reason: 'passcode' };
  }
  const all = computeAllScores_();
  const overall = averagesForPool_(all.latestStudents);

  const areaKeys = AREA_ORDER;
  const areaLabels = AREA_ORDER.map(a => QUESTION_MAP.find(q => q.area === a).areaLabel);
  const categoryLabels = CATEGORY_ORDER;

  const roster = all.latestStudents
    .slice()
    .sort((a, b) => (a.key > b.key ? 1 : -1))
    .map(s => {
      const area = AREA_ORDER.map(a => s.areaScores[a]);
      const category = CATEGORY_ORDER.map(c => s.categoryScores[c]);
      // 가장 낮은 하위영역(=가장 관심이 필요한 부분)을 하나 뽑아 요약으로 제공
      let weakestCat = null, weakestScore = Infinity;
      CATEGORY_ORDER.forEach((c, i) => {
        if (category[i] != null && category[i] < weakestScore) { weakestScore = category[i]; weakestCat = c; }
      });
      return {
        grade: s.grade, cls: s.cls, number: s.number,
        area, category, weakestCat,
        timestamp: s.timestamp instanceof Date ? s.timestamp.toISOString() : String(s.timestamp || ''),
      };
    });

  return {
    ok: true,
    count: all.count,
    unmatchedCount: all.unmatchedCount,
    areaKeys, areaLabels, categoryLabels,
    overallArea: AREA_ORDER.map(a => overall.area[a]),
    overallCategory: CATEGORY_ORDER.map(c => overall.category[c]),
    roster,
  };
}

function averagesForPool_(pool) {
  const area = {}, category = {};
  AREA_ORDER.forEach(a => { area[a] = average_(pool.map(s => s.areaScores[a]).filter(v => v !== null)); });
  CATEGORY_ORDER.forEach(c => { category[c] = average_(pool.map(s => s.categoryScores[c]).filter(v => v !== null)); });
  return { area, category, count: pool.length };
}

/**
 * WebApp.html에서 google.script.run으로 호출한다.
 * 학년/반/번호로 학생 1명을 찾아 5개 대영역 + 6개 하위영역 점수와,
 * 비교에 쓸 수 있도록 전체/같은 학년/같은 반 평균을 한 번에 함께 돌려준다.
 * (화면에서는 비교 기준을 눌러도 서버를 다시 호출하지 않고 바로 전환한다)
 */
function getStudentReport(grade, cls, number) {
  const all = computeAllScores_();
  const key = normalizeIdValue_(grade) + '-' + normalizeIdValue_(cls) + '-' + normalizeIdValue_(number);
  const student = all.latestStudents.find(s => s.key === key);

  if (!student) {
    return { found: false, count: all.count };
  }

  const gradePool = all.latestStudents.filter(s => s.grade === student.grade);
  const classPool = all.latestStudents.filter(s => s.grade === student.grade && s.cls === student.cls);
  const overall = averagesForPool_(all.latestStudents);
  const byGrade = averagesForPool_(gradePool.length >= 2 ? gradePool : all.latestStudents);
  const byClass = averagesForPool_(classPool.length >= 2 ? classPool : all.latestStudents);

  return {
    found: true,
    count: all.count,
    grade: student.grade, cls: student.cls, number: student.number,
    areaLabels: AREA_ORDER.map(a => QUESTION_MAP.find(q => q.area === a).areaLabel),
    areaKeys: AREA_ORDER,
    categoryLabels: CATEGORY_ORDER,
    studentArea: AREA_ORDER.map(a => student.areaScores[a]),
    studentCategory: CATEGORY_ORDER.map(c => student.categoryScores[c]),
    scopes: {
      all: { label: '전체 평균', count: overall.count, area: AREA_ORDER.map(a => overall.area[a]), category: CATEGORY_ORDER.map(c => overall.category[c]) },
      grade: { label: '같은 학년 평균', count: byGrade.count, area: AREA_ORDER.map(a => byGrade.area[a]), category: CATEGORY_ORDER.map(c => byGrade.category[c]) },
      class: { label: '같은 반 평균', count: byClass.count, area: AREA_ORDER.map(a => byClass.area[a]), category: CATEGORY_ORDER.map(c => byClass.category[c]) },
    },
  };
}
