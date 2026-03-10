/**
 * LGD 사전심사자료 자동화 - Google Apps Script (v4)
 * =================================================
 * HtmlService + google.script.run 방식 (CORS 문제 완전 해결)
 *
 * 사용법:
 *   1. 구글 시트 > 확장 프로그램 > Apps Script
 *   2. Code.gs에 이 코드 붙여넣기
 *   3. 파일 추가(+) > HTML > 이름: "Index" > Index.html 내용 붙여넣기
 *   4. [배포 > 새 배포 > 웹 앱] → 실행 주체: "나", 액세스: "모든 사용자"
 *   5. 배포 URL을 브라우저에서 열면 입력 폼이 나타남
 */

// ═══════════════════════════════════════════════
// ★ 설정
// ═══════════════════════════════════════════════

const TEMPLATE_SPREADSHEET_ID = '1fiHBRlwv1W_i4SGcARV7Q8hKidFuhHTa';

const PLACEHOLDERS = ['작성일', '제품명', '색상', '상품명1', '상품명2', '상품명3'];

const CHECKSHEET_BUNDLE = [
  'Class1(도입금지물질) List(23.08.10)',
  '안전보건 Check List(23.08.10)',
  '환경 Check List(25.11.25)',
];

// 시트별 PDF 내보내기 설정
// portrait: true=세로, false=가로 | margins: 인치 단위 (top, bottom, left, right)
// scale: 1=기본, 2=너비맞춤, 3=높이맞춤, 4=페이지맞춤
const SHEET_PDF_CONFIG = {
  'MSDS':             { portrait: true,  scale: 4, top: 0.2, bottom: 0.2, left: 0.2, right: 0.2 },
  '경고표지':          { portrait: false, scale: 4, top: 0.2, bottom: 0.2, left: 0.2, right: 0.2 },
  '구성제품확인서1':    { portrait: true,  scale: 4, top: 0.2, bottom: 0.2, left: 0.2, right: 0.2 },
  '구성제품확인서2':    { portrait: true,  scale: 4, top: 0.2, bottom: 0.2, left: 0.2, right: 0.2 },
  '구성제품확인서3':    { portrait: true,  scale: 4, top: 0.2, bottom: 0.2, left: 0.2, right: 0.2 },
  '작업공정별관리요령':  { portrait: false, scale: 4, top: 0.2, bottom: 0.2, left: 0.2, right: 0.2 },
  '비공개물질확인서':   { portrait: true,  scale: 4, top: 0.2, bottom: 0.2, left: 0.2, right: 0.2 },
};

// ═══════════════════════════════════════════════
// 웹 앱 진입점: HTML 페이지 제공
// ═══════════════════════════════════════════════

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('LGD 사전심사자료 자동화')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ═══════════════════════════════════════════════
// 클라이언트에서 호출하는 함수들 (google.script.run)
// ═══════════════════════════════════════════════

/** 연결 테스트 */
function testConnection() {
  return { ok: true, message: 'Apps Script v4 연결 성공!' };
}

/** 다운로드 대상 목록 반환 */
function getDownloadList(data) {
  const count = data['상품명3'] ? 3 : data['상품명2'] ? 2 : 1;
  const tasks = [];

  // PDF 대상
  const pdfSheets = ['MSDS', '경고표지', '구성제품확인서' + count, '작업공정별관리요령', '비공개물질확인서'];
  for (const name of pdfSheets) {
    tasks.push({ action: 'pdf', sheetName: name, label: name + ' (PDF)' });
  }

  // MSDS 엑셀
  tasks.push({ action: 'xlsx', sheetName: 'MSDS', label: 'MSDS (엑셀)' });

  // 비공개물질 Checksheet 묶음
  tasks.push({ action: 'checksheet', sheetName: null, label: '비공개물질 Checksheet (엑셀)' });

  return tasks;
}

/**
 * 파일 1개 생성 (개별 호출)
 * @param {string} action - 'pdf' | 'xlsx' | 'checksheet'
 * @param {string|null} sheetName - 시트 이름
 * @param {Object} data - { 작성일, 제품명, 색상, 상품명1, 상품명2, 상품명3 }
 * @returns {{ ok, fileData, fileName, mimeType }}
 */
function generateFile(action, sheetName, data) {
  let tempFile = null;

  try {
    tempFile = copyTemplate_(data);
    const tempSS = SpreadsheetApp.open(tempFile);

    if (action === 'pdf') {
      const sheet = tempSS.getSheetByName(sheetName);
      if (!sheet) return { ok: false, error: '시트 없음: ' + sheetName };

      const pdfBytes = exportSheetAsPDF_(tempSS, sheet);
      return {
        ok: true,
        fileData: Utilities.base64Encode(pdfBytes),
        fileName: buildFileName_(data, sheetName, 'pdf'),
        mimeType: 'application/pdf',
      };
    }

    if (action === 'xlsx') {
      const xlsxBytes = exportSingleSheetAsXLSX_(tempSS, sheetName);
      return {
        ok: true,
        fileData: Utilities.base64Encode(xlsxBytes),
        fileName: buildFileName_(data, sheetName, 'xlsx'),
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      };
    }

    if (action === 'checksheet') {
      const xlsxBytes = exportBundleAsXLSX_(tempSS, CHECKSHEET_BUNDLE);
      return {
        ok: true,
        fileData: Utilities.base64Encode(xlsxBytes),
        fileName: 'LT소재_' + (data['제품명'] || '제품명') + '_비공개물질 Checksheet.xlsx',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      };
    }

    return { ok: false, error: '알 수 없는 action: ' + action };

  } catch (err) {
    return { ok: false, error: err.message };
  } finally {
    if (tempFile) {
      try { DriveApp.getFileById(tempFile.getId()).setTrashed(true); } catch (_) {}
    }
  }
}

// ═══════════════════════════════════════════════
// 내부 함수
// ═══════════════════════════════════════════════

function copyTemplate_(data) {
  const templateFile = DriveApp.getFileById(TEMPLATE_SPREADSHEET_ID);
  const tempFile = templateFile.makeCopy('_temp_' + Date.now());
  const tempSS = SpreadsheetApp.open(tempFile);

  for (const sheet of tempSS.getSheets()) {
    for (const ph of PLACEHOLDERS) {
      const value = data[ph] || '';
      if (!value) continue;
      sheet.createTextFinder('[[' + ph + ']]')
        .matchCase(true)
        .matchEntireCell(false)
        .replaceAllWith(value);
    }
  }

  SpreadsheetApp.flush();
  return tempFile;
}

function exportSheetAsPDF_(ss, sheet) {
  const name = sheet.getName();
  const cfg = SHEET_PDF_CONFIG[name] || { portrait: true, scale: 4, top: 0.3, bottom: 0.3, left: 0.3, right: 0.3 };

  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?' +
    'exportFormat=pdf&format=pdf' +
    '&size=A4' +
    '&portrait=' + cfg.portrait +
    '&scale=' + cfg.scale +
    '&top_margin=' + cfg.top +
    '&bottom_margin=' + cfg.bottom +
    '&left_margin=' + cfg.left +
    '&right_margin=' + cfg.right +
    '&sheetnames=false&printtitle=false&pagenumbers=false' +
    '&gridlines=false&fzr=false' +
    '&gid=' + sheet.getSheetId();

  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });

  if (resp.getResponseCode() !== 200) {
    throw new Error('PDF 변환 실패 (' + name + '): HTTP ' + resp.getResponseCode());
  }
  return resp.getContent();
}

function exportSingleSheetAsXLSX_(ss, sheetName) {
  const file = DriveApp.getFileById(ss.getId()).makeCopy('_xlsx_temp_' + Date.now());
  const tempSS = SpreadsheetApp.open(file);

  const target = tempSS.getSheetByName(sheetName);
  if (!target) {
    DriveApp.getFileById(file.getId()).setTrashed(true);
    throw new Error('시트 없음: ' + sheetName);
  }

  for (const s of tempSS.getSheets()) {
    if (s.getSheetId() !== target.getSheetId()) tempSS.deleteSheet(s);
  }
  SpreadsheetApp.flush();

  const bytes = fetchXLSX_(tempSS.getId());
  DriveApp.getFileById(file.getId()).setTrashed(true);
  return bytes;
}

function exportBundleAsXLSX_(ss, sheetNames) {
  const file = DriveApp.getFileById(ss.getId()).makeCopy('_bundle_temp_' + Date.now());
  const tempSS = SpreadsheetApp.open(file);

  const keepIds = new Set();
  for (const name of sheetNames) {
    const s = tempSS.getSheetByName(name);
    if (s) keepIds.add(s.getSheetId());
  }

  for (const s of tempSS.getSheets()) {
    if (!keepIds.has(s.getSheetId())) tempSS.deleteSheet(s);
  }
  SpreadsheetApp.flush();

  const bytes = fetchXLSX_(tempSS.getId());
  DriveApp.getFileById(file.getId()).setTrashed(true);
  return bytes;
}

function fetchXLSX_(ssId) {
  const url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?exportFormat=xlsx';
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error('XLSX 변환 실패: HTTP ' + resp.getResponseCode());
  }
  return resp.getContent();
}

function buildFileName_(data, sheetName, ext) {
  return (data['제품명'] || '제품명') + '_' + sheetName + '.' + ext;
}
