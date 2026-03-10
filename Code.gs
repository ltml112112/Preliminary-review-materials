/**
 * LGD 사전심사자료 자동화 - Google Apps Script (v3)
 * =================================================
 * 플레이스홀더 방식: 시트 안의 [[작성일]], [[제품명]] 등을 찾아 치환
 * CORS 우회: doGet으로 모든 요청 처리 (POST 리다이렉트 CORS 문제 방지)
 *
 * 사용법:
 *   1. 이 코드를 구글 시트의 [확장 프로그램 > Apps Script]에 붙여넣기
 *   2. TEMPLATE_SPREADSHEET_ID를 템플릿 시트 ID로 설정
 *   3. [배포 > 새 배포 > 웹 앱] → 실행 주체: "나", 액세스: "모든 사용자"
 *   4. 배포 URL을 HTML의 APPS_SCRIPT_URL에 붙여넣기
 *
 * ★ 중요: 코드 수정 후 반드시 [배포 > 배포 관리 > 새 버전]으로 재배포!
 */

// ═══════════════════════════════════════════════
// ★ 설정
// ═══════════════════════════════════════════════

/** 원본 템플릿 스프레드시트 ID */
const TEMPLATE_SPREADSHEET_ID = '1fiHBRlwv1W_i4SGcARV7Q8hKidFuhHTa';

/** 플레이스홀더 목록 */
const PLACEHOLDERS = ['작성일', '제품명', '색상', '상품명1', '상품명2', '상품명3'];

/** 비공개물질 Checksheet 묶음 시트 */
const CHECKSHEET_BUNDLE = [
  'Class1(도입금지물질) List(23.08.10)',
  '안전보건 Check List(23.08.10)',
  '환경 Check List(25.11.25)',
];

// ═══════════════════════════════════════════════
// 엔트리 포인트 - GET으로 모든 요청 처리 (CORS 우회)
// ═══════════════════════════════════════════════

function doGet(e) {
  // payload 파라미터가 없으면 연결 테스트
  if (!e.parameter.payload) {
    return jsonResp({ ok: true, message: 'LGD 사전심사자료 Apps Script v3 연결 성공!' });
  }

  return processRequest(JSON.parse(e.parameter.payload));
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    return processRequest(req);
  } catch (err) {
    return jsonResp({ ok: false, error: err.message });
  }
}

// ═══════════════════════════════════════════════
// 요청 처리
// ═══════════════════════════════════════════════

function processRequest(req) {
  let tempFile = null;

  try {
    const action = req.action;     // 'test' | 'pdf' | 'xlsx' | 'checksheet'
    const data = req.data || {};

    if (action === 'test') {
      return jsonResp({ ok: true, message: '연결 성공' });
    }

    // ── 템플릿 복사 → 플레이스홀더 치환 ──
    tempFile = copyTemplate_(data);
    const tempSS = SpreadsheetApp.open(tempFile);

    // ── PDF: 개별 시트 ──
    if (action === 'pdf') {
      const sheetName = req.sheetName;
      const sheet = tempSS.getSheetByName(sheetName);
      if (!sheet) return jsonResp({ ok: false, error: '시트 없음: ' + sheetName });

      const pdfBytes = exportSheetAsPDF_(tempSS, sheet);
      return jsonResp({
        ok: true,
        fileData: Utilities.base64Encode(pdfBytes),
        fileName: buildFileName_(data, sheetName, 'pdf'),
        mimeType: 'application/pdf',
      });
    }

    // ── XLSX: 개별 시트 (MSDS 등) ──
    if (action === 'xlsx') {
      const sheetName = req.sheetName;
      const xlsxBytes = exportSingleSheetAsXLSX_(tempSS, sheetName);
      return jsonResp({
        ok: true,
        fileData: Utilities.base64Encode(xlsxBytes),
        fileName: buildFileName_(data, sheetName, 'xlsx'),
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
    }

    // ── Checksheet: 3개 시트 묶음 엑셀 ──
    if (action === 'checksheet') {
      const xlsxBytes = exportBundleAsXLSX_(tempSS, CHECKSHEET_BUNDLE);
      return jsonResp({
        ok: true,
        fileData: Utilities.base64Encode(xlsxBytes),
        fileName: 'LT소재_' + (data['제품명'] || '제품명') + '_비공개물질 Checksheet.xlsx',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
    }

    return jsonResp({ ok: false, error: '알 수 없는 action: ' + action });

  } catch (err) {
    return jsonResp({ ok: false, error: err.message });
  } finally {
    if (tempFile) {
      try { DriveApp.getFileById(tempFile.getId()).setTrashed(true); } catch (_) {}
    }
  }
}

// ═══════════════════════════════════════════════
// 핵심 함수
// ═══════════════════════════════════════════════

/** 템플릿 복사 + 플레이스홀더 치환 */
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

/** 상품명 개수에 따른 PDF 대상 시트 결정 */
function getPDFTargets(data) {
  const targets = ['MSDS', '경고표지'];

  const count = data['상품명3'] ? 3 : data['상품명2'] ? 2 : 1;
  targets.push('구성제품확인서' + count);

  targets.push('작업공정별관리요령', '비공개물질확인서');
  return targets;
}

/** 특정 시트 → PDF */
function exportSheetAsPDF_(ss, sheet) {
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?' +
    'exportFormat=pdf&format=pdf' +
    '&size=A4&portrait=true&fitw=true' +
    '&sheetnames=false&printtitle=false&pagenumbers=false' +
    '&gridlines=false&fzr=false' +
    '&gid=' + sheet.getSheetId();

  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });

  if (resp.getResponseCode() !== 200) {
    throw new Error('PDF 변환 실패 (' + sheet.getName() + '): HTTP ' + resp.getResponseCode());
  }
  return resp.getContent();
}

/** 특정 시트만 XLSX로 내보내기 */
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

/** 여러 시트를 묶어서 XLSX로 내보내기 */
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

/** 스프레드시트 → XLSX 바이트 */
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

// ═══════════════════════════════════════════════
// 유틸
// ═══════════════════════════════════════════════

function buildFileName_(data, sheetName, ext) {
  return (data['제품명'] || '제품명') + '_' + sheetName + '.' + ext;
}

function jsonResp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
