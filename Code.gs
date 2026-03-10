/**
 * LGD 사전심사자료 자동화 - Google Apps Script (v2)
 * =================================================
 * 플레이스홀더 방식: 시트 안의 [[작성일]], [[제품명]] 등을 찾아 치환
 *
 * 사용법:
 *   1. 이 코드를 구글 시트의 [확장 프로그램 > Apps Script]에 붙여넣기
 *   2. TEMPLATE_SPREADSHEET_ID를 템플릿 시트 ID로 설정
 *   3. [배포 > 새 배포 > 웹 앱] → 실행 주체: "나", 액세스: "모든 사용자"
 *   4. 배포 URL을 HTML의 APPS_SCRIPT_URL에 붙여넣기
 */

// ═══════════════════════════════════════════════
// ★ 설정
// ═══════════════════════════════════════════════

/** 원본 템플릿 스프레드시트 ID (이 파일은 수정하지 않고 복사해서 사용) */
const TEMPLATE_SPREADSHEET_ID = '1fiHBRlwv1W_i4SGcARV7Q8hKidFuhHTa';

/** 플레이스홀더 목록 - 시트에서 이 텍스트를 찾아 치환합니다 */
const PLACEHOLDERS = ['작성일', '제품명', '색상', '상품명1', '상품명2', '상품명3'];

/** 개별 PDF 다운로드 대상 시트 목록 */
const PDF_SHEETS = [
  'MSDS',
  '경고표지',
  '구성제품확인서1',
  '구성제품확인서2',
  '구성제품확인서3',
  '작업공정별관리요령',
  '비공개물질확인서',
];

/** MSDS는 엑셀로도 다운로드 */
const XLSX_SINGLE_SHEETS = ['MSDS'];

/** 비공개물질 Checksheet 묶음 (하나의 엑셀로 다운로드) */
const CHECKSHEET_BUNDLE = [
  'Class1(도입금지물질) List(23.08.10)',
  '안전보건 Check List(23.08.10)',
  '환경 Check List(25.11.25)',
];

// ═══════════════════════════════════════════════
// 엔트리 포인트
// ═══════════════════════════════════════════════

function doGet(e) {
  return jsonResp({ ok: true, message: 'LGD 사전심사자료 Apps Script 연결 성공!', version: 'v2' });
}

function doPost(e) {
  let tempFile = null;

  try {
    const req = JSON.parse(e.postData.contents);
    const action = req.action;  // 'pdf' | 'xlsx' | 'checksheet' | 'test'
    const data = req.data || {}; // { 작성일, 제품명, 색상, 상품명1, 상품명2, 상품명3 }

    // ── 연결 테스트 ──
    if (action === 'test') {
      return jsonResp({ ok: true, message: '연결 성공' });
    }

    // ── 템플릿 복사 → 값 치환 → 작업 ──
    tempFile = copyTemplate_(data);
    const tempSS = SpreadsheetApp.open(tempFile);

    if (action === 'pdf') {
      // 개별 시트 PDF
      const sheetName = req.sheetName;
      const sheet = tempSS.getSheetByName(sheetName);
      if (!sheet) return jsonResp({ ok: false, error: '시트 없음: ' + sheetName });

      const pdfBytes = exportSheetAsPDF_(tempSS, sheet);
      return jsonResp({
        ok: true,
        pdf: Utilities.base64Encode(pdfBytes),
        fileName: buildFileName_(data, sheetName, 'pdf'),
      });
    }

    if (action === 'xlsx') {
      // 개별 시트 엑셀 (MSDS 등)
      const sheetName = req.sheetName;
      const xlsxBytes = exportSingleSheetAsXLSX_(tempSS, sheetName);
      return jsonResp({
        ok: true,
        xlsx: Utilities.base64Encode(xlsxBytes),
        fileName: buildFileName_(data, sheetName, 'xlsx'),
      });
    }

    if (action === 'checksheet') {
      // 비공개물질 Checksheet 묶음 엑셀
      const xlsxBytes = exportBundleAsXLSX_(tempSS, CHECKSHEET_BUNDLE);
      return jsonResp({
        ok: true,
        xlsx: Utilities.base64Encode(xlsxBytes),
        fileName: `LT소재_${data['제품명'] || '제품명'}_비공개물질 Checksheet.xlsx`,
      });
    }

    if (action === 'all') {
      // 전체 일괄 다운로드: 필요한 파일들을 한번에 생성
      const results = [];
      const productCount = getProductCount_(data);

      // PDF 다운로드 목록 결정
      const pdfTargets = getPDFTargets_(productCount);

      for (const sheetName of pdfTargets) {
        const sheet = tempSS.getSheetByName(sheetName);
        if (!sheet) continue;
        const pdfBytes = exportSheetAsPDF_(tempSS, sheet);
        results.push({
          type: 'pdf',
          data: Utilities.base64Encode(pdfBytes),
          fileName: buildFileName_(data, sheetName, 'pdf'),
        });
      }

      // MSDS 엑셀
      const msdsXlsx = exportSingleSheetAsXLSX_(tempSS, 'MSDS');
      results.push({
        type: 'xlsx',
        data: Utilities.base64Encode(msdsXlsx),
        fileName: buildFileName_(data, 'MSDS', 'xlsx'),
      });

      // 비공개물질 Checksheet 엑셀
      const checkXlsx = exportBundleAsXLSX_(tempSS, CHECKSHEET_BUNDLE);
      results.push({
        type: 'xlsx',
        data: Utilities.base64Encode(checkXlsx),
        fileName: `LT소재_${data['제품명'] || '제품명'}_비공개물질 Checksheet.xlsx`,
      });

      return jsonResp({ ok: true, files: results });
    }

    return jsonResp({ ok: false, error: '알 수 없는 action: ' + action });

  } catch (err) {
    return jsonResp({ ok: false, error: err.message, stack: err.stack });
  } finally {
    // 임시 파일 정리
    if (tempFile) {
      try { DriveApp.getFileById(tempFile.getId()).setTrashed(true); } catch (_) {}
    }
  }
}

// ═══════════════════════════════════════════════
// 핵심 함수
// ═══════════════════════════════════════════════

/**
 * 템플릿 스프레드시트를 복사하고 플레이스홀더를 치환
 */
function copyTemplate_(data) {
  const templateFile = DriveApp.getFileById(TEMPLATE_SPREADSHEET_ID);
  const tempFile = templateFile.makeCopy('_temp_' + new Date().getTime());
  const tempSS = SpreadsheetApp.open(tempFile);

  // 모든 시트에서 플레이스홀더 치환
  const sheets = tempSS.getSheets();
  for (const sheet of sheets) {
    for (const placeholder of PLACEHOLDERS) {
      const value = data[placeholder] || '';
      if (!value) continue;

      const finder = sheet.createTextFinder('[[' + placeholder + ']]')
        .matchCase(true)
        .matchEntireCell(false);
      finder.replaceAllWith(value);
    }
  }

  SpreadsheetApp.flush();
  return tempFile;
}

/**
 * 상품명 개수 파악
 */
function getProductCount_(data) {
  if (data['상품명3']) return 3;
  if (data['상품명2']) return 2;
  return 1;
}

/**
 * 상품명 개수에 따른 PDF 출력 대상 시트 결정
 * - 상품명 1개: MSDS, 경고표지, 구성제품확인서1, 작업공정별관리요령, 비공개물질확인서
 * - 상품명 2개: 구성제품확인서1 대신 구성제품확인서2
 * - 상품명 3개: 구성제품확인서1 대신 구성제품확인서3
 */
function getPDFTargets_(productCount) {
  const base = ['MSDS', '경고표지'];

  if (productCount === 1) {
    base.push('구성제품확인서1');
  } else if (productCount === 2) {
    base.push('구성제품확인서2');
  } else {
    base.push('구성제품확인서3');
  }

  base.push('작업공정별관리요령', '비공개물질확인서');
  return base;
}

/**
 * 특정 시트를 PDF로 변환
 */
function exportSheetAsPDF_(ss, sheet) {
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?' +
    'exportFormat=pdf&format=pdf' +
    '&size=A4&portrait=true&fitw=true' +
    '&sheetnames=false&printtitle=false&pagenumbers=false' +
    '&gridlines=false&fzr=false' +
    '&gid=' + sheet.getSheetId();

  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('PDF 변환 실패 (' + sheet.getName() + '): HTTP ' + response.getResponseCode());
  }

  return response.getContent();
}

/**
 * 특정 시트만 엑셀로 내보내기
 * (임시 스프레드시트에서 해당 시트만 남기고 나머지 삭제 후 xlsx 변환)
 */
function exportSingleSheetAsXLSX_(ss, sheetName) {
  // 원본 임시파일을 또 복사해서 시트 정리
  const file = DriveApp.getFileById(ss.getId()).makeCopy('_xlsx_temp_' + Date.now());
  const tempSS = SpreadsheetApp.open(file);

  const targetSheet = tempSS.getSheetByName(sheetName);
  if (!targetSheet) {
    DriveApp.getFileById(file.getId()).setTrashed(true);
    throw new Error('시트 없음: ' + sheetName);
  }

  // 대상 시트 외 모두 삭제
  const allSheets = tempSS.getSheets();
  for (const s of allSheets) {
    if (s.getSheetId() !== targetSheet.getSheetId()) {
      tempSS.deleteSheet(s);
    }
  }
  SpreadsheetApp.flush();

  const xlsxBytes = fetchXLSX_(tempSS.getId());
  DriveApp.getFileById(file.getId()).setTrashed(true);
  return xlsxBytes;
}

/**
 * 특정 시트들만 묶어서 엑셀로 내보내기 (비공개물질 Checksheet)
 */
function exportBundleAsXLSX_(ss, sheetNames) {
  const file = DriveApp.getFileById(ss.getId()).makeCopy('_bundle_temp_' + Date.now());
  const tempSS = SpreadsheetApp.open(file);

  // 번들에 포함된 시트만 남기고 삭제
  const keepIds = new Set();
  for (const name of sheetNames) {
    const s = tempSS.getSheetByName(name);
    if (s) keepIds.add(s.getSheetId());
  }

  const allSheets = tempSS.getSheets();
  for (const s of allSheets) {
    if (!keepIds.has(s.getSheetId())) {
      tempSS.deleteSheet(s);
    }
  }
  SpreadsheetApp.flush();

  const xlsxBytes = fetchXLSX_(tempSS.getId());
  DriveApp.getFileById(file.getId()).setTrashed(true);
  return xlsxBytes;
}

/**
 * 스프레드시트를 xlsx로 다운로드
 */
function fetchXLSX_(ssId) {
  const url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?exportFormat=xlsx';
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('XLSX 변환 실패: HTTP ' + response.getResponseCode());
  }

  return response.getContent();
}

// ═══════════════════════════════════════════════
// 유틸
// ═══════════════════════════════════════════════

function buildFileName_(data, sheetName, ext) {
  const product = data['제품명'] || '제품명';
  return product + '_' + sheetName + '.' + ext;
}

function jsonResp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
