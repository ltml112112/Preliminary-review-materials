/**
 * LGD 사전심사자료 자동화 - Google Apps Script (v5)
 * =================================================
 * HtmlService + google.script.run 방식
 * PDF 병렬 생성 (UrlFetchApp.fetchAll) 으로 속도 최적화
 */

// ═══════════════════════════════════════════════
// ★ 설정
// ═══════════════════════════════════════════════

const TEMPLATE_SPREADSHEET_ID = '1kh2oBZYKXaadIJoZQJ5OPYZHlwZftiFpuIT45v2SjTk';

const PLACEHOLDERS = ['작성일', '제품명', '색상', '상품명1', '상품명2', '상품명3'];

const CHECKSHEET_BUNDLE = [
  'Class1(도입금지물질) List(23.08.10)',
  '안전보건 Check List(23.08.10)',
  '환경 Check List(25.11.25)',
];

// 시트별 PDF 내보내기 설정
// portrait: true=세로, false=가로 | margins: 인치 단위 (top, bottom, left, right)
// scale: 1=기본, 2=너비맞춤, 3=높이맞춤, 4=페이지맞춤
// hAlign: CENTER/LEFT/RIGHT, vAlign: TOP/MIDDLE/BOTTOM
const SHEET_PDF_CONFIG = {
  'MSDS':             { portrait: true,  scale: 2, top: 0.75, bottom: 0.75, left: 0.7, right: 0.7, hAlign: 'CENTER', vAlign: 'TOP' },
  '경고표지':          { portrait: false, scale: 3, top: 0, bottom: 0, left: 0.24, right: 0.24, hAlign: 'CENTER', vAlign: 'MIDDLE' },
  '구성제품확인서1':    { portrait: true,  scale: 2, top: 0.75, bottom: 0, left: 0.24, right: 0.24, hAlign: 'CENTER', vAlign: 'TOP' },
  '구성제품확인서2':    { portrait: true,  scale: 2, top: 0.75, bottom: 0, left: 0.24, right: 0.24, hAlign: 'CENTER', vAlign: 'TOP' },
  '구성제품확인서3':    { portrait: true,  scale: 2, top: 0.75, bottom: 0, left: 0.24, right: 0.24, hAlign: 'CENTER', vAlign: 'TOP' },
  '작업공정별관리요령':  { portrait: false, scale: 4, top: 0, bottom: 0, left: 0, right: 0, hAlign: 'CENTER', vAlign: 'MIDDLE' },
  '비공개물질확인서':   { portrait: true,  scale: 4, top: 0.2, bottom: 0.2, left: 0.2, right: 0.2, hAlign: 'CENTER', vAlign: 'TOP' },
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
  return { ok: true, message: 'Apps Script v5 연결 성공!' };
}

/**
 * 전체 파일 일괄 생성 (서버 1회 호출)
 * 템플릿 1번 복사 → PDF 병렬 생성 + XLSX 병렬 생성
 */
function generateAllFiles(data) {
  let tempFile = null;

  try {
    const count = data['상품명3'] ? 3 : data['상품명2'] ? 2 : 1;
    tempFile = copyTemplate_(data);
    const tempSS = SpreadsheetApp.open(tempFile);
    const results = [];

    // ── PDF 5개 병렬 생성 (UrlFetchApp.fetchAll) ──
    const pdfSheetNames = ['MSDS', '경고표지', '구성제품확인서' + count, '작업공정별관리요령', '비공개물질확인서'];
    const pdfRequests = [];
    const pdfMeta = []; // 요청과 매칭할 메타 정보

    const token = ScriptApp.getOAuthToken();
    for (const name of pdfSheetNames) {
      const sheet = tempSS.getSheetByName(name);
      if (!sheet) {
        results.push({ ok: false, label: name + ' (PDF)', error: '시트 없음' });
        continue;
      }
      const cfg = SHEET_PDF_CONFIG[name] || { portrait: true, scale: 4, top: 0.3, bottom: 0.3, left: 0.3, right: 0.3 };
      const url = 'https://docs.google.com/spreadsheets/d/' + tempSS.getId() + '/export?' +
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
        '&horizontal_alignment=' + (cfg.hAlign || 'CENTER') +
        '&vertical_alignment=' + (cfg.vAlign || 'TOP') +
        '&gid=' + sheet.getSheetId();

      pdfRequests.push({ url: url, headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
      pdfMeta.push({ name: name });
    }

    // 병렬 fetch!
    if (pdfRequests.length > 0) {
      const pdfResponses = UrlFetchApp.fetchAll(pdfRequests);
      for (let i = 0; i < pdfResponses.length; i++) {
        const resp = pdfResponses[i];
        const name = pdfMeta[i].name;
        if (resp.getResponseCode() === 200) {
          results.push({
            ok: true, label: name + ' (PDF)',
            fileData: Utilities.base64Encode(resp.getContent()),
            fileName: buildFileName_(data, name, 'pdf'),
            mimeType: 'application/pdf',
          });
        } else {
          results.push({ ok: false, label: name + ' (PDF)', error: 'HTTP ' + resp.getResponseCode() });
        }
      }
    }

    // ── XLSX 2개 병렬 생성 ──
    // MSDS 단일시트용 복사본 + Checksheet 묶음용 복사본을 동시 생성
    const mainFileId = tempSS.getId();
    const xlsxCopy1 = DriveApp.getFileById(mainFileId).makeCopy('_xlsx1_' + Date.now());
    const xlsxCopy2 = DriveApp.getFileById(mainFileId).makeCopy('_xlsx2_' + Date.now());

    try {
      // MSDS 단일시트 - 불필요 시트 삭제
      const ss1 = SpreadsheetApp.open(xlsxCopy1);
      const msdsSheet = ss1.getSheetByName('MSDS');
      if (msdsSheet) {
        for (const s of ss1.getSheets()) {
          if (s.getSheetId() !== msdsSheet.getSheetId()) ss1.deleteSheet(s);
        }
      }

      // Checksheet 묶음 - 불필요 시트 삭제
      const ss2 = SpreadsheetApp.open(xlsxCopy2);
      const keepIds = new Set();
      for (const bname of CHECKSHEET_BUNDLE) {
        const s = ss2.getSheetByName(bname);
        if (s) keepIds.add(s.getSheetId());
      }
      for (const s of ss2.getSheets()) {
        if (!keepIds.has(s.getSheetId())) ss2.deleteSheet(s);
      }

      SpreadsheetApp.flush();

      // 2개 XLSX 병렬 fetch
      const xlsxResponses = UrlFetchApp.fetchAll([
        { url: 'https://docs.google.com/spreadsheets/d/' + xlsxCopy1.getId() + '/export?exportFormat=xlsx',
          headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true },
        { url: 'https://docs.google.com/spreadsheets/d/' + xlsxCopy2.getId() + '/export?exportFormat=xlsx',
          headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true },
      ]);

      // MSDS 엑셀
      if (xlsxResponses[0].getResponseCode() === 200) {
        results.push({
          ok: true, label: 'MSDS (엑셀)',
          fileData: Utilities.base64Encode(xlsxResponses[0].getContent()),
          fileName: buildFileName_(data, 'MSDS', 'xlsx'),
          mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
      } else {
        results.push({ ok: false, label: 'MSDS (엑셀)', error: 'HTTP ' + xlsxResponses[0].getResponseCode() });
      }

      // Checksheet 묶음
      if (xlsxResponses[1].getResponseCode() === 200) {
        results.push({
          ok: true, label: '비공개물질 Checksheet (엑셀)',
          fileData: Utilities.base64Encode(xlsxResponses[1].getContent()),
          fileName: 'LT소재_' + (data['제품명'] || '제품명') + '_비공개물질 Checksheet.xlsx',
          mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
      } else {
        results.push({ ok: false, label: '비공개물질 Checksheet (엑셀)', error: 'HTTP ' + xlsxResponses[1].getResponseCode() });
      }
    } finally {
      try { DriveApp.getFileById(xlsxCopy1.getId()).setTrashed(true); } catch (_) {}
      try { DriveApp.getFileById(xlsxCopy2.getId()).setTrashed(true); } catch (_) {}
    }

    return { ok: true, files: results };

  } catch (err) {
    return { ok: false, error: err.message, files: [] };
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

function buildFileName_(data, sheetName, ext) {
  return (data['제품명'] || '제품명') + '_' + sheetName + '.' + ext;
}
