/**
 * LGD 사전심사자료 자동화 - Google Apps Script (v6)
 * =================================================
 * 설정 시트 기반 동적 구성 - 시트 이름/PDF 옵션 변경 시 코드 수정 불필요
 *
 * 사용법:
 *   1. 구글 시트 > 확장 프로그램 > Apps Script
 *   2. Code.gs에 이 코드 붙여넣기
 *   3. 파일 추가(+) > HTML > 이름: "Index" > Index.html 내용 붙여넣기
 *   4. [배포 > 새 배포 > 웹 앱] → 실행 주체: "나", 액세스: "모든 사용자"
 *   5. 배포 URL을 브라우저에서 열면 입력 폼이 나타남
 *
 * 템플릿 스프레드시트에 "설정" 시트가 있어야 합니다.
 * 시트 이름이나 PDF 옵션을 변경할 때는 설정 시트만 수정하세요.
 */

// ═══════════════════════════════════════════════
// ★ 설정 (코드에서 관리하는 것은 이 2개뿐)
// ═══════════════════════════════════════════════

const TEMPLATE_SPREADSHEET_ID = '1kh2oBZYKXaadIJoZQJ5OPYZHlwZftiFpuIT45v2SjTk';

const PLACEHOLDERS = ['작성일', '제품명', '색상', '상품명1', '상품명2', '상품명3'];

// ═══════════════════════════════════════════════
// 웹 앱 진입점
// ═══════════════════════════════════════════════

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('LGD 사전심사자료 자동화')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ═══════════════════════════════════════════════
// 클라이언트 호출 함수
// ═══════════════════════════════════════════════

function testConnection() {
  return { ok: true, message: 'Apps Script v6 연결 성공!' };
}

/**
 * 전체 파일 일괄 생성 (서버 1회 호출)
 * 설정 시트에서 구성을 읽어 PDF 병렬 생성 + XLSX 병렬 생성
 */
function generateAllFiles(data) {
  let tempFile = null;

  try {
    const count = data['상품명3'] ? 3 : data['상품명2'] ? 2 : 1;
    tempFile = copyTemplate_(data);
    const tempSS = SpreadsheetApp.open(tempFile);

    // ★ 설정 시트에서 구성 읽기 (1회)
    const config = readConfig_(tempSS);
    const results = [];
    const token = ScriptApp.getOAuthToken();

    // ── PDF 병렬 생성 (UrlFetchApp.fetchAll) ──
    const pdfRequests = [];
    const pdfMeta = [];

    for (const entry of config.pdfSheets) {
      // 구성제품그룹 필터: group이 있으면 count와 일치할 때만 출력
      if (entry.group > 0 && entry.group !== count) continue;

      const sheet = tempSS.getSheetByName(entry.name);
      if (!sheet) {
        results.push({ ok: false, label: entry.name + ' (PDF)', error: '시트 없음: ' + entry.name });
        continue;
      }

      pdfRequests.push({
        url: buildPdfUrl_(tempSS.getId(), sheet.getSheetId(), entry),
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true,
      });
      pdfMeta.push({ name: entry.name });
    }

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

    // ── XLSX 생성 ──
    const xlsxCopies = [];
    const xlsxMeta = []; // { type: 'single'|'bundle', name, names }

    // XLSX 단일시트들
    for (const entry of config.xlsxSingle) {
      const copy = DriveApp.getFileById(tempSS.getId()).makeCopy('_xlsx_' + Date.now() + '_' + entry.name);
      const ss = SpreadsheetApp.open(copy);
      const target = ss.getSheetByName(entry.name);
      if (target) {
        for (const s of ss.getSheets()) {
          if (s.getSheetId() !== target.getSheetId()) ss.deleteSheet(s);
        }
      }
      xlsxCopies.push(copy);
      xlsxMeta.push({ type: 'single', name: entry.name });
    }

    // XLSX 묶음
    if (config.xlsxBundle.length > 0) {
      const copy = DriveApp.getFileById(tempSS.getId()).makeCopy('_bundle_' + Date.now());
      const ss = SpreadsheetApp.open(copy);
      const keepIds = new Set();
      for (const entry of config.xlsxBundle) {
        const s = ss.getSheetByName(entry.name);
        if (s) keepIds.add(s.getSheetId());
      }
      for (const s of ss.getSheets()) {
        if (!keepIds.has(s.getSheetId())) ss.deleteSheet(s);
      }
      xlsxCopies.push(copy);
      xlsxMeta.push({ type: 'bundle', name: '비공개물질 Checksheet', names: config.xlsxBundle.map(e => e.name) });
    }

    SpreadsheetApp.flush();

    // XLSX 병렬 fetch
    if (xlsxCopies.length > 0) {
      const xlsxRequests = xlsxCopies.map(copy => ({
        url: 'https://docs.google.com/spreadsheets/d/' + copy.getId() + '/export?exportFormat=xlsx',
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true,
      }));

      const xlsxResponses = UrlFetchApp.fetchAll(xlsxRequests);

      for (let i = 0; i < xlsxResponses.length; i++) {
        const resp = xlsxResponses[i];
        const meta = xlsxMeta[i];

        if (resp.getResponseCode() === 200) {
          const fileName = meta.type === 'bundle'
            ? 'LT소재_' + (data['제품명'] || '제품명') + '_비공개물질 Checksheet.xlsx'
            : buildFileName_(data, meta.name, 'xlsx');
          results.push({
            ok: true,
            label: meta.name + ' (엑셀)',
            fileData: Utilities.base64Encode(resp.getContent()),
            fileName: fileName,
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          });
        } else {
          results.push({ ok: false, label: meta.name + ' (엑셀)', error: 'HTTP ' + resp.getResponseCode() });
        }
      }

      // 임시 복사본 삭제
      for (const copy of xlsxCopies) {
        try { DriveApp.getFileById(copy.getId()).setTrashed(true); } catch (_) {}
      }
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
// 설정 시트 읽기
// ═══════════════════════════════════════════════

/**
 * 설정 시트에서 내보내기 구성을 읽어온다.
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @returns {{ pdfSheets: Array, xlsxSingle: Array, xlsxBundle: Array }}
 */
function readConfig_(ss) {
  const configSheet = ss.getSheetByName('설정');
  if (!configSheet) {
    throw new Error('설정 시트를 찾을 수 없습니다. 템플릿에 "설정" 시트를 추가하세요.');
  }

  const SCALE_MAP = { '기본': 1, '너비맞춤': 2, '높이맞춤': 3, '페이지맞춤': 4 };
  const HALIGN_MAP = { '가운데': 'CENTER', '왼쪽': 'LEFT', '오른쪽': 'RIGHT' };
  const VALIGN_MAP = { '위': 'TOP', '중간': 'MIDDLE', '아래': 'BOTTOM' };

  const data = configSheet.getDataRange().getValues();
  const config = { pdfSheets: [], xlsxSingle: [], xlsxBundle: [] };

  // 행 1 = 헤더, 행 2부터 데이터
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = String(row[0] || '').trim();
    if (!name) continue;

    const types = String(row[1] || '').trim();
    if (!types) continue;

    const isPdf = types.indexOf('PDF') !== -1;
    const isXlsxSingle = types.indexOf('XLSX단일') !== -1;
    const isXlsxBundle = types.indexOf('XLSX묶음') !== -1;

    if (isPdf) {
      const direction = String(row[2] || '').trim();
      const scaleText = String(row[3] || '').trim();
      config.pdfSheets.push({
        name: name,
        group: (row[10] !== '' && row[10] !== null && row[10] !== undefined) ? Number(row[10]) : 0,
        portrait: direction !== '가로',
        scale: SCALE_MAP[scaleText] || 4,
        top: Number(row[4]) || 0,
        bottom: Number(row[5]) || 0,
        left: Number(row[6]) || 0,
        right: Number(row[7]) || 0,
        hAlign: HALIGN_MAP[String(row[8] || '').trim()] || 'CENTER',
        vAlign: VALIGN_MAP[String(row[9] || '').trim()] || 'TOP',
      });
    }

    if (isXlsxSingle) {
      config.xlsxSingle.push({ name: name });
    }

    if (isXlsxBundle) {
      config.xlsxBundle.push({ name: name });
    }
  }

  return config;
}

// ═══════════════════════════════════════════════
// 내부 함수
// ═══════════════════════════════════════════════

function copyTemplate_(data) {
  const templateFile = DriveApp.getFileById(TEMPLATE_SPREADSHEET_ID);
  const tempFile = templateFile.makeCopy('_temp_' + Date.now());
  const tempSS = SpreadsheetApp.open(tempFile);

  for (const sheet of tempSS.getSheets()) {
    if (sheet.getName() === '설정') continue; // 설정 시트는 건너뛰기
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

function buildPdfUrl_(ssId, sheetGid, cfg) {
  return 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?' +
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
    '&horizontal_alignment=' + cfg.hAlign +
    '&vertical_alignment=' + cfg.vAlign +
    '&gid=' + sheetGid;
}

function buildFileName_(data, sheetName, ext) {
  return (data['제품명'] || '제품명') + '_' + sheetName + '.' + ext;
}
