/**
 * LGD 사전심사자료 자동화 - Google Apps Script
 * ================================================
 * 사용법:
 *   1. 이 코드를 구글 시트의 [확장 프로그램 > Apps Script]에 붙여넣기
 *   2. CELL_MAP 에서 각 필드의 셀 주소를 실제 시트에 맞게 수정
 *   3. [배포 > 웹 앱으로 배포] → 액세스: "모든 사용자" 로 설정
 *   4. 배포 URL을 HTML의 APPS_SCRIPT_URL 에 붙여넣기
 */

// ═══════════════════════════════════════════════
// ★ 설정 구역: 실제 시트 구조에 맞게 수정하세요
// ═══════════════════════════════════════════════

const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'; // 구글 시트 URL의 /d/XXXX/edit 부분

// 각 시트(문서)별 셀 매핑
// key: 필드명, value: 셀 주소 (A1 표기법)
const SHEET_CONFIG = {

  'MSDS': {
    sheetName: 'MSDS',       // 구글 시트의 탭 이름
    printRange: null,        // null = 기본 인쇄 영역, 또는 'A1:J50'
    cells: {
      '작성일':  'H4',        // ← 실제 셀 주소로 수정
      '제품명':  'D5',
      '색상':    'D6',
    }
  },

  '경고표지': {
    sheetName: '경고표지',
    printRange: null,
    cells: {
      '작성일':  'H3',
      '제품명':  'B4',
      '색상':    'B5',
    }
  },

  '작업공정별관리요령': {
    sheetName: '작업공정별관리요령',
    printRange: null,
    cells: {
      '작성일':  'H3',
      '제품명':  'C4',
      '색상':    'C5',
    }
  },

  '비공개물질확인서': {
    sheetName: '비공개물질확인서',
    printRange: null,
    cells: {
      '작성일':  'H3',
      '제품명':  'C4',
      '색상':    'C5',
    }
  },

  '구성제품확인서1': {
    sheetName: '구성제품확인서1',
    printRange: null,
    cells: {
      '작성일':  'H3',
      '제품명':  'C4',
      '색상':    'C5',
      '상품명1': 'C8',
    }
  },

  '구성제품확인서2': {
    sheetName: '구성제품확인서2',
    printRange: null,
    cells: {
      '작성일':  'H3',
      '제품명':  'C4',
      '색상':    'C5',
      '상품명2': 'C8',
    }
  },

  '구성제품확인서3': {
    sheetName: '구성제품확인서3',
    printRange: null,
    cells: {
      '작성일':  'H3',
      '제품명':  'C4',
      '색상':    'C5',
      '상품명3': 'C8',
    }
  },

};

// ═══════════════════════════════════════════════
// 핵심 로직 (수정 불필요)
// ═══════════════════════════════════════════════

function doPost(e) {
  try {
    const req  = JSON.parse(e.postData.contents);
    const doc  = req.doc;       // 'MSDS', '경고표지', ...
    const data = req.data;      // { 제품명, 색상, 작성일, 상품명1, ... }

    if (!SHEET_CONFIG[doc]) {
      return jsonResponse({ ok: false, error: '알 수 없는 문서: ' + doc });
    }

    const cfg    = SHEET_CONFIG[doc];
    const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet  = ss.getSheetByName(cfg.sheetName);

    if (!sheet) {
      return jsonResponse({ ok: false, error: '시트를 찾을 수 없음: ' + cfg.sheetName });
    }

    // 셀 값 채우기
    for (const [field, cellAddr] of Object.entries(cfg.cells)) {
      const val = data[field] ?? '';
      sheet.getRange(cellAddr).setValue(val);
    }

    // 변경 사항 즉시 저장
    SpreadsheetApp.flush();

    // PDF 변환 (해당 시트만)
    const pdfBytes = exportSheetAsPDF(ss, sheet, cfg.printRange);
    const b64      = Utilities.base64Encode(pdfBytes);

    return jsonResponse({ ok: true, pdf: b64, sheetName: cfg.sheetName });

  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// GET 요청: 연결 테스트용
function doGet(e) {
  return jsonResponse({ ok: true, message: 'LGD Apps Script 연결 성공!', sheets: Object.keys(SHEET_CONFIG) });
}

// 특정 시트를 PDF로 변환
function exportSheetAsPDF(ss, sheet, printRange) {
  const ssId  = ss.getId();
  const gId   = sheet.getSheetId();

  // 구글 드라이브 export URL 파라미터
  const params = [
    'exportFormat=pdf',
    'format=pdf',
    'size=A4',
    'portrait=true',
    'fitw=true',          // 너비에 맞춤
    'sheetnames=false',
    'printtitle=false',
    'pagenumbers=false',
    'gridlines=false',
    'fzr=false',          // 행 고정 반복 없음
    `gid=${gId}`,         // 특정 시트만
  ];

  if (printRange) {
    // 특정 범위만 출력 (예: 'A1:J50')
    params.push(`r1=0&c1=0`); // 필요시 범위 지정
  }

  const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?` + params.join('&');

  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('PDF 변환 실패: HTTP ' + response.getResponseCode());
  }

  return response.getContent();
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════
// 유틸: 셀 주소 자동 탐색 도우미 (선택 사항)
// 아래 함수를 Apps Script 에디터에서 직접 실행하면
// 현재 시트의 비어있지 않은 셀 목록을 로그에 출력합니다.
// ═══════════════════════════════════════════════
function listNonEmptyCells() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ss.getSheets().forEach(sheet => {
    Logger.log('=== 시트: ' + sheet.getName() + ' ===');
    const data = sheet.getDataRange().getValues();
    data.forEach((row, r) => {
      row.forEach((val, c) => {
        if (val !== '') {
          const addr = sheet.getRange(r+1, c+1).getA1Notation();
          Logger.log(`  ${addr}: ${String(val).substring(0,40)}`);
        }
      });
    });
  });
}
