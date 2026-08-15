/*
  팡이케어 예약 데이터를 저장하는 구글 시트용 Apps Script입니다.
  구글 스프레드시트 > 확장 프로그램 > Apps Script 에 이 파일 내용을 그대로 붙여넣고,
  웹 앱으로 배포한 뒤 발급되는 URL을 sheet-config.js 의 SHEET_API_URL 에 입력하세요.

  시트에는 아래 순서 그대로 한 줄씩 사람이 읽기 쉽게 기록됩니다.
  접수시간 · 이름 · 연락처 · 에어컨 종류 · 견적 · 주소지 · 날짜 · 방문시간 · 수량 · 문의사항
  (상태/관리자메모/처리이력/id 는 예약 승인·취소·삭제 기능을 위해 뒤쪽에 함께 기록됩니다)

  "상태" 칸은 시트에서 셀을 클릭하면 드롭다운으로 바로 바꿀 수 있습니다.
  예약대기 / 예약확정 / 작업진행중 / 작업완료 / 취소요청(고객) / 예약취소
  여기서 직접 바꾼 상태는 홈페이지·관리자 페이지에도 그대로 반영됩니다.
*/

const SHEET_NAME = "reservations";

// 관리자만 아는 비밀번호로 반드시 바꿔주세요.
const ADMIN_KEY = "여기에_관리자_비밀번호_설정";

const COLUMNS = [
    { key: "createdAt", label: "접수시간" },
    { key: "name",       label: "이름" },
    { key: "phone",      label: "연락처" },
    { key: "type",       label: "에어컨 종류" },
    { key: "price",      label: "견적" },
    { key: "address",    label: "주소지" },
    { key: "date",       label: "날짜" },
    { key: "time",       label: "방문시간" },
    { key: "count",      label: "수량" },
    { key: "memo",       label: "문의사항" },
    { key: "status",     label: "상태" },
    { key: "adminMemo",  label: "관리자 메모" },
    { key: "history",    label: "처리 이력" },
    { key: "id",         label: "id" }
];

// 시트의 "상태" 칸에는 아래 한글 값이 그대로 표시/입력됩니다. 시트에서 직접 클릭해 바꿀 수 있습니다.
const STATUS_LABELS = {
    pending: "예약대기",
    confirmed: "예약확정",
    work: "작업진행중",
    completed: "작업완료",
    cancel_requested: "취소요청(고객)",
    cancelled: "예약취소"
};

const STATUS_KEYS_BY_LABEL = Object.keys(STATUS_LABELS).reduce((acc, key) => {
    acc[STATUS_LABELS[key]] = key;
    return acc;
}, {});

function statusToLabel_(key){
    return STATUS_LABELS[key] || key;
}

function statusToKey_(label){
    return STATUS_KEYS_BY_LABEL[label] || label;
}

function getSheet_(){
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if(!sheet){
        sheet = ss.insertSheet(SHEET_NAME);
        sheet.appendRow(COLUMNS.map(c => c.label));
        sheet.setFrozenRows(1);
    }

    // 날짜/시간 같은 값을 시트가 자기 마음대로 날짜 형식으로 바꿔버리는 것을 막습니다.
    applyPlainTextFormat_(sheet);
    applyStatusDropdown_(sheet);

    return sheet;
}

// 모든 데이터 칸을 "일반 텍스트" 서식으로 고정합니다. (날짜/시간 값이 자동 변환되는 것 방지)
function applyPlainTextFormat_(sheet){
    sheet.getRange(2, 1, 998, COLUMNS.length).setNumberFormat("@");
}

// "상태" 열에 드롭다운을 걸어서, 시트에서 직접 클릭 몇 번으로 취소/작업중/완료 처리를 할 수 있게 합니다.
function applyStatusDropdown_(sheet){
    const statusCol = COLUMNS.findIndex(c => c.key === "status") + 1;

    const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(Object.values(STATUS_LABELS), true)
        .setAllowInvalid(false)
        .build();

    sheet.getRange(2, statusCol, 998, 1).setDataValidation(rule);
}

function nowText_(){
    return Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
}

function idColumnIndex_(){
    return COLUMNS.findIndex(c => c.key === "id");
}

function doGet(e){
    const sheet = getSheet_();
    const rows = sheet.getDataRange().getValues();
    const idIndex = idColumnIndex_();

    const list = rows.slice(1)
        .filter(row => row[idIndex])
        .map(row => rowToObject_(row));

    return jsonOutput_({ ok: true, reservations: list });
}

function rowToObject_(row){
    const obj = {};

    COLUMNS.forEach((c,i) => { obj[c.key] = row[i]; });

    obj.count = Number(obj.count) || 1;
    obj.status = statusToKey_(obj.status);

    try{ obj.history = JSON.parse(obj.history || "[]"); }
    catch(err){ obj.history = []; }

    return obj;
}

function doPost(e){
    const body = JSON.parse(e.postData.contents);
    const sheet = getSheet_();

    if(body.action === "create"){
        return handleCreate_(sheet, body.data);
    }

    if(body.action === "update"){
        return handleUpdate_(sheet, body.id, body.patch, body.adminKey);
    }

    if(body.action === "delete"){
        return handleDelete_(sheet, body.id, body.adminKey);
    }

    return jsonOutput_({ ok: false, error: "알 수 없는 요청입니다." });
}

function handleCreate_(sheet, data){
    const id = Utilities.getUuid();

    const row = COLUMNS.map(c => {
        if(c.key === "id") return id;
        if(c.key === "createdAt") return nowText_();
        if(c.key === "history") return JSON.stringify((data && data.history) || []);
        if(c.key === "status") return statusToLabel_((data && data.status) || "pending");
        return (data && data[c.key] !== undefined) ? data[c.key] : "";
    });

    sheet.appendRow(row);
    return jsonOutput_({ ok: true, id: id });
}

function handleUpdate_(sheet, id, patch, adminKey){
    // 고객이 스스로 "취소 요청"만 보내는 경우는 비밀번호 없이 허용합니다.
    const keys = Object.keys(patch || {});
    const isSelfCancelRequest =
        patch && patch.status === "cancel_requested" &&
        keys.every(k => k === "status" || k === "history");

    if(!isSelfCancelRequest && adminKey !== ADMIN_KEY){
        return jsonOutput_({ ok: false, error: "관리자 인증이 필요합니다." });
    }

    const rowIndex = findRowById_(sheet, id);
    if(rowIndex === -1) return jsonOutput_({ ok: false, error: "예약을 찾을 수 없습니다." });

    COLUMNS.forEach((c,i) => {
        if(patch[c.key] !== undefined){
            let value = patch[c.key];
            if(c.key === "history") value = JSON.stringify(value);
            if(c.key === "status") value = statusToLabel_(value);
            sheet.getRange(rowIndex, i+1).setValue(value);
        }
    });

    return jsonOutput_({ ok: true });
}

function handleDelete_(sheet, id, adminKey){
    if(adminKey !== ADMIN_KEY){
        return jsonOutput_({ ok: false, error: "관리자 인증이 필요합니다." });
    }

    const rowIndex = findRowById_(sheet, id);
    if(rowIndex === -1) return jsonOutput_({ ok: false, error: "예약을 찾을 수 없습니다." });

    sheet.deleteRow(rowIndex);
    return jsonOutput_({ ok: true });
}

function findRowById_(sheet, id){
    const lastRow = sheet.getLastRow();
    if(lastRow < 2) return -1;

    const idIndex = idColumnIndex_();
    const ids = sheet.getRange(2, idIndex+1, lastRow-1, 1).getValues();

    for(let i=0;i<ids.length;i++){
        if(ids[i][0] === id) return i+2;
    }

    return -1;
}

function jsonOutput_(obj){
    return ContentService.createTextOutput(JSON.stringify(obj))
        .setMimeType(ContentService.MimeType.JSON);
}
