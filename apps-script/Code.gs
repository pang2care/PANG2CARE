const SHEET_NAME = "reservations";

// 관리자만 아는 비밀번호로 반드시 바꿔주세요.
const ADMIN_KEY = "1354";

// 새 예약이 들어오면 카카오톡 "나에게 보내기"로 알림을 보내기 위한 설정입니다.
// developers.kakao.com 에서 발급받은 REST API 키를 아래에 붙여넣으세요.
// (주의: "플랫폼 키 > REST API 키 > 클라이언트 시크릿"은 반드시 "사용 안 함"으로 꺼두세요.
//  이 값은 진짜 비밀번호급 보안 코드라 공개 저장소인 이 코드에는 넣지 않습니다.)
const KAKAO_REST_API_KEY = "9ab7849cbdc03e9778ae1429f0ea938a";

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
    applyMonthColors_(sheet);

    return sheet;
}

// 예약 날짜(월)를 기준으로 행 전체에 은은한 배경색을 입혀서 시트에서 달별로 한눈에 구분되게 합니다.
const MONTH_COLORS = [
    "#fdeceb", // 1월
    "#fdf3e6", // 2월
    "#fdfbe6", // 3월
    "#eef8ea", // 4월
    "#e9f7f0", // 5월
    "#e7f6f7", // 6월
    "#e8f1fb", // 7월
    "#ecebfa", // 8월
    "#f5eafa", // 9월
    "#faeaf3", // 10월
    "#f2e9e4", // 11월
    "#eceeee"  // 12월
];

function monthColor_(dateText){
    const parts = String(dateText || "").split("-");
    if(parts.length !== 3) return "#ffffff";

    const month = Number(parts[1]);
    if(isNaN(month) || month < 1 || month > 12) return "#ffffff";

    return MONTH_COLORS[month - 1];
}

function applyMonthColors_(sheet){
    const lastRow = sheet.getLastRow();
    if(lastRow < 2) return;

    const dateIdx = COLUMNS.findIndex(c => c.key === "date");
    const dateValues = sheet.getRange(2, dateIdx + 1, lastRow - 1, 1).getValues();

    const backgrounds = dateValues.map(row => {
        const color = monthColor_(row[0]);
        return new Array(COLUMNS.length).fill(color);
    });

    sheet.getRange(2, 1, lastRow - 1, COLUMNS.length).setBackgrounds(backgrounds);
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
    const params0 = (e && e.parameter) || {};

    // 카카오 로그인 인증 후 돌아오는 콜백(코드에 ?code=... 가 붙어서 옴)입니다.
    // 최초 1회 카카오 연동 인증을 할 때만 발생하고, 평소 예약 조회 요청에는 없습니다.
    if(params0.code){
        return handleKakaoOAuthCallback_(params0.code);
    }

    // 관리자가 admin.html "후기 등록"에서 추가한 후기 목록입니다.
    // index.html에서도 로그인 없이 불러와 보여줘야 해서 별도 인증 없이 공개합니다.
    if(params0.resource === "reviews"){
        return handleReviewList_();
    }

    const sheet = getSheet_();

    autoUpdateStatuses_(sheet);

    const rows = sheet.getDataRange().getValues();
    const idIndex = idColumnIndex_();

    let list = rows.slice(1)
        .filter(row => row[idIndex])
        .map(row => rowToObject_(row));

    const params = (e && e.parameter) || {};

    if(params.adminKey){
        // 관리자 비밀번호가 맞을 때만 이름/연락처/주소 등 전체 정보를 반환합니다.
        if(params.adminKey !== ADMIN_KEY){
            return jsonOutput_({ ok: false, error: "관리자 인증이 필요합니다." });
        }
    }else if(params.phone){
        // 고객 본인 조회: 연락처가 일치하는 예약만 반환합니다.
        const phoneFilter = String(params.phone).replace(/\D/g,"");
        list = list.filter(r => String(r.phone || "").replace(/\D/g,"") === phoneFilter);
    }else{
        // 비로그인 상태(예약 가능 시간 확인용)에는 개인정보 없이 날짜/시간/상태만 반환합니다.
        list = list.map(r => ({ date: r.date, time: r.time, status: r.status }));
    }

    return jsonOutput_({ ok: true, reservations: list });
}

// 예약확정은 방문 시간이 되면 자동으로 작업진행중으로, 작업진행중은 방문 시간+3시간이 지나면 자동으로 작업완료로 바뀝니다.
// 트리거 등록: Apps Script 편집기 왼쪽 시계 아이콘 -> 트리거 추가 -> 함수: autoUpdateStatuses -> 시간 기반 -> 분 단위 타이머
function autoUpdateStatuses(){
    autoUpdateStatuses_(getSheet_());
}

function autoUpdateStatuses_(sheet){
    const rows = sheet.getDataRange().getValues();
    const now = new Date();

    const dateIdx = COLUMNS.findIndex(c => c.key === "date");
    const timeIdx = COLUMNS.findIndex(c => c.key === "time");
    const statusIdx = COLUMNS.findIndex(c => c.key === "status");
    const historyIdx = COLUMNS.findIndex(c => c.key === "history");

    for(let r = 1; r < rows.length; r++){
        const row = rows[r];
        const key = statusToKey_(row[statusIdx]);

        if(key !== "confirmed" && key !== "work") continue;

        const scheduled = parseScheduledDateTime_(row[dateIdx], row[timeIdx]);
        if(!scheduled) continue;

        if(key === "confirmed" && now.getTime() >= scheduled.getTime()){
            setStatusWithHistory_(sheet, r+1, statusIdx, historyIdx, row[historyIdx],
                "work", "자동 처리: 방문 시간이 되어 작업진행중으로 변경");
        }else if(key === "work"){
            const threeHoursLater = scheduled.getTime() + (3 * 60 * 60 * 1000);
            if(now.getTime() >= threeHoursLater){
                setStatusWithHistory_(sheet, r+1, statusIdx, historyIdx, row[historyIdx],
                    "completed", "자동 처리: 방문 시간으로부터 3시간이 지나 작업완료로 변경");
            }
        }
    }
}

function parseScheduledDateTime_(dateText, timeText){
    if(!dateText || !timeText) return null;

    const dateParts = String(dateText).split("-");
    const timeParts = String(timeText).split(":");

    if(dateParts.length !== 3 || timeParts.length < 2) return null;

    const y = Number(dateParts[0]);
    const m = Number(dateParts[1]) - 1;
    const d = Number(dateParts[2]);
    const hh = Number(timeParts[0]);
    const mm = Number(timeParts[1]);

    if([y,m,d,hh,mm].some(n => isNaN(n))) return null;

    return new Date(y, m, d, hh, mm, 0);
}

function setStatusWithHistory_(sheet, rowNumber, statusIdx, historyIdx, historyRaw, newStatusKey, actionLabel){
    let history = [];
    try{ history = JSON.parse(historyRaw || "[]"); }
    catch(err){ history = []; }

    history.push({ action: actionLabel, at: nowText_() });

    sheet.getRange(rowNumber, statusIdx+1).setValue(statusToLabel_(newStatusKey));
    sheet.getRange(rowNumber, historyIdx+1).setValue(JSON.stringify(history));
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

    if(body.action === "createReview"){
        return handleReviewCreate_(body.data);
    }

    if(body.action === "deleteReview"){
        return handleReviewDelete_(body.id, body.adminKey);
    }

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
    const required = ["name", "phone", "date", "time"];
    const missing = required.filter(key => !(data && String(data[key] || "").trim()));

    if(missing.length > 0){
        return jsonOutput_({ ok: false, error: "필수 항목이 누락되었습니다: " + missing.join(", ") });
    }

    const id = Utilities.getUuid();

    const row = COLUMNS.map(c => {
        if(c.key === "id") return id;
        if(c.key === "createdAt") return nowText_();
        if(c.key === "history") return JSON.stringify((data && data.history) || []);
        if(c.key === "status") return statusToLabel_((data && data.status) || "pending");
        return (data && data[c.key] !== undefined) ? data[c.key] : "";
    });

    // appendRow()로 값을 먼저 쓰면 "09:00" 같은 값이 시간으로 자동 인식되어
    // 앞자리 0이 사라지는 문제가 있어(예: "9:00"), 값을 쓰기 전에 먼저
    // 텍스트 서식을 지정한 뒤 setValues()로 씁니다.
    const newRow = sheet.getLastRow() + 1;
    const range = sheet.getRange(newRow, 1, 1, COLUMNS.length);
    range.setNumberFormat("@");
    range.setValues([row]);

    // 카카오톡 알림이 실패하더라도(연동 전, 토큰 만료 등) 예약 저장 자체는 항상 성공해야 합니다.
    try{
        sendKakaoNotification_(data);
    }catch(err){
        // 무시: 알림은 부가 기능이라 예약 처리를 막지 않습니다.
    }

    return jsonOutput_({ ok: true, id: id });
}

// ===== 카카오톡 "나에게 보내기" 알림 =====
// 최초 1회, 아래 authorize URL을 브라우저에서 열어 카카오 로그인으로 인증하면
// access_token/refresh_token이 스크립트 속성에 저장되고 그 다음부터는 자동으로 동작합니다.
// authorize URL: https://kauth.kakao.com/oauth/authorize?client_id=REST_API_키&redirect_uri=이_웹앱_URL&response_type=code&scope=talk_message

function kakaoRedirectUri_(){
    return ScriptApp.getService().getUrl();
}

function handleKakaoOAuthCallback_(code){
    const resp = UrlFetchApp.fetch("https://kauth.kakao.com/oauth/token", {
        method: "post",
        payload: {
            grant_type: "authorization_code",
            client_id: KAKAO_REST_API_KEY,
            redirect_uri: kakaoRedirectUri_(),
            code: code
        },
        muteHttpExceptions: true
    });

    const result = JSON.parse(resp.getContentText());

    if(result.access_token){
        const props = PropertiesService.getScriptProperties();
        props.setProperty("KAKAO_ACCESS_TOKEN", result.access_token);
        props.setProperty("KAKAO_REFRESH_TOKEN", result.refresh_token);
        return HtmlService.createHtmlOutput("카카오톡 알림 연동이 완료되었습니다. 이 창은 닫으셔도 됩니다.");
    }

    return HtmlService.createHtmlOutput("연동 실패: " + resp.getContentText());
}

function refreshKakaoAccessToken_(){
    const props = PropertiesService.getScriptProperties();
    const refreshToken = props.getProperty("KAKAO_REFRESH_TOKEN");
    if(!refreshToken) return null;

    const resp = UrlFetchApp.fetch("https://kauth.kakao.com/oauth/token", {
        method: "post",
        payload: {
            grant_type: "refresh_token",
            client_id: KAKAO_REST_API_KEY,
            refresh_token: refreshToken
        },
        muteHttpExceptions: true
    });

    const result = JSON.parse(resp.getContentText());
    if(!result.access_token) return null;

    props.setProperty("KAKAO_ACCESS_TOKEN", result.access_token);
    if(result.refresh_token){
        props.setProperty("KAKAO_REFRESH_TOKEN", result.refresh_token);
    }

    return result.access_token;
}

function sendKakaoNotification_(data){
    const props = PropertiesService.getScriptProperties();
    const accessToken = props.getProperty("KAKAO_ACCESS_TOKEN");
    if(!accessToken){
        return "건너뜀: KAKAO_ACCESS_TOKEN이 저장되어 있지 않음";
    }

    const text = "[팡이케어 새 예약]\n" +
        "이름: " + (data.name || "") + "\n" +
        "연락처: " + (data.phone || "") + "\n" +
        "일시: " + (data.date || "") + " " + (data.time || "") + "\n" +
        "주소: " + (data.address || "");

    const templateObject = {
        object_type: "text",
        text: text,
        link: {
            web_url: "https://pang2care.github.io/PANG2CARE/admin.html",
            mobile_web_url: "https://pang2care.github.io/PANG2CARE/admin.html"
        },
        button_title: "관리자 페이지 열기"
    };

    const callKakao_ = function(token){
        return UrlFetchApp.fetch("https://kapi.kakao.com/v2/api/talk/memo/default/send", {
            method: "post",
            headers: { Authorization: "Bearer " + token },
            payload: { template_object: JSON.stringify(templateObject) },
            muteHttpExceptions: true
        });
    };

    let resp = callKakao_(accessToken);
    let info = "1차 응답 " + resp.getResponseCode() + ": " + resp.getContentText();

    if(resp.getResponseCode() === 401){
        // 액세스 토큰이 만료된 경우 갱신 후 한 번 더 시도합니다.
        const refreshed = refreshKakaoAccessToken_();
        info += " | 토큰갱신: " + (refreshed ? "성공" : "실패");
        if(refreshed){
            const resp2 = callKakao_(refreshed);
            info += " | 2차 응답 " + resp2.getResponseCode() + ": " + resp2.getContentText();
        }
    }

    return info;
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
            // "09:00" 같은 값이 시간으로 자동 인식되지 않도록 값을 쓰기 전에
            // 텍스트 서식을 먼저 지정합니다.
            const cell = sheet.getRange(rowIndex, i+1);
            cell.setNumberFormat("@");
            cell.setValue(value);
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

// ===== 관리자가 등록하는 고객 후기 (admin.html "후기 등록") =====
// 기존 "reservations" 시트와 별개로 "reviews" 시트를 자동 생성해 쌓습니다.
// 등록해도 기존 후기는 절대 지워지지 않고, 삭제 버튼을 눌러야만 지워집니다.

const REVIEW_SHEET_NAME = "reviews";

const REVIEW_COLUMNS = [
    { key: "id",        label: "id" },
    { key: "createdAt", label: "등록일시" },
    { key: "author",    label: "작성자" },
    { key: "rating",    label: "별점" },
    { key: "content",   label: "내용" }
];

function getReviewSheet_(){
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(REVIEW_SHEET_NAME);

    if(!sheet){
        sheet = ss.insertSheet(REVIEW_SHEET_NAME);
        sheet.appendRow(REVIEW_COLUMNS.map(c => c.label));
        sheet.setFrozenRows(1);
    }

    return sheet;
}

function reviewIdColumnIndex_(){
    return REVIEW_COLUMNS.findIndex(c => c.key === "id");
}

function handleReviewList_(){
    const sheet = getReviewSheet_();
    const rows = sheet.getDataRange().getValues();
    const idIdx = reviewIdColumnIndex_();

    const list = rows.slice(1)
        .filter(row => row[idIdx])
        .map(row => {
            const obj = {};
            REVIEW_COLUMNS.forEach((c,i) => { obj[c.key] = row[i]; });
            obj.rating = Number(obj.rating) || 5;
            return obj;
        });

    return jsonOutput_({ ok: true, reviews: list });
}

function handleReviewCreate_(data){
    const required = ["author", "content"];
    const missing = required.filter(key => !(data && String(data[key] || "").trim()));

    if(missing.length > 0){
        return jsonOutput_({ ok: false, error: "필수 항목이 누락되었습니다: " + missing.join(", ") });
    }

    const sheet = getReviewSheet_();
    const id = Utilities.getUuid();

    const row = REVIEW_COLUMNS.map(c => {
        if(c.key === "id") return id;
        if(c.key === "createdAt") return nowText_();
        if(c.key === "rating") return Number(data.rating) || 5;
        return (data && data[c.key] !== undefined) ? data[c.key] : "";
    });

    sheet.appendRow(row);

    return jsonOutput_({ ok: true, id: id });
}

function handleReviewDelete_(id, adminKey){
    if(adminKey !== ADMIN_KEY){
        return jsonOutput_({ ok: false, error: "관리자 인증이 필요합니다." });
    }

    const sheet = getReviewSheet_();
    const lastRow = sheet.getLastRow();
    if(lastRow < 2) return jsonOutput_({ ok: false, error: "후기를 찾을 수 없습니다." });

    const idIdx = reviewIdColumnIndex_();
    const ids = sheet.getRange(2, idIdx+1, lastRow-1, 1).getValues();

    let rowIndex = -1;
    for(let i=0;i<ids.length;i++){
        if(ids[i][0] === id){ rowIndex = i+2; break; }
    }

    if(rowIndex === -1) return jsonOutput_({ ok: false, error: "후기를 찾을 수 없습니다." });

    sheet.deleteRow(rowIndex);
    return jsonOutput_({ ok: true });
}

function jsonOutput_(obj){
    return ContentService.createTextOutput(JSON.stringify(obj))
        .setMimeType(ContentService.MimeType.JSON);
}
