/*
  index.html(고객 홈페이지)과 admin.html(관리자 달력)이 함께 사용하는
  구글 시트(Apps Script 웹앱) 연동 함수 모음입니다.
  sheet-config.js 에서 정의한 SHEET_API_URL 을 사용합니다.
*/

const TIME_SLOTS = ["09:00","10:00","11:00","12:00","13:00","14:00","15:00","16:00"];

function escapeHtml(value){
    return String(value == null ? "" : value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

const STATUS_LABEL = {
    pending: "예약대기",
    confirmed: "예약확정",
    work: "작업중",
    completed: "작업완료",
    cancel_requested: "취소요청",
    cancelled: "취소완료"
};

const STATUS_COLOR = {
    pending: "#f2ad35",
    confirmed: "#087f8c",
    work: "#1b9c68",
    completed: "#1677d2",
    cancel_requested: "#e07b39",
    cancelled: "#999999"
};

async function apiList(){
    const res = await fetch(SHEET_API_URL, { method: "GET" });
    const json = await res.json();
    if(!json.ok) throw new Error(json.error || "예약 목록을 불러오지 못했습니다.");
    return json.reservations;
}

async function apiCreate(data){
    const res = await fetch(SHEET_API_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ action: "create", data })
    });
    const json = await res.json();
    if(!json.ok) throw new Error(json.error || "예약 등록에 실패했습니다.");
    return json.id;
}

async function apiUpdate(id, patch, adminKey){
    const res = await fetch(SHEET_API_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ action: "update", id, patch, adminKey: adminKey || "" })
    });
    const json = await res.json();
    if(!json.ok) throw new Error(json.error || "예약 수정에 실패했습니다.");
    return json;
}

async function apiDelete(id, adminKey){
    const res = await fetch(SHEET_API_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ action: "delete", id, adminKey: adminKey || "" })
    });
    const json = await res.json();
    if(!json.ok) throw new Error(json.error || "예약 삭제에 실패했습니다.");
    return json;
}
