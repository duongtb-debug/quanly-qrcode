const DEPLOY_API_URL = "https://script.google.com/macros/s/AKfycby20w2G6TDUBUJWIHiv8IhoHUga1hGHxRCBlD5RdcBgrLF9eGgSztVSgsVLGtmnQ6-pdg/exec";

let currentTabId = "nhansu";
let masterCacheData = [];
let dataHeaders = [];

const appModules = {
    nhansu: {
        title: "QUẢN LÝ NHÂN SỰ CHUYÊN NGHIỆP A-Z (BẢNG FILE EXCEL)",
        getAct: "get_nhansu", saveAct: "save_nhansu", btnLabel: "UPLOAD EXCEL NHÂN SỰ",
        defaultFields: ["MA_NV", "HO_TEN", "GIOI_TINH", "NGAY_SINH", "CCCD", "SDT", "MA_DV", "CHUC_DANH", "NGAY_VAO_LAM", "TRANG_THAI"]
    },
    thietbi: {
        title: "QUẢN LÝ THIẾT BỊ TÀI SẢN CHUYÊN NGHIỆP",
        getAct: "get_thietbi", saveAct: "save_thietbi", btnLabel: "IMPORT EXCEL THIẾT BỊ",
        defaultFields: ["MA_TB", "TEN_THIET_BI", "LOAI_TB", "NGAY_MUA", "NGUYEN_GIA", "MA_NV_SUDUNG", "MA_DV_QUANLY", "TRANG_THAI"]
    },
    nhacviec: {
        title: "HỆ THỐNG NHẮC VIỆC & THEO DÕI DEADLINE",
        getAct: "get_nhacviec", saveAct: "save_nhacviec", btnLabel: "IMPORT EXCEL NHẮC VIỆC",
        defaultFields: ["MA_TASK", "TEN_CONG_VIEC", "NGUOI_THUC_HIEN", "NGAY_BAT_DAU", "HAN_CHOT", "MUC_DO_UT", "TIEN_DO", "GHI_CHU"]
    },
    muasam: {
        title: "HỒ SƠ QUẢN LÝ MUA SẮM SỬA CHỮA A-Z",
        getAct: "get_muasam", saveAct: "save_muasam", btnLabel: "IMPORT EXCEL MUA SẮM",
        defaultFields: ["MA_HS", "LOAI_YEU_CAU", "MA_TB", "NOI_DUNG", "CHI_PHI", "NGUOI_DE_XUAT", "NGAY_DE_XUAT", "TRANG_THAI_DUYET"]
    }
};

window.onload = () => {
    switchTab("nhansu");
    initPomodoroEngine();
};

function switchTab(tabId) {
    currentTabId = tabId;
    document.querySelectorAll(".nav-btn").forEach(btn => {
        btn.className = "nav-btn w-full flex items-center gap-3 px-4 py-3 rounded-xl text-xs font-bold transition-all text-
