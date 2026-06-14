// Đường dẫn kết nối API Google Apps Script của anh
const DEPLOY_API_URL = "https://script.google.com/macros/s/AKfycby20w2G6TDUBUJWIHiv8IhoHUga1hGHxRCBlD5RdcBgrLF9eGgSztVSgsVLGtmnQ6-pdg/exec";

let currentTabId = "nhansu";
let masterCacheData = [];
let dataHeaders = [];

const appModules = {
    nhansu: {
        title: "QUẢN LÝ NHÂN SỰ CHUYÊN NGHIỆP A-Z (BẢNG FILE EXCEL)",
        getAct: "get_nhansu", saveAct: "save_nhansu", btnLabel: "<i class='fa-solid fa-cloud-arrow-up'></i> UPLOAD EXCEL NHÂN SỰ",
        defaultFields: ["MA_NV", "HO_TEN", "GIOI_TINH", "NGAY_SINH", "CCCD", "SDT", "MA_DV", "CHUC_DANH", "NGAY_VAO_LAM", "TRANG_THAI"]
    },
    thietbi: {
        title: "QUẢN LÝ THIẾT BỊ TÀI SẢN CHUYÊN NGHIỆP",
        getAct: "get_thietbi", saveAct: "save_thietbi", btnLabel: "<i class='fa-solid fa-file-excel'></i> IMPORT EXCEL THIẾT BỊ",
        defaultFields: ["MA_TB", "TEN_THIET_BI", "LOAI_TB", "NGAY_MUA", "NGUYEN_GIA", "MA_NV_SUDUNG", "MA_DV_QUANLY", "TRANG_THAI"]
    },
    nhacviec: {
        title: "HỆ THỐNG NHẮC VIỆC & THEO DÕI DEADLINE",
        getAct: "get_nhacviec", saveAct: "save_nhacviec", btnLabel: "<i class='fa-solid fa-file-excel'></i> IMPORT EXCEL NHẮC VIỆC",
        defaultFields: ["MA_TASK", "TEN_CONG_VIEC", "NGUOI_THUC_HIEN", "NGAY_BAT_DAU", "HAN_CHOT", "MUC_DO_UT", "TIEN_DO", "GHI_CHU"]
    },
    muasam: {
        title: "HỒ SƠ QUẢN LÝ MUA SẮM SỬA CHỮA A-Z",
        getAct: "get_muasam", saveAct: "save_muasam", btnLabel: "<i class='fa-solid fa-file-excel'></i> IMPORT EXCEL MUA SẮM",
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
        btn.className = "nav-btn w-full flex items-center gap-3 px-4 py-3 rounded-xl text-xs font-bold transition-all text-slate-400 hover:bg-slate-800 hover:text-white";
    });
    const activeBtn = document.getElementById(`btn-${tabId}`);
    if(activeBtn) activeBtn.className = "nav-btn w-full flex items-center gap-3 px-4 py-3 rounded-xl text-xs font-bold transition-all bg-indigo-600 text-white shadow-md shadow-indigo-600/10";
    
    if(tabId === 'lich') {
        document.getElementById("tab-title").innerText = "LỊCH VẠN NIÊN & CHẾ ĐỘ THỜI GIAN KHỎE MẠNH";
        document.getElementById("data-view").classList.remove("active");
        document.getElementById("calendar-view").classList.add("active");
        document.getElementById("action-group").classList.add("hidden");
        renderVanNienData();
    } else {
        const module = appModules[tabId];
        document.getElementById("tab-title").innerText = module.title;
        document.getElementById("btn-upload-excel").innerHTML = module.btnLabel;
        document.getElementById("data-view").classList.add("active");
        document.getElementById("calendar-view").classList.remove("active");
        document.getElementById("action-group").classList.remove("hidden");
        
        if (tabId === "nhansu") {
            document.getElementById("form-allowed").classList.add("hidden");
            document.getElementById("form-disabled").classList.remove("hidden");
        } else {
            document.getElementById("form-allowed").classList.remove("hidden");
            document.getElementById("form-disabled").classList.add("hidden");
            buildDynamicForm(module.defaultFields);
        }
        
        syncDataFromServer();
    }
}

function buildDynamicForm(fields) {
    const form = document.getElementById("dynamic-form");
    form.innerHTML = fields.map(f => `
        <div>
            <label class="block text-[10px] font-bold text-slate-400 uppercase mb-1">${f.replace(/_/g, ' ')}</label>
            <input type="text" name="${f}" required class="w-full border border-slate-200 rounded-xl px-3 py-1.5 text-xs font-semibold focus:outline-none focus:border-indigo-500">
        </div>
    `).join('');
}

async function syncDataFromServer() {
    setLoadingState(true, "Đang đồng bộ dữ liệu...");
    const module = appModules[currentTabId];
    try {
        const response = await fetch(`${DEPLOY_API_URL}?action=${module.getAct}`);
        masterCacheData = await response.json();
        
        if (masterCacheData && masterCacheData.length > 0) {
            dataHeaders = Object.keys(masterCacheData[0]);
        } else {
            dataHeaders = module.defaultFields;
        }
        
        setLoadingState(false);
        triggerSearch();
    } catch (err) {
        document.getElementById("tbody-node").innerHTML = `<tr><td colspan="100" class="py-12 text-center text-rose-500 font-bold">LỖI KẾT NỐI API. Vui lòng kiểm tra lại cấu hình mở khóa Deploy link Script.</td></tr>`;
    }
}

function importExcel(input) {
    if (!input.files.length) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const dataBytes = new Uint8Array(e.target.result);
            const workbook = XLSX.read(dataBytes, { type: 'array', raw: false });
            const rawJson = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: "" });
            
            if(!rawJson.length) return alert("Tệp Excel trống.");
            
            const excelHeaders = Object.keys(rawJson[0]);
            const cleanData = rawJson.map(row => {
                let cleanRow = {};
                excelHeaders.forEach(h => {
                    cleanRow[h] = row[h] !== undefined && row[h] !== null ? String(row[h]).trim() : "";
                });
                return cleanRow;
            });
            
            pushDataToCloud(cleanData);
        } catch(err) { alert("Lỗi xử lý file: " + err.message); }
    };
    reader.readAsArrayBuffer(input.files[0]);
    input.value = "";
}

function handleFormSubmit(e) {
    e.preventDefault();
    const formData = new FormData(e.target);
    let newRow = {};
    dataHeaders.forEach(f => {
        newRow[f] = (formData.get(f) || "").trim();
    });
    let updatePayload = [newRow, ...masterCacheData];
    pushDataToCloud(updatePayload);
    e.target.reset();
}

async function pushDataToCloud(dataset) {
    setLoadingState(true, "Đang đẩy dữ liệu lên hệ thống...");
    try {
        await fetch(DEPLOY_API_URL, {
            method: "POST",
            mode: "no-cors",
            body: JSON.stringify({ action: appModules[currentTabId].saveAct, data: dataset })
        });
        setTimeout(() => { syncDataFromServer(); }, 1500);
    } catch (err) {
        alert("Đường truyền tải bị lỗi.");
        setLoadingState(false);
    }
}

function triggerSearch() {
    const kw = document.getElementById("search-box").value.trim().toLowerCase();
    const filtered = masterCacheData.filter(row => {
        if(!kw) return true;
        return dataHeaders.some(f => String(row[f] || "").toLowerCase().includes(kw));
    });
    renderGrid(filtered);
}

function renderGrid(dataset) {
    document.getElementById("row-counter").innerText = `TỔNG SỐ BẢN GHI: ${dataset.length} / ${masterCacheData.length}`;
    const head = document.getElementById("thead-node");
    const body = document.getElementById("tbody-node");
    
    if(!dataset.length) {
        head.innerHTML = "";
        body.innerHTML = `<tr><td class="py-20 text-center text-slate-400 font-medium">Không tìm thấy dữ liệu phù hợp.</td></tr>`;
        return;
    }

    let headHtml = `<tr class="bg-slate-100 border-b border-slate-200"><th class="p-3 text-center w-12 border-r border-slate-200/50">STT</th>`;
    dataHeaders.forEach(f => headHtml += `<th class="p-3 uppercase border-r border-slate-200/50 tracking-wider text-slate-700">${f.replace(/_/g, ' ')}</th>`);
    head.innerHTML = headHtml + `</tr>`;

    body.innerHTML = dataset.map((row, index) => {
        let rowHtml = `<tr class="hover:bg-slate-50 transition-colors"><td class="p-2.5 text-center font-mono text-slate-400 border-r border-slate-200/40">${index + 1}</td>`;
        dataHeaders.forEach(f => {
            let val = String(row[f] || '').trim();
            let isCode = f.toUpperCase().includes('MA') || f.toUpperCase().includes('CCCD') || f.toUpperCase().includes('SDT') || f.toUpperCase().includes('BHXH');
            rowHtml += `<td class="p-2.5 border-r border-slate-200/40 ${isCode ? 'font-mono font-bold text-slate-900' : 'text-slate-600'}">${val || '-'}</td>`;
        });
        return rowHtml + `</tr>`;
    }).join('');
}

function setLoadingState(isLoading, msg = "") {
    const dot = document.getElementById("status-dot");
    if(isLoading) {
        dot.className = "w-2.5 h-2.5 rounded-full bg-rose-500 animate-ping";
        document.getElementById("tbody-node").innerHTML = `<tr><td colspan="100" class="py-20 text-center text-indigo-600 font-bold animate-pulse">${msg}</td></tr>`;
    } else {
        dot.className = "w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse";
    }
}

let timerTimeLeft = 25 * 60;
let timerIsWorking = true;

function initPomodoroEngine() {
    setInterval(() => {
        if (timerTimeLeft > 0) {
            timerTimeLeft--;
            let m = Math.floor(timerTimeLeft / 60).toString().padStart(2, '0');
            let s = (timerTimeLeft % 60).toString().padStart(2, '0');
            document.getElementById("time-display").innerText = `${m}:${s}`;
        } else {
            if (timerIsWorking) {
                alert("🚨 HẾT GIỜ LÀM VIỆC LIÊN TỤC TẬP TRUNG!\nAnh hãy đứng dậy thư giãn 5 phút để bảo vệ sức khỏe.");
                timerTimeLeft = 5 * 60;
                timerIsWorking = false;
                document.getElementById("timer-label").innerText = "ĐANG NGHỈ NGƠI";
            } else {
                alert("🔔 HẾT GIỜ NGHỈ NGƠI!\nQuay trở lại làm việc thôi anh.");
                timerTimeLeft = parseInt(document.getElementById("input-work-min").value || 25) * 60;
                timerIsWorking = true;
                document.getElementById("timer-label").innerText = "ĐANG LÀM VIỆC";
            }
        }
    }, 1000);
}

function applyNewTimerConfig() {
    let m = parseInt(document.getElementById("input-work-min").value || 25);
    timerTimeLeft = m * 60;
    timerIsWorking = true;
    document.getElementById("timer-label").innerText = "ĐANG LÀM VIỆC";
    alert(`Đã cấu hình chu kỳ làm việc mới: ${m} phút.`);
}

function renderVanNienData() {
    const today = new Date();
    document.getElementById("cal-solar-date").innerText = today.getDate();
    document.getElementById("cal-solar-month").innerText = `Tháng ${String(today.getMonth() + 1).padStart(2, '0')}, Năm ${today.getFullYear()}`;
}
