// Đường dẫn kết nối API Google Apps Script của anh
const DEPLOY_API_URL = "https://script.google.com/macros/s/AKfycby20w2G6TDUBUJWIHiv8IhoHUga1hGHxRCBlD5RdcBgrLF9eGgSztVSgsVLGtmnQ6-pdg/exec";

let currentTabId = "nhansu";
let masterCacheData = [];
let dataHeaders = [];

const appModules = {
    nhansu: {
        title: "QUẢN LÝ NHÂN SỰ CHUYÊN NGHIỆP A-Z (BẢNG FILE EXCEL)",
        getAct: "get_nhansu", saveAct: "save_nhansu", btnLabel: "<i class='fa-solid fa-cloud-arrow-up'></i> UPLOAD EXCEL NHÂN SỰ",
        defaultFields: ["Stt", "Họ và tên", "Tên", "Số sổ BHXH", "Số CCCD", "An toàn viên", "NNĐH", "Năm về hưu", "Mã đơn vị", "Năm", "Năm sinh", "Nữ (x)", "Chức danh công việc"]
    },
    thietbi: {
        title: "QUẢN LÝ THIẾT BỊ TÀI SẢN CHUYÊN NGHIỆP",
        getAct: "get_thietbi", saveAct: "save_thietbi", btnLabel: "<i class='fa-solid fa-file-excel'></i> IMPORT EXCEL THIẾT BỊ",
        defaultFields: ["Mã TB", "Tên thiết bị", "Loại TB", "Ngày mua", "Nguyên giá", "Người sử dụng", "Đơn vị quản lý", "Trạng thái"]
    },
    nhacviec: {
        title: "HỆ THỐNG NHẮC VIỆC & THEO DÕI DEADLINE",
        getAct: "get_nhacviec", saveAct: "save_nhacviec", btnLabel: "<i class='fa-solid fa-file-excel'></i> IMPORT EXCEL NHẮC VIỆC",
        defaultFields: ["Mã công việc", "Tên công việc", "Người thực hiện", "Ngày bắt đầu", "Hạn chót", "Mức độ UT", "Tiến độ", "Ghi chú"]
    },
    muasam: {
        title: "HỒ SƠ QUẢN LÝ MUA SẮM SỬA CHỮA A-Z",
        getAct: "get_muasam", saveAct: "save_muasam", btnLabel: "<i class='fa-solid fa-file-excel'></i> IMPORT EXCEL MUA SẮM",
        defaultFields: ["Mã HS", "Loại yêu cầu", "Mã TB", "Nội dung", "Chi phí", "Người đề xuất", "Ngày đề xuất", "Trạng thái duyệt"]
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
        
        const uploadBtnText = document.getElementById("btn-upload-excel-text");
        if(uploadBtnText) uploadBtnText.innerHTML = module.btnLabel;
        
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
    if(!form) return;
    form.innerHTML = fields.map(f => `
        <div>
            <label class="block text-[10px] font-bold text-slate-400 uppercase mb-1">${f}</label>
            <input type="text" name="${f}" required class="w-full border border-slate-200 rounded-xl px-3 py-1.5 text-xs font-semibold focus:outline-none focus:border-indigo-500">
        </div>
    `).join('');
}

async function syncDataFromServer() {
    setLoadingState(true, "Đang đồng bộ dữ liệu từ Google Sheet...");
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

// XỬ LÝ THANH LỌC FILE EXCEL ĐỂ ĐỒNG BỘ TUYỆT ĐỐI VÀO SHEET
function importExcel(input) {
    if (!input.files.length) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const dataBytes = new Uint8Array(e.target.result);
            const workbook = XLSX.read(dataBytes, { type: 'array', raw: false });
            
            // Đọc dạng mảng hai chiều (Arrays) thay vì JSON trực tiếp để loại bỏ dòng trống bẫy ban đầu
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
            
            // 1. Tìm vị trí dòng tiêu đề thực sự (Dòng chứa chữ "Stt" hoặc "Họ và tên")
            let headerIndex = 0;
            for (let i = 0; i < rawRows.length; i++) {
                if (rawRows[i].some(cell => String(cell).trim().toLowerCase() === "stt" || String(cell).trim().toLowerCase() === "họ và tên")) {
                    headerIndex = i;
                    break;
                }
            }
            
            // Làm sạch tiêu đề cột: Loại bỏ khoảng xuống dòng (\n) và khoảng cách thừa
            const rawHeaders = rawRows[headerIndex];
            const cleanHeaders = rawHeaders.map(h => String(h || "").replace(/\r?\n|\r/g, " ").replace(/\s+/g, " ").trim());
            
            // 2. Gom dữ liệu từ các hàng phía sau hàng tiêu đề
            const cleanData = [];
            for (let i = headerIndex + 1; i < rawRows.length; i++) {
                const rowData = rawRows[i];
                // Kiểm tra nếu hàng rỗng hoàn toàn thì bỏ qua
                if (rowData.every(cell => String(cell).trim() === "")) continue;
                
                let rowObj = {};
                cleanHeaders.forEach((header, colIdx) => {
                    if (header) { // Chỉ lấy các cột có tiêu đề rõ ràng
                        let cellVal = rowData[colIdx];
                        rowObj[header] = cellVal !== undefined && cellVal !== null ? String(cellVal).trim() : "";
                    }
                });
                cleanData.push(rowObj);
            }
            
            if(!cleanData.length) return alert("Không tìm thấy dữ liệu nhân sự hợp lệ trong tệp.");
            
            // Áp cấu trúc cột mới vừa tối ưu sạch sẽ
            dataHeaders = cleanHeaders.filter(h => h !== "");
            pushDataToCloud(cleanData);
        } catch(err) { alert("Lỗi phân tích file Excel: " + err.message); }
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
    const counter = document.getElementById("row-counter");
    if(counter) counter.innerText = `TỔNG SỐ BẢN GHI: ${dataset.length} / ${masterCacheData.length}`;
    
    const head = document.getElementById("thead-node");
    const body = document.getElementById("tbody-node");
    
    if(!head || !body) return;
    
    if(!dataset.length) {
        head.innerHTML = "";
        body.innerHTML = `<tr><td class="py-20 text-center text-slate-400 font-medium">Không tìm thấy dữ liệu phù hợp hoặc bảng trống.</td></tr>`;
        return;
    }

    let headHtml = `<tr class="bg-slate-100 border-b border-slate-200"><th class="p-3 text-center w-12 border-r border-slate-200/50">STT</th>`;
    dataHeaders.forEach(f => headHtml += `<th class="p-3 uppercase border-r border-slate-200/50 tracking-wider text-slate-700">${f}</th>`);
    head.innerHTML = headHtml + `</tr>`;

    body.innerHTML = dataset.map((row, index) => {
        let rowHtml = `<tr class="hover:bg-slate-50 transition-colors"><td class="p-2.5 text-center font-mono text-slate-400 border-r border-slate-200/40">${index + 1}</td>`;
        dataHeaders.forEach(f => {
            let val = String(row[f] || '').trim();
            let isCode = f.toUpperCase().includes('MÃ') || f.toUpperCase().includes('CCCD') || f.toUpperCase().includes('BHXH') || f.toUpperCase().includes('SĐT') || f.toUpperCase().includes('SDT');
            rowHtml += `<td class="p-2.5 border-r border-slate-200/40 ${isCode ? 'font-mono font-bold text-slate-900' : 'text-slate-600'}">${val || '-'}</td>`;
        });
        return rowHtml + `</tr>`;
    }).join('');
}

function setLoadingState(isLoading, msg = "") {
    const dot = document.getElementById("status-dot");
    const tbody = document.getElementById("tbody-node");
    if(!dot || !tbody) return;
    
    if(isLoading) {
        dot.className = "w-2.5 h-2.5 rounded-full bg-rose-500 animate-ping";
        tbody.innerHTML = `<tr><td colspan="100" class="py-20 text-center text-indigo-600 font-bold animate-pulse">${msg}</td></tr>`;
    } else {
        dot.className = "w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse";
    }
}

let timerTimeLeft = 25 * 60;
let timerIsWorking = true;

function initPomodoroEngine() {
    setInterval(() => {
        const display = document.getElementById("time-display");
        const label = document.getElementById("timer-label");
        if(!display || !label) return;

        if (timerTimeLeft > 0) {
            timerTimeLeft--;
            let m = Math.floor(timerTimeLeft / 60).toString().padStart(2, '0');
            let s = (timerTimeLeft % 60).toString().padStart(2, '0');
            display.innerText = `${m}:${s}`;
        } else {
            if (timerIsWorking) {
                alert("🚨 HẾT GIỜ LÀM VIỆC LIÊN TỤC TẬP TRUNG!\nAnh hãy đứng dậy thư giãn 5 phút để bảo vệ sức khỏe.");
                timerTimeLeft = 5 * 60;
                timerIsWorking = false;
                label.innerText = "ĐANG NGHỈ NGƠI";
            } else {
                alert("🔔 HẾT GIỜ NGHỈ NGƠI!\nQuay trở lại làm việc thôi anh.");
                const workInput = document.getElementById("input-work-min");
                timerTimeLeft = parseInt(workInput ? workInput.value : 25) * 60;
                timerIsWorking = true;
                label.innerText = "ĐANG LÀM VIỆC";
            }
        }
    }, 1000);
}

function applyNewTimerConfig() {
    const workInput = document.getElementById("input-work-min");
    let m = parseInt(workInput ? workInput.value : 25);
    timerTimeLeft = m * 60;
    timerIsWorking = true;
    const label = document.getElementById("timer-label");
    if(label) label.innerText = "ĐANG LÀM VIỆC";
    alert(`Đã cấu hình chu kỳ làm việc mới: ${m} phút.`);
}

function renderVanNienData() {
    const today = new Date();
    const calDate = document.getElementById("cal-solar-date");
    const calMonth = document.getElementById("cal-solar-month");
    if(calDate) calDate.innerText = today.getDate();
    if(calMonth) calMonth.innerText = `Tháng ${String(today.getMonth() + 1).padStart(2, '0')}, Năm ${today.getFullYear()}`;
}
