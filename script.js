const DEPLOY_API_URL = "https://script.google.com/macros/s/AKfycbyeUvO1Dj-_CbZSVz3hy3gsedlm9tuXBtHisSEAFD5zwwlJNnmeAI_sWnuX5tGMuPgzGQ/exec";

let currentTabId = "nhansu";
let masterCacheData = [];
let dataHeaders = []; // Sẽ lưu đúng thứ tự cột gốc từ trái qua phải

const appModules = {
    nhansu: {
        title: "QUẢN LÝ NHÂN SỰ CHUYÊN NGHIỆP A-Z (TỰ ĐỘNG CẤU HÌNH ĐỘNG CỘT)",
        getAct: "get_nhansu", saveAct: "save_nhansu", btnLabel: "<i class='fa-solid fa-cloud-arrow-up'></i> UPLOAD EXCEL NHÂN SỰ ĐỘNG"
    },
    thietbi: {
        title: "QUẢN LÝ THIẾT BỊ TÀI SẢN CHUYÊN NGHIỆP",
        getAct: "get_thietbi", saveAct: "save_thietbi", btnLabel: "<i class='fa-solid fa-file-excel'></i> IMPORT EXCEL THIẾT BỊ"
    },
    nhacviec: {
        title: "HỆ THỐNG NHẮC VIỆC & THEO DÕI DEADLINE",
        getAct: "get_nhacviec", saveAct: "save_nhacviec", btnLabel: "<i class='fa-solid fa-file-excel'></i> IMPORT EXCEL NHẮC VIỆC"
    },
    muasam: {
        title: "HỒ SƠ QUẢN LÝ MUA SẮM SỬA CHỮA A-Z",
        getAct: "get_muasam", saveAct: "save_muasam", btnLabel: "<i class='fa-solid fa-file-excel'></i> IMPORT EXCEL MUA SẮM"
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
        
        document.getElementById("form-allowed").classList.add("hidden");
        document.getElementById("form-disabled").classList.remove("hidden");
        
        syncDataFromServer();
    }
}

async function syncDataFromServer() {
    setLoadingState(true, "Đang đồng bộ dữ liệu động từ Google Sheet...");
    const module = appModules[currentTabId];
    try {
        const response = await fetch(`${DEPLOY_API_URL}?action=${module.getAct}`);
        const resData = await response.json();
        
        // Nhận dữ liệu bóc tách từ Server
        if (resData && resData.data && resData.data.length > 0) {
            masterCacheData = resData.data;
            // GIỮ NGUYÊN THỨ TỰ CỘT: Ưu tiên lấy mảng headers gốc từ API trả về, không dùng Object.keys nữa
            dataHeaders = resData.headers || Object.keys(masterCacheData[0]);
        } else if (Array.isArray(resData) && resData.length > 0) {
            masterCacheData = resData;
            dataHeaders = Object.keys(masterCacheData[0]);
        } else {
            masterCacheData = [];
            dataHeaders = ["Trạng thái hệ thống"];
        }
        
        // Lọc sạch tiêu đề trống
        dataHeaders = dataHeaders.filter(h => h && h.trim() !== "" && h !== "undefined");
        
        setLoadingState(false);
        triggerSearch();
    } catch (err) {
        document.getElementById("tbody-node").innerHTML = `<tr><td colspan="100" class="py-12 text-center text-rose-500 font-bold">LỖI KẾT NỐI API. Vui lòng kiểm tra cấu hình hoặc deploy lại Apps Script.</td></tr>`;
    }
}

function importExcel(input) {
    if (!input.files.length) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const dataBytes = new Uint8Array(e.target.result);
            const workbook = XLSX.read(dataBytes, { type: 'array', raw: false });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
            
            let headerIndex = 0;
            for (let i = 0; i < rawRows.length; i++) {
                if (rawRows[i].some(cell => String(cell).trim().toLowerCase() === "stt" || String(cell).trim().toLowerCase() === "họ và tên")) {
                    headerIndex = i;
                    break;
                }
            }
            
            const rawHeaders = rawRows[headerIndex];
            // Làm sạch tên tiêu đề (xóa bỏ \n xuống hàng trong ô của file Excel)
            const cleanHeaders = rawHeaders.map(h => String(h || "").replace(/\r?\n|\r/g, " ").replace(/\s+/g, " ").trim());
            // Lọc danh sách cột thực tế, loại bỏ cột rỗng
            const finalHeaders = cleanHeaders.filter(h => h && h.trim() !== "");
            
            const cleanData = [];
            for (let i = headerIndex + 1; i < rawRows.length; i++) {
                const rowData = rawRows[i];
                if (rowData.every(cell => String(cell).trim() === "")) continue; // Bỏ qua dòng trống rỗng
                
                let rowObj = {};
                cleanHeaders.forEach((header, colIdx) => {
                    if (header && header.trim() !== "") { 
                        let cellVal = rowData[colIdx];
                        rowObj[header] = cellVal !== undefined && cellVal !== null ? String(cellVal).trim() : "";
                    }
                });
                cleanData.push(rowObj);
            }
            
            if(!cleanData.length) return alert("Không tìm thấy dữ liệu hợp lệ.");
            
            dataHeaders = finalHeaders;
            pushDataToCloud({ headers: finalHeaders, data: cleanData });
        } catch(err) { alert("Lỗi phân tích file Excel: " + err.message); }
    };
    reader.readAsArrayBuffer(input.files[0]);
    input.value = "";
}

async function pushDataToCloud(payload) {
    setLoadingState(true, "Hệ thống đang đồng bộ lên Google Sheet...");
    try {
        await fetch(DEPLOY_API_URL, {
            method: "POST",
            mode: "no-cors",
            body: JSON.stringify({ action: appModules[currentTabId].saveAct, headers: payload.headers, data: payload.data })
        });
        setTimeout(() => { syncDataFromServer(); }, 1500);
    } catch (err) {
        alert("Đường truyền dữ liệu gặp sự cố.");
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
    if(counter) counter.innerText = `TỔNG SỐ HỒ SƠ: ${dataset.length} / ${masterCacheData.length}`;
    
    const head = document.getElementById("thead-node");
    const body = document.getElementById("tbody-node");
    if(!head || !body) return;
    
    if(!dataset.length) {
        head.innerHTML = "";
        body.innerHTML = `<tr><td class="py-20 text-center text-slate-400 font-medium">Bảng trống hoặc không tìm thấy kết quả phù hợp.</td></tr>`;
        return;
    }

    // Thiết lập đầu bảng chính xác theo thứ tự tiêu đề
    let headHtml = `<tr class="bg-slate-100 border-b border-slate-200"><th class="p-3 text-center w-12 border-r border-slate-200/50">STT</th>`;
    dataHeaders.forEach(f => {
        headHtml += `<th class="p-3 uppercase border-r border-slate-200/50 tracking-wider text-slate-700 font-bold text-xs">${f}</th>`;
    });
    head.innerHTML = headHtml + `</tr>`;

    // Thiết lập nội dung từng hàng
    body.innerHTML = dataset.map((row, index) => {
        let rowHtml = `<tr class="hover:bg-slate-50 transition-colors"><td class="p-2.5 text-center font-mono text-slate-400 border-r border-slate-200/40 text-xs">${index + 1}</td>`;
        
        dataHeaders.forEach(f => {
            let val = row[f];
            
            if (typeof val === 'object' && val !== null) {
                val = val.toString ? val.toString() : JSON.stringify(val);
            } else {
                val = String(val || '').trim();
            }
            
            if (val === "[object Object]") val = "-";

            // ĐÃ SỬA: Bộ lọc xử lý định dạng Ngày tháng ISO dài dòng vỡ bảng
            if (val.includes('T') && val.includes('-') && val.length >= 10) {
                let checkDate = val.split('T')[0];
                if (!isNaN(Date.parse(checkDate))) {
                    val = checkDate; // Chỉ giữ lại yyyy-mm-dd cho gọn gàng sạch sẽ
                }
            }

            let isCode = f.toUpperCase().includes('MÃ') || f.toUpperCase().includes('CCCD') || f.toUpperCase().includes('BHXH') || f.toUpperCase().includes('SĐT') || f.toUpperCase().includes('SDT');
            rowHtml += `<td class="p-2.5 text-xs border-r border-slate-200/40 ${isCode ? 'font-mono font-bold text-slate-900' : 'text-slate-600'}">${val || '-'}</td>`;
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
function initPomodoroEngine() {
    setInterval(() => {
        const display = document.getElementById("time-display");
        if(!display) return;
        if (timerTimeLeft > 0) timerTimeLeft--;
    }, 1000);
}

function renderVanNienData() {
    const today = new Date();
    const calDate = document.getElementById("cal-solar-date");
    if(calDate) calDate.innerText = today.getDate();
}
