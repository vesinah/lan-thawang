/**
 * ===================================================
 * ลานท่าวัง - Main Application JavaScript
 * ===================================================
 * ระบบลงทะเบียนคัมภีร์ใบลาน
 * เชื่อมต่อ Google Sheets ผ่าน Google Apps Script API
 * =================================================== */

// ===== CONFIG =====
// ⚠️ เปลี่ยน URL นี้เป็น URL ของ Google Apps Script Web App ที่ deploy แล้ว
const API_URL = 'https://script.google.com/macros/s/AKfycbz-ovCtyNHHztPpGJltsBKQiLI4Pkjau4ZuJ5JYpBpkHqRiCNm4PpBDYe7DMfXzTrw/exec';

// Polling interval (ms) สำหรับ real-time sync
const SYNC_INTERVAL = 30000; // 30 วินาที

// Records per page
const RECORDS_PER_PAGE = 20;

// ===== DROPDOWN OPTIONS (fallback ถ้าดึงจาก API ไม่ได้) =====
const DEFAULT_DROPDOWNS = {
    หมวดคัมภีร์: [
        'พระวินัยปิฎก', 'พระสุตตันตปิฎก', 'พระอภิธรรมปิฎก',
        'ปกรณ์วิเสส', 'คัมภีร์ประเพณี', 'วรรณกรรมท้องถิ่น',
        'คัมภีร์บรรณรักษ์', 'ยังไม่ระบุหมวด', 'อื่นๆ'
    ],
    ยุคสมัย: ['อยุธยา', 'รัตนโกสินทร์', 'ยังไม่ระบุสมัย'],
    อักษร: ['ขอม', 'ไทย', 'ขอม-ไทย', 'อื่นๆ'],
    ภาษา: ['บาลี', 'ไทย', 'บาลี-ไทย', 'อื่นๆ'],
    เส้น: ['จาร', 'เขียน', 'จาร-เขียน', 'อื่นๆ'],
    ฉบับ: ['ลานยาว', 'ลานสั้น', 'อื่นๆ'],
    ชนิดไม้ประกับ: [
        'ไม้ธรรมดา', 'ไม้แกะสลัก', 'ไม้ลงรักปิดทอง', 'ไม้ลงรัก',
        'ไม้ประดับกระจก', 'ไม่มีไม้ประกับ', 'อื่นๆ'
    ],
    สภาพคัมภีร์: ['ดี', 'พอใช้', 'ชำรุดเล็กน้อย', 'ชำรุดมาก', 'เสียหายหนัก'],
    Digitize: ['แล้ว', 'ยังไม่ได้', 'อยู่ระหว่างดำเนินการ']
};

// ===== APP STATE =====
const state = {
    sheet1Data: [],
    sheet2Data: [],
    dashboardData: null,
    dropdowns: DEFAULT_DROPDOWNS,
    currentPage: 'dashboard',
    formStep: 1,
    totalSteps: 5,
    // Records table
    allRecords: [],
    filteredRecords: [],
    currentTablePage: 1,
    sortColumn: null,
    sortDirection: 'asc',
    // Charts
    charts: {},
    // Sync
    syncTimer: null,
    isOnline: true
};

// ===== CHART COLORS =====
const CHART_COLORS = [
    '#d4a017', '#8B4513', '#2563eb', '#059669', '#dc2626',
    '#7c3aed', '#db2777', '#ea580c', '#0891b2', '#4f46e5',
    '#84cc16', '#f59e0b', '#6366f1', '#14b8a6'
];

// ===== INITIALIZATION =====
document.addEventListener('DOMContentLoaded', () => {
    initializeApp();
});

async function initializeApp() {
    setupEventListeners();
    populateDropdowns(DEFAULT_DROPDOWNS);
    
    // ลองดึงข้อมูลจาก API
    if (API_URL && API_URL !== 'YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL') {
        await loadAllData();
        startAutoSync();
    } else {
        // ถ้ายังไม่ได้ตั้ง API ให้ใช้ข้อมูลตัวอย่าง
        loadDemoData();
        showToast('info', 'โหมดสาธิต - กรุณาตั้งค่า API URL ในไฟล์ app.js');
    }
    
    hideLoading();
}

// ===== API FUNCTIONS =====

async function apiGet(action) {
    try {
        const response = await fetch(`${API_URL}?action=${action}`);
        const data = await response.json();
        if (data.success) {
            updateSyncStatus(true);
            return data.data;
        } else {
            throw new Error(data.error || 'API Error');
        }
    } catch (err) {
        console.error(`API GET Error (${action}):`, err);
        updateSyncStatus(false);
        throw err;
    }
}

async function apiPost(action, payload) {
    try {
        const response = await fetch(API_URL, {
            method: 'POST',
            redirect: 'follow',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ action, ...payload })
        });
        const data = await response.json();
        if (data.success) {
            updateSyncStatus(true);
            return data.data;
        } else {
            throw new Error(data.error || 'API Error');
        }
    } catch (err) {
        console.error(`API POST Error (${action}):`, err);
        updateSyncStatus(false);
        throw err;
    }
}

async function loadAllData() {
    // 1. แสดงข้อมูลจาก localStorage ทันที (ไม่ต้องรอ API)
    const cachedDashboard = localStorage.getItem('lt_dashboard');
    const cachedSheet1 = localStorage.getItem('lt_sheet1');
    const cachedSheet2 = localStorage.getItem('lt_sheet2');
    
    if (cachedDashboard) {
        try {
            state.dashboardData = JSON.parse(cachedDashboard);
            renderDashboard();
            updateLastSyncTime();
            console.log('Dashboard loaded from cache (instant)');
        } catch(e) { /* ignore parse errors */ }
    }
    if (cachedSheet1 && cachedSheet2) {
        try {
            state.sheet1Data = JSON.parse(cachedSheet1);
            state.sheet2Data = JSON.parse(cachedSheet2);
            mergeAndDisplayRecords();
            console.log('Records loaded from cache (instant)');
        } catch(e) { /* ignore */ }
    }
    
    // 2. ซ่อน loading overlay ทันที (ถ้ามี cache)
    if (cachedDashboard) {
        hideLoading();
    }
    
    // 3. Sync จาก API ใน background
    refreshFromAPI();
}

async function refreshFromAPI() {
    try {
        // โหลด dashboard
        const dashboard = await apiGet('getDashboard');
        state.dashboardData = dashboard;
        localStorage.setItem('lt_dashboard', JSON.stringify(dashboard));
        renderDashboard();
        updateLastSyncTime();
        hideLoading();
    } catch (err) {
        console.error('Error loading dashboard:', err);
        if (!state.dashboardData) {
            showToast('error', 'ไม่สามารถโหลดข้อมูลได้: ' + err.message);
        }
        hideLoading();
    }
    
    // โหลด full data
    loadFullDataBackground();
}

async function loadFullDataBackground() {
    try {
        const result = await apiGet('getAllData');
        state.sheet1Data = result.sheet1 || [];
        state.sheet2Data = result.sheet2 || [];
        localStorage.setItem('lt_sheet1', JSON.stringify(state.sheet1Data));
        localStorage.setItem('lt_sheet2', JSON.stringify(state.sheet2Data));
        mergeAndDisplayRecords();
        console.log('Full data synced:', state.sheet1Data.length + state.sheet2Data.length, 'records');
    } catch (err) {
        console.error('Error loading full data:', err);
    }
}

function hideLoading() {
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) overlay.style.display = 'none';
}

function loadDemoData() {
    // ข้อมูลตัวอย่างสำหรับ demo mode
    state.sheet2Data = [
        { ที่: '1', ที่เก็บรักษา: 'วัดบูรณาราม', เลขมัด: 'บล.001', รหัสคัมภีร์: 'NS-BNR-PL-001-01', ชื่อเรื่อง: 'พระอภิธรรม 7 คัมภีร์', หมวด: 'พระอภิธรรมปิฎก', ยุคสมัย: 'รัตนโกสินทร์', อักษร: 'ขอม', ภาษา: 'บาลี', _source: 'sheet2' },
        { ที่: '2', ที่เก็บรักษา: 'วัดบูรณาราม', เลขมัด: 'บล.002', รหัสคัมภีร์: 'NS-BNR-PL-002-01', ชื่อเรื่อง: 'พุรอภิธมุมตถสงฺคหวิภาโคน', หมวด: 'พระอภิธรรมปิฎก', ยุคสมัย: 'รัตนโกสินทร์', อักษร: 'ขอม', ภาษา: 'บาลี', _source: 'sheet2' },
        { ที่: '3', ที่เก็บรักษา: 'วัดบูรณาราม', เลขมัด: 'บล.003', รหัสคัมภีร์: 'NS-BNR-PL-003-01', ชื่อเรื่อง: 'สังคินี', หมวด: 'พระอภิธรรมปิฎก', ยุคสมัย: 'อยุธยา', อักษร: 'ขอม', ภาษา: 'บาลี', _source: 'sheet2' },
        { ที่: '4', ที่เก็บรักษา: 'วัดบูรณาราม', เลขมัด: 'บล.004', รหัสคัมภีร์: 'NS-BNR-PL-004-01', ชื่อเรื่อง: 'วิภังค์', หมวด: 'พระอภิธรรมปิฎก', ยุคสมัย: 'รัตนโกสินทร์', อักษร: 'ขอม', ภาษา: 'บาลี-ไทย', _source: 'sheet2' },
        { ที่: '5', ที่เก็บรักษา: 'วัดท่าโพธิ์', เลขมัด: 'บล.005', รหัสคัมภีร์: 'NS-THP-PL-005-01', ชื่อเรื่อง: 'ธมฺมจกฺกปฺปวตฺตนสุตฺต', หมวด: 'พระสุตตันตปิฎก', ยุคสมัย: 'อยุธยา', อักษร: 'ขอม', ภาษา: 'บาลี', _source: 'sheet2' },
        { ที่: '6', ที่เก็บรักษา: 'วัดท่าโพธิ์', เลขมัด: 'บล.006', รหัสคัมภีร์: 'NS-THP-PL-006-01', ชื่อเรื่อง: 'มงฺคลสุตฺต', หมวด: 'พระสุตตันตปิฎก', ยุคสมัย: 'รัตนโกสินทร์', อักษร: 'ไทย', ภาษา: 'ไทย', _source: 'sheet2' },
        { ที่: '7', ที่เก็บรักษา: 'วัดวังตะวันตก', เลขมัด: 'บล.007', รหัสคัมภีร์: 'NS-WWT-PL-007-01', ชื่อเรื่อง: 'พระปาฏิโมกข์', หมวด: 'พระวินัยปิฎก', ยุคสมัย: 'รัตนโกสินทร์', อักษร: 'ขอม', ภาษา: 'บาลี', _source: 'sheet2' },
        { ที่: '8', ที่เก็บรักษา: 'วัดมเหยงค์', เลขมัด: 'บล.008', รหัสคัมภีร์: 'NS-MHK-PL-008-01', ชื่อเรื่อง: 'ตำนานเมืองนคร', หมวด: 'วรรณกรรมท้องถิ่น', ยุคสมัย: 'อยุธยา', อักษร: 'ไทย', ภาษา: 'ไทย', _source: 'sheet2' },
    ];

    // สร้าง dashboard data จาก demo
    state.dashboardData = computeDashboardFromData();
    
    renderDashboard();
    mergeAndDisplayRecords();
}

function computeDashboardFromData() {
    const s1 = state.sheet1Data;
    const s2 = state.sheet2Data;
    
    const categoryCount = {};
    const eraCount = {};
    const scriptCount = {};
    const languageCount = {};
    const locationCount = {};
    
    s1.forEach(r => {
        const cat = r['หมวดคัมภีร์'] || r['หมวด'] || 'ไม่ระบุ';
        categoryCount[cat] = (categoryCount[cat] || 0) + 1;
        const era = r['ยุคสมัย'] || 'ไม่ระบุ';
        eraCount[era] = (eraCount[era] || 0) + 1;
        const sc = r['อักษร'] || 'ไม่ระบุ';
        scriptCount[sc] = (scriptCount[sc] || 0) + 1;
        const lang = r['ภาษา'] || 'ไม่ระบุ';
        languageCount[lang] = (languageCount[lang] || 0) + 1;
    });
    
    s2.forEach(r => {
        const cat = r['หมวด'] || r['หมวดคัมภีร์'] || 'ไม่ระบุ';
        categoryCount[cat] = (categoryCount[cat] || 0) + 1;
        const era = r['ยุคสมัย'] || 'ไม่ระบุ';
        eraCount[era] = (eraCount[era] || 0) + 1;
        const sc = r['อักษร'] || 'ไม่ระบุ';
        scriptCount[sc] = (scriptCount[sc] || 0) + 1;
        const lang = r['ภาษา'] || 'ไม่ระบุ';
        languageCount[lang] = (languageCount[lang] || 0) + 1;
        const loc = r['ที่เก็บรักษา'] || 'ไม่ระบุ';
        locationCount[loc] = (locationCount[loc] || 0) + 1;
    });
    
    if (s1.length > 0) {
        locationCount['วัดวังตะวันตก'] = (locationCount['วัดวังตะวันตก'] || 0) + s1.length;
    }
    
    return {
        summary: {
            totalAll: s1.length + s2.length,
            totalSheet1: s1.length,
            totalSheet2: s2.length,
            digitized: s1.filter(r => r['Digitize'] === 'แล้ว').length
        },
        categoryCount,
        eraCount,
        scriptCount,
        languageCount,
        locationCount
    };
}

// ===== AUTO SYNC =====

function startAutoSync() {
    if (state.syncTimer) clearInterval(state.syncTimer);
    state.syncTimer = setInterval(() => {
        loadAllData();
    }, SYNC_INTERVAL);
}

function updateSyncStatus(isOnline) {
    state.isOnline = isOnline;
    const el = document.getElementById('syncStatus');
    if (isOnline) {
        el.innerHTML = '<i class="fas fa-circle sync-dot"></i><span>เชื่อมต่อแล้ว</span>';
        el.classList.remove('error');
    } else {
        el.innerHTML = '<i class="fas fa-circle sync-dot"></i><span>ขาดการเชื่อมต่อ</span>';
        el.classList.add('error');
    }
}

function updateLastSyncTime() {
    const now = new Date();
    const time = now.toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' });
    document.getElementById('lastSyncTime').textContent = `อัพเดทล่าสุด: ${time}`;
}

// ===== UI HELPERS =====

function hideLoading() {
    setTimeout(() => {
        document.getElementById('loadingOverlay').classList.add('hidden');
    }, 800);
}

function showToast(type, message) {
    const container = document.getElementById('toastContainer');
    const iconMap = { success: 'fa-check-circle', error: 'fa-exclamation-circle', info: 'fa-info-circle' };
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.innerHTML = `<i class="fas ${iconMap[type]}"></i><span>${message}</span>`;
    container.appendChild(toast);
    setTimeout(() => toast.remove(), 3000);
}

// ===== NAVIGATION =====

function navigateTo(page) {
    // Hide all pages
    document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
    // Show target page
    const target = document.getElementById(page + 'Page');
    if (target) target.classList.add('active');
    
    // Update nav links
    document.querySelectorAll('.nav-link').forEach(link => {
        link.classList.toggle('active', link.dataset.page === page);
    });
    
    // Update page title
    const titles = {
        dashboard: 'แดชบอร์ด',
        register: 'ลงทะเบียนคัมภีร์ใบลาน',
        records: 'รายการข้อมูลทั้งหมด'
    };
    document.getElementById('pageTitle').textContent = titles[page] || page;
    
    state.currentPage = page;
    
    // Close sidebar on mobile
    closeSidebar();
}

function closeSidebar() {
    document.getElementById('sidebar').classList.remove('open');
    const backdrop = document.querySelector('.sidebar-backdrop');
    if (backdrop) backdrop.classList.remove('active');
}

// ===== EVENT LISTENERS =====

function setupEventListeners() {
    // Navigation
    document.querySelectorAll('.nav-link').forEach(link => {
        link.addEventListener('click', (e) => {
            e.preventDefault();
            navigateTo(link.dataset.page);
        });
    });
    
    // Sidebar toggle (mobile)
    document.getElementById('menuToggle').addEventListener('click', () => {
        const sidebar = document.getElementById('sidebar');
        sidebar.classList.toggle('open');
        
        // Create/toggle backdrop
        let backdrop = document.querySelector('.sidebar-backdrop');
        if (!backdrop) {
            backdrop = document.createElement('div');
            backdrop.className = 'sidebar-backdrop';
            document.body.appendChild(backdrop);
            backdrop.addEventListener('click', closeSidebar);
        }
        backdrop.classList.toggle('active', sidebar.classList.contains('open'));
    });
    
    document.getElementById('sidebarClose').addEventListener('click', closeSidebar);
    
    // Refresh button
    document.getElementById('refreshBtn').addEventListener('click', async () => {
        const btn = document.getElementById('refreshBtn');
        btn.querySelector('i').classList.add('fa-spin');
        if (API_URL && API_URL !== 'YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL') {
            await loadAllData();
        } else {
            state.dashboardData = computeDashboardFromData();
            renderDashboard();
            mergeAndDisplayRecords();
        }
        btn.querySelector('i').classList.remove('fa-spin');
        showToast('success', 'รีเฟรชข้อมูลเรียบร้อย');
        updateLastSyncTime();
    });
    
    // Quick Add button
    document.getElementById('quickAddBtn').addEventListener('click', () => {
        navigateTo('register');
    });
    
    // Form step navigation
    document.getElementById('nextStepBtn').addEventListener('click', nextFormStep);
    document.getElementById('prevStepBtn').addEventListener('click', prevFormStep);
    
    // Form step indicators click
    document.querySelectorAll('.form-steps .step').forEach(step => {
        step.addEventListener('click', () => {
            const targetStep = parseInt(step.dataset.step);
            goToFormStep(targetStep);
        });
    });
    
    // Form submission
    document.getElementById('registrationForm').addEventListener('submit', handleFormSubmit);
    
    // Search & Filter
    document.getElementById('searchInput').addEventListener('input', debounce(filterRecords, 300));
    document.getElementById('sourceFilter').addEventListener('change', filterRecords);
    document.getElementById('categoryFilter').addEventListener('change', filterRecords);
    document.getElementById('eraFilter').addEventListener('change', filterRecords);
    
    // Table sorting
    document.querySelectorAll('.data-table th.sortable').forEach(th => {
        th.addEventListener('click', () => {
            const sortKey = th.dataset.sort;
            if (state.sortColumn === sortKey) {
                state.sortDirection = state.sortDirection === 'asc' ? 'desc' : 'asc';
            } else {
                state.sortColumn = sortKey;
                state.sortDirection = 'asc';
            }
            filterRecords();
        });
    });
    
    // Pagination
    document.getElementById('prevPageBtn').addEventListener('click', () => {
        if (state.currentTablePage > 1) {
            state.currentTablePage--;
            renderRecordsTable();
        }
    });
    document.getElementById('nextPageBtn').addEventListener('click', () => {
        const totalPages = Math.ceil(state.filteredRecords.length / RECORDS_PER_PAGE);
        if (state.currentTablePage < totalPages) {
            state.currentTablePage++;
            renderRecordsTable();
        }
    });
    
    // Modal close
    document.getElementById('modalClose').addEventListener('click', closeModal);
    document.getElementById('modalCloseBtn').addEventListener('click', closeModal);
    document.getElementById('detailModal').addEventListener('click', (e) => {
        if (e.target === e.currentTarget) closeModal();
    });

    // ===== AUTO-GENERATE FIELDS =====
    setupAutoGenerateFields();
}

function debounce(fn, delay) {
    let timer;
    return function (...args) {
        clearTimeout(timer);
        timer = setTimeout(() => fn.apply(this, args), delay);
    };
}

// ===== AUTO-GENERATE FIELD LOGIC =====

function setupAutoGenerateFields() {
    const เลขมัดInput = document.getElementById('เลขมัดInput');
    const เลขคัมภีร์Input = document.getElementById('เลขคัมภีร์Input');
    const หมวดSelect = document.getElementById('หมวดคัมภีร์');
    const หมวดCustomInput = document.getElementById('หมวดคัมภีร์Custom');

    // เลขมัด: พิมพ์ตัวเลข → แสดงตัวอย่าง บล.XXX
    if (เลขมัดInput) {
        เลขมัดInput.addEventListener('input', () => {
            updateAutoFields();
        });
    }

    // เลขคัมภีร์: พิมพ์ตัวเลข → แสดงตัวอย่าง
    if (เลขคัมภีร์Input) {
        เลขคัมภีร์Input.addEventListener('input', () => {
            updateAutoFields();
        });
    }

    // Generic custom dropdown handler
    const customDropdowns = [
        { selectId: 'หมวดคัมภีร์', customId: 'หมวดคัมภีร์Custom' },
        { selectId: 'ยุคสมัย', customId: 'ยุคสมัยCustom' },
        { selectId: 'อักษร', customId: 'อักษรCustom' },
        { selectId: 'ภาษา', customId: 'ภาษาCustom' },
        { selectId: 'เส้น', customId: 'เส้นCustom' },
        { selectId: 'ฉบับ', customId: 'ฉบับCustom' },
        { selectId: 'ชนิดไม้ประกับ', customId: 'ชนิดไม้ประกับCustom' },
        { selectId: 'ฉลากคัมภีร์', customId: 'ฉลากคัมภีร์Custom' },
        { selectId: 'ผ้าห่อคัมภีร์', customId: 'ผ้าห่อคัมภีร์Custom' },
        { selectId: 'สภาพ', customId: 'สภาพCustom' }
    ];

    customDropdowns.forEach(({ selectId, customId }) => {
        const select = document.getElementById(selectId);
        const customInput = document.getElementById(customId);
        if (select && customInput) {
            select.addEventListener('change', () => {
                if (select.value === '__custom__') {
                    customInput.style.display = 'block';
                    customInput.focus();
                } else {
                    customInput.style.display = 'none';
                    customInput.value = '';
                }
            });
        }
    });
}

function updateAutoFields() {
    const เลขมัดInput = document.getElementById('เลขมัดInput');
    const เลขคัมภีร์Input = document.getElementById('เลขคัมภีร์Input');
    const เลขมัดHidden = document.getElementById('เลขมัด');
    const เลขคัมภีร์Hidden = document.getElementById('เลขคัมภีร์');
    const รหัสField = document.getElementById('รหัสคัมภีร์');
    const เลขมัดPreview = document.getElementById('เลขมัดPreview');
    const เลขคัมภีร์Preview = document.getElementById('เลขคัมภีร์Preview');

    const มัดNum = parseInt(เลขมัดInput.value) || 0;
    const คัมภีร์Num = parseInt(เลขคัมภีร์Input.value) || 0;

    // จัดรูปแบบเลขมัด → บล.XXX
    if (มัดNum > 0 && มัดNum <= 999) {
        const มัดFormatted = 'บล.' + String(มัดNum).padStart(3, '0');
        เลขมัดHidden.value = มัดFormatted;
        เลขมัดPreview.textContent = '→ ' + มัดFormatted;
        เลขมัดPreview.style.color = 'var(--emerald)';
    } else {
        เลขมัดHidden.value = '';
        เลขมัดPreview.textContent = '';
    }

    // จัดรูปแบบเลขคัมภีร์
    if (คัมภีร์Num > 0 && คัมภีร์Num <= 99) {
        เลขคัมภีร์Hidden.value = String(คัมภีร์Num);
        เลขคัมภีร์Preview.textContent = '→ คัมภีร์ที่ ' + คัมภีร์Num;
        เลขคัมภีร์Preview.style.color = 'var(--emerald)';
    } else {
        เลขคัมภีร์Hidden.value = '';
        เลขคัมภีร์Preview.textContent = '';
    }

    // Auto-generate รหัสคัมภีร์: NS-VTT-PL-XXX-XX
    if (มัดNum > 0 && คัมภีร์Num > 0) {
        const มัดPad = String(มัดNum).padStart(3, '0');
        const คัมภีร์Pad = String(คัมภีร์Num).padStart(2, '0');
        รหัสField.value = `NS-VTT-PL-${มัดPad}-${คัมภีร์Pad}`;
    } else if (มัดNum > 0) {
        const มัดPad = String(มัดNum).padStart(3, '0');
        รหัสField.value = `NS-VTT-PL-${มัดPad}-__`;
    } else {
        รหัสField.value = '';
    }
}

function resetAutoFields() {
    const previews = ['เลขมัดPreview', 'เลขคัมภีร์Preview'];
    previews.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.textContent = '';
    });
    const hiddens = ['เลขมัด', 'เลขคัมภีร์'];
    hiddens.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = '';
    });
    const รหัส = document.getElementById('รหัสคัมภีร์');
    if (รหัส) รหัส.value = '';
    // Reset all custom dropdown inputs
    const customInputIds = ['หมวดคัมภีร์Custom', 'ยุคสมัยCustom', 'อักษรCustom', 'ภาษาCustom', 'เส้นCustom', 'ฉบับCustom', 'ชนิดไม้ประกับCustom', 'ฉลากคัมภีร์Custom', 'ผ้าห่อคัมภีร์Custom', 'สภาพCustom'];
    customInputIds.forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.style.display = 'none';
            el.value = '';
        }
    });
}

// ===== DASHBOARD =====

function renderDashboard() {
    const data = state.dashboardData;
    if (!data) return;
    
    // Animate stat numbers (มัด/เรื่อง/ผูก/Digitized)
    animateNumber('statBundles', data.summary.totalBundles || 0);
    animateNumber('statStories', data.summary.totalAll || 0);
    animateNumber('statVolumes', data.summary.totalVolumes || 0);
    animateNumber('statDigitized', data.summary.digitized || 0);
    
    // Render progress bars (สถิติการทำสำเนา)
    renderDigitizeProgress(data);
    
    // Render charts
    renderCategoryChart(data.categoryCount);
    renderEraChart(data.eraCount);
    renderScriptChart(data.scriptCount);
    renderDigitizeDonut(data);
    
    // Recent records table
    renderRecentTable();
}

function renderDigitizeProgress(data) {
    const s = data.summary;
    
    // มัด
    const bundlesTotal = s.totalBundles || 1;
    const bundlesDigi = s.digitizedBundles || 0;
    const bundlesPercent = Math.round((bundlesDigi / bundlesTotal) * 100);
    document.getElementById('digitizeBundles').textContent = `${bundlesDigi}/${bundlesTotal}`;
    document.getElementById('digitizeBundlesBar').style.width = `${bundlesPercent}%`;
    document.getElementById('digitizeBundlesPercent').textContent = `${bundlesPercent}%`;
    
    // เรื่อง
    const storiesTotal = s.totalAll || 1;
    const storiesDigi = s.digitizedStories || 0;
    const storiesPercent = Math.round((storiesDigi / storiesTotal) * 100);
    document.getElementById('digitizeStories').textContent = `${storiesDigi}/${storiesTotal}`;
    document.getElementById('digitizeStoriesBar').style.width = `${storiesPercent}%`;
    document.getElementById('digitizeStoriesPercent').textContent = `${storiesPercent}%`;
    
    // ผูก
    const volumesTotal = s.totalVolumes || 1;
    const volumesDigi = s.digitizedVolumes || 0;
    const volumesPercent = Math.round((volumesDigi / volumesTotal) * 100);
    document.getElementById('digitizeVolumes').textContent = `${volumesDigi}/${volumesTotal}`;
    document.getElementById('digitizeVolumesBar').style.width = `${volumesPercent}%`;
    document.getElementById('digitizeVolumesPercent').textContent = `${volumesPercent}%`;
}

function renderDigitizeDonut(data) {
    const ctx = document.getElementById('digitizeChart');
    if (!ctx) return;
    
    if (state.charts.digitize) state.charts.digitize.destroy();
    
    const digitized = data.summary.digitized || 0;
    const notDigitized = (data.summary.totalAll || 0) - digitized;
    
    state.charts.digitize = new Chart(ctx.getContext('2d'), {
        type: 'doughnut',
        data: {
            labels: ['Digitized แล้ว', 'ยังไม่ได้'],
            datasets: [{
                data: [digitized, notDigitized],
                backgroundColor: ['#C4A44B', '#E8DCC8'],
                borderWidth: 2,
                borderColor: '#fff'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '60%',
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: { font: { family: 'Sarabun', size: 12 } }
                }
            }
        }
    });
}

function animateNumber(elementId, target) {
    const element = document.getElementById(elementId);
    const start = parseInt(element.textContent) || 0;
    const duration = 800;
    const startTime = performance.now();
    
    function update(currentTime) {
        const elapsed = currentTime - startTime;
        const progress = Math.min(elapsed / duration, 1);
        const easeOut = 1 - Math.pow(1 - progress, 3);
        element.textContent = Math.round(start + (target - start) * easeOut).toLocaleString('th-TH');
        if (progress < 1) requestAnimationFrame(update);
    }
    requestAnimationFrame(update);
}

function renderCategoryChart(data) {
    const ctx = document.getElementById('categoryChart').getContext('2d');
    if (state.charts.category) state.charts.category.destroy();
    
    const labels = Object.keys(data);
    const values = Object.values(data);
    
    state.charts.category = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels,
            datasets: [{
                data: values,
                backgroundColor: CHART_COLORS.slice(0, labels.length),
                borderWidth: 2,
                borderColor: '#fff',
                hoverOffset: 8
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'right',
                    labels: {
                        font: { family: 'Sarabun', size: 11 },
                        padding: 12,
                        usePointStyle: true,
                        pointStyleWidth: 10
                    }
                }
            }
        }
    });
}

function renderEraChart(data) {
    const ctx = document.getElementById('eraChart').getContext('2d');
    if (state.charts.era) state.charts.era.destroy();
    
    const labels = Object.keys(data);
    const values = Object.values(data);
    
    state.charts.era = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: 'จำนวนคัมภีร์',
                data: values,
                backgroundColor: CHART_COLORS.slice(0, labels.length).map(c => c + 'cc'),
                borderColor: CHART_COLORS.slice(0, labels.length),
                borderWidth: 2,
                borderRadius: 8,
                barPercentage: 0.6
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: { font: { family: 'Sarabun' }, precision: 0 },
                    grid: { color: 'rgba(0,0,0,0.05)' }
                },
                x: {
                    ticks: { font: { family: 'Sarabun', size: 11 } },
                    grid: { display: false }
                }
            }
        }
    });
}

function renderScriptChart(data) {
    const ctx = document.getElementById('scriptChart').getContext('2d');
    if (state.charts.script) state.charts.script.destroy();
    
    const labels = Object.keys(data);
    const values = Object.values(data);
    
    state.charts.script = new Chart(ctx, {
        type: 'polarArea',
        data: {
            labels,
            datasets: [{
                data: values,
                backgroundColor: CHART_COLORS.slice(0, labels.length).map(c => c + '88'),
                borderColor: CHART_COLORS.slice(0, labels.length),
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'right',
                    labels: {
                        font: { family: 'Sarabun', size: 11 },
                        padding: 12,
                        usePointStyle: true
                    }
                }
            }
        }
    });
}

function renderLocationChart(data) {
    const ctx = document.getElementById('locationChart').getContext('2d');
    if (state.charts.location) state.charts.location.destroy();
    
    const labels = Object.keys(data);
    const values = Object.values(data);
    
    state.charts.location = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: 'จำนวนคัมภีร์',
                data: values,
                backgroundColor: CHART_COLORS.slice(2, 2 + labels.length).map(c => c + 'cc'),
                borderColor: CHART_COLORS.slice(2, 2 + labels.length),
                borderWidth: 2,
                borderRadius: 8,
                barPercentage: 0.5
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: 'y',
            plugins: {
                legend: { display: false }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    ticks: { font: { family: 'Sarabun' }, precision: 0 },
                    grid: { color: 'rgba(0,0,0,0.05)' }
                },
                y: {
                    ticks: { font: { family: 'Sarabun', size: 11 } },
                    grid: { display: false }
                }
            }
        }
    });
}

function renderRecentTable() {
    const tbody = document.getElementById('recentTableBody');
    const recentData = state.sheet1Data.slice(-10).reverse();
    
    if (recentData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="8" class="empty-message">ยังไม่มีข้อมูลในทะเบียนวัดวังตะวันตก</td></tr>';
        return;
    }
    
    tbody.innerHTML = recentData.map(row => `
        <tr>
            <td>${escapeHtml(row['ที่'] || '')}</td>
            <td>${escapeHtml(row['เลขมัด'] || '')}</td>
            <td>${escapeHtml(row['รหัสคัมภีร์'] || '')}</td>
            <td>${escapeHtml(row['ชื่อเรื่องบนคัมภีร์'] || row['ชื่อเรื่อง'] || '')}</td>
            <td>${escapeHtml(row['หมวดคัมภีร์'] || row['หมวด'] || '')}</td>
            <td>${escapeHtml(row['ยุคสมัย'] || '')}</td>
            <td>${escapeHtml(row['อักษร'] || '')}</td>
            <td>${escapeHtml(row['ภาษา'] || '')}</td>
        </tr>
    `).join('');
}

// ===== RECORDS TABLE (FULL) =====

function mergeAndDisplayRecords() {
    // Merge sheet1 + sheet2 records
    const s1 = state.sheet1Data.map(r => ({
        ...r,
        _source: 'sheet1',
        _sourceName: 'วัดวังตะวันตก',
        ชื่อเรื่อง: r['ชื่อเรื่องบนคัมภีร์'] || r['ชื่อเรื่อง'] || '',
        หมวด: r['หมวดคัมภีร์'] || r['หมวด'] || ''
    }));
    
    const s2 = state.sheet2Data.map(r => ({
        ...r,
        _source: 'sheet2',
        _sourceName: r['ที่เก็บรักษา'] || 'เมืองนคร',
        ชื่อเรื่อง: r['ชื่อเรื่อง'] || r['ชื่อเรื่องบนคัมภีร์'] || '',
        หมวด: r['หมวด'] || r['หมวดคัมภีร์'] || ''
    }));
    
    state.allRecords = [...s1, ...s2];
    
    // Populate filter dropdowns
    populateFilterDropdowns();
    
    filterRecords();
}

function populateFilterDropdowns() {
    // Category filter
    const categories = [...new Set(state.allRecords.map(r => r['หมวด']).filter(Boolean))];
    const catFilter = document.getElementById('categoryFilter');
    catFilter.innerHTML = '<option value="all">ทุกหมวด</option>';
    categories.forEach(cat => {
        catFilter.innerHTML += `<option value="${escapeHtml(cat)}">${escapeHtml(cat)}</option>`;
    });
    
    // Era filter
    const eras = [...new Set(state.allRecords.map(r => r['ยุคสมัย']).filter(Boolean))];
    const eraFilter = document.getElementById('eraFilter');
    eraFilter.innerHTML = '<option value="all">ทุกยุคสมัย</option>';
    eras.forEach(era => {
        eraFilter.innerHTML += `<option value="${escapeHtml(era)}">${escapeHtml(era)}</option>`;
    });
}

function filterRecords() {
    const search = document.getElementById('searchInput').value.toLowerCase().trim();
    const source = document.getElementById('sourceFilter').value;
    const category = document.getElementById('categoryFilter').value;
    const era = document.getElementById('eraFilter').value;
    
    let records = [...state.allRecords];
    
    // Filter by source
    if (source !== 'all') {
        records = records.filter(r => r._source === source);
    }
    
    // Filter by category
    if (category !== 'all') {
        records = records.filter(r => r['หมวด'] === category);
    }
    
    // Filter by era
    if (era !== 'all') {
        records = records.filter(r => r['ยุคสมัย'] === era);
    }
    
    // Search
    if (search) {
        records = records.filter(r => {
            return (r['ชื่อเรื่อง'] || '').toLowerCase().includes(search) ||
                   (r['รหัสคัมภีร์'] || '').toLowerCase().includes(search) ||
                   (r['เลขมัด'] || '').toLowerCase().includes(search) ||
                   (r['_sourceName'] || '').toLowerCase().includes(search);
        });
    }
    
    // Sort
    if (state.sortColumn) {
        records.sort((a, b) => {
            let valA = a[state.sortColumn] || '';
            let valB = b[state.sortColumn] || '';
            
            // Try numeric sort
            const numA = parseFloat(valA);
            const numB = parseFloat(valB);
            if (!isNaN(numA) && !isNaN(numB)) {
                return state.sortDirection === 'asc' ? numA - numB : numB - numA;
            }
            
            // String sort
            valA = valA.toString().toLowerCase();
            valB = valB.toString().toLowerCase();
            if (valA < valB) return state.sortDirection === 'asc' ? -1 : 1;
            if (valA > valB) return state.sortDirection === 'asc' ? 1 : -1;
            return 0;
        });
    }
    
    state.filteredRecords = records;
    state.currentTablePage = 1;
    renderRecordsTable();
}

function renderRecordsTable() {
    const tbody = document.getElementById('recordsTableBody');
    const records = state.filteredRecords;
    const totalPages = Math.max(1, Math.ceil(records.length / RECORDS_PER_PAGE));
    const start = (state.currentTablePage - 1) * RECORDS_PER_PAGE;
    const end = start + RECORDS_PER_PAGE;
    const pageRecords = records.slice(start, end);
    
    if (pageRecords.length === 0) {
        tbody.innerHTML = '<tr><td colspan="10" class="empty-message">ไม่พบข้อมูล</td></tr>';
    } else {
        tbody.innerHTML = pageRecords.map((row, i) => `
            <tr data-index="${start + i}" style="cursor:pointer;" onclick="showRecordDetail(${start + i})">
                <td>${escapeHtml(row['ที่'] || '')}</td>
                <td><span class="source-badge ${row._source}">${escapeHtml(row._sourceName)}</span></td>
                <td>${escapeHtml(row['เลขมัด'] || '')}</td>
                <td>${escapeHtml(row['รหัสคัมภีร์'] || '')}</td>
                <td>${escapeHtml(row['ชื่อเรื่อง'] || '')}</td>
                <td>${escapeHtml(row['หมวด'] || '')}</td>
                <td>${escapeHtml(row['ยุคสมัย'] || '')}</td>
                <td>${escapeHtml(row['อักษร'] || '')}</td>
                <td>${escapeHtml(row['ภาษา'] || '')}</td>
                <td>
                    <div class="row-actions">
                        <button class="btn-icon" onclick="event.stopPropagation(); showRecordDetail(${start + i})" title="ดูรายละเอียด">
                            <i class="fas fa-eye"></i>
                        </button>
                        ${row._source === 'sheet1' ? `
                        <button class="btn-icon" onclick="event.stopPropagation(); editRecord(${start + i})" title="แก้ไข">
                            <i class="fas fa-edit"></i>
                        </button>
                        <button class="btn-icon btn-delete" onclick="event.stopPropagation(); deleteRecordConfirm(${start + i})" title="ลบ">
                            <i class="fas fa-trash-alt"></i>
                        </button>` : ''}
                    </div>
                </td>
            </tr>
        `).join('');
    }
    
    // Update pagination
    document.getElementById('recordCount').textContent = `แสดง ${records.length} รายการ`;
    document.getElementById('pageInfo').textContent = `หน้า ${state.currentTablePage} / ${totalPages}`;
    document.getElementById('prevPageBtn').disabled = state.currentTablePage <= 1;
    document.getElementById('nextPageBtn').disabled = state.currentTablePage >= totalPages;
}

// ===== RECORD DETAIL MODAL =====

function showRecordDetail(index) {
    const record = state.filteredRecords[index];
    if (!record) return;
    
    const modal = document.getElementById('detailModal');
    const body = document.getElementById('modalBody');
    
    // Build detail fields
    const fields = [
        { label: 'ลำดับ', value: record['ที่'] },
        { label: 'แหล่งข้อมูล', value: record._sourceName },
        { label: 'เลขมัด', value: record['เลขมัด'] },
        { label: 'รหัสคัมภีร์', value: record['รหัสคัมภีร์'] },
        { label: 'ชื่อเรื่อง', value: record['ชื่อเรื่อง'], fullWidth: true },
        { label: 'หมวด', value: record['หมวด'] },
        { label: 'ยุคสมัย', value: record['ยุคสมัย'] },
        { label: 'อักษร', value: record['อักษร'] },
        { label: 'ภาษา', value: record['ภาษา'] },
        { label: 'จำนวนผูก', value: record['จำนวนผูก'] },
        { label: 'ผูกที่มี', value: record['ผูกที่มี'] },
        { label: 'เส้น', value: record['เส้น'] },
        { label: 'ฉบับ', value: record['ฉบับ'] },
        { label: 'ชนิดไม้ประกับ', value: record['ชนิดไม้ประกับ'] },
        { label: 'ขนาดคัมภีร์', value: record['ขนาดคัมภีร์'] },
        { label: 'สภาพคัมภีร์', value: record['สภาพคัมภีร์'] },
        { label: 'Digitize', value: record['Digitize'] },
        { label: 'ผู้บันทึก', value: record['ผู้บันทึก'] },
        { label: 'หมายเหตุ', value: record['หมายเหตุ'], fullWidth: true },
        { label: 'ประวัติ/บริบท', value: record['ประวัติ/บริบท'], fullWidth: true },
    ].filter(f => f.value && f.value.toString().trim());
    
    body.innerHTML = `<div class="detail-grid">
        ${fields.map(f => `
            <div class="detail-item ${f.fullWidth ? 'full-width' : ''}">
                <div class="detail-label">${escapeHtml(f.label)}</div>
                <div class="detail-value">${escapeHtml(f.value)}</div>
            </div>
        `).join('')}
    </div>`;
    
    // Show/hide edit/delete buttons
    const editBtn = document.getElementById('modalEditBtn');
    const deleteBtn = document.getElementById('modalDeleteBtn');
    
    if (record._source === 'sheet1') {
        editBtn.style.display = '';
        deleteBtn.style.display = '';
        editBtn.onclick = () => { closeModal(); editRecord(index); };
        deleteBtn.onclick = () => { closeModal(); deleteRecord(index); };
    } else {
        editBtn.style.display = 'none';
        deleteBtn.style.display = 'none';
    }
    
    modal.classList.add('active');
}

function closeModal() {
    document.getElementById('detailModal').classList.remove('active');
}

// ===== FORM MANAGEMENT =====

function populateDropdowns(options) {
    // dropdown ที่ hardcode ไว้ใน HTML แล้ว ไม่ต้อง populate
    const hardcodedDropdowns = ['หมวดคัมภีร์', 'ยุคสมัย', 'อักษร', 'ภาษา', 'เส้น', 'ฉบับ', 'ชนิดไม้ประกับ', 'ฉลากคัมภีร์', 'ผ้าห่อคัมภีร์', 'สภาพ'];
    Object.keys(options).forEach(fieldName => {
        if (hardcodedDropdowns.includes(fieldName)) return;
        
        const select = document.getElementById(fieldName);
        if (!select) return;
        
        // Keep first option (placeholder)
        const placeholder = select.options[0];
        select.innerHTML = '';
        select.appendChild(placeholder);
        
        options[fieldName].forEach(optVal => {
            const opt = document.createElement('option');
            opt.value = optVal;
            opt.textContent = optVal;
            select.appendChild(opt);
        });
    });
}

function goToFormStep(step) {
    if (step < 1 || step > state.totalSteps) return;
    
    state.formStep = step;
    
    // Update step indicators
    document.querySelectorAll('.form-steps .step').forEach(s => {
        const sStep = parseInt(s.dataset.step);
        s.classList.remove('active', 'completed');
        if (sStep === step) s.classList.add('active');
        else if (sStep < step) s.classList.add('completed');
    });
    
    // Show/hide step content
    document.querySelectorAll('.form-step-content').forEach(c => {
        c.classList.toggle('active', parseInt(c.dataset.step) === step);
    });
    
    // Update buttons
    document.getElementById('prevStepBtn').style.display = step > 1 ? '' : 'none';
    document.getElementById('nextStepBtn').style.display = step < state.totalSteps ? '' : 'none';
    document.getElementById('submitBtn').style.display = step === state.totalSteps ? '' : 'none';
}

function nextFormStep() {
    // Validate current step required fields
    const currentContent = document.querySelector(`.form-step-content[data-step="${state.formStep}"]`);
    const requiredFields = currentContent.querySelectorAll('[required]');
    let valid = true;
    requiredFields.forEach(field => {
        if (!field.value.trim()) {
            field.style.borderColor = '#dc2626';
            valid = false;
        } else {
            field.style.borderColor = '';
        }
    });
    
    if (!valid) {
        showToast('error', 'กรุณากรอกข้อมูลที่จำเป็น');
        return;
    }
    
    goToFormStep(state.formStep + 1);
}

function prevFormStep() {
    goToFormStep(state.formStep - 1);
}

async function handleFormSubmit(e) {
    e.preventDefault();
    
    const form = e.target;
    const formData = new FormData(form);
    const record = {};
    
    for (const [key, value] of formData.entries()) {
        record[key] = value;
    }
    
    // แปลงค่า "ระบุ..." เป็นค่าจาก custom input สำหรับทุก dropdown
    const customFields = [
        { name: 'หมวดคัมภีร์', customId: 'หมวดคัมภีร์Custom' },
        { name: 'ยุคสมัย', customId: 'ยุคสมัยCustom' },
        { name: 'อักษร', customId: 'อักษรCustom' },
        { name: 'ภาษา', customId: 'ภาษาCustom' },
        { name: 'เส้น', customId: 'เส้นCustom' },
        { name: 'ฉบับ', customId: 'ฉบับCustom' },
        { name: 'ชนิดไม้ประกับ', customId: 'ชนิดไม้ประกับCustom' },
        { name: 'ฉลากคัมภีร์', customId: 'ฉลากคัมภีร์Custom' },
        { name: 'ผ้าห่อคัมภีร์', customId: 'ผ้าห่อคัมภีร์Custom' },
        { name: 'สภาพ', customId: 'สภาพCustom' }
    ];
    customFields.forEach(({ name, customId }) => {
        if (record[name] === '__custom__') {
            const customVal = document.getElementById(customId).value.trim();
            record[name] = customVal || 'ไม่ระบุ';
        }
    });

    // Validate required fields
    if (!record['เลขมัด'] || !record['รหัสคัมภีร์'] || !record['ชื่อเรื่องบนคัมภีร์'] || !record['ผู้บันทึก']) {
        showToast('error', 'กรุณากรอกข้อมูลที่จำเป็นให้ครบ');
        return;
    }
    
    if (API_URL && API_URL !== 'YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL') {
        try {
            const submitBtn = document.getElementById('submitBtn');
            submitBtn.disabled = true;
            submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> กำลังบันทึก...';
            
            await apiPost('addRecord', { record });
            
            showToast('success', 'บันทึกข้อมูลเรียบร้อยแล้ว');
            form.reset();
            resetAutoFields();
            goToFormStep(1);
            
            submitBtn.disabled = false;
            submitBtn.innerHTML = '<i class="fas fa-save"></i> บันทึกข้อมูล';
            
            // Sync background (ไม่ block UI)
            loadFullDataBackground();
        } catch (err) {
            showToast('error', 'เกิดข้อผิดพลาด: ' + err.message);
            document.getElementById('submitBtn').disabled = false;
            document.getElementById('submitBtn').innerHTML = '<i class="fas fa-save"></i> บันทึกข้อมูล';
        }
    } else {
        // Demo mode - add to local state
        record['ที่'] = (state.sheet1Data.length + 1).toString();
        state.sheet1Data.push(record);
        showToast('success', 'บันทึกข้อมูลเรียบร้อย (โหมดสาธิต)');
        form.reset();
        resetAutoFields();
        goToFormStep(1);
        
        state.dashboardData = computeDashboardFromData();
        renderDashboard();
        mergeAndDisplayRecords();
    }
}

// ===== EDIT & DELETE =====

function editRecord(index) {
    const record = state.filteredRecords[index];
    if (!record || record._source !== 'sheet1') {
        showToast('error', 'สามารถแก้ไขได้เฉพาะข้อมูลวัดวังตะวันตก');
        return;
    }
    
    // Navigate to form page
    navigateTo('register');
    
    // Fill form with record data
    const fieldMappings = [
        'เลขมัด', 'รหัสคัมภีร์', 'ชื่อเรื่องบนคัมภีร์', 'หมวดคัมภีร์',
        'เรื่อง', 'จำนวนผูก', 'ผูกที่มี', 'ยุคสมัย', 'อักษร', 'ภาษา',
        'เส้น', 'ฉบับ', 'ชนิดไม้ประกับ', 'ฉลากคัมภีร์', 'ผ้าห่อคัมภีร์',
        'ขนาดคัมภีร์', 'ขนาดไม้ประกับ', 'ตู้ที่', 'ชั้นที่',
        'ปีที่สร้าง', 'ชื่อผู้สร้าง', 'ประวัติ/บริบท', 'สภาพคัมภีร์',
        'Digitize', 'หมายเหตุสภาพ', 'หมายเหตุ', 'ประวัติการจัดการ', 'ผู้บันทึก'
    ];
    
    fieldMappings.forEach(field => {
        const el = document.getElementById(field);
        if (el && record[field]) {
            el.value = record[field];
        }
    });
    
    goToFormStep(1);
    showToast('info', 'แก้ไขข้อมูล - กรุณาตรวจสอบแล้วกดบันทึก');
}

async function deleteRecord(index) {
    const record = state.filteredRecords[index];
    if (!record || record._source !== 'sheet1') {
        showToast('error', 'สามารถลบได้เฉพาะข้อมูลวัดวังตะวันตก');
        return;
    }
    
    if (!confirm(`ต้องการลบรายการ "${record['ชื่อเรื่อง'] || record['รหัสคัมภีร์']}" หรือไม่?`)) {
        return;
    }
    
    if (API_URL && API_URL !== 'YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL') {
        try {
            // Optimistic: ลบจาก local state ทันที
            state.sheet1Data = state.sheet1Data.filter(r => r._rowIndex !== record._rowIndex);
            mergeAndDisplayRecords();
            showToast('success', 'ลบรายการเรียบร้อย');
            
            // ส่ง API ใน background (ไม่ block UI)
            apiPost('deleteRecord', { rowIndex: record._rowIndex }).then(() => {
                loadFullDataBackground(); // sync จริงทีหลัง
            }).catch(err => {
                showToast('error', 'เกิดข้อผิดพลาด: ' + err.message);
                loadFullDataBackground(); // reload ถ้า error
            });
        } catch (err) {
            showToast('error', 'เกิดข้อผิดพลาด: ' + err.message);
        }
    } else {
        // Demo mode
        state.sheet1Data = state.sheet1Data.filter(r => r['ที่'] !== record['ที่']);
        state.dashboardData = computeDashboardFromData();
        renderDashboard();
        mergeAndDisplayRecords();
        showToast('success', 'ลบรายการเรียบร้อย (โหมดสาธิต)');
    }
}

function deleteRecordConfirm(index) {
    deleteRecord(index);
}

// ===== UTILITIES =====

function escapeHtml(str) {
    if (!str) return '';
    const div = document.createElement('div');
    div.textContent = str.toString();
    return div.innerHTML;
}
