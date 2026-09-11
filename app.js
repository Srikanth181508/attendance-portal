const DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const PERIOD_TIMES = ['9:30–10:30', '10:30–11:25', '11:35–12:30', '1:30–2:30', '2:30–3:25', '3:35–4:30'];
const NUM_PERIODS = 6;

// Empty string sets API root dynamically for both Render and Localhost
// Auto-detect: Live Server (5500) வழியாகத் திறந்தாலும் Port 3000 Node சர்வருக்கு அனுப்பும்
const API_BASE = (window.location.port && window.location.port !== '3000' && !window.location.hostname.includes('render.com') && !window.location.hostname.includes('ngrok'))
  ? 'http://localhost:3000'
  : '';
// Helper to bypass browser warning on standard JSON fetch requests
async function secureFetch(url, options = {}) {
  const headers = {
    'ngrok-skip-browser-warning': 'true',
    ...(options.headers || {})
  };
  return fetch(url, { ...options, headers });
}

let state = {
  view: 'login',
  role: null, // 'admin' | 'it' | 'advisor'
  dept: 'aids', // 'aids' | 'it'
  year: null,
  currentStaff: null
};

const AVAILABLE_YEARS = [1, 2, 3, 4];

function getDeptLabel(dept = state.dept) {
  return dept === 'it' ? 'Information Technology (IT)' : 'AI & DS (Artificial Intelligence & Data Science)';
}

function getDeptShort(dept = state.dept) {
  return dept === 'it' ? 'IT' : 'AI & DS';
}

function getYearKey(dept = state.dept, year = state.year) {
  const d = (dept || 'aids').toLowerCase().replace(/[^a-z0-9]/g, '');
  const y = year || 1;
  return `${d}_year${y}`;
}

function getYearLabel(year) {
  if (year === 1) return 'First';
  if (year === 2) return 'Second';
  if (year === 3) return 'Third';
  return 'Fourth';
}

function emptyTimetable() {
  return DAYS.map(d => ({
    day: d,
    periods: Array.from({ length: NUM_PERIODS }, () => ({ subject: '', staff: '' }))
  }));
}

function showToast(msg) {
  const t = document.getElementById('toast');
  if (!t) return;
  const msgEl = document.getElementById('toastMsg');
  if (msgEl) msgEl.textContent = msg;
  t.classList.add('show');
  clearTimeout(window._toastTimer);
  window._toastTimer = setTimeout(() => t.classList.remove('show'), 2400);
}

// ---------------- EXCEL EXPORT HELPERS (SheetJS) ---------------- //
function exportDailyToExcel(dept, year, date, studentList, studentMap) {
  if (typeof XLSX === 'undefined') {
    return showToast('SheetJS library not loaded.');
  }
  if (!studentList || studentList.length === 0) {
    return showToast('No student records available to export.');
  }

  const wsData = [
    [`${getDeptLabel(dept)} - Year ${year} - Daily Attendance Report`],
    [`Date: ${date}`],
    [],
    ['Sl', 'Student Name', 'Reg No', 'P1', 'P2', 'P3', 'P4', 'P5', 'P6']
  ];

  studentList.forEach((s, idx) => {
    wsData.push([
      idx + 1,
      s.name,
      s.reg,
      studentMap[s.reg]?.periods['P1'] || '-',
      studentMap[s.reg]?.periods['P2'] || '-',
      studentMap[s.reg]?.periods['P3'] || '-',
      studentMap[s.reg]?.periods['P4'] || '-',
      studentMap[s.reg]?.periods['P5'] || '-',
      studentMap[s.reg]?.periods['P6'] || '-'
    ]);
  });

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = [{ wch: 6 }, { wch: 24 }, { wch: 16 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, ws, `Daily_${date}`);
  XLSX.writeFile(wb, `${getDeptShort(dept)}_Year${year}_Daily_Attendance_${date}.xlsx`);
  showToast('Daily Attendance Excel downloaded!');
}

function exportMonthlyToExcel(dept, year, monthVal, list) {
  if (typeof XLSX === 'undefined') {
    return showToast('SheetJS library not loaded.');
  }
  if (!list || list.length === 0) {
    return showToast('No data available to export.');
  }

  const wsData = [
    [`${getDeptLabel(dept)} - Year ${year} - Monthly Cumulative Attendance Report`],
    [`Month: ${monthVal}`],
    [],
    ['Sl', 'Student Name', 'Total Hours', 'Present', 'OD', 'Absent', 'Cumulative %', 'Eligibility (<75%)']
  ];

  list.forEach((st, idx) => {
    wsData.push([
      idx + 1,
      st.studentName,
      st.totalPeriods,
      st.present,
      st.od,
      st.absent,
      `${st.percentage}%`,
      st.isShortage ? 'Shortage (<75%)' : 'Eligible'
    ]);
  });

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = [{ wch: 6 }, { wch: 25 }, { wch: 14 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 16 }, { wch: 20 }];
  XLSX.utils.book_append_sheet(wb, ws, `Monthly_${monthVal}`);
  XLSX.writeFile(wb, `${getDeptShort(dept)}_Year${year}_Cumulative_Report_${monthVal}.xlsx`);
  showToast('Monthly Cumulative Excel downloaded!');
}

// ---------------- API CALLS ---------------- //
async function loadTimetable(year = state.year, dept = state.dept) {
  const targetKey = getYearKey(dept, year);
  try {
    const res = await secureFetch(`${API_BASE}/api/timetable/${targetKey}`);
    if (!res.ok) return { grid: null, pdf: null };
    return await res.json();
  } catch (e) {
    return { grid: null, pdf: null };
  }
}

async function saveTimetable(grid, year = state.year, dept = state.dept) {
  const targetKey = getYearKey(dept, year);
  try {
    const res = await secureFetch(`${API_BASE}/api/timetable/${targetKey}`, {
      method: 'POST',
      headers: { 
        'Content-Type': 'application/json' 
      },
      body: JSON.stringify(grid)
    });
    
    if (!res.ok) {
      const errText = await res.text();
      console.error('Save Timetable Server Error:', errText);
      return false;
    }
    return true;
  } catch (e) {
    console.error('Save Timetable Fetch Error:', e);
    return false;
  }
}

// Fixed FormData fetch: lets the browser auto-set multipart boundary
async function uploadPdf(file, year = state.year, dept = state.dept) {
  const targetKey = getYearKey(dept, year);
  try {
    const formData = new FormData();
    formData.append('pdf', file);
    
    const res = await fetch(`${API_BASE}/api/timetable/pdf/${targetKey}`, {
      method: 'POST',
      headers: {
        'ngrok-skip-browser-warning': 'true'
      },
      body: formData
    });

    if (!res.ok) {
      const errText = await res.text();
      return { success: false, error: `Server error (${res.status}): ${errText}` };
    }
    return await res.json();
  } catch (e) {
    console.error('PDF Upload Error:', e);
    return { success: false, error: 'Connection failed with server' };
  }
}

async function loadYearStudents(year = state.year, dept = state.dept) {
  const targetKey = getYearKey(dept, year);
  try {
    const res = await secureFetch(`${API_BASE}/api/students/${targetKey}`);
    if (!res.ok) return { students: [], pdf: null };
    const data = await res.json();
    return {
      students: Array.isArray(data.students) ? data.students : [],
      pdf: data.pdf || null
    };
  } catch (e) {
    return { students: [], pdf: null };
  }
}

async function saveYearStudents(studentsList, year = state.year, dept = state.dept) {
  const targetKey = getYearKey(dept, year);
  try {
    const res = await secureFetch(`${API_BASE}/api/students/${targetKey}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(studentsList)
    });
    return res.ok;
  } catch (e) {
    return false;
  }
}

// Fixed FormData fetch for Student Namelist PDF
async function uploadStudentPdf(file, year = state.year, dept = state.dept) {
  const targetKey = getYearKey(dept, year);
  try {
    const formData = new FormData();
    formData.append('pdf', file);

    const res = await fetch(`${API_BASE}/api/students/pdf/${targetKey}`, {
      method: 'POST',
      headers: {
        'ngrok-skip-browser-warning': 'true'
      },
      body: formData
    });

    if (!res.ok) {
      const errText = await res.text();
      return { success: false, error: `Server error (${res.status}): ${errText}` };
    }
    return await res.json();
  } catch (e) {
    console.error('Student PDF Upload Error:', e);
    return { success: false, error: 'Connection failed with server' };
  }
}

async function staffLoginAPI(staffName, password, newPassword = null) {
  try {
    const res = await secureFetch(`${API_BASE}/api/staff/login`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ staffName, password, newPassword })
    });
    return await res.json();
  } catch (e) {
    return { success: false, error: 'Server connection failed' };
  }
}

async function loadAttendance(dateStr, periodIdx, year = state.year, dept = state.dept) {
  const targetKey = getYearKey(dept, year);
  try {
    const res = await secureFetch(`${API_BASE}/api/attendance/${targetKey}?date=${dateStr}&period=${periodIdx + 1}`);
    if (!res.ok) return null;
    return await res.json();
  } catch (e) {
    return null;
  }
}

async function saveAttendance(dateStr, periodIdx, payload, year = state.year, dept = state.dept) {
  const targetKey = getYearKey(dept, year);
  try {
    const res = await secureFetch(`${API_BASE}/api/attendance/${targetKey}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ dateStr, periodIdx, payload })
    });
    return res.ok;
  } catch (e) {
    return false;
  }
}

function todayInfo() {
  const now = new Date();
  const jsDay = now.getDay();
  const dayName = jsDay === 0 ? null : DAYS[jsDay - 1];
  const dateStr = now.toISOString().slice(0, 10);
  return { now, dayName, dateStr };
}

function parsePeriodRange(str) {
  const [a, b] = str.split('–');
  function toMinutes(t) {
    let [h, m] = t.split(':').map(Number);
    if (h < 9) h += 12;
    return h * 60 + m;
  }
  return [toMinutes(a), toMinutes(b)];
}

function initials(name) {
  if (!name || !name.trim()) return 'IT';
  return name.trim().split(/\s+/).map(w => w[0]).join('').slice(0, 2).toUpperCase();
}

function attachTilt(el) {
  if (!el) return;
  el.style.setProperty('--mx', '50%');
  el.style.setProperty('--my', '50%');
  el.addEventListener('mousemove', (e) => {
    const r = el.getBoundingClientRect();
    const px = (e.clientX - r.left) / r.width;
    const py = (e.clientY - r.top) / r.height;
    const rx = (py - 0.5) * -10;
    const ry = (px - 0.5) * 10;
    el.style.transform = `perspective(700px) rotateX(${rx}deg) rotateY(${ry}deg) translateZ(6px)`;
    el.style.setProperty('--mx', (px * 100) + '%');
    el.style.setProperty('--my', (py * 100) + '%');
  });
  el.addEventListener('mouseleave', () => {
    el.style.transform = 'perspective(700px) rotateX(0) rotateY(0) translateZ(0)';
  });
}

function setCrumbs() {
  const parts = [];
  if (state.role === 'admin') parts.push('Admin');
  else if (state.role === 'it') parts.push('Faculty Portal');
  else if (state.role === 'advisor') parts.push('Class Advisor');

  if (state.dept) parts.push(getDeptShort());
  if (state.year) parts.push('Year ' + state.year);
  if (state.currentStaff) parts.push(state.currentStaff);

  document.getElementById('crumbs').innerHTML = parts.length
    ? parts.map((p, i) => (i > 0 ? '<span>/</span>' : '') + '<b>' + p + '</b>').join(' ')
    : '';
  const right = document.getElementById('topbarRight');
  right.innerHTML = state.view !== 'login'
    ? '<button class="btn-ghost" onclick="goBack()">← Back</button><button class="btn-ghost" onclick="logout()">Logout</button>'
    : '';
}

function goBack() {
  if (state.role === 'admin') {
    if (state.view === 'admin') { logout(); return; }
    state.view = 'admin';
    state.year = null;
    render();
    return;
  }
  if (state.role === 'advisor') {
    if (state.view === 'advisorDashboard') {
      state.view = 'advisorYear';
      state.year = null;
      render();
      return;
    }
    if (state.view === 'advisorYear') { logout(); return; }
  }
  if (state.view === 'attendance') {
    state.view = 'staffSelect';
    state.currentStaff = null;
    render();
    return;
  }
  if (state.view === 'staffSelect' || state.view === 'timetable' || state.view === 'students' || state.view === 'manageStudents') {
    state.view = state.role === 'admin' ? 'admin' : 'year';
    state.currentStaff = null;
    render();
    return;
  }
  if (state.view === 'year') { logout(); return; }
  render();
}

function logout() {
  state = { view: 'login', role: null, dept: 'aids', year: null, currentStaff: null };
  render();
}

function render() {
  setCrumbs();
  const el = document.getElementById('view');
  el.classList.remove('view');
  void el.offsetWidth;
  el.classList.add('view');

  if (state.view === 'login') return renderLogin(el);
  if (state.view === 'year') return renderYearSelect(el);
  if (state.view === 'advisorYear') return renderAdvisorYearSelect(el);
  if (state.view === 'advisorDashboard') return renderAdvisorDashboard(el);
  if (state.view === 'staffSelect') return renderStaffSelect(el);
  if (state.view === 'attendance') return renderAttendance(el);
  if (state.view === 'timetable') return renderTimetableView(el);
  if (state.view === 'upload') return renderUpload(el);
  if (state.view === 'manageStudents') return renderManageStudents(el);
  if (state.view === 'admin') return renderAdmin(el);
  if (state.view === 'students') return renderStudents(el);
}

/* 1. LOGIN CHOICE */
function renderLogin(el) {
  el.innerHTML = `
    <div class="eyebrow">Sign in</div>
    <h1 class="title">Choose your console</h1>
    <p class="lede">Configure timetables & student rosters, mark live period attendance, or monitor cumulative analytics.</p>
    <div class="choice-grid cols-3">
      <div class="tilt-card" id="card-admin">
        <span class="tag">01</span>
        <div class="card-icon">🛡️</div>
        <h3>Admin Console</h3>
        <p>Manage timetable slots, PDFs, and add Year-wise student lists & PDFs.</p>
      </div>
      <div class="tilt-card" id="card-it">
        <span class="tag">02</span>
        <div class="card-icon">🎓</div>
        <h3>Faculty Portal</h3>
        <p>Select Department &rarr; Year &rarr; Staff Profile &rarr; Mark live attendance (Present / Absent / OD).</p>
      </div>
      <div class="tilt-card advisor-card" id="card-advisor">
        <span class="tag">03</span>
        <div class="card-icon" style="background: rgba(168, 85, 247, 0.2); color: #c084fc;">📊</div>
        <h3 style="color: #c084fc;">Class Advisor</h3>
        <p>Daily period sheets, monthly cumulative % calculation, and Excel export.</p>
      </div>
    </div>
  `;

  document.getElementById('card-admin').onclick = () => { state.role = 'admin'; state.view = 'admin'; render(); };
  document.getElementById('card-it').onclick = () => { state.role = 'it'; state.view = 'year'; render(); };
  document.getElementById('card-advisor').onclick = () => { state.role = 'advisor'; state.view = 'advisorYear'; render(); };
  [...el.querySelectorAll('.tilt-card')].forEach(attachTilt);
}

/* 2. DEPARTMENT & YEAR SELECTION (FACULTY PORTAL) */
function renderYearSelect(el) {
  el.innerHTML = `
    <div class="eyebrow">Faculty Portal</div>
    <h1 class="title">Select Department & Year</h1>
    <p class="lede">Choose your academic department to expand and select your class year.</p>
    
    <!-- DEPARTMENT CARDS GRID -->
    <div class="choice-grid cols-2" id="deptSelectorGrid">
      <div class="tilt-card ${state.dept === 'aids' ? 'active-dept-card' : ''}" data-dept="aids" style="border-top: 4px solid var(--cyan, #06b6d4);">
        <span class="tag">DEPT 01</span>
        <div class="card-icon" style="font-size: 30px;">🤖</div>
        <h3>AI & DS</h3>
        <p style="color:var(--cyan); font-weight:600;">Artificial Intelligence & Data Science</p>
        <p style="font-size:12px; color:var(--text-dimmer); margin-top:8px;">Click to view 1st, 2nd, 3rd & 4th Years &rarr;</p>
      </div>

      <div class="tilt-card ${state.dept === 'it' ? 'active-dept-card' : ''}" data-dept="it" style="border-top: 4px solid var(--violet, #8b5cf6);">
        <span class="tag">DEPT 02</span>
        <div class="card-icon" style="font-size: 30px;">💻</div>
        <h3>IT</h3>
        <p style="color:var(--violet); font-weight:600;">Information Technology</p>
        <p style="font-size:12px; color:var(--text-dimmer); margin-top:8px;">Click to view 1st, 2nd, 3rd & 4th Years &rarr;</p>
      </div>
    </div>

    <!-- EXPANDABLE YEARS BOX GRID -->
    <div class="panel" id="yearBoxContainer" style="margin-top: 25px; display: block;">
      <div class="panel-head" style="margin-bottom: 15px;">
        <div style="font-weight:700; font-family:'Space Grotesk',sans-serif;" id="yearHeaderLabel">
          Select Year for ${getDeptLabel()}
        </div>
      </div>

      <div class="choice-grid cols-4" id="yearCardsGrid">
        ${AVAILABLE_YEARS.map(y => `
          <div class="tilt-card" data-year="${y}">
            <span class="tag">YR ${y}</span>
            <div class="year-badge">0${y}</div>
            <h3>${getYearLabel(y)} Year</h3>
            <p>Assigned faculty & periods for Year ${y}.</p>
          </div>
        `).join('')}
      </div>
    </div>
  `;

  el.querySelectorAll('[data-dept]').forEach(card => {
    attachTilt(card);
    card.onclick = () => {
      state.dept = card.dataset.dept;
      renderYearSelect(el);
    };
  });

  el.querySelectorAll('[data-year]').forEach(card => {
    attachTilt(card);
    card.onclick = () => {
      state.year = parseInt(card.dataset.year);
      state.view = 'staffSelect';
      render();
    };
  });
}

/* 3. CLASS ADVISOR TARGET SELECTION */
function renderAdvisorYearSelect(el) {
  el.innerHTML = `
    <div class="eyebrow">Class Advisor Console</div>
    <h1 class="title">Select Department & Target Class</h1>
    <p class="lede">Choose department and year to view daily logs and monthly cumulative attendance.</p>
    
    <div class="choice-grid cols-2" style="margin-bottom: 25px;">
      <div class="tilt-card advisor-card ${state.dept === 'aids' ? 'active-dept-card' : ''}" data-adv-dept="aids">
        <span class="tag">DEPT 01</span>
        <div class="card-icon" style="background: rgba(168, 85, 247, 0.2); color: #c084fc;">🤖</div>
        <h3 style="color: #c084fc;">AI & DS</h3>
        <p>Artificial Intelligence & Data Science</p>
      </div>

      <div class="tilt-card advisor-card ${state.dept === 'it' ? 'active-dept-card' : ''}" data-adv-dept="it">
        <span class="tag">DEPT 02</span>
        <div class="card-icon" style="background: rgba(168, 85, 247, 0.2); color: #c084fc;">💻</div>
        <h3 style="color: #c084fc;">IT</h3>
        <p>Information Technology</p>
      </div>
    </div>

    <div class="panel">
      <div class="panel-head" style="margin-bottom: 15px;">
        <div style="font-weight:700; color:#c084fc; font-family:'Space Grotesk',sans-serif;">
          ${getDeptShort()} · Select Target Year
        </div>
      </div>
      <div class="choice-grid cols-4">
        ${AVAILABLE_YEARS.map(y => `
          <div class="tilt-card advisor-card" data-year="${y}">
            <span class="tag">YR ${y}</span>
            <div class="year-badge" style="color: #c084fc; border-color: rgba(192, 132, 252, 0.4);">0${y}</div>
            <h3 style="color: #c084fc;">${getYearLabel(y)} Year</h3>
            <p>Daily period logs and eligibility stats.</p>
          </div>
        `).join('')}
      </div>
    </div>
  `;

  el.querySelectorAll('[data-adv-dept]').forEach(card => {
    attachTilt(card);
    card.onclick = () => {
      state.dept = card.dataset.advDept;
      renderAdvisorYearSelect(el);
    };
  });

  el.querySelectorAll('[data-year]').forEach(card => {
    attachTilt(card);
    card.onclick = () => {
      state.year = parseInt(card.dataset.year);
      state.view = 'advisorDashboard';
      render();
    };
  });
}

/* 4. CLASS ADVISOR DASHBOARD */
function renderAdvisorDashboard(el) {
  const today = new Date().toISOString().slice(0, 10);
  const currentMonth = today.slice(5, 7);
  const currentYear = today.slice(0, 4);

  el.innerHTML = `
    <div class="eyebrow">Class Advisor · ${getDeptShort()} Department</div>
    <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom: 15px;">
      <h1 class="title" style="margin:0;">Year ${state.year} Analytics</h1>
      <div style="display:flex; gap:10px;">
        <button class="btn-solid" id="tabDailyBtn">📅 Daily Attendance</button>
        <button class="btn-outline" id="tabMonthlyBtn">📈 Monthly Cumulative %</button>
      </div>
    </div>

    <!-- DAILY ATTENDANCE TAB -->
    <div class="panel" id="advisorDailyPanel">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom: 20px; flex-wrap:wrap; gap:10px;">
        <div style="display:flex; align-items:center; gap:12px; flex-wrap:wrap;">
          <label style="font-family:'JetBrains Mono'; font-size:12px; color:var(--text-dimmer);">SELECT DATE:</label>
          <input type="date" id="advDailyDate" value="${today}" class="modal-input" style="width: auto; padding: 8px 12px;" />
          <button class="btn-solid" id="fetchDailyBtn">Fetch Day Report</button>
        </div>
        <button class="btn-outline" id="exportDailyExcelBtn" style="border-color: #22c55e; color: #22c55e; font-weight:600;">
          📥 Download Daily Excel (.xlsx)
        </button>
      </div>
      <div id="advDailyTableContainer">
        <div class="empty-state"><div class="glyph">⟳</div><p>Loading Daily Attendance...</p></div>
      </div>
    </div>

    <!-- MONTHLY CUMULATIVE TAB -->
    <div class="panel" id="advisorMonthlyPanel" style="display:none;">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom: 20px; flex-wrap:wrap; gap:10px;">
        <div style="display:flex; align-items:center; gap:12px; flex-wrap:wrap;">
          <label style="font-family:'JetBrains Mono'; font-size:12px; color:var(--text-dimmer);">SELECT MONTH & YEAR:</label>
          <input type="month" id="advMonthInput" value="${currentYear}-${currentMonth}" class="modal-input" style="width: auto; padding: 8px 12px;" />
          <button class="btn-solid" id="fetchMonthlyBtn">Calculate Cumulative %</button>
        </div>
        <button class="btn-outline" id="exportMonthlyExcelBtn" style="border-color: #22c55e; color: #22c55e; font-weight:600;">
          📥 Download Monthly Excel (.xlsx)
        </button>
      </div>
      <div id="advMonthlyTableContainer">
        <div class="empty-state"><div class="glyph">⟳</div><p>Select month to calculate cumulative attendance...</p></div>
      </div>
    </div>
  `;

  const dailyPanel = document.getElementById('advisorDailyPanel');
  const monthlyPanel = document.getElementById('advisorMonthlyPanel');
  const tabDailyBtn = document.getElementById('tabDailyBtn');
  const tabMonthlyBtn = document.getElementById('tabMonthlyBtn');

  tabDailyBtn.onclick = () => {
    dailyPanel.style.display = 'block';
    monthlyPanel.style.display = 'none';
    tabDailyBtn.className = 'btn-solid';
    tabMonthlyBtn.className = 'btn-outline';
    loadAdvisorDailyData();
  };

  tabMonthlyBtn.onclick = () => {
    dailyPanel.style.display = 'none';
    monthlyPanel.style.display = 'block';
    tabDailyBtn.className = 'btn-outline';
    tabMonthlyBtn.className = 'btn-solid';
    loadAdvisorMonthlyData();
  };

  document.getElementById('fetchDailyBtn').onclick = loadAdvisorDailyData;
  document.getElementById('fetchMonthlyBtn').onclick = loadAdvisorMonthlyData;
  document.getElementById('advDailyDate').onchange = loadAdvisorDailyData;
  document.getElementById('advMonthInput').onchange = loadAdvisorMonthlyData;

  loadAdvisorDailyData();
}

async function loadAdvisorDailyData() {
  const date = document.getElementById('advDailyDate').value;
  const container = document.getElementById('advDailyTableContainer');
  const exportBtn = document.getElementById('exportDailyExcelBtn');
  container.innerHTML = `<div class="empty-state"><div class="glyph">⟳</div><p>Loading records for ${date}...</p></div>`;

  try {
    const { students: currentStudentList } = await loadYearStudents(state.year, state.dept);
    
    if (!currentStudentList || currentStudentList.length === 0) {
      container.innerHTML = `
        <div class="empty-state">
          <div class="glyph">👥</div>
          <h3>No Students Added for ${getDeptShort()} Year ${state.year}</h3>
          <p>Please go to Admin Console &rarr; "Manage Students & PDF" to add student names.</p>
        </div>
      `;
      exportBtn.onclick = () => showToast('No students enrolled to export.');
      return;
    }

    const targetKey = getYearKey(state.dept, state.year);
    const res = await secureFetch(`${API_BASE}/api/advisor/daily/${targetKey}?date=${date}`);
    const data = await res.json();
    const periods = data.periodsData || {};

    const studentMap = {};
    currentStudentList.forEach(s => {
      studentMap[s.reg] = { name: s.name, periods: {} };
      for (let p = 1; p <= NUM_PERIODS; p++) studentMap[s.reg].periods[`P${p}`] = '-';
    });

    for (let p = 1; p <= NUM_PERIODS; p++) {
      const pPayload = periods[`P${p}`];
      if (pPayload) {
        let recs = pPayload.records || pPayload;
        if (typeof recs === 'object' && !Array.isArray(recs)) {
          Object.keys(recs).forEach(reg => {
            if (studentMap[reg]) {
              const val = recs[reg];
              const st = typeof val === 'object' ? val.status : val;
              studentMap[reg].periods[`P${p}`] = (st || '-').toUpperCase();
            }
          });
        }
      }
    }

    exportBtn.onclick = () => exportDailyToExcel(state.dept, state.year, date, currentStudentList, studentMap);

    let html = `
      <div style="overflow-x:auto;">
        <table class="report-table">
          <thead>
            <tr>
              <th style="width: 50px;">Sl</th>
              <th>Student Name</th>
              <th>Reg No</th>
              ${Array.from({ length: NUM_PERIODS }, (_, i) => `<th style="text-align:center;">P${i + 1}<br><span style="font-size:9px; color:var(--text-dimmer);">${PERIOD_TIMES[i]}</span></th>`).join('')}
            </tr>
          </thead>
          <tbody>
            ${currentStudentList.map((s, idx) => `
              <tr>
                <td style="color:var(--text-dimmer); font-family:'JetBrains Mono';">${String(idx + 1).padStart(2, '0')}</td>
                <td style="font-weight:600; color:var(--text);">${s.name}</td>
                <td style="font-family:'JetBrains Mono'; font-size:12px; color:var(--text-dim);">${s.reg}</td>
                ${Array.from({ length: NUM_PERIODS }, (_, i) => {
                  const st = studentMap[s.reg]?.periods[`P${i + 1}`] || '-';
                  let color = 'var(--text-dimmer)';
                  if (st === 'PRESENT') color = 'var(--emerald, #22c55e)';
                  else if (st === 'ABSENT') color = 'var(--rose, #ef4444)';
                  else if (st === 'OD') color = 'var(--amber, #f97316)';
                  return `<td style="text-align:center; font-weight:700; color:${color}; font-family:'JetBrains Mono';">${st}</td>`;
                }).join('')}
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    `;
    container.innerHTML = html;
  } catch (err) {
    container.innerHTML = `<div class="empty-state"><p style="color:var(--rose);">Error loading daily report: ${err.message}</p></div>`;
  }
}

async function loadAdvisorMonthlyData() {
  const monthVal = document.getElementById('advMonthInput').value;
  if (!monthVal) return;
  const [fullYear, month] = monthVal.split('-');

  const container = document.getElementById('advMonthlyTableContainer');
  const exportBtn = document.getElementById('exportMonthlyExcelBtn');
  container.innerHTML = `<div class="empty-state"><div class="glyph">⟳</div><p>Calculating monthly cumulative attendance for ${monthVal}...</p></div>`;

  try {
    const targetKey = getYearKey(state.dept, state.year);
    const res = await secureFetch(`${API_BASE}/api/attendance/cumulative/${targetKey}?fullYear=${fullYear}&month=${month}`);
    const data = await res.json();
    const list = data.studentsReport || [];

    if (list.length === 0) {
      container.innerHTML = `<div class="empty-state"><p style="color:var(--amber);">No attendance records found for ${monthVal}.</p></div>`;
      exportBtn.onclick = () => showToast('No records found for this month.');
      return;
    }

    exportBtn.onclick = () => exportMonthlyToExcel(state.dept, state.year, monthVal, list);

    let html = `
      <div style="overflow-x:auto;">
        <table class="report-table">
          <thead>
            <tr>
              <th style="width: 50px;">Sl</th>
              <th>Student Name</th>
              <th style="text-align:center;">Total Hours</th>
              <th style="text-align:center; color: #22c55e;">Present</th>
              <th style="text-align:center; color: #f97316;">OD</th>
              <th style="text-align:center; color: #ef4444;">Absent</th>
              <th style="text-align:center;">Cumulative %</th>
              <th style="text-align:center;">Eligibility (<75%)</th>
            </tr>
          </thead>
          <tbody>
            ${list.map((st, idx) => `
              <tr>
                <td style="color:var(--text-dimmer); font-family:'JetBrains Mono';">${String(idx + 1).padStart(2, '0')}</td>
                <td style="font-weight:600; color:var(--text);">${st.studentName}</td>
                <td style="text-align:center; font-family:'JetBrains Mono';">${st.totalPeriods}</td>
                <td style="text-align:center; font-weight:600; color:#22c55e; font-family:'JetBrains Mono';">${st.present}</td>
                <td style="text-align:center; font-weight:600; color:#f97316; font-family:'JetBrains Mono';">${st.od}</td>
                <td style="text-align:center; font-weight:600; color:#ef4444; font-family:'JetBrains Mono';">${st.absent}</td>
                <td style="text-align:center; font-size:14px; font-weight:700; color:${st.percentage >= 75 ? 'var(--cyan, #38bdf8)' : '#ef4444'}; font-family:'JetBrains Mono';">
                  ${st.percentage}%
                </td>
                <td style="text-align:center;">
                  <span class="badge ${st.isShortage ? 'badge-danger' : 'badge-success'}">
                    ${st.isShortage ? 'Shortage (<75%)' : 'Eligible'}
                  </span>
                </td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    `;
    container.innerHTML = html;
  } catch (err) {
    container.innerHTML = `<div class="empty-state"><p style="color:var(--rose);">Error calculating monthly report: ${err.message}</p></div>`;
  }
}

/* 5. STAFF PROFILE SELECTION */
async function renderStaffSelect(el) {
  el.innerHTML = `<div class="panel"><div class="empty-state"><div class="glyph">⟳</div><p>Fetching ${getDeptShort()} Year ${state.year} Faculty...</p></div></div>`;
  
  const { grid, pdf } = await loadTimetable(state.year, state.dept);
  
  const staffMap = {};
  if (grid) {
    grid.forEach(day => {
      day.periods.forEach(p => {
        if (p.staff && p.staff.trim()) {
          const sName = p.staff.trim();
          if (!staffMap[sName]) staffMap[sName] = [];
          if (p.subject && !staffMap[sName].includes(p.subject.trim())) {
            staffMap[sName].push(p.subject.trim());
          }
        }
      });
    });
  }

  const staffList = Object.keys(staffMap);

  el.innerHTML = `
    <div class="eyebrow">${getDeptShort()} · Year ${state.year} Faculty Portal</div>
    <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom: 8px;">
      <h1 class="title" style="margin:0;">Select Your Profile</h1>
      <div style="display:flex; gap:10px;">
        ${pdf ? `<a href="${pdf.url}" target="_blank" class="btn-outline" style="text-decoration:none;">📄 View Timetable PDF</a>` : ''}
        <button class="btn-outline" id="viewFullTTBtn">View Timetable Grid</button>
      </div>
    </div>
    <p class="lede">Click your faculty profile card to log in and take class attendance.</p>

    ${staffList.length === 0 ? `
      <div class="panel">
        <div class="empty-state">
          <div class="glyph">👨‍🏫</div>
          <h3>No Faculty Mapped for ${getDeptShort()} Year ${state.year}</h3>
          <p>Please contact admin to upload timetable slots and assign staff members.</p>
          <button class="btn-solid" onclick="state.view='year'; render();">Back to Years</button>
        </div>
      </div>
    ` : `
      <div class="choice-grid cols-3" id="staffGrid">
        ${staffList.map((staff) => `
          <div class="tilt-card" data-staff="${staff}">
            <div class="card-icon" style="background: linear-gradient(135deg, var(--cyan), var(--violet)); color:#04070f; font-weight:700;">
              ${initials(staff)}
            </div>
            <h3>${staff}</h3>
            <p style="color:var(--cyan); margin-bottom:4px;">${staffMap[staff].join(', ') || 'Faculty'}</p>
            <p style="font-size:11.5px; color:var(--text-dimmer);">🔒 Password Protected</p>
          </div>
        `).join('')}
      </div>
    `}
  `;

  document.getElementById('viewFullTTBtn').onclick = () => {
    state.view = 'timetable';
    render();
  };

  [...el.querySelectorAll('[data-staff]')].forEach(card => {
    attachTilt(card);
    card.onclick = () => openStaffAuthModal(card.dataset.staff);
  });
}

/* MODAL FOR STAFF PASSWORD AUTHENTICATION */
async function openStaffAuthModal(staffName) {
  const oldModal = document.getElementById('authModal');
  if (oldModal) oldModal.remove();

  const checkRes = await staffLoginAPI(staffName, '');
  const isFirstTime = checkRes.requiresSetup;

  const modalHtml = `
    <div class="modal-overlay" id="authModal">
      <div class="modal-card">
        <div style="display:flex; align-items:center; gap:12px;">
          <div class="student-avatar" style="width:36px; height:36px;">${initials(staffName)}</div>
          <div>
            <h3 class="modal-title">${staffName}</h3>
            <p class="modal-sub">${isFirstTime ? 'First Time Setup · Create Password' : 'Enter Password to Login'}</p>
          </div>
        </div>
        
        <div style="margin-top: 15px;">
          <label style="font-size:11px; font-family:'JetBrains Mono'; color:var(--text-dimmer); display:block; margin-bottom:6px;">
            ${isFirstTime ? 'SET NEW PASSWORD / PIN' : 'ENTER PASSWORD'}
          </label>
          <input type="password" id="modalPasswordInput" class="modal-input" placeholder="••••••••" autocomplete="off" />
        </div>

        <div id="modalError" style="font-size:12px; color:var(--rose); font-family:'JetBrains Mono'; display:none; margin-top:8px;"></div>

        <div class="modal-actions" style="margin-top:20px;">
          <button class="btn-outline" id="modalCancelBtn">Cancel</button>
          <button class="btn-solid" id="modalSubmitBtn">${isFirstTime ? 'Set & Login' : 'Login'}</button>
        </div>
      </div>
    </div>
  `;

  document.body.insertAdjacentHTML('beforeend', modalHtml);

  const pwdInput = document.getElementById('modalPasswordInput');
  const submitBtn = document.getElementById('modalSubmitBtn');
  const cancelBtn = document.getElementById('modalCancelBtn');
  const errorDiv = document.getElementById('modalError');

  pwdInput.focus();

  const handleAuth = async () => {
    const val = pwdInput.value.trim();
    if (!val) {
      errorDiv.textContent = 'Please enter password!';
      errorDiv.style.display = 'block';
      return;
    }

    submitBtn.textContent = 'Verifying…';
    submitBtn.disabled = true;

    if (isFirstTime) {
      const setRes = await staffLoginAPI(staffName, null, val);
      if (setRes.success) {
        document.getElementById('authModal').remove();
        showToast('Password configured successfully!');
        state.currentStaff = staffName;
        state.view = 'attendance';
        render();
      } else {
        errorDiv.textContent = setRes.error || 'Setup failed!';
        errorDiv.style.display = 'block';
        submitBtn.textContent = 'Set & Login';
        submitBtn.disabled = false;
      }
    } else {
      const loginRes = await staffLoginAPI(staffName, val);
      if (loginRes.success) {
        document.getElementById('authModal').remove();
        showToast(`Welcome back, ${staffName}!`);
        state.currentStaff = staffName;
        state.view = 'attendance';
        render();
      } else {
        errorDiv.textContent = '❌ Invalid Password! Please try again.';
        errorDiv.style.display = 'block';
        submitBtn.textContent = 'Login';
        submitBtn.disabled = false;
        pwdInput.value = '';
        pwdInput.focus();
      }
    }
  };

  submitBtn.onclick = handleAuth;
  pwdInput.onkeydown = (e) => { if (e.key === 'Enter') handleAuth(); };
  cancelBtn.onclick = () => document.getElementById('authModal').remove();
}

/* 6. ATTENDANCE SHEET & LIVE MARKING */
async function renderAttendance(el) {
  el.innerHTML = `<div class="panel"><div class="empty-state"><div class="glyph">⟳</div><p>Loading Attendance Console…</p></div></div>`;
  
  const { grid: tt } = await loadTimetable(state.year, state.dept);
  const { students: classStudents } = await loadYearStudents(state.year, state.dept);

  if (!tt) {
    el.innerHTML = `
      <div class="panel">
        <div class="empty-state">
          <h3>No Schedule Configured</h3>
          <button class="btn-solid" onclick="state.view='staffSelect'; render();">Back</button>
        </div>
      </div>`;
    return;
  }

  const { dayName, dateStr } = todayInfo();
  const activeDay = dayName || 'Mon';
  buildAttendanceUI(el, tt, classStudents, activeDay, dateStr, !dayName);
}

async function buildAttendanceUI(el, tt, classStudents, activeDay, dateStr, isFallbackDay) {
  const dayRow = tt.find(d => d.day === activeDay);
  const periods = dayRow ? dayRow.periods : [];
  const nowMins = (() => { const n = new Date(); return n.getHours() * 60 + n.getMinutes(); })();
  
  let currentIdx = periods.findIndex((p, i) => {
    if (!p.subject.trim()) return false;
    if (state.currentStaff && p.staff.trim().toLowerCase() !== state.currentStaff.toLowerCase()) return false;
    const [s, e] = parsePeriodRange(PERIOD_TIMES[i]);
    return nowMins >= s && nowMins <= e;
  });

  if (currentIdx === -1 && state.currentStaff) {
    currentIdx = periods.findIndex(p => p.staff.trim().toLowerCase() === state.currentStaff.toLowerCase());
  }
  if (currentIdx === -1) currentIdx = 0;

  el.innerHTML = `
    <div class="eyebrow">${getDeptShort()} · Year ${state.year} · Staff: ${state.currentStaff || 'Faculty'}</div>
    <h1 class="title">${activeDay === todayInfo().dayName ? "Today's" : activeDay + "'s"} Class Attendance</h1>
    <p class="lede">${isFallbackDay ? 'Showing periods for attendance.' : 'Select a period to mark students present/absent/OD.'}</p>
    <div class="panel">
      <div class="day-tabs" id="dayTabs">
        ${DAYS.map(d => `<button class="day-tab ${d === activeDay ? 'active' : ''}" data-day="${d}">${d}</button>`).join('')}
      </div>
      <div class="period-tabs" id="periodTabs">
        ${periods.map((p, i) => `
          <button class="period-tab ${i === currentIdx ? 'active' : ''} ${!(p.subject || '').trim() ? 'free' : ''}" data-idx="${i}">
            <span class="pt-num">P${i + 1}</span>
            <span class="pt-time">${PERIOD_TIMES[i]}</span>
            <span class="pt-subj">${(p.subject || '').trim() || 'Free'}</span>
            <span style="font-size:9px; color:var(--text-dimmer);">${p.staff || ''}</span>
          </button>
        `).join('')}
      </div>
      <div id="attBody"></div>
    </div>
  `;

  [...el.querySelectorAll('.day-tab')].forEach(btn => {
    btn.onclick = () => buildAttendanceUI(el, tt, classStudents, btn.dataset.day, dateStr, false);
  });

  const periodBtns = [...el.querySelectorAll('.period-tab')];
  periodBtns.forEach(btn => {
    btn.onclick = () => {
      periodBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      loadPeriodAttendance(el, periods, classStudents, parseInt(btn.dataset.idx), dateStr, activeDay);
    };
  });

  if (periods.length) {
    loadPeriodAttendance(el, periods, classStudents, currentIdx, dateStr, activeDay);
  } else {
    document.getElementById('attBody').innerHTML = `<div class="empty-state"><p>No periods configured for ${activeDay}.</p></div>`;
  }
}

async function loadPeriodAttendance(el, periods, classStudents, idx, dateStr, activeDay) {
  const body = document.getElementById('attBody');
  const p = periods[idx];
  if (!p || !(p.subject || '').trim()) {
    body.innerHTML = `<div class="empty-state"><p>This is a free period — no attendance needed.</p></div>`;
    return;
  }

  if (!classStudents || classStudents.length === 0) {
    body.innerHTML = `
      <div class="empty-state">
        <div class="glyph">👥</div>
        <h3>No Students Enrolled</h3>
        <p>Contact admin to add students for ${getDeptShort()} Year ${state.year}.</p>
      </div>
    `;
    return;
  }

  body.innerHTML = `<div class="empty-state"><p>Loading students…</p></div>`;
  const existing = await loadAttendance(dateStr, idx, state.year, state.dept);
  const records = {};
  
  classStudents.forEach(s => { 
    if (existing && existing.records && existing.records[s.reg]) {
      records[s.reg] = existing.records[s.reg];
    } else {
      records[s.reg] = 'present'; 
    }
  });

  function summary() {
    const pCount = Object.values(records).filter(v => v === 'present').length;
    const aCount = Object.values(records).filter(v => v === 'absent').length;
    const oCount = Object.values(records).filter(v => v === 'od').length;
    return { pCount, aCount, oCount };
  }

  function paint() {
    const { pCount, aCount, oCount } = summary();
    body.innerHTML = `
      <div class="att-headrow">
        <div class="att-subj-info">
          <span class="subj-pill">${p.subject}</span>
          <span class="staff-pill">Faculty: ${p.staff || 'Unassigned'}</span>
          ${existing ? '<span class="saved-pill">Previously Saved</span>' : ''}
        </div>
        <div class="att-summary">
          <span class="sum-p">${pCount} Present</span>
          <span class="sum-a">${aCount} Absent</span>
          <span class="sum-o">${oCount} OD</span>
        </div>
      </div>
      <div class="quick-actions">
        <button class="btn-outline" id="allPresent">Mark All Present</button>
        <button class="btn-outline" id="allAbsent">Mark All Absent</button>
        <button class="btn-outline" id="allOD">Mark All OD</button>
      </div>
      <div class="student-att-list" id="studentAttList">
        ${classStudents.map(s => `
          <div class="att-row" data-reg="${s.reg}">
            <div class="att-who">
              <span class="student-avatar small">${initials(s.name)}</span>
              <div>
                <div class="student-name">${s.name}</div>
                <div class="student-reg">${s.reg}</div>
              </div>
            </div>
            <div class="att-toggle">
              <button class="pa-btn present ${records[s.reg] === 'present' ? 'active' : ''}" data-status="present">Present</button>
              <button class="pa-btn absent ${records[s.reg] === 'absent' ? 'active' : ''}" data-status="absent">Absent</button>
              <button class="pa-btn od ${records[s.reg] === 'od' ? 'active' : ''}" data-status="od">OD</button>
            </div>
          </div>
        `).join('')}
      </div>
      <div class="save-bar">
        <button class="btn-solid" id="saveAttBtn">Save Attendance & Send Email</button>
      </div>
    `;

    body.querySelectorAll('.att-row').forEach(row => {
      const reg = row.dataset.reg;
      row.querySelectorAll('.pa-btn').forEach(btn => {
        btn.onclick = () => {
          records[reg] = btn.dataset.status;
          row.querySelectorAll('.pa-btn').forEach(b => b.classList.remove('active'));
          btn.classList.add('active');
          const { pCount: pc, aCount: ac, oCount: oc } = summary();
          body.querySelector('.sum-p').textContent = pc + ' Present';
          body.querySelector('.sum-a').textContent = ac + ' Absent';
          body.querySelector('.sum-o').textContent = oc + ' OD';
        };
      });
    });

    document.getElementById('allPresent').onclick = () => { classStudents.forEach(s => records[s.reg] = 'present'); paint(); };
    document.getElementById('allAbsent').onclick = () => { classStudents.forEach(s => records[s.reg] = 'absent'); paint(); };
    document.getElementById('allOD').onclick = () => { classStudents.forEach(s => records[s.reg] = 'od'); paint(); };
    
    document.getElementById('saveAttBtn').onclick = async () => {
      const btn = document.getElementById('saveAttBtn');
      btn.textContent = 'Saving & Sending Email…'; 
      btn.disabled = true;

      const payloadStudentsList = classStudents.map(st => ({
        regNo: st.reg,
        name: st.name,
        status: (records[st.reg] || 'present').toUpperCase(),
        remarks: records[st.reg] === 'od' ? 'On Duty' : ''
      }));

      const ok = await saveAttendance(dateStr, idx, {
        day: activeDay,
        period: idx + 1,
        subject: p.subject,
        staff: p.staff,
        date: dateStr,
        records,
        students: payloadStudentsList,
        savedAt: new Date().toISOString()
      }, state.year, state.dept);

      btn.textContent = 'Save Attendance & Send Email'; 
      btn.disabled = false;
      showToast(ok ? `Attendance saved & Email sent for P${idx + 1} (${p.subject})` : 'Could not save.');
    };
  }
  paint();
}

/* 7. TIMETABLE GRID VIEW */
function renderTimetableView(el) {
  loadTimetable(state.year, state.dept).then(({ grid: data, pdf }) => {
    el.innerHTML = `
      <div class="eyebrow">${getDeptShort()} · Year ${state.year} Schedule</div>
      <h1 class="title">Timetable Grid</h1>
      ${pdf ? `<div style="margin-bottom:16px;"><a href="${pdf.url}" target="_blank" class="btn-solid" style="text-decoration:none;">📄 View / Download Timetable PDF</a></div>` : ''}
      <div class="panel">
        <div class="tt-grid">
          <div class="tt-cell tt-head">Day</div>
          ${PERIOD_TIMES.map((t, i) => `<div class="tt-cell tt-time">P${i + 1}<br>${t}</div>`).join('')}
          ${(data || emptyTimetable()).map(d => `
            <div class="tt-cell tt-head">${d.day}</div>
            ${d.periods.map(p => `
              <div class="tt-cell tt-slot ${(p.subject || '').trim() ? '' : 'empty'}">
                <div class="subj">${(p.subject || '').trim() || 'Free'}</div>
                ${(p.subject || '').trim() ? `<div class="staff"><span>${(p.staff || '').trim() || 'Unassigned'}</span></div>` : ''}
              </div>
            `).join('')}
          `).join('')}
        </div>
      </div>
    `;
    [...el.querySelectorAll('.tt-slot')].forEach(attachTilt);
  });
}

/* 8. ADMIN DASHBOARD */
function renderAdmin(el) {
  el.innerHTML = `
    <div class="eyebrow">Admin Console</div>
    <h1 class="title">Academic Management & Schedules</h1>
    <p class="lede">Manage timetable slots, uploaded PDFs, and department-wise student lists.</p>
    
    <!-- Admin Dept Switcher Tabs -->
    <div class="day-tabs" style="margin-bottom: 20px;">
      <button class="day-tab ${state.dept === 'aids' ? 'active' : ''}" id="adminDeptAidsBtn">🤖 AI & DS Department</button>
      <button class="day-tab ${state.dept === 'it' ? 'active' : ''}" id="adminDeptItBtn">💻 IT Department</button>
    </div>

    <div class="panel">
      <div class="panel-head" style="margin-bottom: 15px;">
        <div style="font-weight: 700; font-family:'Space Grotesk',sans-serif;">${getDeptLabel()} Classes</div>
      </div>
      <div class="admin-list" id="adminList">
        ${AVAILABLE_YEARS.map(y => `
          <div class="admin-row" data-year="${y}">
            <div class="l">
              <div class="admin-year">0${y}</div>
              <div>
                <div style="font-family:'Space Grotesk',sans-serif; font-weight:600; font-size:15px;">Year ${y} Management (${getDeptShort()})</div>
                <div class="admin-status status-pending" id="status-${y}">Checking status…</div>
              </div>
            </div>
            <div style="display:flex; gap:8px; flex-wrap:wrap;">
              <button class="btn-solid" data-add-tt="${y}">📅 Timetable Slots</button>
              <button class="btn-outline" data-manage-students="${y}">👥 Manage Students & PDF</button>
              <button class="btn-outline" data-view-tt="${y}">View Live</button>
            </div>
          </div>
        `).join('')}
      </div>
    </div>
  `;

  document.getElementById('adminDeptAidsBtn').onclick = () => { state.dept = 'aids'; renderAdmin(el); };
  document.getElementById('adminDeptItBtn').onclick = () => { state.dept = 'it'; renderAdmin(el); };

  AVAILABLE_YEARS.forEach(async y => {
    const res = await loadTimetable(y, state.dept);
    const { students: stList, pdf: stPdf } = await loadYearStudents(y, state.dept);
    const hasGrid = res.grid && res.grid.some(d => d.periods.some(p => (p.subject || '').trim()));
    const hasPdf = !!res.pdf;
    const s = document.getElementById(`status-${y}`);

    let statusText = [];
    if (hasGrid || hasPdf) statusText.push('Timetable Active');
    if (stList && stList.length > 0) statusText.push(`${stList.length} Students Added`);
    if (stPdf) statusText.push('Student PDF Attached');

    if (statusText.length > 0) {
      s.textContent = `● ${statusText.join(' · ')}`;
      s.className = 'admin-status status-live';
    } else {
      s.textContent = '○ Not configured yet (No students/timetable)';
      s.className = 'admin-status status-pending';
    }
  });

  el.querySelectorAll('[data-add-tt]').forEach(btn => {
    btn.onclick = () => { state.year = parseInt(btn.dataset.addTt); state.view = 'upload'; render(); };
  });
  el.querySelectorAll('[data-manage-students]').forEach(btn => {
    btn.onclick = () => { state.year = parseInt(btn.dataset.manageStudents); state.view = 'manageStudents'; render(); };
  });
  el.querySelectorAll('[data-view-tt]').forEach(btn => {
    btn.onclick = () => { state.year = parseInt(btn.dataset.viewTt); state.view = 'timetable'; render(); };
  });
}

/* 9. ADMIN TIMETABLE UPLOAD FORM (WITH PDF AUTO-PARSE POPULATION) */
function renderUpload(el, prefill) {
  if (!state.year) state.year = 1;

  loadTimetable(state.year, state.dept).then(({ grid: savedGrid, pdf }) => {
    const data = prefill || savedGrid || emptyTimetable();
    el.innerHTML = `
      <div class="eyebrow">Admin Mode · ${getDeptShort()} · Year ${state.year}</div>
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:12px;">
        <h1 class="title" style="margin:0;">Configure Year ${state.year} Timetable</h1>
        <div class="day-tabs" style="margin:0;">
          ${AVAILABLE_YEARS.map(y => `<button class="day-tab ${y === state.year ? 'active' : ''}" data-switch-year="${y}">Year ${y}</button>`).join('')}
        </div>
      </div>

      <div class="panel" style="margin-bottom: 24px;">
        <div class="panel-head">
          <div style="font-weight:600;">📄 Step 1: Upload Timetable PDF (Auto-Detects Periods & Staff)</div>
          ${pdf ? `<a href="${pdf.url}" target="_blank" class="btn-outline" style="text-decoration:none;">View PDF</a>` : ''}
        </div>
        <div style="border: 2px dashed var(--border); padding: 20px; border-radius: 12px; text-align: center;">
          <input type="file" id="pdfFileInput" accept="application/pdf" style="display:none;" />
          <div id="pdfSelectedName" style="font-size:13px; color:var(--cyan); margin-bottom:10px;">
            ${pdf ? `Current: ${pdf.originalName}` : 'No PDF attached'}
          </div>
          <button class="btn-outline" onclick="document.getElementById('pdfFileInput').click()">Choose File</button>
          <button class="btn-solid" id="uploadPdfBtn" style="margin-left: 8px;">Upload & Auto-Parse PDF</button>
        </div>
      </div>

      <div class="panel">
        <div class="panel-head"><div style="font-weight:600;">⏱️ Step 2: Staff & Subject Mapping</div></div>
        <div class="form-grid">
          <div></div>
          ${Array.from({ length: NUM_PERIODS }, (_, i) => `<div class="form-head">P${i + 1}<br><span style="font-size:9px; color:var(--text-dimmer);">${PERIOD_TIMES[i]}</span></div>`).join('')}
          ${data.map((d, di) => `
            <div class="form-day">${d.day}</div>
            ${d.periods.map((p, pi) => `
              <div class="form-cell">
                <input type="text" placeholder="Subject" value="${(p.subject || '').replace(/"/g, '&quot;')}" data-d="${di}" data-p="${pi}" data-f="subject">
                <input type="text" placeholder="Staff Name" value="${(p.staff || '').replace(/"/g, '&quot;')}" data-d="${di}" data-p="${pi}" data-f="staff">
              </div>
            `).join('')}
          `).join('')}
        </div>
        <div style="display:flex; gap:12px; justify-content:flex-end; margin-top: 15px;">
          <button class="btn-outline" id="clearBtn">Clear</button>
          <button class="btn-solid" id="saveGridBtn">Save Timetable Slots</button>
        </div>
      </div>
    `;

    el.querySelectorAll('[data-switch-year]').forEach(btn => {
      btn.onclick = () => { state.year = parseInt(btn.dataset.switchYear); renderUpload(el); };
    });

    const pdfFileInput = document.getElementById('pdfFileInput');
    pdfFileInput.onchange = () => {
      if (pdfFileInput.files[0]) document.getElementById('pdfSelectedName').textContent = `Ready: ${pdfFileInput.files[0].name}`;
    };

    // Auto-parse integration: updates grid slots immediately upon upload
    document.getElementById('uploadPdfBtn').onclick = async () => {
      const file = pdfFileInput.files[0];
      if (!file) return showToast('Please select a PDF file first!');
      const btn = document.getElementById('uploadPdfBtn');
      btn.textContent = 'Auto-Parsing PDF…'; 
      btn.disabled = true;

      const res = await uploadPdf(file, state.year, state.dept);
      btn.textContent = 'Upload & Auto-Parse PDF'; 
      btn.disabled = false;

      if (res && res.success) {
        showToast('PDF uploaded & Timetable auto-filled successfully!');
        renderUpload(el, res.parsedGrid);
      } else {
        showToast(res.error || 'Upload failed');
      }
    };

    const gridData = data;
    el.querySelectorAll('.form-cell input').forEach(inp => {
      inp.addEventListener('input', () => {
        const { d, p, f } = inp.dataset;
        gridData[d].periods[p][f] = inp.value;
      });
    });

    document.getElementById('clearBtn').onclick = () => renderUpload(el, emptyTimetable());
    document.getElementById('saveGridBtn').onclick = async () => {
      const btn = document.getElementById('saveGridBtn');
      btn.textContent = 'Saving…'; btn.disabled = true;
      const ok = await saveTimetable(gridData, state.year, state.dept);
      btn.textContent = 'Save Timetable Slots'; btn.disabled = false;
      if (ok) {
        showToast('Timetable saved successfully!');
        state.view = 'timetable';
        render();
      } else {
        showToast('Could not save schedule.');
      }
    };
  });
}

/* 10. ADMIN STUDENT ROSTER & STUDENT PDF MANAGEMENT */
async function renderManageStudents(el) {
  if (!state.year) state.year = 1;
  el.innerHTML = `<div class="panel"><div class="empty-state"><div class="glyph">⟳</div><p>Loading Students Roster...</p></div></div>`;

  const { students: loadedStudents, pdf: studentPdf } = await loadYearStudents(state.year, state.dept);
  let studentList = [...loadedStudents];

  function paintStudentView() {
    el.innerHTML = `
      <div class="eyebrow">Admin Mode · ${getDeptShort()} · Year ${state.year} Roster</div>
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:15px;">
        <h1 class="title" style="margin:0;">Manage Year ${state.year} Students</h1>
        <div class="day-tabs" style="margin:0;">
          ${AVAILABLE_YEARS.map(y => `<button class="day-tab ${y === state.year ? 'active' : ''}" data-switch-year="${y}">Year ${y}</button>`).join('')}
        </div>
      </div>

      <!-- 1. UPLOAD STUDENT LIST PDF -->
      <div class="panel" style="margin-bottom: 20px;">
        <div class="panel-head">
          <div style="font-weight:600;">📑 Step 1: Upload Student Namelist PDF (Optional)</div>
          ${studentPdf ? `<a href="${studentPdf.url}" target="_blank" class="btn-outline" style="text-decoration:none;">📄 View Uploaded PDF</a>` : ''}
        </div>
        <div style="border: 2px dashed var(--border); padding: 20px; border-radius: 12px; text-align: center;">
          <input type="file" id="stPdfFileInput" accept="application/pdf" style="display:none;" />
          <div id="stPdfSelectedName" style="font-size:13px; color:var(--cyan); margin-bottom:10px;">
            ${studentPdf ? `Current PDF: ${studentPdf.originalName}` : 'No Student PDF attached'}
          </div>
          <button class="btn-outline" onclick="document.getElementById('stPdfFileInput').click()">Choose Student PDF</button>
          <button class="btn-solid" id="uploadStPdfBtn" style="margin-left: 8px;">Upload Student PDF</button>
        </div>
      </div>

      <!-- 2. ADD STUDENT (SINGLE & BULK) -->
      <div class="panel" style="margin-bottom: 20px;">
        <div class="panel-head"><div style="font-weight:600;">➕ Step 2: Add Students (Manual / Bulk Paste)</div></div>
        
        <div style="display: grid; grid-template-columns: 1fr 1fr auto; gap: 12px; margin-bottom: 15px; align-items: flex-end;">
          <div>
            <label style="font-family:'JetBrains Mono'; font-size:11px; color:var(--text-dimmer); display:block; margin-bottom:4px;">REGISTER NUMBER</label>
            <input type="text" id="newRegInput" placeholder="e.g. 621823205001" class="modal-input" />
          </div>
          <div>
            <label style="font-family:'JetBrains Mono'; font-size:11px; color:var(--text-dimmer); display:block; margin-bottom:4px;">STUDENT NAME</label>
            <input type="text" id="newNameInput" placeholder="e.g. M.K.Aakash" class="modal-input" />
          </div>
          <button class="btn-solid" id="addSingleStudentBtn" style="height: 42px;">+ Add Student</button>
        </div>

        <details style="background: rgba(15, 23, 42, 0.6); padding: 12px; border-radius: 8px; border: 1px solid var(--border);">
          <summary style="cursor: pointer; font-weight: 600; color: var(--cyan); font-size: 13px;">📋 Bulk Add (Paste Multiple Lines)</summary>
          <p style="font-size: 12px; color: var(--text-dim); margin: 8px 0;">Paste one per line (Format: <code>RegNo, Name</code> or <code>RegNo Name</code>):</p>
          <textarea id="bulkStudentsInput" rows="4" class="modal-input" placeholder="621823205001, M.K.Aakash&#10;621823205002, R.Akilan"></textarea>
          <div style="margin-top: 8px; text-align: right;">
            <button class="btn-outline" id="processBulkBtn">Process Bulk List</button>
          </div>
        </details>
      </div>

      <!-- 3. STUDENT LIST TABLE -->
      <div class="panel">
        <div class="panel-head">
          <div style="font-weight:600;">👥 Current Enrolled Students (${studentList.length})</div>
          <button class="btn-solid" id="saveStudentsListBtn" style="background: #22c55e;">💾 Save Student Roster</button>
        </div>

        ${studentList.length === 0 ? `
          <div class="empty-state" style="padding: 24px 0;">
            <div class="glyph">👥</div>
            <h3>No Students Added Yet</h3>
            <p>Add students using the form above and click "Save Student Roster".</p>
          </div>
        ` : `
          <div style="overflow-x:auto;">
            <table class="report-table">
              <thead>
                <tr>
                  <th style="width: 50px;">Sl</th>
                  <th>Register Number</th>
                  <th>Student Name</th>
                  <th style="text-align:center; width: 80px;">Action</th>
                </tr>
              </thead>
              <tbody id="studentTableBody">
                ${studentList.map((st, i) => `
                  <tr>
                    <td style="color:var(--text-dimmer); font-family:'JetBrains Mono';">${String(i + 1).padStart(2, '0')}</td>
                    <td style="font-family:'JetBrains Mono'; font-weight:600; color:var(--text);">${st.reg}</td>
                    <td style="color:var(--text);">${st.name}</td>
                    <td style="text-align:center;">
                      <button class="btn-outline" data-del-idx="${i}" style="padding: 4px 8px; font-size: 11px; border-color: #ef4444; color: #ef4444;">Delete</button>
                    </td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        `}
      </div>
    `;

    el.querySelectorAll('[data-switch-year]').forEach(btn => {
      btn.onclick = () => { state.year = parseInt(btn.dataset.switchYear); renderManageStudents(el); };
    });

    const stPdfInput = document.getElementById('stPdfFileInput');
    stPdfInput.onchange = () => {
      if (stPdfInput.files[0]) document.getElementById('stPdfSelectedName').textContent = `Ready: ${stPdfInput.files[0].name}`;
    };

    document.getElementById('uploadStPdfBtn').onclick = async () => {
      const file = stPdfInput.files[0];
      if (!file) return showToast('Please select a student list PDF first!');
      const btn = document.getElementById('uploadStPdfBtn');
      btn.textContent = 'Uploading…'; btn.disabled = true;
      const res = await uploadStudentPdf(file, state.year, state.dept);
      btn.textContent = 'Upload Student PDF'; btn.disabled = false;
      if (res && res.success) {
        showToast('Student PDF uploaded successfully!');
        renderManageStudents(el);
      } else {
        showToast(res.error || 'Upload failed');
      }
    };

    document.getElementById('addSingleStudentBtn').onclick = () => {
      const reg = document.getElementById('newRegInput').value.trim();
      const name = document.getElementById('newNameInput').value.trim();
      if (!reg || !name) return showToast('Please enter both Register Number and Name!');
      studentList.push({ reg, name });
      paintStudentView();
      showToast(`Added ${name}! Click "Save Student Roster" to persist.`);
    };

    document.getElementById('processBulkBtn').onclick = () => {
      const text = document.getElementById('bulkStudentsInput').value.trim();
      if (!text) return showToast('Bulk input is empty!');
      const lines = text.split('\n');
      let count = 0;
      lines.forEach(l => {
        const parts = l.split(/[,\t]+/).map(s => s.trim()).filter(Boolean);
        if (parts.length >= 2) {
          studentList.push({ reg: parts[0], name: parts.slice(1).join(' ') });
          count++;
        }
      });
      paintStudentView();
      showToast(`Added ${count} students! Click "Save Student Roster".`);
    };

    el.querySelectorAll('[data-del-idx]').forEach(btn => {
      btn.onclick = () => {
        const idx = parseInt(btn.dataset.delIdx);
        studentList.splice(idx, 1);
        paintStudentView();
      };
    });

    document.getElementById('saveStudentsListBtn').onclick = async () => {
      const btn = document.getElementById('saveStudentsListBtn');
      btn.textContent = 'Saving…'; btn.disabled = true;
      const ok = await saveYearStudents(studentList, state.year, state.dept);
      btn.textContent = '💾 Save Student Roster'; btn.disabled = false;
      if (ok) {
        showToast(`${getDeptShort()} Year ${state.year} student roster saved to database!`);
      } else {
        showToast('Failed to save student roster.');
      }
    };
  }

  paintStudentView();
}

function renderStudents(el) {
  loadYearStudents(state.year, state.dept).then(({ students: list }) => {
    el.innerHTML = `
      <div class="eyebrow">${getDeptShort()} · Year ${state.year} Class Roster</div>
      <h1 class="title">Enrolled Students (${list.length})</h1>
      <div class="panel">
        ${list.length === 0 ? `
          <div class="empty-state">
            <div class="glyph">👥</div>
            <h3>No Students Enrolled</h3>
            <p>Admin has not configured students for ${getDeptShort()} Year ${state.year} yet.</p>
          </div>
        ` : `
          <div class="student-grid">
            ${list.map((s, i) => `
              <div class="tilt-card student-card">
                <div class="student-avatar">${initials(s.name)}</div>
                <div class="student-info">
                  <div class="student-name">${s.name}</div>
                  <div class="student-reg">${s.reg}</div>
                </div>
                <div class="student-sl">${String(i + 1).padStart(2, '0')}</div>
              </div>
            `).join('')}
          </div>
        `}
      </div>
    `;
    [...el.querySelectorAll('.student-card')].forEach(attachTilt);
  });
}

// Initial Boot
render();