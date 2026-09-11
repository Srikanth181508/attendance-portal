require('dotenv').config();

const express = require('express');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const cors = require('cors');
const mysql = require('mysql2/promise');
const { Resend } = require('resend');
const pdfParse = require('pdf-parse');

const app = express();
const PORT = process.env.PORT || 3000;
const UPLOAD_DIR = path.join(__dirname, 'uploads');

if (!fs.existsSync(UPLOAD_DIR)) {
  fs.mkdirSync(UPLOAD_DIR, { recursive: true });
}

// 0. Universal Key Normalizer
function getYearKey(y) {
  if (!y) return 'aids_year1';
  const str = String(y).toLowerCase().trim();

  if (str.startsWith('aids_') || str.startsWith('it_')) {
    return str;
  }

  const cleaned = str.replace(/[^0-9]/g, '');
  return cleaned ? `aids_year${cleaned}` : 'aids_year1';
}

// Helper: Format Display Title for UI and Emails
function getDisplayTitle(yearKey) {
  const isIT = yearKey.startsWith('it_');
  const yNum = yearKey.replace(/[^0-9]/g, '') || '1';
  return {
    dept: isIT ? 'Information Technology (IT)' : 'Artificial Intelligence & Data Science (AI & DS)',
    deptShort: isIT ? 'IT' : 'AI & DS',
    year: `${yNum}${yNum === '1' ? 'st' : yNum === '2' ? 'nd' : yNum === '3' ? 'rd' : 'th'} Year`
  };
}

// Helper: Safe JSON stringifier
function safeStringify(data) {
  if (data === null || data === undefined) return null;
  if (typeof data === 'string') return data;
  return JSON.stringify(data);
}

// Helper: Auto-extract timetable structure from PDF text
function parseTimetableText(rawText) {
  const DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const grid = DAYS.map(d => ({
    day: d,
    periods: Array.from({ length: 6 }, () => ({ subject: '', staff: '' }))
  }));

  if (!rawText || typeof rawText !== 'string') return grid;

  const lines = rawText.split('\n').map(l => l.trim()).filter(Boolean);

  DAYS.forEach(day => {
    const dayRegex = new RegExp(`^(${day}|${day.toUpperCase()}|${day.toLowerCase()})\\b`, 'i');
    const matchedLine = lines.find(l => dayRegex.test(l));

    if (matchedLine) {
      const tokens = matchedLine.replace(dayRegex, '').trim().split(/\s{2,}|\t|\|/);
      tokens.slice(0, 6).forEach((cell, idx) => {
        if (!cell) return;
        const match = cell.match(/(.*?)\s*[\/\(\-]\s*(.*?)[\)]?$/);
        const dayItem = grid.find(g => g.day === day);
        if (dayItem && dayItem.periods[idx]) {
          if (match) {
            dayItem.periods[idx].subject = match[1].trim();
            dayItem.periods[idx].staff = match[2].trim();
          } else {
            dayItem.periods[idx].subject = cell.trim();
          }
        }
      });
    }
  });

  return grid;
}

// 1. MySQL Connection Pool (Credentials via Environment Variable - GitHub Safe)
const DATABASE_URL = process.env.DATABASE_URL || 'mysql://root:@localhost:3306/attendance_db';
const parsedDbUrl = new URL(DATABASE_URL);
const db = mysql.createPool({
  host: parsedDbUrl.hostname,
  port: Number(parsedDbUrl.port) || 4000,
  user: decodeURIComponent(parsedDbUrl.username),
  password: decodeURIComponent(parsedDbUrl.password),
  database: parsedDbUrl.pathname.replace(/^\//, '') || 'test',
  ssl: {
    minVersion: 'TLSv1.2',
    rejectUnauthorized: true
  },
  waitForConnections: true,
  connectionLimit: 10,
  queueLimit: 0,
  connectTimeout: 60000,   
  enableKeepAlive: true,
  keepAliveInitialDelay: 10000
});

// 2. Email Configuration (Secure HTTPS API via Resend - GitHub Safe)
const RESEND_API_KEY = process.env.RESEND_API_KEY || '';
const resend = RESEND_API_KEY ? new Resend(RESEND_API_KEY) : null;

// Helper: Parse Students Status List
function parseAttendancePayload(payload) {
  let presentList = [];
  let absentList = [];
  let odList = [];

  let data = payload;
  if (typeof data === 'string') {
    try { data = JSON.parse(data); } catch (e) {}
  }

  if (data && !Array.isArray(data) && typeof data === 'object') {
    if (Array.isArray(data.students)) data = data.students;
    else if (Array.isArray(data.attendance)) data = data.attendance;
    else if (Array.isArray(data.records)) data = data.records;
    else if (Array.isArray(data.list)) data = data.list;
  }

  if (Array.isArray(data)) {
    data.forEach(item => {
      if (!item) return;
      const studentInfo = item.name || item.studentName || item.student_name || item.student_id || item.regNo || 'Student';
      const rawStatus = (item.status || item.attendance || '').toString().trim().toUpperCase();
      const reason = item.remarks || item.reason || 'On Duty';

      if (rawStatus === 'PRESENT' || rawStatus === 'P') presentList.push(studentInfo);
      else if (rawStatus === 'ABSENT' || rawStatus === 'A') absentList.push(studentInfo);
      else if (rawStatus === 'OD' || rawStatus === 'ON DUTY') odList.push(`${studentInfo} (${reason})`);
    });
  } else if (data && typeof data === 'object') {
    Object.keys(data).forEach(key => {
      const val = data[key];
      let status = '';
      let reason = 'On Duty';

      if (typeof val === 'string') status = val.toUpperCase().trim();
      else if (typeof val === 'object' && val !== null) {
        status = (val.status || val.state || '').toUpperCase().trim();
        reason = val.remarks || val.reason || 'On Duty';
      }

      if (status === 'PRESENT' || status === 'P') presentList.push(key);
      else if (status === 'ABSENT' || status === 'A') absentList.push(key);
      else if (status === 'OD' || status === 'ON DUTY') odList.push(`${key} (${reason})`);
    });
  }

  return { presentList, absentList, odList };
}

// Helper: Send Attendance Email via Resend HTTPS REST API (Port 443)
async function sendAttendanceEmail(yearKey, dateStr, periodIdx, presentList, absentList, odList) {
  if (!resend || !RESEND_API_KEY) {
    console.warn('⚠️ Resend API Key is missing. Skipping email notification.');
    return;
  }

  const { dept, deptShort, year } = getDisplayTitle(yearKey);

  try {
    const htmlContent = `
      <div style="font-family: Arial, sans-serif; padding: 22px; border: 1px solid #e2e8f0; border-radius: 8px; max-width: 650px;">
        <h2 style="color: #1e3a8a; margin-top: 0;">Attendance Submission Report</h2>
        <div style="background-color: #f1f5f9; padding: 12px 16px; border-radius: 6px; margin-bottom: 15px;">
          <p style="margin: 4px 0;"><strong>Department:</strong> ${dept}</p>
          <p style="margin: 4px 0;"><strong>Academic Year:</strong> ${year}</p>
          <p style="margin: 4px 0;"><strong>Date:</strong> ${dateStr}</p>
          <p style="margin: 4px 0;"><strong>Period:</strong> Period ${periodIdx + 1}</p>
        </div>
        
        <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
          <thead>
            <tr style="background-color: #f8fafc;">
              <th style="padding: 10px; border: 1px solid #cbd5e1; text-align: left;">Status</th>
              <th style="padding: 10px; border: 1px solid #cbd5e1; text-align: center; width: 70px;">Count</th>
              <th style="padding: 10px; border: 1px solid #cbd5e1; text-align: left;">Students List</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td style="padding: 10px; border: 1px solid #cbd5e1; color: #16a34a; font-weight: bold;">PRESENT</td>
              <td style="padding: 10px; border: 1px solid #cbd5e1; text-align: center; font-weight: bold;">${presentList.length}</td>
              <td style="padding: 10px; border: 1px solid #cbd5e1;">${presentList.join(', ') || 'None'}</td>
            </tr>
            <tr>
              <td style="padding: 10px; border: 1px solid #cbd5e1; color: #dc2626; font-weight: bold;">ABSENT</td>
              <td style="padding: 10px; border: 1px solid #cbd5e1; text-align: center; font-weight: bold; color: #dc2626;">${absentList.length}</td>
              <td style="padding: 10px; border: 1px solid #cbd5e1;">${absentList.join(', ') || 'None'}</td>
            </tr>
            <tr>
              <td style="padding: 10px; border: 1px solid #cbd5e1; color: #ea580c; font-weight: bold;">OD (On Duty)</td>
              <td style="padding: 10px; border: 1px solid #cbd5e1; text-align: center; font-weight: bold; color: #ea580c;">${odList.length}</td>
              <td style="padding: 10px; border: 1px solid #cbd5e1;">${odList.join(', ') || 'None'}</td>
            </tr>
          </tbody>
        </table>
        <p style="margin-top: 20px; font-size: 12px; color: #64748b;">Automated attendance notification delivered via secure cloud API.</p>
      </div>
    `;

    const { data, error } = await resend.emails.send({
      from: 'Attendance Portal <onboarding@resend.dev>',
      to: 'itaids2327@gmail.com',
      subject: `Attendance Report | ${deptShort} ${year} | Date: ${dateStr} | Period: ${periodIdx + 1}`,
      html: htmlContent
    });

    if (error) {
      console.error('📧 Resend API Error:', error.message);
    } else {
      console.log('📧 Attendance Email sent successfully via Secure HTTPS API! ID:', data.id);
    }
  } catch (err) {
    console.error('📧 Resend Exception:', err.message);
  }
}

// 3. Auto-create Database Tables
async function initDB() {
  try {
    const connection = await db.getConnection();
    
    await connection.query(`
      CREATE TABLE IF NOT EXISTS timetables (
        year_key VARCHAR(50) PRIMARY KEY,
        grid_data JSON,
        pdf_info JSON,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
      )
    `);

    await connection.query(`
      CREATE TABLE IF NOT EXISTS year_students (
        year_key VARCHAR(50) PRIMARY KEY,
        students_data JSON,
        pdf_info JSON,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
      )
    `);

    await connection.query(`
      CREATE TABLE IF NOT EXISTS staff_auth (
        staff_key VARCHAR(100) PRIMARY KEY,
        password VARCHAR(255) NOT NULL,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
      )
    `);

    await connection.query(`
      CREATE TABLE IF NOT EXISTS attendance_records (
        attendance_key VARCHAR(100) PRIMARY KEY,
        year_key VARCHAR(50) NOT NULL,
        date_str VARCHAR(20) NOT NULL,
        period_idx INT NOT NULL,
        payload JSON NOT NULL,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
      )
    `);

    connection.release();
    console.log(`TiDB Cloud MySQL Database & Tables initialized successfully!`);
  } catch (err) {
    console.error('MySQL Connection Error:', err.message);
  }
}
initDB();

// 4. Multer Storage Configuration
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOAD_DIR),
  filename: (req, file, cb) => {
    const yearKey = getYearKey(req.params.year);
    const safeName = file.originalname.replace(/[^a-zA-Z0-9.]/g, '_');
    cb(null, `${file.fieldname}-${yearKey}-${Date.now()}-${safeName}`);
  }
});

const upload = multer({
  storage,
  limits: { fileSize: 15 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (file.mimetype === 'application/pdf' || ext === '.pdf') {
      cb(null, true);
    } else {
      cb(new Error('Only PDF files are allowed!'));
    }
  }
});

// Middleware configuration
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));
app.use('/uploads', express.static(UPLOAD_DIR));

// ==========================================
// 5. Timetable Endpoints
// ==========================================
app.get(['/api/timetable/:year', '/api/timetable/:year/'], async (req, res) => {
  try {
    const yearKey = getYearKey(req.params.year);
    const [rows] = await db.query('SELECT grid_data, pdf_info FROM timetables WHERE year_key = ?', [yearKey]);
    if (rows.length > 0) {
      let grid = rows[0].grid_data;
      let pdf = rows[0].pdf_info;
      while (typeof grid === 'string') { try { grid = JSON.parse(grid); } catch (e) { break; } }
      while (typeof pdf === 'string') { try { pdf = JSON.parse(pdf); } catch (e) { break; } }
      res.json({ grid: grid || null, pdf: pdf || null });
    } else {
      res.json({ grid: null, pdf: null });
    }
  } catch (err) { 
    console.error('Fetch Timetable Error:', err.message);
    res.status(500).json({ error: err.message }); 
  }
});

app.post(['/api/timetable/:year', '/api/timetable/:year/'], async (req, res) => {
  try {
    const yearKey = getYearKey(req.params.year);
    const gridPayload = safeStringify(req.body);

    await db.query(
      `INSERT INTO timetables (year_key, grid_data) 
       VALUES (?, ?) 
       ON DUPLICATE KEY UPDATE grid_data = ?`,
      [yearKey, gridPayload, gridPayload]
    );

    console.log(`✅ Timetable slots saved successfully for: ${yearKey}`);
    res.json({ success: true, message: 'Timetable grid saved successfully!' });
  } catch (err) { 
    console.error('❌ Save Timetable Error:', err.message);
    res.status(500).json({ success: false, error: err.message }); 
  }
});

// Auto-Parse & Auto-Save Timetable PDF Endpoint
app.post(['/api/timetable/pdf/:year', '/api/timetable/pdf/:year/'], (req, res) => {
  upload.single('pdf')(req, res, async (err) => {
    if (err) {
      console.error('Multer Upload Error:', err.message);
      return res.status(400).json({ success: false, error: err.message });
    }
    if (!req.file) {
      return res.status(400).json({ success: false, error: 'No PDF provided' });
    }

    try {
      const yearKey = getYearKey(req.params.year);
      const filePath = req.file.path;

      const pdfInfo = {
        originalName: req.file.originalname,
        filename: req.file.filename,
        url: `/uploads/${req.file.filename}`,
        uploadedAt: new Date().toISOString()
      };

      let parsedGrid = null;

      try {
        const dataBuffer = fs.readFileSync(filePath);
        const pdfData = await pdfParse(dataBuffer);
        if (pdfData && pdfData.text) {
          parsedGrid = parseTimetableText(pdfData.text);
        }
      } catch (parseErr) {
        console.warn('⚠️ PDF text parsing warning (Fallback to manual slots):', parseErr.message);
      }

      if (parsedGrid) {
        const gridJson = safeStringify(parsedGrid);
        const pdfJson = JSON.stringify(pdfInfo);
        await db.query(
          `INSERT INTO timetables (year_key, pdf_info, grid_data) 
           VALUES (?, ?, ?) 
           ON DUPLICATE KEY UPDATE pdf_info = ?, grid_data = ?`,
          [yearKey, pdfJson, gridJson, pdfJson, gridJson]
        );
      } else {
        const pdfJson = JSON.stringify(pdfInfo);
        await db.query(
          `INSERT INTO timetables (year_key, pdf_info) 
           VALUES (?, ?) 
           ON DUPLICATE KEY UPDATE pdf_info = ?`,
          [yearKey, pdfJson, pdfJson]
        );
      }

      console.log(`✅ Timetable PDF processed successfully for ${yearKey}`);
      res.json({ success: true, pdf: pdfInfo, parsedGrid });
    } catch (dbErr) { 
      console.error('❌ Database save error on PDF upload:', dbErr.message);
      res.status(500).json({ success: false, error: dbErr.message }); 
    }
  });
});

// ==========================================
// 6. Year Student Management Endpoints
// ==========================================
app.get(['/api/students/:year', '/api/students/:year/'], async (req, res) => {
  try {
    const yearKey = getYearKey(req.params.year);
    const [rows] = await db.query('SELECT students_data, pdf_info FROM year_students WHERE year_key = ?', [yearKey]);
    if (rows.length > 0) {
      let students = rows[0].students_data;
      let pdf = rows[0].pdf_info;
      while (typeof students === 'string') { try { students = JSON.parse(students); } catch (e) { break; } }
      while (typeof pdf === 'string') { try { pdf = JSON.parse(pdf); } catch (e) { break; } }
      res.json({ students: students || [], pdf: pdf || null });
    } else {
      res.json({ students: [], pdf: null });
    }
  } catch (err) { 
    console.error('Fetch Students Error:', err.message);
    res.status(500).json({ error: err.message }); 
  }
});

app.post(['/api/students/:year', '/api/students/:year/'], async (req, res) => {
  try {
    const yearKey = getYearKey(req.params.year);
    const rawData = req.body && req.body.students !== undefined ? req.body.students : req.body;
    const studentsPayload = safeStringify(rawData);

    await db.query(
      `INSERT INTO year_students (year_key, students_data) 
       VALUES (?, ?) 
       ON DUPLICATE KEY UPDATE students_data = ?`,
      [yearKey, studentsPayload, studentsPayload]
    );

    const count = Array.isArray(rawData) ? rawData.length : 0;
    console.log(`✅ Student roster saved successfully for: ${yearKey} (${count} students)`);
    res.json({ success: true, message: 'Student roster saved successfully!' });
  } catch (err) { 
    console.error('❌ Save Students Error:', err.message);
    res.status(500).json({ success: false, error: err.message }); 
  }
});

app.post(['/api/students/pdf/:year', '/api/students/pdf/:year/'], (req, res) => {
  upload.single('pdf')(req, res, async (err) => {
    if (err) return res.status(400).json({ success: false, error: err.message });
    if (!req.file) return res.status(400).json({ success: false, error: 'No PDF provided' });

    try {
      const yearKey = getYearKey(req.params.year);
      const pdfInfo = {
        originalName: req.file.originalname,
        filename: req.file.filename,
        url: `/uploads/${req.file.filename}`,
        uploadedAt: new Date().toISOString()
      };
      const pdfJson = JSON.stringify(pdfInfo);

      await db.query(
        `INSERT INTO year_students (year_key, pdf_info) 
         VALUES (?, ?) 
         ON DUPLICATE KEY UPDATE pdf_info = ?`,
        [yearKey, pdfJson, pdfJson]
      );
      res.json({ success: true, pdf: pdfInfo });
    } catch (dbErr) { 
      res.status(500).json({ success: false, error: dbErr.message }); 
    }
  });
});

// ==========================================
// 7. Staff Authentication
// ==========================================
app.post(['/api/staff/login', '/api/staff/login/'], async (req, res) => {
  try {
    const { staffName, password, newPassword } = req.body;
    if (!staffName) return res.status(400).json({ success: false, error: 'Staff name is required' });

    const key = staffName.trim().toLowerCase();
    const [rows] = await db.query('SELECT password FROM staff_auth WHERE staff_key = ?', [key]);

    if (rows.length === 0) {
      if (newPassword) {
        await db.query('INSERT INTO staff_auth (staff_key, password) VALUES (?, ?)', [key, newPassword]);
        return res.json({ success: true, message: 'Password created successfully' });
      }
      return res.json({ requiresSetup: true, message: 'Create your access PIN/Password' });
    }

    if (rows[0].password === password) {
      return res.json({ success: true });
    } else {
      return res.status(401).json({ success: false, error: 'Invalid Staff Password' });
    }
  } catch (err) { res.status(500).json({ success: false, error: err.message }); }
});

// ==========================================
// 8. Attendance Endpoints (Email Notification)
// ==========================================
app.get(['/api/attendance/:year', '/api/attendance/:year/'], async (req, res) => {
  try {
    const yearKey = getYearKey(req.params.year);
    const { date, period } = req.query;
    const key = `attendance:${yearKey}:${date}:P${period}`;
    const [rows] = await db.query('SELECT payload FROM attendance_records WHERE attendance_key = ?', [key]);
    if (rows.length > 0) {
      let payloadData = rows[0].payload;
      while (typeof payloadData === 'string') { try { payloadData = JSON.parse(payloadData); } catch (e) { break; } }
      res.json(payloadData);
    } else {
      res.json(null);
    }
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post(['/api/attendance/:year', '/api/attendance/:year/'], async (req, res) => {
  try {
    const yearKey = getYearKey(req.params.year);
    const { dateStr, periodIdx, payload } = req.body;
    const key = `attendance:${yearKey}:${dateStr}:P${periodIdx + 1}`;
    const recordPayload = safeStringify(payload);

    await db.query(
      `INSERT INTO attendance_records (attendance_key, year_key, date_str, period_idx, payload)
       VALUES (?, ?, ?, ?, ?)
       ON DUPLICATE KEY UPDATE payload = ?`,
      [key, yearKey, dateStr, periodIdx + 1, recordPayload, recordPayload]
    );

    const { presentList, absentList, odList } = parseAttendancePayload(payload);
    sendAttendanceEmail(yearKey, dateStr, periodIdx, presentList, absentList, odList);

    res.json({ success: true, message: 'Attendance saved and email report sent!' });
  } catch (err) { 
    console.error('Save Attendance Error:', err.message);
    res.status(500).json({ success: false, error: err.message }); 
  }
});

// ==========================================
// 9. Class Advisor Endpoints
// ==========================================
app.get(['/api/advisor/daily/:year', '/api/advisor/daily/:year/'], async (req, res) => {
  try {
    const yearKey = getYearKey(req.params.year);
    const { date } = req.query;
    const [rows] = await db.query(
      `SELECT period_idx, payload FROM attendance_records WHERE year_key = ? AND date_str = ? ORDER BY period_idx ASC`,
      [yearKey, date]
    );

    const periodsData = {};
    rows.forEach(r => {
      let d = r.payload;
      while (typeof d === 'string') { try { d = JSON.parse(d); } catch (e) { break; } }
      periodsData[`P${r.period_idx}`] = d;
    });

    res.json({ success: true, date, periodsData });
  } catch (err) { res.status(500).json({ success: false, error: err.message }); }
});

// Dual Endpoint: Cumulative & Monthly-Summary
app.get([
  '/api/attendance/cumulative/:year', 
  '/api/attendance/cumulative/:year/',
  '/api/attendance/monthly-summary/:year',
  '/api/attendance/monthly-summary/:year/'
], async (req, res) => {
  try {
    const yearKey = getYearKey(req.params.year);
    const { month, fullYear } = req.query;
    const formattedMonth = month ? String(month).padStart(2, '0') : '';
    const datePrefix = fullYear && formattedMonth ? `${fullYear}-${formattedMonth}` : '';

    const [rows] = await db.query(
      `SELECT date_str, period_idx, payload FROM attendance_records WHERE year_key = ? AND date_str LIKE ? ORDER BY date_str ASC, period_idx ASC`,
      [yearKey, `${datePrefix}%`]
    );

    if (rows.length === 0) {
      return res.json({ 
        success: true, 
        totalRecordsParsed: 0, 
        totalWorkingDays: 0, 
        studentsReport: [] 
      });
    }

    const uniqueDates = [...new Set(rows.map(r => r.date_str))];
    const totalWorkingDays = uniqueDates.length;

    const studentStats = {};
    const studentDays = {};
    const IGNORED_KEYS = ['day', 'date', 'staff', 'period', 'subject', 'savedat', 'students', 'records', 'attendance', 'list'];

    rows.forEach(row => {
      const date = row.date_str;
      let data = row.payload;
      while (typeof data === 'string') { try { data = JSON.parse(data); } catch (e) { break; } }

      let recordsList = [];
      if (Array.isArray(data)) recordsList = data;
      else if (data && typeof data === 'object') {
        if (Array.isArray(data.students)) recordsList = data.students;
        else if (Array.isArray(data.records)) recordsList = data.records;
        else if (data.records && typeof data.records === 'object') {
          Object.keys(data.records).forEach(k => recordsList.push({ regNo: k, name: k, status: data.records[k] }));
        } else {
          Object.keys(data).forEach(k => {
            if (!IGNORED_KEYS.includes(k.toLowerCase())) recordsList.push({ regNo: k, name: k, status: data[k] });
          });
        }
      }

      recordsList.forEach(item => {
        if (!item) return;
        const name = item.name || item.studentName || item.regNo || item.student_id || 'Student';
        if (IGNORED_KEYS.includes(name.toLowerCase())) return;

        let rawStatus = typeof item === 'object' ? (item.status || item.attendance || '') : item;
        rawStatus = (rawStatus || '').toString().trim().toUpperCase();

        // 1. Period-wise calculation
        if (!studentStats[name]) {
          studentStats[name] = { present: 0, absent: 0, od: 0, total: 0 };
        }
        studentStats[name].total += 1;
        if (rawStatus === 'PRESENT' || rawStatus === 'P') studentStats[name].present += 1;
        else if (rawStatus === 'ABSENT' || rawStatus === 'A') studentStats[name].absent += 1;
        else if (rawStatus === 'OD' || rawStatus === 'ON DUTY' || rawStatus === 'ONDUTY') studentStats[name].od += 1;

        // 2. Day-wise calculation
        if (!studentDays[name]) studentDays[name] = {};
        if (!studentDays[name][date]) studentDays[name][date] = { attended: 0, total: 0 };
        studentDays[name][date].total += 1;
        if (rawStatus === 'PRESENT' || rawStatus === 'P' || rawStatus === 'OD' || rawStatus === 'ON DUTY' || rawStatus === 'ONDUTY') {
          studentDays[name][date].attended += 1;
        }
      });
    });

    const studentsReport = Object.keys(studentStats).map(studentName => {
      const stats = studentStats[studentName];
      const attendedPeriods = stats.present + stats.od;

      let daysPresent = 0;
      uniqueDates.forEach(date => {
        const dayRecord = studentDays[studentName] ? studentDays[studentName][date] : null;
        if (dayRecord && dayRecord.total > 0) {
          const ratio = dayRecord.attended / dayRecord.total;
          if (ratio >= 0.5) daysPresent += 1;
          else if (ratio > 0) daysPresent += 0.5;
        }
      });
      const daysAbsent = Math.max(0, totalWorkingDays - daysPresent);

      // Percentage calculation based on working days
      const percentage = totalWorkingDays > 0 
        ? ((daysPresent / totalWorkingDays) * 100).toFixed(2) 
        : (stats.total > 0 ? ((attendedPeriods / stats.total) * 100).toFixed(2) : '0.00');

      return {
        studentName,
        totalPeriods: stats.total,
        present: stats.present,
        absent: stats.absent,
        od: stats.od,
        attendedPeriods,
        totalWorkingDays,
        daysPresent,
        daysAbsent,
        percentage: parseFloat(percentage),
        isShortage: parseFloat(percentage) < 75.0
      };
    });

    const display = getDisplayTitle(yearKey);
    res.json({ 
      success: true, 
      department: display.dept,
      year: display.year, 
      month: `${fullYear}-${formattedMonth}`, 
      totalRecordsParsed: rows.length, 
      totalWorkingDays,
      studentsReport 
    });
  } catch (err) { 
    console.error('Cumulative Attendance Error:', err.message);
    res.status(500).json({ success: false, error: err.message }); 
  }
});

// ==========================================
// 10. Static Files & Root Route
// ==========================================
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static(__dirname));

app.get('*', (req, res) => {
  const publicPath = path.join(__dirname, 'public', 'index.html');
  const rootPath = path.join(__dirname, 'index.html');

  if (fs.existsSync(publicPath)) {
    res.sendFile(publicPath);
  } else if (fs.existsSync(rootPath)) {
    res.sendFile(rootPath);
  } else {
    res.send('<h2>index.html not found!</h2>');
  }
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server running globally on local network at port ${PORT}`);
});