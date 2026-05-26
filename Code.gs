// ================================================================
//  SMART CAREER GUIDANCE PLATFORM — Google Apps Script Backend
//  Paste this in script.google.com, save, then deploy as Web App
//  Google Sheet ID: 1NWPLMMy7gNM3DKAbrgNpTZI2E1ryxGvsJeHYi6BYvtE
// ================================================================

const SHEET_ID = '1NWPLMMy7gNM3DKAbrgNpTZI2E1ryxGvsJeHYi6BYvtE';
const ADMIN_EMAIL = 'YOUR_ADMIN_EMAIL@gmail.com'; // <-- Change this

// ----------------------------------------------------------------
//  ENTRY POINT — handles POST requests from the platform
// ----------------------------------------------------------------
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const result = processRegistration(data);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ----------------------------------------------------------------
//  MAIN PROCESSING LOGIC
// ----------------------------------------------------------------
function processRegistration(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  // Ensure sheets exist
  ensureSheet(ss, 'Registered_Users');
  ensureSheet(ss, 'Duplicate_Attempts');
  ensureSheet(ss, 'Analytics');
  
  const registeredSheet = ss.getSheetByName('Registered_Users');
  const duplicateSheet  = ss.getSheetByName('Duplicate_Attempts');
  
  // Check duplicate
  const isDuplicate = checkDuplicate(registeredSheet, data);
  
  const timestamp = new Date().toISOString();
  const row = buildRow(data, timestamp);
  
  if (isDuplicate) {
    // Store silently in Duplicate_Attempts — DO NOT alert user
    if (duplicateSheet.getLastRow() === 0 || duplicateSheet.getLastRow() === 1) {
      setupDuplicateHeaders(duplicateSheet);
    }
    duplicateSheet.appendRow([...row, 'DUPLICATE']);
    updateAnalytics(ss, data, true);
    
    // Still send success to user — no interruption
    return { status: 'success', registered: false, duplicate: true };
  }
  
  // Setup headers if first entry
  if (registeredSheet.getLastRow() === 0 || registeredSheet.getLastRow() === 1) {
    setupRegisteredHeaders(registeredSheet);
  }
  
  // Save to Registered_Users
  registeredSheet.appendRow(row);
  
  // Update analytics
  updateAnalytics(ss, data, false);
  
  // Send welcome email to student
  sendWelcomeEmail(data);
  
  // Send admin notification
  sendAdminNotification(data, timestamp);
  
  return { status: 'success', registered: true, duplicate: false };
}

// ----------------------------------------------------------------
//  DUPLICATE DETECTION
//  Considers duplicate if: Name+Phone OR Name+Email OR Phone+Email match
// ----------------------------------------------------------------
function checkDuplicate(sheet, data) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false; // No data yet (row 1 = header)
  
  const range = sheet.getRange(2, 1, lastRow - 1, 20);
  const values = range.getValues();
  
  const newName  = (data.name  || '').toLowerCase().trim();
  const newPhone = (data.phone || '').trim();
  const newEmail = (data.email || '').toLowerCase().trim();
  
  for (const row of values) {
    const existingName  = (row[0] || '').toString().toLowerCase().trim();
    const existingPhone = (row[5] || '').toString().trim();
    const existingEmail = (row[6] || '').toString().toLowerCase().trim();
    
    const namePhoneMatch = newName && newPhone && existingName === newName && existingPhone === newPhone;
    const nameEmailMatch = newName && newEmail && existingName === newName && existingEmail === newEmail;
    const phoneEmailMatch = newPhone && newEmail && existingPhone === newPhone && existingEmail === newEmail;
    
    if (namePhoneMatch || nameEmailMatch || phoneEmailMatch) return true;
  }
  return false;
}

// ----------------------------------------------------------------
//  BUILD ROW DATA
// ----------------------------------------------------------------
function buildRow(data, timestamp) {
  return [
    data.name        || '',
    data.age         || '',
    data.gender      || '',
    data.district    || '',
    data.school      || '',
    data.phone       || '',
    data.email       || '',
    data.educationLevel || '',
    data.hscGroup    || '',
    data.selectedCareer || '',
    data.total       || '',
    data.percentage  || '',
    // Subject marks
    data.tamil       || data.maths || '',
    data.english     || '',
    data.maths       || '',
    data.physics     || '',
    data.chemistry   || '',
    data.biology     || data.computer_science || '',
    timestamp,
    data.source      || 'CareerIQ'
  ];
}

// ----------------------------------------------------------------
//  SHEET HEADERS
// ----------------------------------------------------------------
function setupRegisteredHeaders(sheet) {
  const headers = [
    'Name','Age','Gender','District','School',
    'Phone','Email','Education Level','HSC Group','Selected Career',
    'Total Marks','Percentage',
    'Subject 1','Subject 2 (English)','Maths','Physics','Chemistry','Bio/CS',
    'Timestamp','Source'
  ];
  const headerRow = sheet.getRange(1, 1, 1, headers.length);
  headerRow.setValues([headers]);
  headerRow.setBackground('#1a3a5c');
  headerRow.setFontColor('#00c8ff');
  headerRow.setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function setupDuplicateHeaders(sheet) {
  const headers = [
    'Name','Age','Gender','District','School',
    'Phone','Email','Education Level','HSC Group','Selected Career',
    'Total Marks','Percentage',
    'Subject 1','Subject 2','Maths','Physics','Chemistry','Bio/CS',
    'Timestamp','Source','Status'
  ];
  const headerRow = sheet.getRange(1, 1, 1, headers.length);
  headerRow.setValues([headers]);
  headerRow.setBackground('#3a1a1a');
  headerRow.setFontColor('#ff6b35');
  headerRow.setFontWeight('bold');
  sheet.setFrozenRows(1);
}

// ----------------------------------------------------------------
//  ANALYTICS UPDATE
// ----------------------------------------------------------------
function updateAnalytics(ss, data, isDuplicate) {
  const analyticsSheet = ss.getSheetByName('Analytics');
  
  // Find or create analytics row
  const today = new Date().toDateString();
  const lastRow = analyticsSheet.getLastRow();
  
  if (lastRow === 0) {
    // Setup headers
    analyticsSheet.getRange(1,1,1,8).setValues([[
      'Date','Total Registrations','Engineering','Medical','Commerce','Arts/Other','Govt/UPSC','Duplicate Attempts'
    ]]);
    analyticsSheet.getRange(1,1,1,8).setBackground('#0a1a2e').setFontColor('#00ff88').setFontWeight('bold');
  }
  
  const career = (data.selectedCareer || '').toLowerCase();
  const stream = career.includes('engineer') || career.includes('cse') || career.includes('ece') ? 'Engineering'
    : career.includes('mbbs') || career.includes('medical') || career.includes('bds') ? 'Medical'
    : career.includes('commerce') || career.includes('ca') || career.includes('bank') ? 'Commerce'
    : career.includes('upsc') || career.includes('government') || career.includes('tnpsc') ? 'Government'
    : 'Arts/Other';

  // Find today's row
  let foundRow = -1;
  for (let r = 2; r <= lastRow; r++) {
    if (analyticsSheet.getRange(r, 1).getValue() === today) {
      foundRow = r; break;
    }
  }
  
  if (foundRow === -1) {
    analyticsSheet.appendRow([today, 0, 0, 0, 0, 0, 0, 0]);
    foundRow = analyticsSheet.getLastRow();
  }
  
  const colMap = {total:2, Engineering:3, Medical:4, Commerce:5, 'Arts/Other':6, Government:7, duplicates:8};
  
  if (!isDuplicate) {
    const totalCell = analyticsSheet.getRange(foundRow, 2);
    totalCell.setValue(totalCell.getValue() + 1);
    
    if (colMap[stream]) {
      const streamCell = analyticsSheet.getRange(foundRow, colMap[stream]);
      streamCell.setValue(streamCell.getValue() + 1);
    }
  } else {
    const dupCell = analyticsSheet.getRange(foundRow, 8);
    dupCell.setValue(dupCell.getValue() + 1);
  }
}

// ----------------------------------------------------------------
//  WELCOME EMAIL to Student
// ----------------------------------------------------------------
function sendWelcomeEmail(data) {
  try {
    if (!data.email) return;
    
    const name = data.name || 'Student';
    const career = data.selectedCareer || 'your chosen career';
    const group = data.hscGroup || '';
    
    const tipMap = {
      'cse': 'Focus on Data Structures & Algorithms. Build projects on GitHub. Practice LeetCode daily.',
      'ece': 'Master Embedded C and VLSI concepts. Work on hardware projects using Arduino/Raspberry Pi.',
      'mbbs': 'Start NEET preparation immediately. Focus on NCERT Biology, Physics, Chemistry.',
      'default': 'Stay consistent in your studies. Set clear goals and work towards them every day.'
    };
    
    const tipKey = Object.keys(tipMap).find(k => career.toLowerCase().includes(k)) || 'default';
    const tip = tipMap[tipKey];
    
    const subject = `🎯 Welcome to Smart Career Guidance Platform, ${name}!`;
    
    const htmlBody = `
<!DOCTYPE html>
<html>
<head><style>
  body{font-family:'Segoe UI',sans-serif;background:#040b14;color:#e8f4ff;margin:0;padding:0}
  .container{max-width:600px;margin:0 auto;background:#071220;border:1px solid rgba(0,200,255,0.2);border-radius:16px;overflow:hidden}
  .header{background:linear-gradient(135deg,#0a1a2e,#102540);padding:2.5rem;text-align:center;border-bottom:1px solid rgba(0,200,255,0.15)}
  .logo{font-size:2rem;font-weight:800;background:linear-gradient(135deg,#00c8ff,#00ff88);-webkit-background-clip:text;-webkit-text-fill-color:transparent}
  .content{padding:2rem}
  .highlight{background:rgba(0,200,255,0.08);border:1px solid rgba(0,200,255,0.2);border-radius:10px;padding:1.25rem;margin:1rem 0}
  .tip{background:rgba(0,255,136,0.08);border-left:3px solid #00ff88;padding:1rem 1.25rem;margin:1rem 0;border-radius:0 8px 8px 0}
  .badge{display:inline-block;background:rgba(0,200,255,0.15);color:#00c8ff;border:1px solid rgba(0,200,255,0.3);padding:0.3rem 0.8rem;border-radius:20px;font-size:0.85rem;font-weight:600}
  .footer{background:#040b14;padding:1.5rem;text-align:center;color:#4a6a82;font-size:0.82rem;border-top:1px solid rgba(0,200,255,0.1)}
  h1,h2,h3{font-family:'Segoe UI',sans-serif}
  p{line-height:1.7;color:#b0c8d8}
</style></head>
<body>
<div class="container">
  <div class="header">
    <div class="logo">🎯 CareerIQ</div>
    <p style="color:#7a9bb5;margin-top:0.5rem;font-size:0.9rem">Smart Career Guidance Platform</p>
  </div>
  <div class="content">
    <h2 style="color:#00c8ff">Welcome, ${name}! 🎉</h2>
    <p>Congratulations on completing your career assessment. Your personalized roadmap is ready!</p>
    
    <div class="highlight">
      <strong style="color:#00c8ff">Your Selected Career Path:</strong><br>
      <span class="badge" style="margin-top:0.5rem">${career}</span>
    </div>
    
    <div class="tip">
      <strong style="color:#00ff88">💡 Personal Tip for You:</strong><br>
      <p style="margin-top:0.5rem">${tip}</p>
    </div>
    
    <p>Your profile details:</p>
    <ul style="color:#b0c8d8;line-height:2">
      <li><strong style="color:#e8f4ff">District:</strong> ${data.district || 'Tamil Nadu'}</li>
      <li><strong style="color:#e8f4ff">Education:</strong> ${data.educationLevel || 'HSC'}</li>
      ${group ? `<li><strong style="color:#e8f4ff">Group:</strong> ${group}</li>` : ''}
    </ul>
    
    <p style="margin-top:1.5rem">We believe in your potential. Stay focused, stay consistent, and success will follow!</p>
    <br>
    <p>With best wishes,<br><strong style="color:#00c8ff">Kalidass</strong><br><em style="color:#4a6a82">Smart Career Guidance Platform</em></p>
  </div>
  <div class="footer">
    © ${new Date().getFullYear()} Smart Career Guidance Platform · Tamil Nadu, India<br>
    <em>Empowering students through intelligent career guidance</em>
  </div>
</div>
</body>
</html>`;
    
    GmailApp.sendEmail(data.email, subject, 
      `Welcome ${name}! Your career path: ${career}. Personal tip: ${tip} — Kalidass, Smart Career Guidance Platform`,
      { htmlBody: htmlBody }
    );
  } catch (err) {
    console.log('Email error:', err.message);
  }
}

// ----------------------------------------------------------------
//  ADMIN NOTIFICATION
// ----------------------------------------------------------------
function sendAdminNotification(data, timestamp) {
  try {
    const subject = `[CareerIQ] New Registration: ${data.name} — ${data.selectedCareer || 'Career Analysis'}`;
    
    const body = `
New student registered on Smart Career Guidance Platform.

Details:
━━━━━━━━━━━━━━━━━━━━━━━━━━
Name        : ${data.name || 'N/A'}
Age         : ${data.age || 'N/A'}
Gender      : ${data.gender || 'N/A'}
District    : ${data.district || 'N/A'}
School      : ${data.school || 'N/A'}
Phone       : ${data.phone || 'N/A'}
Email       : ${data.email || 'N/A'}
━━━━━━━━━━━━━━━━━━━━━━━━━━
Education   : ${data.educationLevel || 'N/A'}
HSC Group   : ${data.hscGroup || 'N/A'}
Career      : ${data.selectedCareer || 'N/A'}
━━━━━━━━━━━━━━━━━━━━━━━━━━
Timestamp   : ${timestamp}
━━━━━━━━━━━━━━━━━━━━━━━━━━

View full data: https://docs.google.com/spreadsheets/d/${SHEET_ID}
    `;
    
    GmailApp.sendEmail(ADMIN_EMAIL, subject, body);
  } catch (err) {
    console.log('Admin notify error:', err.message);
  }
}

// ----------------------------------------------------------------
//  UTILITY: Ensure Sheet Exists
// ----------------------------------------------------------------
function ensureSheet(ss, name) {
  if (!ss.getSheetByName(name)) {
    ss.insertSheet(name);
  }
}

// ----------------------------------------------------------------
//  TEST FUNCTION — run manually to verify
// ----------------------------------------------------------------
function testIntegration() {
  const testData = {
    name: 'Test Student',
    age: 17,
    gender: 'Male',
    district: 'Madurai',
    school: 'Test School',
    phone: '9876543210',
    email: 'test@example.com',
    educationLevel: 'HSC',
    hscGroup: 'biomaths',
    selectedCareer: 'CSE Engineering',
    total: 580,
    percentage: 96.7,
    source: 'CareerIQ Test'
  };
  
  const result = processRegistration(testData);
  Logger.log('Test result: ' + JSON.stringify(result));
}
