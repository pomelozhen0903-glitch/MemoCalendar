/**
 * Code.gs - V18 (Calendar List & UI Fixes)
 */

const GOOGLE_API_KEY = 'AIzaSyAlD_nXkX_V9MDlsqgHoHS_ejdNOtpGvuM';
const S_ID = '1JfbkK440m5pxDxZRX4YRI8YGpbTX2ixdho1fnJzmhck';
const SHEET_NAME = '工作表1'; 

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('VoicePro V18')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function getSheet() {
  try {
    const ss = SpreadsheetApp.openById(S_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.getSheets()[0];
    return sheet;
  } catch (e) {
    throw new Error('DB Error: ' + e.message);
  }
}

// --- 新增：取得使用者所有行事曆 ---
function getUserCalendars() {
  try {
    const calendars = CalendarApp.getAllCalendars();
    return calendars.map(cal => ({
      id: cal.getId(),
      name: cal.getName(),
      isPrimary: cal.isMyPrimaryCalendar()
    }));
  } catch (e) {
    return [];
  }
}

// --- 讀取功能 ---
function getNotes() {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const values = sheet.getDataRange().getValues();
    const rawData = values.slice(1);
    const safe = (row, i) => (row[i] === undefined || row[i] === null) ? '' : row[i];

    const notes = rawData.map(row => {
      let rawDate = safe(row, 3);
      let dateStr = '';
      
      // 處理日期格式 YYYY/MM/DD
      if (rawDate instanceof Date) {
        if (rawDate.getFullYear() > 1950) {
           const y = rawDate.getFullYear();
           const m = (rawDate.getMonth()+1).toString().padStart(2,'0');
           const d = rawDate.getDate().toString().padStart(2,'0');
           dateStr = `${y}/${m}/${d}`; // V18: 改為斜線
        }
      } else if (typeof rawDate === 'string') {
         dateStr = rawDate.replace(/-/g, '/');
      }

      return {
        id: safe(row, 0).toString(),
        originalText: safe(row, 1).toString(),
        title: safe(row, 2).toString(),
        dateStr: dateStr,
        timeStr: safe(row, 4).toString(),
        timestamp: safe(row, 5),
        displayDate: dateStr, 
        priority: safe(row, 7) || 'Normal',
        remarks: safe(row, 8).toString(),
        isPinned: String(safe(row, 9)) === 'true',
        isDone: String(safe(row, 10)) === 'true'
      };
    });

    return notes.filter(n => n.id !== '').reverse();
  } catch (e) {
    throw new Error(e.message);
  }
}

// --- 寫入邏輯 ---
function processVoiceNote(text, clientTimeStr) {
  const sheet = getSheet();
  const now = new Date();
  const timestamp = now.getTime();
  const id = timestamp.toString();
  
  let ai = { title: text, priority: 'Normal', dateStr: '', timeStr: '' };
  
  if (GOOGLE_API_KEY) {
    try {
      const prompt = `Current Time: ${clientTimeStr}\nInput: "${text}"\nExtract Title, Date, Time (HH:mm), Priority. Return JSON.`;
      const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=${GOOGLE_API_KEY}`;
      const payload = { contents: [{ parts: [{ text: prompt }] }] };
      const res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
      const json = JSON.parse(res.getContentText()).candidates[0].content.parts[0].text.replace(/```json|```/g, '').trim();
      ai = JSON.parse(json);
    } catch (e) {}
  }

  // 自動加入預設行事曆 (語音輸入時仍保持自動，但只加預設)
  let autoCalMsg = false;
  if (ai.dateStr && ai.timeStr) {
     const calRes = addToGoogleCalendar(ai.title, ai.dateStr, ai.timeStr, null); // null = default calendar
     if(calRes.success) autoCalMsg = true;
  }

  const newRow = [id, text, ai.title || text, ai.dateStr||'', ai.timeStr||'', timestamp, '', ai.priority || 'Normal', '', 'false', 'false'];
  if(sheet.getLastColumn() < 11) sheet.getRange(1,1,1,11).setValues([['ID','Text','Title','Date','Time','TS','Disp','Prio','Rem','Pin','Done']]);
  sheet.appendRow(newRow);
  
  const res = formatNoteObject(newRow);
  res.autoCal = autoCalMsg;
  return res;
}

function addManualNote(data) {
  const sheet = getSheet();
  const now = new Date();
  const newRow = [data.id || now.getTime().toString(), data.originalText||data.title, data.title, data.dateStr, data.timeStr, now.getTime(), '', data.priority, data.remarks, String(data.isPinned), 'false'];
  sheet.appendRow(newRow);
  return formatNoteObject(newRow);
}

function updateNote(data) {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === data.id.toString()) {
      const r = i + 1;
      sheet.getRange(r, 3).setValue(data.title);
      sheet.getRange(r, 4).setValue(data.dateStr);
      sheet.getRange(r, 5).setValue(data.timeStr);
      sheet.getRange(r, 8).setValue(data.priority);
      sheet.getRange(r, 9).setValue(data.remarks);
      sheet.getRange(r, 10).setValue(String(data.isPinned));
      return { success: true };
    }
  }
  return { success: false };
}

function toggleDone(id, isDone) {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === id.toString()) {
      sheet.getRange(i + 1, 11).setValue(String(isDone)); 
      return true;
    }
  }
  return false;
}

function deleteNotes(ids) {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (ids.includes(rows[i][0].toString())) sheet.deleteRow(i + 1);
  }
  return { success: true };
}

// --- V18: 支援指定行事曆 ID ---
function addToGoogleCalendar(title, dateStr, timeStr, calendarId) {
  try {
    // 如果有指定 ID 就用指定的，不然用預設
    const calendar = calendarId ? CalendarApp.getCalendarById(calendarId) : CalendarApp.getDefaultCalendar();
    
    if (!calendar) throw new Error('找不到指定的行事曆');

    let start;
    if (dateStr && timeStr) start = new Date(dateStr.replace(/-/g, '/') + ' ' + timeStr);
    else if (dateStr) start = new Date(dateStr.replace(/-/g, '/') + ' 09:00:00');
    else { start = new Date(); start.setHours(start.getHours() + 1); }
    
    const end = new Date(start.getTime() + 60 * 60 * 1000);
    calendar.createEvent(title, start, end).addPopupReminder(0);
    
    return { success: true, msg: `✅ 已加入「${calendar.getName()}」` };
  } catch (e) { 
    return { success: false, msg: '錯誤: ' + e.toString() }; 
  }
}

function formatNoteObject(row) {
  const safe = (i) => (row[i]===undefined||row[i]===null)?'':row[i].toString();
  let d = safe(3);
  if(d.includes('-')) d = d.replace(/-/g, '/'); // 統一轉為 / 
  return { id: safe(0), originalText: safe(1), title: safe(2), dateStr: d, timeStr: safe(4), timestamp: safe(5), displayDate: d, priority: safe(7)||'Normal', remarks: safe(8), isPinned: String(safe(9))==='true', isDone: String(safe(10))==='true' };
}
