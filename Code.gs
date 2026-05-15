// ─────────────────────────────────────────────────────────────────
//  MANTRA  —  Daily Scheduler v2
//  Paste this entire file into Apps Script
//  (Extensions → Apps Script in your Google Sheet)
//
//  Your email and settings live in the Config sheet — no need
//  to edit this script directly.
// ─────────────────────────────────────────────────────────────────


// ── Subject line pool ──────────────────────────────────────────────
const SUBJECTS = [
  '✦ Your mantra',
  'A word for now',
  'Something to carry with you',
  'For this moment',
  'Before you continue',
  'Pause.',
  'This one is yours',
  'A reminder',
  'Keep this close today',
  'One thing',
  'Let this land',
  'Quietly, for you',
  'Worth remembering',
  'Right now',
  'Hold this',
  'Breathe first',
  'A small truth',
  'Before the noise',
  'Just this',
  'For the next few hours',
  'An anchor',
  'Words first',
  'Before you react',
  'Still with you',
  'A thread to follow',
  'Soft reminder',
  'In the middle of your day',
  'Take a second',
  'Something real',
  'Return to this',
  'A quiet word',
  'Only for a moment',
  'Today\'s compass',
  'Keep going',
  'From your own words',
  'What you wrote',
  'You said this once',
  'A note to self',
  'Your words, back to you',
  'Written for a day like this',
];


// ── Load config from Config sheet ────────────────────────────────
function loadConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');
  if (!sheet) throw new Error('Config sheet not found. Please check sheet tab name.');

  const data = sheet.getDataRange().getValues();
  const config = { email: '', defaultMax: 2, themeMax: {} };

  // General settings rows 4–7 (0-indexed: 3–6)
  for (let i = 3; i < 7; i++) {
    const label = String(data[i]?.[0] || '').trim();
    const val   = String(data[i]?.[1] || '').trim();
    if (label.toLowerCase().includes('email'))       config.email = val;
    if (label.toLowerCase().includes('default max')) config.defaultMax = parseInt(val) || 2;
  }

  // Theme limits — find the table starting after "PER-THEME DAILY LIMITS" header
  let inThemeTable = false;
  for (let i = 0; i < data.length; i++) {
    const cell = String(data[i]?.[0] || '').trim();
    if (cell.toUpperCase().includes('PER-THEME')) { inThemeTable = true; continue; }
    if (inThemeTable && cell && cell.toLowerCase() !== 'theme') {
      const limit = parseInt(data[i]?.[1]);
      if (cell && !isNaN(limit)) config.themeMax[cell] = limit;
    }
  }

  if (!config.email || !config.email.includes('@')) {
    throw new Error('No valid email found in Config sheet. Please set Your Email Address.');
  }

  return config;
}


// ── Main function — triggered daily at midnight ──────────────────
function scheduleDailyMantras() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const config = loadConfig();
  const today  = getTodayKey();   // YYYY-MM-DD
  const dayName = getDayName();   // Mon, Tue, etc.

  // ── Duplicate guard ──
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('scheduled_date') === today) {
    Logger.log('Already scheduled for today (' + today + '). Skipping.');
    return;
  }

  // ── Read all mantras (skip blank rows, no early exit) ──
  const mantrasSheet = ss.getSheetByName('Mantras');
  if (!mantrasSheet) throw new Error('Mantras sheet not found.');

  const lastRow = mantrasSheet.getLastRow();
  const allRows = lastRow >= 5
    ? mantrasSheet.getRange(5, 1, lastRow - 4, 11).getValues()
    : [];

  const now = new Date();

  // ── Per-theme count tracking ──
  const themeCount = {};

  // ── Build list of mantras to schedule ──
  const toSchedule = [];

  // Shuffle rows for equal distribution (random order before applying limits)
  const shuffled = allRows.slice().sort(() => Math.random() - 0.5);

  for (const row of shuffled) {
    const text       = String(row[0] || '').trim();
    const theme      = String(row[1] || '').trim();
    const active     = String(row[2] || '').trim().toUpperCase();
    const winStart   = row[3] || '';
    const winEnd     = row[4] || '';
    const fixedTime  = row[5] || '';
    const daysRaw    = String(row[6] || '').trim();
    const pauseUntil = String(row[7] || '').trim();

    // Skip blank mantra rows
    if (!text) continue;

    // Skip inactive
    if (active !== 'YES') continue;

    // Skip if paused until future date
    if (pauseUntil) {
      const pauseDate = new Date(pauseUntil);
      if (!isNaN(pauseDate) && pauseDate > now) continue;
    }

    // Skip if today not in delivery days
    if (daysRaw) {
      const days = daysRaw.split(',').map(d => d.trim());
      if (!days.includes(dayName)) continue;
    }

    // Check per-theme limit
    const maxForTheme = (theme && config.themeMax[theme] !== undefined)
      ? config.themeMax[theme]
      : config.defaultMax;
    const currentCount = themeCount[theme] || 0;
    if (currentCount >= maxForTheme) continue;

    // Resolve delivery time
    const deliveryTime = resolveDeliveryTime(fixedTime, winStart, winEnd);
    if (!deliveryTime) continue;

    themeCount[theme] = currentCount + 1;
    toSchedule.push({ text, theme, deliveryTime, row });
  }

  if (toSchedule.length === 0) {
    Logger.log('No mantras to schedule today.');
    props.setProperty('scheduled_date', today);
    return;
  }

  // ── Schedule each mantra ──
  const subject = SUBJECTS[Math.floor(Math.random() * SUBJECTS.length)];
  const scheduledItems = [];

  for (const item of toSchedule) {
    createTimedTrigger(item.text, item.theme, item.deliveryTime, subject);
    scheduledItems.push({
      text: item.text,
      theme: item.theme,
      time: formatTime(item.deliveryTime)
    });
    Logger.log('Scheduled: "' + item.text.slice(0, 40) + '" at ' + formatTime(item.deliveryTime));
  }

  // ── Mark today as scheduled ──
  props.setProperty('scheduled_date', today);

  // ── Update Analytics today's schedule ──
  updateTodaySchedule(ss, scheduledItems);

  Logger.log('Done. ' + toSchedule.length + ' mantra(s) scheduled.');
}


// ── Create a one-time trigger for a single mantra ──────────────────
function createTimedTrigger(text, theme, deliveryTime, subject) {
  const props = PropertiesService.getScriptProperties();

  // Create the trigger first, then use its unique ID as the key
  const trigger = ScriptApp.newTrigger('sendPendingMantras')
    .timeBased()
    .at(deliveryTime)
    .create();

  const key = 'pending_' + trigger.getUniqueId();
  props.setProperty(key, JSON.stringify({ text, theme, subject }));
}


// ── Called by each one-time trigger ──────────────────────────────
function sendPendingMantras(e) {
  const props    = PropertiesService.getScriptProperties();
  const allProps = props.getProperties();
  const config   = loadConfig();
  const ss       = SpreadsheetApp.getActiveSpreadsheet();

  // Use the trigger's unique ID to find exactly the right mantra
  // Note: e.triggerUid is a number, getUniqueId() returns a string — compare both ways
  const triggerId = e && e.triggerUid ? String(e.triggerUid) : null;
  const exactKey  = triggerId ? 'pending_' + triggerId : null;

  Logger.log('Trigger fired. triggerUid=' + triggerId + ' exactKey=' + exactKey);
  Logger.log('All pending keys: ' + Object.keys(allProps).filter(k => k.startsWith('pending_')).join(', '));

  // Build list of keys to process — exact match first, fallback to all pending
  const keysToProcess = exactKey && allProps[exactKey]
    ? [exactKey]
    : Object.keys(allProps).filter(k => k.startsWith('pending_'));

  Logger.log('Keys to process: ' + keysToProcess.join(', '));

  for (const key of keysToProcess) {
    if (!allProps[key]) continue;
    try {
      const mantra = JSON.parse(allProps[key]);
      const body   = buildEmailBody(mantra.text, mantra.theme);
      GmailApp.sendEmail(config.email, mantra.subject, '', { htmlBody: body });
      logDelivery(ss, mantra.text, mantra.theme, formatTime(new Date()));
      updateMantraRow(ss, mantra.text);
      Logger.log('Sent: "' + mantra.text.slice(0, 40) + '"');
    } catch(err) {
      Logger.log('Send error: ' + err.message);
    }
    props.deleteProperty(key);
  }

  cleanupTrigger(triggerId);
}


// ── Log a delivery to the Delivery Log sheet ──────────────────────
function logDelivery(ss, text, theme, time) {
  const sheet = ss.getSheetByName('Delivery Log');
  if (!sheet) return;

  const today    = getTodayKey();
  const lastRow  = sheet.getLastRow();
  const insertAt = lastRow < 3 ? 4 : lastRow + 1;

  sheet.getRange(insertAt, 1, 1, 6).setValues([[
    today,
    time,
    text,
    theme,
    '—',
    'No'
  ]]);
}


// ── Update delivery count + last sent in Mantras sheet ────────────
function updateMantraRow(ss, text) {
  const sheet   = ss.getSheetByName('Mantras');
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 5) return;

  const texts = sheet.getRange(5, 1, lastRow - 4, 1).getValues();
  for (let i = 0; i < texts.length; i++) {
    if (String(texts[i][0]).trim() === text) {
      const dataRow = i + 5;
      const countCell = sheet.getRange(dataRow, 9);
      const lastSentCell = sheet.getRange(dataRow, 10);
      countCell.setValue((parseInt(countCell.getValue()) || 0) + 1);
      lastSentCell.setValue(getTodayKey());
      return;
    }
  }
}


// ── Update today's schedule preview in Analytics ──────────────────
function updateTodaySchedule(ss, items) {
  const sheet = ss.getSheetByName('Analytics');
  if (!sheet) return;

  // Clear rows 5–9
  sheet.getRange(5, 1, 5, 4).clearContent();

  items.slice(0, 5).forEach((item, i) => {
    sheet.getRange(5 + i, 1).setValue(item.text);
    sheet.getRange(5 + i, 2).setValue(item.theme);
    sheet.getRange(5 + i, 3).setValue(item.time);
    sheet.getRange(5 + i, 4).setValue('Scheduled ✓');
  });
}


// ── Clean up only the trigger that just fired ─────────────────────
function cleanupTrigger(triggerId) {
  if (!triggerId) return;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (String(t.getUniqueId()) === String(triggerId)) {
      try { ScriptApp.deleteTrigger(t); } catch(e) {}
    }
  });
}


// ── Resolve delivery time ─────────────────────────────────────────
function resolveDeliveryTime(fixedTime, winStart, winEnd) {
  const now = new Date();

  if (fixedTime) {
    const d = parseTimeToday(fixedTime);
    return (d && d > now) ? d : null;
  }

  if (winStart && winEnd) {
    const start = parseTimeToday(winStart);
    const end   = parseTimeToday(winEnd);
    if (!start || !end || end <= start) return null;

    const startMs = start.getTime();
    const endMs   = end.getTime();
    const nowMs   = now.getTime();

    if (endMs <= nowMs) return null;

    const from = Math.max(startMs, nowMs + 60000);
    if (from >= endMs) return null;

    const randomMs = from + Math.floor(Math.random() * (endMs - from));
    return new Date(randomMs);
  }

  // Fallback
  return resolveDeliveryTime('', '09:00', '20:00');
}


// ── Helpers ───────────────────────────────────────────────────────
function parseTimeToday(timeVal) {
  // Google Sheets may return a Date object for time cells
  if (timeVal instanceof Date && !isNaN(timeVal)) {
    const d = new Date();
    d.setHours(timeVal.getHours(), timeVal.getMinutes(), 0, 0);
    return d;
  }

  const timeStr = String(timeVal).trim();
  const parts = timeStr.split(':').map(Number);
  if (parts.length < 2 || isNaN(parts[0]) || isNaN(parts[1])) return null;
  const d = new Date();
  d.setHours(parts[0], parts[1], 0, 0);
  return d;
}

function formatTime(date) {
  return date.toLocaleTimeString('en-US', {
    hour: '2-digit', minute: '2-digit', hour12: true,
    timeZone: 'America/Los_Angeles'
  });
}

function getTodayKey() {
  return new Date().toLocaleDateString('en-CA', { timeZone: 'America/Los_Angeles' }); // YYYY-MM-DD
}

function getDayName() {
  return new Date().toLocaleDateString('en-US', { weekday: 'short', timeZone: 'America/Los_Angeles' });
  // Returns "Mon", "Tue" etc.
}

function buildEmailBody(text, theme) {
  const label = (theme && theme.trim() && theme !== 'undefined') ? theme : 'Your mantra';
  return `
    <div style="font-family:Georgia,serif;max-width:480px;margin:40px auto;padding:48px 40px;background:#faf8f5;border-radius:10px;">
      <p style="font-size:10px;letter-spacing:4px;color:#c9a96e;text-transform:uppercase;margin:0 0 28px;">
        ${label}
      </p>
      <p style="font-size:22px;color:#2c2420;line-height:1.7;margin:0 0 36px;font-style:italic;">
        &ldquo;${text}&rdquo;
      </p>
      <p style="font-size:10px;color:#c4b5a8;letter-spacing:1px;margin:0;">
        Breathe it in. Let it guide you.
      </p>
    </div>
  `;
}


// ── Debug: check what's pending right now ────────────────────────
function debugPending() {
  const props    = PropertiesService.getScriptProperties();
  const allProps = props.getProperties();
  const pending  = Object.keys(allProps).filter(k => k.startsWith('pending_'));
  Logger.log('Pending mantras: ' + pending.length);
  pending.forEach(k => {
    try {
      const m = JSON.parse(allProps[k]);
      Logger.log('  Key: ' + k + ' → "' + m.text + '"');
    } catch(e) { Logger.log('  Key: ' + k + ' → (parse error)'); }
  });
  Logger.log('scheduled_date: ' + props.getProperty('scheduled_date'));
  Logger.log('Config email: ' + loadConfig().email);
}


// ── Manual test: send one mantra immediately ──────────────────────
function testSendNow() {
  const config = loadConfig();
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = ss.getSheetByName('Mantras');
  const rows   = sheet.getRange(5, 1, sheet.getLastRow() - 4, 3).getValues();

  // Find first active mantra
  const active = rows.find(r => String(r[0]).trim() && String(r[2]).trim().toUpperCase() === 'YES');
  if (!active) { Logger.log('No active mantras found.'); return; }

  const text = String(active[0]).trim();
  const theme = String(active[1]).trim();
  const body  = buildEmailBody(text, theme);
  const subject = SUBJECTS[Math.floor(Math.random() * SUBJECTS.length)];

  GmailApp.sendEmail(config.email, subject, '', { htmlBody: body });
  Logger.log('Test email sent to ' + config.email + ': "' + text + '"');
}


// ── Test: schedule one mantra 2 minutes from now ─────────────────
function testScheduledDelivery() {
  const config = loadConfig();
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = ss.getSheetByName('Mantras');
  const rows   = sheet.getRange(5, 1, sheet.getLastRow() - 4, 3).getValues();

  const active = rows.find(r => String(r[0]).trim() && String(r[2]).trim().toUpperCase() === 'YES');
  if (!active) { Logger.log('No active mantras found.'); return; }

  const text    = String(active[0]).trim();
  const theme   = String(active[1]).trim();
  const subject = SUBJECTS[Math.floor(Math.random() * SUBJECTS.length)];

  // Deliver 2 minutes from now
  const deliveryTime = new Date(Date.now() + 2 * 60 * 1000);

  const props   = PropertiesService.getScriptProperties();
  const trigger = ScriptApp.newTrigger('sendPendingMantras')
    .timeBased()
    .at(deliveryTime)
    .create();

  const key = 'pending_' + trigger.getUniqueId();
  props.setProperty(key, JSON.stringify({ text, theme, subject }));

  Logger.log('Test scheduled: "' + text + '"');
  Logger.log('Trigger ID: ' + trigger.getUniqueId());
  Logger.log('Key stored: ' + key);
  Logger.log('Delivery at: ' + deliveryTime.toLocaleTimeString());
  Logger.log('Email will go to: ' + config.email);
}


// ── Setup yearly December 1st trigger (run this once manually) ────
function setupYearlyReviewTrigger() {
  // Uses monthly trigger on 1st of month, checks if December
  ScriptApp.newTrigger('checkIfDecember1')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();
  Logger.log('Monthly trigger created. Will send year review every December 1st.');
}

function checkIfDecember1() {
  const now = new Date();
  if (now.getMonth() === 11) { // 11 = December
    sendYearReview();
  }
}


// ── Build year stats from Delivery Log ───────────────────────────
function buildYearStats(year) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Delivery Log');
  if (!sheet) return null;

  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return null;

  const rows = sheet.getRange(4, 1, lastRow - 3, 4).getValues();
  const yearStr = String(year);

  // Filter to this year
  const yearRows = rows.filter(r => String(r[0]).startsWith(yearStr) && String(r[2]).trim());
  if (yearRows.length === 0) return null;

  // Total sent
  const totalSent = yearRows.length;

  // Per-mantra count
  const mantraCounts = {};
  yearRows.forEach(r => {
    const text = String(r[2]).trim();
    mantraCounts[text] = (mantraCounts[text] || 0) + 1;
  });

  // Per-theme count
  const themeCounts = {};
  yearRows.forEach(r => {
    const theme = String(r[3]).trim() || 'General';
    themeCounts[theme] = (themeCounts[theme] || 0) + 1;
  });

  // Most delivered mantra
  const topMantra = Object.entries(mantraCounts).sort((a,b) => b[1]-a[1])[0];

  // Most active theme
  const topTheme = Object.entries(themeCounts).sort((a,b) => b[1]-a[1])[0];

  // Per-month breakdown
  const monthCounts = {};
  yearRows.forEach(r => {
    const month = String(r[0]).slice(0,7); // YYYY-MM
    monthCounts[month] = (monthCounts[month] || 0) + 1;
  });
  const topMonth = Object.entries(monthCounts).sort((a,b) => b[1]-a[1])[0];
  const quietMonth = Object.entries(monthCounts).sort((a,b) => a[1]-b[1])[0];

  // Unique mantras received
  const uniqueMantras = Object.keys(mantraCounts).length;

  // Avg per active month
  const activeMonths = Object.keys(monthCounts).length;
  const avgPerMonth = activeMonths > 0 ? (totalSent / activeMonths).toFixed(1) : 0;

  // Streak calculation
  const deliveryDates = [...new Set(yearRows.map(r => String(r[0]).slice(0,10)))].sort();
  let longestStreak = 1, currentStreak = 1;
  for (let i = 1; i < deliveryDates.length; i++) {
    const prev = new Date(deliveryDates[i-1]);
    const curr = new Date(deliveryDates[i]);
    const diff = (curr - prev) / (1000 * 60 * 60 * 24);
    if (diff === 1) {
      currentStreak++;
      longestStreak = Math.max(longestStreak, currentStreak);
    } else {
      currentStreak = 1;
    }
  }

  // Most active time of day
  const hourCounts = {};
  yearRows.forEach(r => {
    const time = String(r[1]).trim();
    const hour = time ? parseInt(time.split(':')[0]) : null;
    if (hour !== null && !isNaN(hour)) {
      const label = hour < 12 ? 'Morning' : hour < 17 ? 'Afternoon' : 'Evening';
      hourCounts[label] = (hourCounts[label] || 0) + 1;
    }
  });
  const topTimeOfDay = Object.entries(hourCounts).sort((a,b) => b[1]-a[1])[0];

  return {
    year, totalSent, uniqueMantras, avgPerMonth,
    topMantra: topMantra ? { text: topMantra[0], count: topMantra[1] } : null,
    topTheme:  topTheme  ? { name: topTheme[0],  count: topTheme[1]  } : null,
    topMonth:  topMonth  ? { month: topMonth[0],  count: topMonth[1]  } : null,
    quietMonth: quietMonth ? { month: quietMonth[0], count: quietMonth[1] } : null,
    longestStreak, activeMonths, topTimeOfDay: topTimeOfDay ? topTimeOfDay[0] : null,
    themeCounts, mantraCounts, monthCounts
  };
}


// ── Build the year review email HTML ─────────────────────────────
function buildYearReviewEmail(stats) {
  const monthNames = {
    '01':'January','02':'February','03':'March','04':'April',
    '05':'May','06':'June','07':'July','08':'August',
    '09':'September','10':'October','11':'November','12':'December'
  };

  const fmtMonth = (ym) => {
    const parts = ym.split('-');
    return (monthNames[parts[1]] || parts[1]) + ' ' + parts[0];
  };

  // Theme rows
  const themeRows = Object.entries(stats.themeCounts)
    .sort((a,b) => b[1]-a[1])
    .map(([theme, count]) => `
      <tr>
        <td style="padding:8px 0;color:#2c2420;font-size:14px;">${theme}</td>
        <td style="padding:8px 0;color:#c9a96e;font-size:14px;text-align:right;">${count}</td>
      </tr>`).join('');

  // Month rows
  const monthRows = Object.entries(stats.monthCounts)
    .sort((a,b) => a[0].localeCompare(b[0]))
    .map(([ym, count]) => `
      <tr>
        <td style="padding:6px 0;color:#2c2420;font-size:13px;">${fmtMonth(ym)}</td>
        <td style="padding:6px 0;color:#7a6e68;font-size:13px;text-align:right;">${count}</td>
      </tr>`).join('');

  return `
    <div style="font-family:Georgia,serif;max-width:560px;margin:40px auto;background:#faf8f5;border-radius:12px;overflow:hidden;">

      <!-- Header -->
      <div style="background:#0e0c0a;padding:48px 40px;text-align:center;">
        <p style="font-size:10px;letter-spacing:5px;color:#c9a96e;text-transform:uppercase;margin:0 0 16px;">Your year in words</p>
        <p style="font-size:42px;color:#e8e0d5;margin:0;letter-spacing:6px;text-transform:uppercase;font-weight:300;">${stats.year}</p>
      </div>

      <!-- Big number -->
      <div style="padding:40px;text-align:center;border-bottom:1px solid #ede8e2;">
        <p style="font-size:72px;color:#c9a96e;margin:0;font-weight:300;">${stats.totalSent}</p>
        <p style="font-size:12px;letter-spacing:3px;color:#7a6e68;text-transform:uppercase;margin:8px 0 0;">mantras delivered</p>
        <p style="font-size:14px;color:#7a6e68;margin:16px 0 0;font-style:italic;">
          ${stats.uniqueMantras} unique mantras · ${stats.activeMonths} active months · ${stats.avgPerMonth} avg per month
        </p>
      </div>

      <div style="padding:40px;">

        <!-- Top mantra -->
        ${stats.topMantra ? `
        <div style="margin-bottom:32px;">
          <p style="font-size:10px;letter-spacing:3px;color:#c9a96e;text-transform:uppercase;margin:0 0 12px;">Most delivered mantra</p>
          <p style="font-size:20px;color:#2c2420;font-style:italic;line-height:1.6;margin:0 0 8px;">&ldquo;${stats.topMantra.text}&rdquo;</p>
          <p style="font-size:12px;color:#7a6e68;margin:0;">Delivered ${stats.topMantra.count} time${stats.topMantra.count !== 1 ? 's' : ''}</p>
        </div>` : ''}

        <!-- Stats row -->
        <div style="display:flex;gap:16px;margin-bottom:32px;">
          ${stats.longestStreak > 1 ? `
          <div style="flex:1;background:#f2ede6;border-radius:8px;padding:20px;text-align:center;">
            <p style="font-size:32px;color:#c9a96e;margin:0;">${stats.longestStreak}</p>
            <p style="font-size:10px;letter-spacing:2px;color:#7a6e68;text-transform:uppercase;margin:6px 0 0;">day streak</p>
          </div>` : ''}
          ${stats.topTimeOfDay ? `
          <div style="flex:1;background:#f2ede6;border-radius:8px;padding:20px;text-align:center;">
            <p style="font-size:22px;color:#c9a96e;margin:0;">${stats.topTimeOfDay}</p>
            <p style="font-size:10px;letter-spacing:2px;color:#7a6e68;text-transform:uppercase;margin:6px 0 0;">peak time</p>
          </div>` : ''}
          ${stats.topTheme ? `
          <div style="flex:1;background:#f2ede6;border-radius:8px;padding:20px;text-align:center;">
            <p style="font-size:18px;color:#c9a96e;margin:0;">${stats.topTheme.name}</p>
            <p style="font-size:10px;letter-spacing:2px;color:#7a6e68;text-transform:uppercase;margin:6px 0 0;">top theme</p>
          </div>` : ''}
        </div>

        <!-- Theme breakdown -->
        ${themeRows ? `
        <div style="margin-bottom:32px;">
          <p style="font-size:10px;letter-spacing:3px;color:#c9a96e;text-transform:uppercase;margin:0 0 12px;">By theme</p>
          <table style="width:100%;border-collapse:collapse;border-top:1px solid #ede8e2;">
            ${themeRows}
          </table>
        </div>` : ''}

        <!-- Month breakdown -->
        ${monthRows ? `
        <div style="margin-bottom:32px;">
          <p style="font-size:10px;letter-spacing:3px;color:#c9a96e;text-transform:uppercase;margin:0 0 12px;">By month</p>
          <table style="width:100%;border-collapse:collapse;border-top:1px solid #ede8e2;">
            ${monthRows}
          </table>
        </div>` : ''}

        <!-- Footer -->
        <div style="text-align:center;padding-top:24px;border-top:1px solid #ede8e2;">
          <p style="font-size:11px;color:#c4b5a8;letter-spacing:1px;margin:0;">
            Every word you wrote this year was written for a reason.<br>Keep going.
          </p>
        </div>

      </div>
    </div>
  `;
}


// ── Send year review email ────────────────────────────────────────
function sendYearReview() {
  const config = loadConfig();
  const year   = new Date().getFullYear();
  const stats  = buildYearStats(year);

  if (!stats) {
    Logger.log('No delivery data found for ' + year);
    return;
  }

  const html    = buildYearReviewEmail(stats);
  const subject = '✦ Your ' + year + ' in mantras';
  GmailApp.sendEmail(config.email, subject, '', { htmlBody: html });
  Logger.log('Year review sent for ' + year + '. Total delivered: ' + stats.totalSent);
}


// ── Test: send year review right now with current data ───────────
function sendYearReviewNow() {
  sendYearReview();
  Logger.log('Year review sent — check your inbox.');
}

