/**
 * SMART MAINTAIN – Bharat Petroleum, Gonda LPG Plant
 * Google Apps Script backend for shared equipment/AMC data and automatic email reminders.
 *
 * IMPORTANT ACCESS NOTE:
 * This version accepts website sync writes without an OTP because the current website
 * has no viewer/editor login. Therefore the Web App endpoint must be restricted to
 * "Anyone within Bharat Petroleum" if available. If deployed as public "Anyone",
 * anyone who obtains the endpoint can change the shared register. Obtain BPCL IT approval.
 *
 * INSTALL:
 * 1. Replace Code.gs in the EXISTING Apps Script project connected to the website URL.
 * 2. Save.
 * 3. Run setupSmartMaintain() once and approve requested Google permissions.
 * 4. Run sendTestEmail() once and approve MailApp permission. It sends a real test email
 *    to all five configured recipients.
 * 5. Deploy > Manage deployments > Edit > New version > Deploy.
 * 6. Keep Execute as: Me. Keep the same /exec URL in index.html.
 * 7. Project Settings > Time zone: Asia/Kolkata.
 */

const RECIPIENTS = [
  'rithikeshrathi@bharatpetroleum.in',
  'gauravmadan@bharatpetroleum.in',
  'patilvishalravi@bharatpetroleum.in',
  'atharavsanjayka@bharatpetroleum.in',
  'frootyrithikesh@gmail.com'
];

const REMINDER_DAYS_DEFAULT = [60, 45, 30, 15, 10, 5, 2, 0];
const TZ = 'Asia/Kolkata';
const PROP_KEY = 'SMART_MAINTAIN_SPREADSHEET_ID';
const REMINDER_DAYS_KEY = 'SMART_MAINTAIN_REMINDER_DAYS';
const REMINDERS_ENABLED_KEY = 'SMART_MAINTAIN_REMINDERS_ENABLED';

function setupSmartMaintain() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId = props.getProperty(PROP_KEY);
  let ss;

  if (spreadsheetId) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  } else {
    ss = SpreadsheetApp.create('SMART MAINTAIN – Gonda LPG Plant AMC Register');
    props.setProperty(PROP_KEY, ss.getId());
  }

  ensureSheet_(ss, 'Equipment Register', ['Equipment ID', 'Equipment JSON']);
  ensureSheet_(ss, 'AMC Contracts', ['AMC ID', 'Contract JSON']);
  ensureSheet_(ss, 'Maintenance Logs', ['Log ID', 'Log JSON']);
  ensureSheet_(ss, 'Plant Data', ['Data Key', 'Data JSON']);
  ensureSheet_(ss, 'Engineering Advisories', ['Advisory ID', 'Advisory JSON']);
  ensureSheet_(ss, 'Sync Status', ['Key', 'Value']);
  ensureSheet_(ss, 'Email Log', [
    'Sent At (IST)', 'AMC ID', 'Contract Ref', 'Expiry Date', 'Reminder Days', 'Recipients'
  ]);

  // Use the requested reminder pattern as the initial default.
  if (!props.getProperty(REMINDER_DAYS_KEY)) {
    props.setProperty(REMINDER_DAYS_KEY, JSON.stringify(REMINDER_DAYS_DEFAULT));
  }
  if (!props.getProperty(REMINDERS_ENABLED_KEY)) {
    props.setProperty(REMINDERS_ENABLED_KEY, 'true');
  }

  installDailyReminderTrigger_();

  Logger.log('SMART MAINTAIN database: ' + ss.getUrl());
  Logger.log('Daily AMC reminder trigger installed. Time zone: ' + TZ);
  Logger.log('Reminder recipients: ' + RECIPIENTS.join(', '));
  Logger.log('Existing spreadsheet data has been retained.');
  return ss.getUrl();
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.appendRow(headers);
  return sheet;
}

function installDailyReminderTrigger_() {
  const handlers = ['sendDailyMandatoryAmcMail', 'retryDailyMandatoryAmcMail', 'checkAmcExpiry'];
  // Remove duplicate/old triggers for these handlers and reinstall cleanly.
  ScriptApp.getProjectTriggers().forEach(t => {
    if (handlers.includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });

  // Primary mandatory report: approximately 8:00 AM IST every day.
  ScriptApp.newTrigger('sendDailyMandatoryAmcMail')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .inTimezone(TZ)
    .create();

  // Safety retry: if the 8 AM email failed/missed, retry every 15 minutes.
  ScriptApp.newTrigger('retryDailyMandatoryAmcMail')
    .timeBased()
    .everyMinutes(15)
    .create();

  // Optional expiry-date reminders remain independent of the mandatory daily report.
  ScriptApp.newTrigger('checkAmcExpiry')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .inTimezone(TZ)
    .create();
}

function reinstallDailyReminderTrigger() {
  installDailyReminderTrigger_();
  Logger.log('Mandatory daily AMC report trigger installed for 8 AM IST, with 15-minute retry backup.');
}

function getDb_() {
  const id = PropertiesService.getScriptProperties().getProperty(PROP_KEY);
  if (!id) throw new Error('Database is not initialized. Run setupSmartMaintain() first.');
  return SpreadsheetApp.openById(id);
}

function jsonResponse_(result, callback) {
  const json = JSON.stringify(result);
  if (callback && /^[a-zA-Z_$][0-9a-zA-Z_$\.]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  const callback = params.callback || '';
  const action = params.action || '';

  try {
    if (action === 'get') {
      const ss = getDb_();
      return jsonResponse_({
        ok: true,
        assets: readJsonRows_(ss.getSheetByName('Equipment Register')),
        contracts: readJsonRows_(ss.getSheetByName('AMC Contracts')),
        logs: readJsonRows_(ss.getSheetByName('Maintenance Logs')),
        plantData: readPlantData_(ss),
        syncMeta: getSyncMeta_(ss)
      }, callback);
    }

    if (action === 'health') {
      const ss = getDb_();
      const triggers = ScriptApp.getProjectTriggers()
        .filter(t => ['checkAmcExpiry', 'sendDailyMandatoryAmcMail', 'retryDailyMandatoryAmcMail'].includes(t.getHandlerFunction()));
      return jsonResponse_({
        ok: true,
        databaseReady: true,
        triggerInstalled: triggers.length > 0,
        triggerCount: triggers.length,
        timeZone: Session.getScriptTimeZone(),
        contractCount: readJsonRows_(ss.getSheetByName('AMC Contracts')).length,
        syncMeta: getSyncMeta_(ss),
        reminderDays: getReminderDays_(),
        remindersEnabled: remindersEnabled_()
      }, callback);
    }

    throw new Error('Unsupported action.');
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err.message || err) }, callback);
  }
}

/**
 * The website posts JSON as text/plain with mode:no-cors.
 * This endpoint intentionally accepts sync without a login token to match the current
 * open-edit website. Restrict the Apps Script deployment to BPCL users where possible.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    lock.waitLock(25000);
    locked = true;
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!body || typeof body !== 'object') throw new Error('Invalid JSON request body.');
    const ss = getDb_();
    ensureSheet_(ss, 'Equipment Register', ['Equipment ID', 'Equipment JSON']);
    ensureSheet_(ss, 'AMC Contracts', ['AMC ID', 'Contract JSON']);
    ensureSheet_(ss, 'Maintenance Logs', ['Log ID', 'Log JSON']);
    ensureSheet_(ss, 'Plant Data', ['Data Key', 'Data JSON']);
    ensureSheet_(ss, 'Engineering Advisories', ['Advisory ID', 'Advisory JSON']);
    ensureSheet_(ss, 'Sync Status', ['Key', 'Value']);
    ensureSheet_(ss, 'Email Log', [
      'Sent At (IST)', 'AMC ID', 'Contract Ref', 'Expiry Date', 'Reminder Days', 'Recipients'
    ]);

    if (body.action === 'saveAdvisory') {
      if (!body.advisory || !body.advisory.id) throw new Error('Advisory payload or ID is missing.');
      const sheet = ss.getSheetByName('Engineering Advisories');
      const merged = mergeJsonRecordsById_(readJsonRows_(sheet), [body.advisory]);
      writeJsonRows_(sheet, merged, 'id');
      return ContentService.createTextOutput(JSON.stringify({
        ok: true, action: 'saveAdvisory', savedId: String(body.advisory.id), count: merged.length
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (body.action !== 'sync') throw new Error('Unsupported action: ' + String(body.action || '(missing)'));
    if (!Array.isArray(body.contracts)) {
      throw new Error('Website sync rejected: contracts must be an array. The existing cloud AMC register was not changed.');
    }

    const assetsSheet = ss.getSheetByName('Equipment Register');
    const contractsSheet = ss.getSheetByName('AMC Contracts');
    const logsSheet = ss.getSheetByName('Maintenance Logs');
    const mergedAssets = mergeJsonRecordsById_(readJsonRows_(assetsSheet), Array.isArray(body.assets) ? body.assets : []);
    const mergedContracts = mergeJsonRecordsById_(readJsonRows_(contractsSheet), body.contracts);
    const mergedLogs = mergeJsonRecordsById_(readJsonRows_(logsSheet), Array.isArray(body.logs) ? body.logs : []);
    const deletedLogIds = new Set((Array.isArray(body.deletedLogIds) ? body.deletedLogIds : []).map(String));
    const finalLogs = deletedLogIds.size ? mergedLogs.filter(x => !deletedLogIds.has(String(x.id || ''))) : mergedLogs;

    writeJsonRows_(assetsSheet, mergedAssets, 'id');
    writeJsonRows_(contractsSheet, mergedContracts, 'id');
    writeJsonRows_(logsSheet, finalLogs, 'id');

    if (body.plantData && typeof body.plantData === 'object') {
      Object.keys(body.plantData).forEach(key => {
        if (key === 'engineeringAdvisories' && Array.isArray(body.plantData[key])) {
          const sh = ss.getSheetByName('Engineering Advisories');
          writeJsonRows_(sh, mergeJsonRecordsById_(readJsonRows_(sh), body.plantData[key]), 'id');
        } else {
          savePlantData_(ss, key, body.plantData[key]);
        }
      });
    }

    if (body.reminderPrefs) {
      if (body.reminderPrefs.days !== undefined) {
        const parsedDays = parseReminderDays_(body.reminderPrefs.days);
        PropertiesService.getScriptProperties().setProperty(REMINDER_DAYS_KEY, JSON.stringify(parsedDays));
      }
      if (body.reminderPrefs.enabled !== undefined) {
        PropertiesService.getScriptProperties().setProperty(
          REMINDERS_ENABLED_KEY,
          body.reminderPrefs.enabled === true || String(body.reminderPrefs.enabled) === 'true' ? 'true' : 'false'
        );
      }
    }

    const storedAssets = readJsonRows_(assetsSheet).length;
    const storedContracts = readJsonRows_(contractsSheet).length;
    const storedLogs = readJsonRows_(logsSheet).length;
    saveSyncMeta_(ss, {
      requestId: String(body.requestId || ''),
      syncedAt: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
      assets: storedAssets,
      contracts: storedContracts,
      logs: storedLogs,
      source: 'SMART MAINTAIN website sync'
    });

    return ContentService.createTextOutput(JSON.stringify({
      ok: true, action: 'sync', requestId: String(body.requestId || ''),
      assets: storedAssets, contracts: storedContracts, logs: storedLogs,
      syncMeta: getSyncMeta_(ss)
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error('SMART MAINTAIN POST failed: ' + String(err.stack || err));
    return ContentService.createTextOutput(JSON.stringify({
      ok: false, error: String(err.message || err)
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    if (locked) lock.releaseLock();
  }
}

function mergeJsonRecordsById_(existingRows, incomingRows) {
  const byId = new Map();
  const add = row => {
    if (!row || row.id === undefined || row.id === null || String(row.id).trim() === '') return;
    const id = String(row.id);
    const old = byId.get(id);
    if (!old) { byId.set(id, row); return; }
    const oldTime = Date.parse(old.updatedAt || '') || 0;
    const newTime = Date.parse(row.updatedAt || '') || 0;
    if (newTime >= oldTime) byId.set(id, row);
  };
  (Array.isArray(existingRows) ? existingRows : []).forEach(add);
  (Array.isArray(incomingRows) ? incomingRows : []).forEach(add);
  return Array.from(byId.values());
}

function getSyncMeta_(ss) {
  const sheet = ss.getSheetByName('Sync Status');
  const out = {};
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(r => {
      if (r[0]) out[String(r[0])] = r[1];
    });
  }
  return {
    requestId: out.requestId || '',
    updatedAt: out.syncedAt || '',
    contractCount: Number(out.contracts || 0),
    assetCount: Number(out.assets || 0),
    logCount: Number(out.logs || 0),
    source: out.source || 'No website sync recorded'
  };
}

function saveSyncMeta_(ss, meta) {
  const sheet = ensureSheet_(ss, 'Sync Status', ['Key', 'Value']);
  const values = Object.keys(meta).map(k => [k, meta[k]]);
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
  if (values.length) sheet.getRange(2, 1, values.length, 2).setValues(values);
}

function readPlantData_(ss) {
  const out = {};
  const sheet = ss.getSheetByName('Plant Data');
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(row => {
      const key = String(row[0] || '');
      if (!key) return;
      try { out[key] = JSON.parse(row[1]); } catch (_) { out[key] = row[1]; }
    });
  }
  // Engineering advisories are returned in the exact shape expected by the website.
  out.engineeringAdvisories = readJsonRows_(ss.getSheetByName('Engineering Advisories'));
  return out;
}

function savePlantData_(ss, key, value) {
  const sheet = ensureSheet_(ss, 'Plant Data', ['Data Key', 'Data JSON']);
  const last = sheet.getLastRow();
  if (last > 1) {
    const keys = sheet.getRange(2, 1, last - 1, 1).getValues().flat();
    const found = keys.findIndex(x => String(x) === String(key));
    if (found >= 0) {
      sheet.getRange(found + 2, 2).setValue(JSON.stringify(value));
      return;
    }
  }
  sheet.appendRow([String(key), JSON.stringify(value)]);
}

function readJsonRows_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues()
    .map(row => {
      try { return JSON.parse(row[0]); } catch (_) { return null; }
    })
    .filter(Boolean);
}

function writeJsonRows_(sheet, items, key) {
  const safeItems = Array.isArray(items) ? items : [];
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
  }
  if (!safeItems.length) return;
  const rows = safeItems.map(item => [String(item[key] || ''), JSON.stringify(item)]);
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
}

function parseReminderDays_(value) {
  const values = Array.isArray(value) ? value : String(value || '').split(',');
  const days = values.map(x => Number(String(x).trim()));
  if (!days.length || days.some(x => !Number.isInteger(x) || x < 0 || x > 3650)) {
    throw new Error('Reminder days must be whole numbers from 0 to 3650.');
  }
  return [...new Set(days)].sort((a, b) => b - a);
}

function getReminderDays_() {
  const saved = PropertiesService.getScriptProperties().getProperty(REMINDER_DAYS_KEY);
  if (!saved) return REMINDER_DAYS_DEFAULT;
  try { return parseReminderDays_(JSON.parse(saved)); }
  catch (_) { return REMINDER_DAYS_DEFAULT; }
}

function remindersEnabled_() {
  return PropertiesService.getScriptProperties().getProperty(REMINDERS_ENABLED_KEY) !== 'false';
}

function parseExpiryDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TZ, 'yyyy-MM-dd');
  }
  const raw = String(value || '').trim();
  if (!raw) return null;

  // Preferred website format: yyyy-MM-dd.
  let m = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + String(m[2]).padStart(2, '0') + '-' + String(m[3]).padStart(2, '0');

  // Also accept dd-MM-yyyy or dd/MM/yyyy from manually imported legacy records.
  m = raw.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (m) return m[3] + '-' + String(m[2]).padStart(2, '0') + '-' + String(m[1]).padStart(2, '0');

  return null;
}

function checkAmcExpiry() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    Logger.log('Another AMC reminder check is already running. This run was skipped.');
    return;
  }

  try {
    if (!remindersEnabled_()) {
      Logger.log('AMC email reminders are disabled in website preferences.');
      return;
    }

    const ss = getDb_();
    const contracts = readJsonRows_(ss.getSheetByName('AMC Contracts'));
    const logs = readJsonRows_(ss.getSheetByName('Maintenance Logs'));
    const emailLog = ensureSheet_(ss, 'Email Log', [
      'Sent At (IST)', 'AMC ID', 'Contract Ref', 'Expiry Date', 'Reminder Days', 'Recipients'
    ]);

    const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    const todayUtc = Date.parse(today + 'T00:00:00Z');
    const reminderDays = getReminderDays_();
    const alreadySent = new Set();

    if (emailLog.getLastRow() > 1) {
      emailLog.getRange(2, 2, emailLog.getLastRow() - 1, 4).getValues().forEach(row => {
        // Columns B-E: AMC ID, Contract Ref, Expiry Date, Reminder Days.
        alreadySent.add([row[0], row[2], row[3]].join('|'));
      });
    }

    const due = [];
    contracts.forEach(c => {
      if (!c || !c.id) return;
      const expiry = parseExpiryDate_(c.end);
      if (!expiry) return;

      const days = Math.round((Date.parse(expiry + 'T00:00:00Z') - todayUtc) / 86400000);

      // Before expiry: only configured reminder dates.
      // On expiry day (0) and every day after expiry (-1, -2, -3 ...): send daily.
      if (days > 0 && !reminderDays.includes(days)) return;

      const dedupeKey = [c.id, expiry, days].join('|');
      if (alreadySent.has(dedupeKey)) return;

      const visitsGiven = logs.filter(l =>
        l && l.type === 'AMC Visit' && String(l.contract || '') === String(c.id)
      ).length;

      const agreed = c.totalVisits === '' || c.totalVisits == null ? '' : Number(c.totalVisits);
      const remaining = agreed === '' ? '' : Math.max(0, agreed - visitsGiven);

      due.push({
        id: c.id,
        ref: c.ref || c.id,
        equipment: c.equipment || '',
        oem: c.oem || c.agency || '',
        start: parseExpiryDate_(c.start) || '',
        expiry: expiry,
        days: days,
        po: c.po || '',
        ola: c.ola || '',
        totalVisits: agreed,
        visitsGiven: visitsGiven,
        visitsRemaining: remaining
      });
    });

    if (!due.length) {
      Logger.log('No AMC reminders due today (' + today + '). Contracts checked: ' + contracts.length);
      return;
    }

    const subject = 'SMART MAINTAIN | AMC expiry reminder | Gonda LPG Plant';
    const lines = [
      'Dear Team Gonda,',
      '',
      'The following AMC contract(s) have reached a configured reminder date.',
      'Reminder dates are calculated from the exact validity end date entered in the contract.',
      ''
    ];

    due.forEach((x, i) => {
      lines.push((i + 1) + '. Equipment: ' + x.equipment);
      lines.push('   Contract reference: ' + x.ref);
      lines.push('   OEM / Agency: ' + x.oem);
      lines.push('   Validity start: ' + x.start);
      lines.push('   Exact validity end: ' + x.expiry);
      lines.push('   Days to expiry: ' + x.days);
      if (x.po) lines.push('   PO: ' + x.po);
      if (x.ola) lines.push('   OLA: ' + x.ola);
      if (x.totalVisits !== '') {
        lines.push('   Visits agreed / given / remaining: ' +
          x.totalVisits + ' / ' + x.visitsGiven + ' / ' + x.visitsRemaining);
      }
      lines.push('');
    });

    lines.push(
      'Please review the contract and arrange the required action.',
      '',
      'Regards,',
      'SMART MAINTAIN – Gonda LPG Plant'
    );

    const remainingQuota = MailApp.getRemainingDailyQuota();
    if (remainingQuota < RECIPIENTS.length) {
      throw new Error('Mail quota is insufficient. Remaining recipient quota: ' + remainingQuota);
    }

    MailApp.sendEmail({
      to: RECIPIENTS.join(','),
      subject: subject,
      body: lines.join('\n'),
      name: 'SMART MAINTAIN – Gonda LPG Plant'
    });

    const sentAt = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
    const rows = due.map(x => [sentAt, x.id, x.ref, x.expiry, x.days, RECIPIENTS.join(', ')]);
    emailLog.getRange(emailLog.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

    Logger.log('SUCCESS: sent one AMC reminder email to all five recipients for ' +
      due.length + ' contract(s).');
  } catch (err) {
    console.error('AMC reminder email failed: ' + String(err.stack || err));
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/**
 * MANDATORY DAILY AMC STATUS EMAIL.
 * Sends every day, even when there are no expiry reminders due.
 * A successful send is recorded by date to prevent duplicate daily emails.
 */
function sendDailyMandatoryAmcMail() {
  return sendDailyMandatoryAmcMail_(false);
}

/** Backup trigger: retries until today's mandatory report has successfully sent. */
function retryDailyMandatoryAmcMail() {
  return sendDailyMandatoryAmcMail_(true);
}

function sendDailyMandatoryAmcMail_(isRetry) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    Logger.log('Mandatory AMC mail: another execution is running; retry will handle it.');
    return;
  }

  try {
    const props = PropertiesService.getScriptProperties();
    const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    const successKey = 'SMART_MAINTAIN_MANDATORY_AMC_MAIL_SENT_' + today;

    if (props.getProperty(successKey) === 'true') {
      Logger.log('Mandatory AMC mail already sent successfully today (' + today + ').');
      return;
    }

    const ss = getDb_();
    const contracts = readJsonRows_(ss.getSheetByName('AMC Contracts'));
    const logs = readJsonRows_(ss.getSheetByName('Maintenance Logs'));
    const emailLog = ensureSheet_(ss, 'Email Log', [
      'Sent At (IST)', 'AMC ID', 'Contract Ref', 'Expiry Date', 'Reminder Days', 'Recipients'
    ]);

    const todayUtc = Date.parse(today + 'T00:00:00Z');
    const rows = contracts.map(c => {
      const expiry = parseExpiryDate_(c.end);
      const days = expiry ? Math.round((Date.parse(expiry + 'T00:00:00Z') - todayUtc) / 86400000) : null;
      const visitsGiven = logs.filter(l =>
        l && l.type === 'AMC Visit' && String(l.contract || '') === String(c.id || '')
      ).length;
      const agreed = c.totalVisits === '' || c.totalVisits == null ? null : Number(c.totalVisits);
      return {
        id: c.id || '',
        ref: c.ref || c.id || '',
        equipment: c.equipment || 'Not specified',
        oem: c.oem || c.agency || 'Not specified',
        start: parseExpiryDate_(c.start) || 'Not specified',
        expiry: expiry || 'Not specified',
        days: days,
        po: c.po || '',
        ola: c.ola || '',
        totalVisits: agreed,
        visitsGiven: visitsGiven,
        visitsRemaining: agreed === null ? null : Math.max(0, agreed - visitsGiven)
      };
    });

    const active = rows.filter(x => x.expiry !== 'Not specified' && x.days >= 0);
    const expired = rows.filter(x => x.expiry !== 'Not specified' && x.days < 0);
    const noExpiry = rows.filter(x => x.expiry === 'Not specified');
    const lines = [
      'Dear Team Gonda,',
      '',
      'MANDATORY DAILY AMC STATUS REPORT – GONDA LPG PLANT',
      'Report date: ' + today,
      'This email is automatically issued every day at 8:00 AM IST.',
      'Website data last synced: ' + (getSyncMeta_(ss).updatedAt || 'NEVER / NOT RECORDED'),
      'Last website sync contract count: ' + getSyncMeta_(ss).contractCount,
      'Data source: shared Google Sheets snapshot received from SMART MAINTAIN website.',
      '',
      'SUMMARY',
      'Total AMC contracts: ' + rows.length,
      'Active / not yet expired: ' + active.length,
      'Expired contracts: ' + expired.length,
      'Contracts with missing/invalid expiry date: ' + noExpiry.length,
      '',
      'REGISTERED AMC EQUIPMENT – EXPIRY DATE LIST (ALL CONTRACTS)',
      'Equipment | AMC / Contract Reference | Valid Till | Status'
    ];

    // Show every registered AMC in a concise expiry list, earliest expiry first.
    const expirySortedRows = rows.slice().sort((a, b) => {
      if (a.expiry === 'Not specified' && b.expiry !== 'Not specified') return 1;
      if (b.expiry === 'Not specified' && a.expiry !== 'Not specified') return -1;
      if (a.expiry === 'Not specified' && b.expiry === 'Not specified') return String(a.equipment).localeCompare(String(b.equipment));
      return a.expiry.localeCompare(b.expiry);
    });

    if (!rows.length) {
      const syncInfo = getSyncMeta_(ss);
      lines.push('WARNING: The cloud register currently contains ZERO AMC contracts.');
      lines.push('This may mean the website AMC records have not synchronized to the shared database.');
      lines.push('Last website sync: ' + (syncInfo.updatedAt || 'NEVER'));
      lines.push('Last website sync reported contract count: ' + syncInfo.contractCount);
      lines.push('ACTION: Open SMART MAINTAIN website, verify AMC list, and press Sync AMC records now.');
    } else {
      expirySortedRows.forEach((x, i) => {
        const status = x.expiry === 'Not specified' ? 'EXPIRY NOT ENTERED' :
          x.days < 0 ? 'EXPIRED' : x.days === 0 ? 'EXPIRES TODAY' : 'ACTIVE';
        lines.push((i + 1) + '. ' + x.equipment + ' | ' + x.ref + ' | ' + x.expiry + ' | ' + status);
      });
      lines.push('', 'FULL AMC CONTRACT DETAILS');
      rows.forEach((x, i) => {
        lines.push('');
        lines.push((i + 1) + '. Equipment: ' + x.equipment);
        lines.push('   Contract / AMC Ref: ' + x.ref);
        lines.push('   OEM / Agency: ' + x.oem);
        lines.push('   Valid from: ' + x.start);
        lines.push('   Valid till: ' + x.expiry);
        if (x.days !== null) {
          lines.push('   Days to expiry: ' + x.days + (x.days < 0 ? ' (EXPIRED)' : x.days === 0 ? ' (EXPIRES TODAY)' : ''));
        }
        if (x.po) lines.push('   PO: ' + x.po);
        if (x.ola) lines.push('   OLA: ' + x.ola);
        if (x.totalVisits !== null) {
          lines.push('   Visits agreed / completed / remaining: ' +
            x.totalVisits + ' / ' + x.visitsGiven + ' / ' + x.visitsRemaining);
        }
      });
    }

    lines.push('', 'Please review the AMC status and take necessary action wherever required.',
      '', 'Regards,', 'SMART MAINTAIN – Gonda LPG Plant');

    const quota = MailApp.getRemainingDailyQuota();
    if (quota < RECIPIENTS.length) {
      throw new Error('Insufficient email quota. Remaining quota: ' + quota);
    }

    MailApp.sendEmail({
      to: RECIPIENTS.join(','),
      subject: 'MANDATORY DAILY AMC STATUS | ' + today + ' | Gonda LPG Plant',
      body: lines.join('\n'),
      name: 'SMART MAINTAIN – Gonda LPG Plant'
    });

    props.setProperty(successKey, 'true');
    props.setProperty('SMART_MAINTAIN_MANDATORY_AMC_LAST_SUCCESS', today);
    props.setProperty('SMART_MAINTAIN_MANDATORY_AMC_LAST_SENT_AT', Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'));

    emailLog.appendRow([
      Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
      'DAILY-MANDATORY-REPORT',
      'Daily AMC Status',
      today,
      'MANDATORY DAILY',
      RECIPIENTS.join(', ')
    ]);

    Logger.log('SUCCESS: Mandatory daily AMC status email sent for ' + today + '. Retry=' + isRetry);
  } catch (err) {
    PropertiesService.getScriptProperties().setProperty(
      'SMART_MAINTAIN_MANDATORY_AMC_LAST_FAILURE',
      Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss') + ' | ' + String(err.message || err)
    );
    console.error('MANDATORY DAILY AMC MAIL FAILED: ' + String(err.stack || err));
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/** Run manually once to authorise MailApp and verify delivery to all five recipients. */
function sendTestEmail() {
  const quota = MailApp.getRemainingDailyQuota();
  if (quota < RECIPIENTS.length) {
    throw new Error('Mail quota is insufficient. Remaining recipient quota: ' + quota);
  }

  MailApp.sendEmail({
    to: RECIPIENTS.join(','),
    subject: 'SMART MAINTAIN – TEST EMAIL – Gonda LPG Plant',
    body: [
      'Dear Team Gonda,',
      '',
      'This is a delivery test from SMART MAINTAIN – Gonda LPG Plant.',
      'If you receive this message, Google Apps Script MailApp is authorised and can send to the configured recipients.',
      '',
      'Recipients:',
      ...RECIPIENTS,
      '',
      'Regards,',
      'SMART MAINTAIN'
    ].join('\n'),
    name: 'SMART MAINTAIN – Gonda LPG Plant'
  });

  Logger.log('TEST EMAIL SENT to: ' + RECIPIENTS.join(', '));
}

/** Manual verification: confirms cloud AMC count and last website sync timestamp. */
function diagnoseAmcSync() {
  const ss = getDb_();
  const contracts = readJsonRows_(ss.getSheetByName('AMC Contracts'));
  const meta = getSyncMeta_(ss);
  Logger.log('Cloud AMC contracts: ' + contracts.length);
  Logger.log('Last website sync: ' + (meta.updatedAt || 'NEVER'));
  Logger.log('Last sync request ID: ' + (meta.requestId || 'none'));
  Logger.log('Last sync reported AMC count: ' + meta.contractCount);
  Logger.log('AMC records in database:');
  contracts.forEach((c, i) => Logger.log((i + 1) + '. ' + (c.equipment || c.id || 'Unknown') + ' | ' + (c.ref || c.id || '') + ' | expiry=' + (c.end || 'missing')));
}

/** Run manually to inspect the database, trigger, and remaining email quota. */
function diagnoseSmartMaintain() {
  const ss = getDb_();
  const triggers = ScriptApp.getProjectTriggers()
    .filter(t => ['checkAmcExpiry', 'sendDailyMandatoryAmcMail', 'retryDailyMandatoryAmcMail'].includes(t.getHandlerFunction()));
  const contracts = readJsonRows_(ss.getSheetByName('AMC Contracts'));
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  Logger.log('Spreadsheet: ' + ss.getUrl());
  Logger.log('Contracts currently stored in cloud: ' + contracts.length);
  Logger.log('Mandatory daily AMC mail trigger count: ' + triggers.filter(t => t.getHandlerFunction() === 'sendDailyMandatoryAmcMail').length);
  Logger.log('Mandatory mail retry trigger count: ' + triggers.filter(t => t.getHandlerFunction() === 'retryDailyMandatoryAmcMail').length);
  Logger.log('AMC expiry reminder trigger count: ' + triggers.filter(t => t.getHandlerFunction() === 'checkAmcExpiry').length);
  Logger.log('Script time zone: ' + Session.getScriptTimeZone());
  Logger.log('Configured mail time zone: ' + TZ);
  Logger.log('Mandatory mail sent today: ' + (props.getProperty('SMART_MAINTAIN_MANDATORY_AMC_MAIL_SENT_' + today) === 'true'));
  Logger.log('Last mandatory mail success: ' + (props.getProperty('SMART_MAINTAIN_MANDATORY_AMC_LAST_SUCCESS') || 'None'));
  Logger.log('Last mandatory mail failure: ' + (props.getProperty('SMART_MAINTAIN_MANDATORY_AMC_LAST_FAILURE') || 'None'));
  Logger.log('Reminder days: ' + JSON.stringify(getReminderDays_()));
  Logger.log('Expiry reminders enabled: ' + remindersEnabled_());
  Logger.log('Remaining daily email recipient quota: ' + MailApp.getRemainingDailyQuota());
  if (triggers.filter(t => t.getHandlerFunction() === 'sendDailyMandatoryAmcMail').length === 0) {
    Logger.log('ACTION REQUIRED: Run setupSmartMaintain() to install the mandatory 8 AM trigger.');
  }
}