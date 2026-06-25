/**
 * Unified Solutions — BD Feed
 * Runs on your Google account, scans Gmail + Calendar, and writes "signals"
 * to your Supabase project. Your dashboard reads those signals and advances
 * cards on its own (reply detected -> Replied, call booked -> Call booked),
 * and can auto-capture new prospects from an optional intake Sheet.
 *
 * SETUP (one time):
 *  1. Run the SQL (provided in chat) in Supabase to create the bd_signals table.
 *  2. Project Settings -> Script Properties, add:
 *       SUPABASE_URL          = https://YOURREF.supabase.co
 *       SUPABASE_SERVICE_KEY  = your SECRET key (sb_secret_...). Server-side only. Never in the website.
 *       USER_ID               = your Supabase auth user id (see chat for how to find it)
 *       INTAKE_SHEET_ID       = (optional) a Google Sheet id for new-prospect intake
 *  3. Run setup() once and authorize Gmail + Calendar when prompted.
 *  4. Run runFeed() once manually to confirm it works, then it runs hourly.
 */

function CONFIG_() {
  var p = PropertiesService.getScriptProperties();
  return {
    url: p.getProperty('SUPABASE_URL'),
    key: p.getProperty('SUPABASE_SERVICE_KEY'),
    userId: p.getProperty('USER_ID'),
    intakeSheetId: p.getProperty('INTAKE_SHEET_ID')
  };
}

function headers_(key) {
  return { 'apikey': key, 'Authorization': 'Bearer ' + key, 'Content-Type': 'application/json' };
}

function getRoster_(c) {
  var url = c.url + '/rest/v1/bd_state?user_id=eq.' + c.userId + '&select=data';
  var res = UrlFetchApp.fetch(url, { headers: headers_(c.key), muteHttpExceptions: true });
  var rows = JSON.parse(res.getContentText() || '[]');
  if (!rows.length || !rows[0].data) return [];
  return rows[0].data.roster || [];
}

function seenDedupes_(c) {
  var url = c.url + '/rest/v1/bd_signals?user_id=eq.' + c.userId + '&select=dedupe&order=created_at.desc&limit=800';
  var res = UrlFetchApp.fetch(url, { headers: headers_(c.key), muteHttpExceptions: true });
  var rows = JSON.parse(res.getContentText() || '[]');
  var set = {};
  rows.forEach(function (r) { if (r.dedupe) set[r.dedupe] = true; });
  return set;
}

function writeSignals_(c, signals) {
  if (!signals.length) return;
  var url = c.url + '/rest/v1/bd_signals?on_conflict=user_id,dedupe';
  var h = headers_(c.key);
  h['Prefer'] = 'resolution=ignore-duplicates,return=minimal';
  UrlFetchApp.fetch(url, { method: 'post', headers: h, payload: JSON.stringify(signals), muteHttpExceptions: true });
}

function parseEmail_(s) {
  var m = String(s).match(/<([^>]+)>/);
  if (m) return m[1].toLowerCase();
  s = String(s).trim();
  return s.indexOf('@') >= 0 ? s.toLowerCase() : null;
}

function runFeed() {
  var c = CONFIG_();
  if (!c.url || !c.key || !c.userId) { Logger.log('Missing Script Properties.'); return; }

  var roster = getRoster_(c);
  var seen = seenDedupes_(c);
  var signals = [];

  var byEmail = {};
  roster.forEach(function (p) { if (p.email) byEmail[p.email.toLowerCase()] = p; });

  // --- Gmail: inbound messages from prospect emails in the last 14 days = a reply/engagement ---
  var threads = GmailApp.search('newer_than:14d -in:chats -in:sent', 0, 100);
  threads.forEach(function (th) {
    th.getMessages().forEach(function (m) {
      var from = parseEmail_(m.getFrom());
      if (!from || !byEmail[from]) return;
      var dedupe = 'reply:' + m.getId();
      if (seen[dedupe]) return;
      signals.push({
        user_id: c.userId, kind: 'reply', email: from,
        hint: th.getFirstMessageSubject(), detail: m.getPlainBody().slice(0, 200), dedupe: dedupe
      });
    });
  });

  // --- Calendar: events in the next 30 days matching a prospect by attendee or title ---
  var now = new Date();
  var end = new Date(now.getTime() + 30 * 24 * 3600 * 1000);
  var events = CalendarApp.getDefaultCalendar().getEvents(now, end);
  events.forEach(function (ev) {
    var guests = ev.getGuestList().map(function (g) { return g.getEmail().toLowerCase(); });
    var matchEmail = null, matchHint = null;
    var title = ev.getTitle() || '';
    roster.forEach(function (p) {
      if (p.email && guests.indexOf(p.email.toLowerCase()) >= 0) matchEmail = p.email;
      if (p.company && title.toLowerCase().indexOf(p.company.toLowerCase()) >= 0) matchHint = p.company;
    });
    if (!matchEmail && !matchHint) return;
    var dedupe = 'meeting:' + ev.getId();
    if (seen[dedupe]) return;
    signals.push({
      user_id: c.userId, kind: 'meeting', email: matchEmail || '',
      hint: matchHint || title, detail: ev.getStartTime().toISOString(), dedupe: dedupe
    });
  });

  // --- Optional: new-prospect intake from a Google Sheet ---
  // Sheet columns (row 1 is a header): Company | Contact | Email | Segment | Trigger | Why | Status
  if (c.intakeSheetId) {
    try {
      var sh = SpreadsheetApp.openById(c.intakeSheetId).getSheets()[0];
      var data = sh.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        if (!row[0] || row[6]) continue; // no company, or already processed
        var dedupe = 'prospect:' + String(row[0]).toLowerCase().trim();
        if (seen[dedupe]) { sh.getRange(i + 1, 7).setValue('dup'); continue; }
        signals.push({
          user_id: c.userId, kind: 'prospect',
          company: row[0], contact: row[1] || '', email: row[2] || '',
          seg: String(row[3] || 'B').toUpperCase().slice(0, 1),
          trigger: row[4] || 'other', detail: row[5] || '', dedupe: dedupe
        });
        sh.getRange(i + 1, 7).setValue('queued');
      }
    } catch (e) { Logger.log('Intake error: ' + e); }
  }

  writeSignals_(c, signals);
  Logger.log('Wrote ' + signals.length + ' signal(s).');
}

function setup() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('runFeed').timeBased().everyHours(1).create();
  Logger.log('Hourly trigger created. Now run runFeed() once to authorize Gmail and Calendar.');
}
