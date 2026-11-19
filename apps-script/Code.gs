/**
 * Kovie Student Actor Signup: backend (Google Apps Script web app).
 *
 * The public page on GitHub Pages calls this web app. It:
 *   - reports which seats are open, held or taken (names only, never emails)
 *   - places a 5 minute hold on a seat while someone fills in the form
 *   - books the seat, writes it to the Google Sheet, sends a confirmation
 *     email and a Google Calendar invite
 *   - lets the actor cancel through the link in their confirmation email
 *
 * The bound Google Sheet is the record. "Signups" is the raw log; "Roster" is
 * rebuilt after every change in the same layout as the original spreadsheet.
 */

var CONFIG = {
  eventTitle: 'Kovie Student Actor Signup',
  dateLine: 'Tuesday, December 1, 2026',
  location: 'Union Mall, near the fountain',
  timeZone: 'America/Chicago',
  contact: 'Kate Carter at 832-773-8129',
  // Replies to any email this sends go to Kate, not the sending Gmail account.
  replyTo: 'klc060@uark.edu',
  // Reminder email to everyone still signed up.
  reminderAt: '2026-11-30T17:00:00-06:00',
  siteUrl: 'https://kea1e.github.io/kovie-actor-signup/',
  calendarName: 'Kovie Actor Signups',
  emailDomain: 'uark.edu',
  seatsPerSlot: 2,
  holdSeconds: 300,
  // Signups close when the event starts.
  closesAt: '2026-12-01T11:00:00-06:00',
  slots: [
    { id: 's1100', label: '11:00 to 11:30 a.m.', start: '2026-12-01T11:00:00-06:00', end: '2026-12-01T11:30:00-06:00' },
    { id: 's1130', label: '11:30 a.m. to 12:00 p.m.', start: '2026-12-01T11:30:00-06:00', end: '2026-12-01T12:00:00-06:00' },
    { id: 's1200', label: '12:00 to 12:30 p.m.', start: '2026-12-01T12:00:00-06:00', end: '2026-12-01T12:30:00-06:00' },
    { id: 's1230', label: '12:30 to 1:00 p.m.', start: '2026-12-01T12:30:00-06:00', end: '2026-12-01T13:00:00-06:00' },
    { id: 's1300', label: '1:00 to 1:30 p.m.', start: '2026-12-01T13:00:00-06:00', end: '2026-12-01T13:30:00-06:00' },
    { id: 's1330', label: '1:30 to 2:00 p.m.', start: '2026-12-01T13:30:00-06:00', end: '2026-12-01T14:00:00-06:00' }
  ]
};

/**
 * Run from the script editor (select "setup", click Run). It creates the sheet
 * tabs, the "Kovie Actor Signups" calendar, schedules the reminder email, and
 * grants the permissions the web app needs. Safe to re-run.
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(CONFIG.timeZone);
  var su = ss.getSheetByName(SIGNUPS) || ss.insertSheet(SIGNUPS);
  su.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#1b3a4b').setFontColor('#ffffff');
  su.setFrozenRows(1);
  su.getRange('F:F').setNumberFormat('mmm d, h:mm am/pm');

  var ro = ss.getSheetByName(ROSTER) || ss.insertSheet(ROSTER, 0);
  ro.getRange(1, 1).setValue(CONFIG.eventTitle + ' | ' + CONFIG.dateLine + ' | 11:00 a.m. to 2:00 p.m. | ' + CONFIG.location).setFontWeight('bold').setFontSize(13);
  ro.getRange(3, 1, 1, 5).setValues([['Performance slot', 'Actor 1: full name', 'Actor 1: UARK email', 'Actor 2: full name', 'Actor 2: UARK email']])
    .setFontWeight('bold').setBackground('#1b3a4b').setFontColor('#ffffff');
  ro.getRange(4, 1, CONFIG.slots.length, 5).setBackground('#fdf1dc');
  ro.setColumnWidth(1, 190);
  [2, 3, 4, 5].forEach(function (c) { ro.setColumnWidth(c, 210); });
  ro.setFrozenRows(3);

  var blank = ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 2 && blank.getLastRow() === 0) ss.deleteSheet(blank);

  calendarId_();
  renderRoster_();
  // Touch MailApp so its permission is granted during this one-time run.
  console.log('Mail quota remaining today: ' + MailApp.getRemainingDailyQuota());
  console.log('Calendar: ' + calendarId_());
  installReminder();
}

var SIGNUPS = 'Signups';
var ROSTER = 'Roster';
var HEADERS = ['Booking ID', 'Slot', 'Seat', 'Full name', 'UARK email', 'Booked at', 'Status', 'Cancelled at', 'Calendar event ID', 'Cancel token', 'Reminder sent'];
var COL = { id: 0, slot: 1, seat: 2, name: 3, email: 4, bookedAt: 5, status: 6, cancelledAt: 7, eventId: 8, token: 9, reminder: 10 };

// ---------------------------------------------------------------- HTTP

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'lookup') return json_(lookup_(p.token));
    return json_(state_());
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (x) {}
  try {
    switch (body.action) {
      case 'hold': return json_(withLock_(function () { return hold_(body); }));
      case 'release': return json_(withLock_(function () { return release_(body); }));
      case 'book': return json_(withLock_(function () { return book_(body); }));
      case 'cancel': return json_(cancel_(body));
      default: return json_({ ok: false, error: 'Unknown action.' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, error: 'The signup sheet is busy. Please try again in a moment.' };
  try { return fn(); } finally { lock.releaseLock(); }
}

// ---------------------------------------------------------------- state

function slotById_(id) {
  for (var i = 0; i < CONFIG.slots.length; i++) if (CONFIG.slots[i].id === id) return CONFIG.slots[i];
  return null;
}

function isClosed_() {
  return new Date() >= new Date(CONFIG.closesAt);
}

function holdKey_(slotId, seat) { return 'hold:' + slotId + ':' + seat; }

function activeHolds_() {
  var keys = [];
  CONFIG.slots.forEach(function (s) {
    for (var seat = 1; seat <= CONFIG.seatsPerSlot; seat++) keys.push(holdKey_(s.id, seat));
  });
  var raw = CacheService.getScriptCache().getAll(keys);
  var now = Date.now();
  var out = {};
  Object.keys(raw).forEach(function (k) {
    try {
      var h = JSON.parse(raw[k]);
      if (h.expiresAt > now) out[k] = h;
    } catch (x) {}
  });
  return out;
}

function sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Missing sheet tab "' + name + '". Run setup() once from the script editor.');
  return sh;
}

function confirmedRows_() {
  var sh = sheet_(SIGNUPS);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  var rows = [];
  values.forEach(function (v, i) {
    if (v[COL.status] === 'Confirmed') rows.push({ row: i + 2, v: v });
  });
  return rows;
}

function state_() {
  var taken = {};
  confirmedRows_().forEach(function (r) { taken[r.v[COL.slot] + ':' + r.v[COL.seat]] = r.v[COL.name]; });
  var holds = activeHolds_();
  return {
    ok: true,
    closed: isClosed_(),
    holdSeconds: CONFIG.holdSeconds,
    slots: CONFIG.slots.map(function (s) {
      var seats = [];
      for (var seat = 1; seat <= CONFIG.seatsPerSlot; seat++) {
        var key = s.id + ':' + seat;
        if (taken[key]) seats.push({ seat: seat, status: 'taken', name: String(taken[key]) });
        else if (holds[holdKey_(s.id, seat)]) seats.push({ seat: seat, status: 'held' });
        else seats.push({ seat: seat, status: 'open' });
      }
      return { id: s.id, label: s.label, seats: seats };
    })
  };
}

// ---------------------------------------------------------------- hold / release

function checkSeat_(slotId, seat) {
  var slot = slotById_(slotId);
  seat = Number(seat);
  if (!slot || !(seat >= 1 && seat <= CONFIG.seatsPerSlot)) throw new Error('That seat does not exist.');
  return { slot: slot, seat: seat };
}

function seatTaken_(slotId, seat) {
  return confirmedRows_().some(function (r) { return r.v[COL.slot] === slotId && Number(r.v[COL.seat]) === seat; });
}

function hold_(body) {
  if (isClosed_()) return { ok: false, error: 'Signups are closed.' };
  var s = checkSeat_(body.slotId, body.seat);
  if (seatTaken_(s.slot.id, s.seat)) return { ok: false, error: 'Someone just took that seat. Pick another one.' };
  var cache = CacheService.getScriptCache();
  var key = holdKey_(s.slot.id, s.seat);
  var existing = activeHolds_()[key];
  if (existing && existing.holdId !== body.holdId) return { ok: false, error: 'Someone is signing up for that seat right now. Pick another one or check back in a few minutes.' };
  var holdId = existing ? existing.holdId : Utilities.getUuid();
  var expiresAt = Date.now() + CONFIG.holdSeconds * 1000;
  cache.put(key, JSON.stringify({ holdId: holdId, expiresAt: expiresAt }), CONFIG.holdSeconds);
  return { ok: true, holdId: holdId, expiresAt: expiresAt, slotLabel: s.slot.label, seat: s.seat };
}

function release_(body) {
  var s = checkSeat_(body.slotId, body.seat);
  var key = holdKey_(s.slot.id, s.seat);
  var existing = activeHolds_()[key];
  if (existing && existing.holdId === body.holdId) CacheService.getScriptCache().remove(key);
  return { ok: true };
}

// ---------------------------------------------------------------- book

function overlaps_(a, b) {
  return new Date(a.start) < new Date(b.end) && new Date(b.start) < new Date(a.end);
}

function cleanName_(n) {
  return String(n || '').replace(/[\u0000-\u001f<>]/g, '').replace(/\s+/g, ' ').trim().slice(0, 80);
}

function book_(body) {
  if (isClosed_()) return { ok: false, error: 'Signups are closed.' };
  var s = checkSeat_(body.slotId, body.seat);
  var name = cleanName_(body.name);
  var email = String(body.email || '').trim().toLowerCase();
  if (name.length < 2 || name.indexOf(' ') < 0) return { ok: false, error: 'Please enter your full name (first and last).' };
  var re = new RegExp('^[a-z0-9._%+-]+@' + CONFIG.emailDomain.replace('.', '\\.') + '$');
  if (!re.test(email)) return { ok: false, error: 'Please use your UARK email (ends in @' + CONFIG.emailDomain + ').' };

  if (seatTaken_(s.slot.id, s.seat)) return { ok: false, error: 'Someone just took that seat. Pick another one.' };
  var key = holdKey_(s.slot.id, s.seat);
  var hold = activeHolds_()[key];
  if (hold && hold.holdId !== body.holdId) return { ok: false, error: 'Your hold on this seat expired and someone else is signing up for it. Pick another seat.' };

  // One person can take several slots, but never two that overlap in time
  // (which includes both seats of the same slot).
  var clash = confirmedRows_().filter(function (r) {
    if (String(r.v[COL.email]).toLowerCase() !== email) return false;
    var other = slotById_(r.v[COL.slot]);
    return other && overlaps_(other, s.slot);
  })[0];
  if (clash) {
    var taken = slotById_(clash.v[COL.slot]);
    return { ok: false, error: 'You are already signed up for ' + taken.label + ', which is the same time. Pick a different time slot.' };
  }

  var id = 'K' + Utilities.formatDate(new Date(), CONFIG.timeZone, 'MMddHHmmss') + Math.floor(Math.random() * 90 + 10);
  var token = Utilities.getUuid().replace(/-/g, '');
  var sh = sheet_(SIGNUPS);
  sh.appendRow([id, s.slot.id, s.seat, name, email, new Date(), 'Confirmed', '', '', token]);
  var row = sh.getLastRow();
  CacheService.getScriptCache().remove(key);

  // Calendar invite and email. A failure here must not undo the booking: the seat
  // is recorded either way and the problem is written into the sheet.
  var notes = [];
  try {
    var eventId = createInvite_(s.slot, name, email, token);
    sh.getRange(row, COL.eventId + 1).setValue(eventId);
  } catch (err) {
    notes.push('invite failed: ' + err);
  }
  try {
    sendConfirmation_(s.slot, s.seat, name, email, token);
  } catch (err) {
    notes.push('email failed: ' + err);
  }
  if (notes.length) sh.getRange(row, COL.cancelledAt + 1).setValue('NOTE ' + notes.join('; '));
  renderRoster_();
  return { ok: true, bookingId: id, slotLabel: s.slot.label, seat: s.seat, name: name, emailed: notes.length === 0 };
}

// ---------------------------------------------------------------- cancel

function findByToken_(token) {
  token = String(token || '');
  if (!/^[a-f0-9]{32}$/.test(token)) return null;
  var sh = sheet_(SIGNUPS);
  var last = sh.getLastRow();
  if (last < 2) return null;
  var values = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  for (var i = 0; i < values.length; i++) if (values[i][COL.token] === token) return { row: i + 2, v: values[i] };
  return null;
}

function lookup_(token) {
  var r = findByToken_(token);
  if (!r) return { ok: false, error: 'That cancel link is not valid.' };
  var slot = slotById_(r.v[COL.slot]);
  return { ok: true, name: r.v[COL.name], slotLabel: slot ? slot.label : '', status: r.v[COL.status] };
}

function cancel_(body) {
  var result = withLock_(function () {
    var r = findByToken_(body.token);
    if (!r) return { ok: false, error: 'That cancel link is not valid.' };
    if (r.v[COL.status] !== 'Confirmed') return { ok: true, already: true };
    var sh = sheet_(SIGNUPS);
    sh.getRange(r.row, COL.status + 1).setValue('Cancelled');
    sh.getRange(r.row, COL.cancelledAt + 1).setValue(new Date());
    renderRoster_();
    return { ok: true, row: r.row, v: r.v };
  });
  if (!result.ok || result.already) return { ok: result.ok, error: result.error, already: !!result.already };
  var v = result.v;
  var slot = slotById_(v[COL.slot]);
  try { if (v[COL.eventId]) cancelInvite_(String(v[COL.eventId])); } catch (err) { console.warn('cancel invite failed: ' + err); }
  try { sendCancellation_(slot, v[COL.name], v[COL.email]); } catch (err) { console.warn('cancel email failed: ' + err); }
  return { ok: true, slotLabel: slot ? slot.label : '' };
}

// ---------------------------------------------------------------- calendar + email

function calendarId_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('CALENDAR_ID');
  if (id) return id;
  var cal = Calendar.Calendars.insert({ summary: CONFIG.calendarName, timeZone: CONFIG.timeZone });
  props.setProperty('CALENDAR_ID', cal.id);
  return cal.id;
}

function eventDescription_(token) {
  return 'Thank you for helping bring Kovie\'s live roommate stories to life!\n\n' +
    'Short scripts will be emailed before the event. There will also be room for improvisation.\n\n' +
    'Questions or schedule changes? Contact ' + CONFIG.contact + '.\n\n' +
    'Need to cancel? ' + CONFIG.siteUrl + '?cancel=' + token;
}

function createInvite_(slot, name, email, token) {
  var ev = Calendar.Events.insert({
    summary: 'Kovie actor slot (' + slot.label + ')',
    location: CONFIG.location,
    description: eventDescription_(token),
    start: { dateTime: slot.start, timeZone: CONFIG.timeZone },
    end: { dateTime: slot.end, timeZone: CONFIG.timeZone },
    attendees: [{ email: email, displayName: name }],
    guestsCanSeeOtherGuests: false,
    reminders: { useDefault: false, overrides: [{ method: 'popup', minutes: 60 }] }
  }, calendarId_(), { sendUpdates: 'all' });
  return ev.id;
}

function cancelInvite_(eventId) {
  Calendar.Events.remove(calendarId_(), eventId, { sendUpdates: 'all' });
}

function esc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function emailShell_(inner) {
  return '<div style="font-family:Arial,Helvetica,sans-serif;max-width:560px;color:#1f2a30">' +
    '<div style="background:#1b3a4b;color:#fff;padding:18px 22px;border-radius:10px 10px 0 0;font-size:20px;font-weight:bold">' + esc_(CONFIG.eventTitle) + '</div>' +
    '<div style="background:#fdf6ea;padding:22px;border-radius:0 0 10px 10px;line-height:1.55">' + inner + '</div></div>';
}

function sendConfirmation_(slot, seat, name, email, token) {
  var cancelUrl = CONFIG.siteUrl + '?cancel=' + token;
  var first = name.split(' ')[0];
  var html = emailShell_(
    '<p>Hi ' + esc_(first) + ',</p>' +
    '<p>You are signed up. This spot is held for you.</p>' +
    '<p style="background:#fff;border-left:4px solid #1b3a4b;padding:12px 14px">' +
    '<b>' + esc_(slot.label) + '</b> (Actor ' + seat + ')<br>' + esc_(CONFIG.dateLine) + '<br>' + esc_(CONFIG.location) + '</p>' +
    '<p>A calendar invite is on its way in a separate email. Short scripts will be emailed before the event, and there will also be room for improvisation.</p>' +
    '<p>Questions or schedule changes? Contact ' + esc_(CONFIG.contact) + '.</p>' +
    '<p style="font-size:13px;color:#5b6870">Can\'t make it? <a href="' + cancelUrl + '">Cancel your slot</a> so someone else can take it.</p>');
  MailApp.sendEmail({
    to: email,
    subject: 'You\'re signed up: Kovie actor slot, ' + slot.label + ', Dec 1',
    htmlBody: html,
    body: 'You are signed up for ' + slot.label + ' (Actor ' + seat + '), ' + CONFIG.dateLine + ', ' + CONFIG.location +
      '.\nQuestions? Contact ' + CONFIG.contact + '.\nCancel: ' + cancelUrl,
    name: 'Kovie',
    replyTo: CONFIG.replyTo
  });
}

function sendCancellation_(slot, name, email) {
  var html = emailShell_(
    '<p>Hi ' + esc_(String(name).split(' ')[0]) + ',</p>' +
    '<p>Your Kovie actor slot (' + esc_(slot ? slot.label : '') + ', ' + esc_(CONFIG.dateLine) + ') is cancelled and the seat is open again.</p>' +
    '<p>Changed your mind? <a href="' + CONFIG.siteUrl + '">Pick a new slot</a>.</p>');
  MailApp.sendEmail({
    to: String(email),
    subject: 'Cancelled: Kovie actor slot, ' + (slot ? slot.label : '') + ', Dec 1',
    htmlBody: html,
    body: 'Your Kovie actor slot is cancelled. Sign up again: ' + CONFIG.siteUrl,
    name: 'Kovie',
    replyTo: CONFIG.replyTo
  });
}

// ---------------------------------------------------------------- reminder

function reminderHtml_(slot, seat, name, token) {
  var cancelUrl = CONFIG.siteUrl + '?cancel=' + token;
  return emailShell_(
    '<p>Hi ' + esc_(String(name).split(' ')[0]) + ',</p>' +
    '<p>Quick reminder: you are performing with Kovie <b>tomorrow</b>.</p>' +
    '<p style="background:#fff;border-left:4px solid #1b3a4b;padding:12px 14px">' +
    '<b>' + esc_(slot.label) + '</b> (Actor ' + seat + ')<br>' + esc_(CONFIG.dateLine) + '<br>' + esc_(CONFIG.location) + '</p>' +
    '<p>Please arrive a few minutes early. There will be room for improvisation, so have fun with it.</p>' +
    '<p>Questions or running late? Contact ' + esc_(CONFIG.contact) + ', or just reply to this email.</p>' +
    '<p style="font-size:13px;color:#5b6870">Can\'t make it anymore? <a href="' + cancelUrl + '">Cancel your slot</a> so someone else can take it.</p>');
}

function sendReminderEmail_(v) {
  var slot = slotById_(v[COL.slot]);
  MailApp.sendEmail({
    to: String(v[COL.email]),
    subject: 'Reminder: Kovie actor slot tomorrow, ' + slot.label,
    htmlBody: reminderHtml_(slot, v[COL.seat], v[COL.name], v[COL.token]),
    body: 'Reminder: you are performing with Kovie tomorrow, ' + slot.label + ' (Actor ' + v[COL.seat] + '), ' + CONFIG.dateLine + ', ' +
      CONFIG.location + '.\nQuestions? Contact ' + CONFIG.contact + ' or reply to this email.\nCancel: ' + CONFIG.siteUrl + '?cancel=' + v[COL.token],
    name: 'Kovie',
    replyTo: CONFIG.replyTo
  });
}

/** Runs from the time trigger. Sends each confirmed actor one reminder per booking. */
function sendReminders() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = sheet_(SIGNUPS);
    var sent = 0;
    confirmedRows_().forEach(function (r) {
      if (r.v[COL.reminder]) return;
      try {
        sendReminderEmail_(r.v);
        sh.getRange(r.row, COL.reminder + 1).setValue(new Date());
        sent++;
      } catch (err) {
        sh.getRange(r.row, COL.reminder + 1).setValue('FAILED ' + err);
      }
    });
    console.log('Reminders sent: ' + sent);
  } finally {
    lock.releaseLock();
  }
}

/** Sends the reminder to Reese's own confirmed test bookings only, without marking anything. */
function testReminderToReese() {
  var n = 0;
  confirmedRows_().forEach(function (r) {
    if (String(r.v[COL.email]).toLowerCase() === '[removed]') { sendReminderEmail_(r.v); n++; }
  });
  console.log('Test reminders sent: ' + n);
}

/**
 * Run once from the script editor. Adds the "Reminder sent" column and schedules
 * sendReminders() for CONFIG.reminderAt. Safe to re-run: it replaces the old trigger.
 */
function installReminder() {
  var su = sheet_(SIGNUPS);
  su.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#1b3a4b').setFontColor('#ffffff');
  su.getRange(1, COL.reminder + 1, su.getMaxRows(), 1).setNumberFormat('mmm d, h:mm am/pm');
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendReminders').timeBased().at(new Date(CONFIG.reminderAt)).create();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    console.log('Trigger: ' + t.getHandlerFunction() + ' ' + t.getEventType() + ' ' + t.getUniqueId());
  });
}

// ---------------------------------------------------------------- sheet layout

function renderRoster_() {
  var sh = sheet_(ROSTER);
  var bySeat = {};
  confirmedRows_().forEach(function (r) { bySeat[r.v[COL.slot] + ':' + r.v[COL.seat]] = r.v; });
  var rows = CONFIG.slots.map(function (s) {
    var out = [s.label];
    for (var seat = 1; seat <= CONFIG.seatsPerSlot; seat++) {
      var v = bySeat[s.id + ':' + seat];
      out.push(v ? v[COL.name] : '', v ? v[COL.email] : '');
    }
    return out;
  });
  sh.getRange(4, 1, rows.length, rows[0].length).setValues(rows);
  sh.getRange(2, 1).setValue('Last updated ' + Utilities.formatDate(new Date(), CONFIG.timeZone, 'MMM d, h:mm a'));
}
