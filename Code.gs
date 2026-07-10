/**
 * Atomic Soup — H2 2026 Sessions backend
 * Google Apps Script Web App. Stores likes / hosts / contributors / nominations
 * in the bound Google Sheet, and emails Jonathan on new hosts and nominations.
 *
 * Setup: see SETUP.md. In short — create a Google Sheet, open
 * Extensions > Apps Script, paste this in, Deploy > New deployment >
 * Web app > Execute as: Me, Who has access: Anyone. Copy the /exec URL
 * into CONFIG.API in index.html.
 */

var NOTIFY_EMAIL = "jonathan.brown@mccann.com";
var EMAIL_ON = { host: true, nominate: true, contribute: false }; // flip to taste
var SHEET_NAME = "Events";
var HEADERS = ["timestamp","action","sessionId","name","topic","nomId","title","date"];

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(HEADERS); }
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  return sh;
}

function doPost(e) {
  try {
    var data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      sheet_().appendRow([
        new Date(), data.action || "", data.sessionId || "",
        data.name || "", data.topic || "", data.nomId || "",
        data.title || "", data.date || ""
      ]);
    } finally { lock.releaseLock(); }

    if (data.action === "host" && EMAIL_ON.host) {
      email_("Atomic Soup — new host: " + (data.title || data.sessionId),
        (data.name || "Someone") + " has offered to host \"" + (data.title || data.sessionId) +
        "\"" + (data.date ? " (" + data.date + ")" : "") + ".");
    }
    if (data.action === "contribute" && EMAIL_ON.contribute) {
      email_("Atomic Soup — new contributor: " + (data.title || data.sessionId),
        (data.name || "Someone") + " wants to contribute to \"" + (data.title || data.sessionId) + "\".");
    }
    if (data.action === "nominate" && EMAIL_ON.nominate) {
      email_("Atomic Soup — session nominated: " + (data.topic || ""),
        (data.name || "Someone") + " nominated a new session: \"" + (data.topic || "") + "\".");
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  var state = getState_();
  var cb = e && e.parameter && e.parameter.callback;
  var body = JSON.stringify(state);
  if (cb) {
    return ContentService.createTextOutput(cb + "(" + body + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body)
    .setMimeType(ContentService.MimeType.JSON);
}

function getState_() {
  var sh = sheet_();
  var rows = sh.getLastRow() > 1
    ? sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).getValues() : [];
  var likes = {}, hosts = {}, contributors = {}, noms = {}, backs = {};
  rows.forEach(function (r) {
    var action = r[1], sid = r[2], name = r[3], topic = r[4], nomId = r[5];
    if (action === "like")      likes[sid] = (likes[sid] || 0) + 1;
    else if (action === "unlike") likes[sid] = Math.max(0, (likes[sid] || 0) - 1);
    else if (action === "host") hosts[sid] = name;
    else if (action === "contribute") {
      contributors[sid] = contributors[sid] || [];
      if (contributors[sid].indexOf(name) === -1) contributors[sid].push(name);
    }
    else if (action === "nominate") noms[nomId] = { id: nomId, topic: topic, name: name, backs: 0 };
    else if (action === "back")   backs[nomId] = (backs[nomId] || 0) + 1;
    else if (action === "unback") backs[nomId] = Math.max(0, (backs[nomId] || 0) - 1);
  });
  var nominations = Object.keys(noms).map(function (k) {
    noms[k].backs = backs[k] || 0; return noms[k];
  });
  return { likes: likes, hosts: hosts, contributors: contributors, nominations: nominations };
}

function email_(subject, body) {
  try {
    MailApp.sendEmail(NOTIFY_EMAIL, subject,
      body + "\n\n— Atomic Soup sessions page");
  } catch (err) { /* MailApp quota / auth — ignore so the write still succeeds */ }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
