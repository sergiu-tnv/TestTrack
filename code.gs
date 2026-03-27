const SHEET_SUITES = "Suites";
const SHEET_CASES  = "Cases";
const SHEET_PLANS  = "Plans";
const SHEET_RUNS   = "TestRuns";

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle("TestTrack")
    .addMetaTag("viewport", "width=device-width,initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function _v(x) {
  if (x === null || x === undefined) return "";
  if (typeof x === "number" || typeof x === "boolean") return x;
  return String(x);
}

function initSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  function ensureSheet(name, headers) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.clearContents();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    autoResize(sh);
    return sh;
  }
  ensureSheet(SHEET_SUITES, ["id","name","description"]);
  ensureSheet(SHEET_CASES,  ["id","suiteId","title","type","priority","status","assignedTo","automationStatus","steps","expected","attachments"]);
  ensureSheet(SHEET_PLANS,  ["id","name","description","createdAt","suiteIds"]);
  ensureSheet(SHEET_RUNS,   ["id","name","description","createdAt","suiteId","suiteName","casesJson"]);
  Logger.log("Sheets initialised.");
}

function getSuites() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sv = ss.getSheetByName(SHEET_SUITES).getDataRange().getValues();
  var cv = ss.getSheetByName(SHEET_CASES).getDataRange().getValues();
  var sh = sv[0], ch = cv[0];
  var suites = sv.slice(1).filter(function(r){ return r[sh.indexOf("id")]; }).map(function(r){
    var obj = {};
    sh.forEach(function(h,i){ obj[h] = r[i]; });
    obj.cases = [];
    return obj;
  });
  cv.slice(1).filter(function(r){ return r[ch.indexOf("id")]; }).forEach(function(r){
    var obj = {};
    ch.forEach(function(h,i){ obj[h] = r[i]; });
    try { obj.attachments = JSON.parse(obj.attachments || "[]"); } catch(e) { obj.attachments = []; }
    var suite = suites.filter(function(s){ return s.id == obj.suiteId; })[0];
    if (suite) suite.cases.push(obj);
  });
  return JSON.stringify(suites);
}

function getPlans() {
  var vals = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PLANS).getDataRange().getValues();
  var headers = vals[0];
  return JSON.stringify(vals.slice(1).filter(function(r){ return r[0]; }).map(function(r){
    var o = {};
    headers.forEach(function(h,i){ o[h] = r[i]; });
    o.selectedSuites = String(o.suiteIds || "").split(",").map(Number).filter(Boolean);
    return o;
  }));
}

function getRuns() {
  var vals = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RUNS).getDataRange().getValues();
  var headers = vals[0];
  return JSON.stringify(vals.slice(1).filter(function(r){ return r[0]; }).map(function(r){
    var o = {};
    headers.forEach(function(h,i){ o[h] = r[i]; });
    try { o.cases = JSON.parse(o.casesJson || "[]"); } catch(e) { o.cases = []; }
    return o;
  }));
}

function saveSuites(json) {
  var suites;
  try { suites = JSON.parse(json); } catch(e) { throw new Error("Invalid suites JSON: " + e.message); }
  if (!Array.isArray(suites)) throw new Error("Suites data must be an array");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sSh = ss.getSheetByName(SHEET_SUITES);
  var cSh = ss.getSheetByName(SHEET_CASES);
  sSh.clearContents();
  sSh.appendRow(["id","name","description"]);
  for (var i = 0; i < suites.length; i++) {
    var s = suites[i];
    if (!s.id) continue;
    sSh.appendRow([_v(s.id), _v(s.name), _v(s.description)]);
  }
  cSh.clearContents();
  cSh.appendRow(["id","suiteId","title","type","priority","status","assignedTo","automationStatus","steps","expected","attachments"]);
  for (var i = 0; i < suites.length; i++) {
    var cases = suites[i].cases || [];
    for (var j = 0; j < cases.length; j++) {
      var c = cases[j];
      if (!c.id) continue;
      cSh.appendRow([
        _v(c.id), _v(c.suiteId), _v(c.title), _v(c.type), _v(c.priority), _v(c.status),
        _v(c.assignedTo), _v(c.automationStatus), _v(c.steps), _v(c.expected),
        JSON.stringify(c.attachments || [])
      ]);
    }
  }
  SpreadsheetApp.flush();
  return true;
}

function savePlans(json) {
  var plans;
  try { plans = JSON.parse(json); } catch(e) { throw new Error("Invalid plans JSON: " + e.message); }
  if (!Array.isArray(plans)) throw new Error("Plans data must be an array");
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PLANS);
  sh.clearContents();
  sh.appendRow(["id","name","description","createdAt","suiteIds"]);
  for (var i = 0; i < plans.length; i++) {
    var p = plans[i];
    if (!p.id) continue;
    sh.appendRow([_v(p.id), _v(p.name), _v(p.description), _v(p.createdAt), (p.selectedSuites||[]).join(",")]);
  }
  SpreadsheetApp.flush();
  return true;
}

function saveRuns(json) {
  var runs;
  try { runs = JSON.parse(json); } catch(e) { throw new Error("Invalid runs JSON: " + e.message); }
  if (!Array.isArray(runs)) throw new Error("Runs data must be an array");
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RUNS);
  sh.clearContents();
  sh.appendRow(["id","name","description","createdAt","suiteId","suiteName","casesJson"]);
  for (var i = 0; i < runs.length; i++) {
    var r = runs[i];
    if (!r.id) continue;
    sh.appendRow([_v(r.id), _v(r.name), _v(r.description), _v(r.createdAt), _v(r.suiteId), _v(r.suiteName), JSON.stringify(r.cases||[])]);
  }
  SpreadsheetApp.flush();
  return true;
}

// Single atomic call — writes all 3 sheets in one round-trip
function saveAll(suitesJson, plansJson, runsJson) {
  try { saveSuites(suitesJson); } catch(e) { throw new Error("Failed saving suites: " + e.message); }
  try { savePlans(plansJson); } catch(e) { throw new Error("Failed saving plans: " + e.message); }
  try { saveRuns(runsJson); } catch(e) { throw new Error("Failed saving runs: " + e.message); }
  return true;
}

function uploadAttachment(base64Data, fileName, mimeType) {
  if (!base64Data || !fileName) throw new Error("Missing attachment data");
  var folder = _getOrCreateFolder("TestTrack Attachments");
  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType || "application/octet-stream", fileName);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return JSON.stringify({
    fileId:   file.getId(),
    viewUrl:  "https://drive.google.com/file/d/" + file.getId() + "/view",
    fileName: file.getName()
  });
}

function deleteAttachment(fileId) {
  if (!fileId) return true;
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch(e) {}
  return true;
}

// Scoped to the same folder as the spreadsheet, not searching all of Drive
function _getOrCreateFolder(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var parents = DriveApp.getFileById(ss.getId()).getParents();
  var parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function autoResize(sh) {
  sh.autoResizeColumns(1, sh.getLastColumn());
}
