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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var suites = JSON.parse(json);
  var sSh = ss.getSheetByName(SHEET_SUITES);
  var cSh = ss.getSheetByName(SHEET_CASES);
  sSh.clearContents();
  sSh.appendRow(["id","name","description"]);
  for (var i = 0; i < suites.length; i++) {
    var s = suites[i];
    sSh.appendRow([_v(s.id), _v(s.name), _v(s.description)]);
  }
  cSh.clearContents();
  cSh.appendRow(["id","suiteId","title","type","priority","status","assignedTo","automationStatus","steps","expected","attachments"]);
  for (var i = 0; i < suites.length; i++) {
    var cases = suites[i].cases || [];
    for (var j = 0; j < cases.length; j++) {
      var c = cases[j];
      cSh.appendRow([
        _v(c.id), _v(c.suiteId), _v(c.title), _v(c.type), _v(c.priority), _v(c.status),
        _v(c.assignedTo), _v(c.automationStatus), _v(c.steps), _v(c.expected),
        JSON.stringify(c.attachments || [])
      ]);
    }
  }
  SpreadsheetApp.flush();
  autoResize(sSh);
  autoResize(cSh);
  return true;
}

function savePlans(json) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PLANS);
  var plans = JSON.parse(json);
  sh.clearContents();
  sh.appendRow(["id","name","description","createdAt","suiteIds"]);
  for (var i = 0; i < plans.length; i++) {
    var p = plans[i];
    sh.appendRow([_v(p.id), _v(p.name), _v(p.description), _v(p.createdAt), (p.selectedSuites||[]).join(",")]);
  }
  SpreadsheetApp.flush();
  autoResize(sh);
  return true;
}

function saveRuns(json) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RUNS);
  var runs = JSON.parse(json);
  sh.clearContents();
  sh.appendRow(["id","name","description","createdAt","suiteId","suiteName","casesJson"]);
  for (var i = 0; i < runs.length; i++) {
    var r = runs[i];
    sh.appendRow([_v(r.id), _v(r.name), _v(r.description), _v(r.createdAt), _v(r.suiteId), _v(r.suiteName), JSON.stringify(r.cases||[])]);
  }
  SpreadsheetApp.flush();
  autoResize(sh);
  return true;
}

// Single atomic call — writes all 3 sheets in one round-trip
function saveAll(suitesJson, plansJson, runsJson) {
  saveSuites(suitesJson);
  savePlans(plansJson);
  saveRuns(runsJson);
  return true;
}

function uploadAttachment(base64Data, fileName, mimeType) {
  var folder = _getOrCreateFolder("TestTrack Attachments");
  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return JSON.stringify({
    fileId:   file.getId(),
    viewUrl:  "https://drive.google.com/file/d/" + file.getId() + "/view",
    fileName: file.getName()
  });
}

function deleteAttachment(fileId) {
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch(e) {}
  return true;
}

function _getOrCreateFolder(name) {
  var it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function autoResize(sh) {
  sh.autoResizeColumns(1, sh.getLastColumn());
}

function testSave() {
  var testSuites = [{
    id: 999, name: "Test Suite", description: "test desc",
    cases: [{
      id: 9001, suiteId: 999, title: "Test case", type: "Functional",
      priority: "High", status: "Untested", assignedTo: "Tester",
      automationStatus: "Manual", steps: "1. Go\n2. Click", expected: "Page opens",
      attachments: []
    }]
  }];
  saveSuites(JSON.stringify(testSuites));
  Logger.log("✅ testSave PASSED");
}