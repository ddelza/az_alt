const SHEET_ID = '1WItERzz5PtYV-T3BGTE7Dfh2N951qWMbn3iKz65T5-w';

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  if (!action) {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('2026 화학 실험 전람회 동료평가')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var result;
  if (action === 'getClasses') {
    result = getClasses();
  } else if (action === 'getTopics') {
    result = getTopics(e.parameter.classNum);
  } else if (action === 'submit') {
    var payload = JSON.parse(decodeURIComponent(e.parameter.payload));
    result = submitEvaluation(payload);
  } else {
    result = { error: 'unknown action' };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function initDataSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('data');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['제출시간','평가자 이름','평가자 학번','피평가자 이름','피평가자 학번','동료평가 내용']);
    sheet.setFrozenRows(1);
  }
}

function getClasses() {
  try {
    var data = SpreadsheetApp.openById(SHEET_ID)
      .getSheetByName('모둠편성')
      .getDataRange().getValues();
    var classes = [];
    for (var i = 1; i < data.length; i++) {
      var c = data[i][0];
      if (c !== '' && classes.indexOf(c) === -1) classes.push(c);
    }
    classes.sort(function(a, b) { return Number(a) - Number(b); });
    return classes;
  } catch(e) {
    return { error: e.message };
  }
}

function getTopics(classNum) {
  try {
    var data = SpreadsheetApp.openById(SHEET_ID)
      .getSheetByName('모둠편성')
      .getDataRange().getValues();
    var topics = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[0]) !== String(classNum)) continue;
      var topic  = String(row[2] || '').trim();
      var silhum = String(row[3] || '').trim();
      var label  = silhum ? topic + ' / ' + silhum : topic;
      var members = [];
      for (var j = 4; j <= 8; j++) {
        var m = String(row[j] || '').trim();
        if (m !== '') members.push(m);
      }
      topics.push({ modum: row[1], label: label, members: members });
    }
    return topics;
  } catch(e) {
    return { error: e.message };
  }
}

function submitEvaluation(payload) {
  try {
    var ss        = SpreadsheetApp.openById(SHEET_ID);
    var dataSheet = ss.getSheetByName('data');
    var stuData   = ss.getSheetByName('학생학번').getDataRange().getValues().slice(1);
    var timestamp = new Date();

    var evaluatorRow = null;
    for (var i = 0; i < stuData.length; i++) {
      if (String(stuData[i][4]).trim() === payload.evaluatorName &&
          String(stuData[i][1]) === String(payload.evaluatorClass)) {
        evaluatorRow = stuData[i]; break;
      }
    }
    var evaluatorId = evaluatorRow ? evaluatorRow[3] : '미상';

    for (var p = 0; p < payload.peers.length; p++) {
      var peer = payload.peers[p];
      var peerRow = null;
      for (var k = 0; k < stuData.length; k++) {
        if (String(stuData[k][4]).trim() === peer.name &&
            String(stuData[k][1]) === String(payload.evaluatorClass)) {
          peerRow = stuData[k]; break;
        }
      }
      var peerId = peerRow ? peerRow[3] : '미상';
      dataSheet.appendRow([timestamp, payload.evaluatorName, evaluatorId, peer.name, peerId, peer.comment]);
    }
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}
