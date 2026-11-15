function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('results') || SpreadsheetApp.getActiveSpreadsheet().insertSheet('results');
    if (sheet.getLastRow() === 0) sheet.appendRow(['서버 기록 시각','응시자 이름','문제셋','문항 수','정답 수','점수','오답 번호','응시 시각']);
    function safe(value) {
      var s = String(value == null ? '' : value);
      return /^[=+@\-]/.test(s) ? "'" + s : s;
    }
    sheet.appendRow([new Date(),safe(data.name).slice(0,30),safe(data.set),Number(data.total)||0,Number(data.correct)||0,Number(data.score)||0,Array.isArray(data.wrongIds)?data.wrongIds.join(','):'',safe(data.date)]);
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(error)})).setMimeType(ContentService.MimeType.JSON);
  }
}