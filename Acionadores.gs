function onOpen() {
  reputacionalMacro()
}

function onEdit(e) {
  var sheet = e.range.getSheet();
  var row = e.range.getRow();
  var col = e.range.getColumn();

  if (sheet.getName() === "PARECER" && row === 6 && col === 2) {
    var numProcesso = sheet.getRange(6, 2).getValue();
    var ww = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CAIXA DE ENTRADA");
    var dataE = ww.getRange("E:E").getValues();
    var dataD = ww.getRange("D:D").getValues();
    var cnpj = "vazio";

    for (var i = 0; i < dataE.length; i++) {
      if (dataE[i][0] == numProcesso) {
        cnpj = dataD[i][0];
        break;
      }
    }

    sheet.getRange(5, 2).setValue(cnpj);
  }
}