function onOpen() {
  reputacionalMacro()
}

function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();

  // Quando o número do processo (B6) de PARECER for editado, busca o CNPJ
  if (sheet.getName() === "PARECER" && range.getRow() === 6 && range.getColumn() === 2) {
    var numeroProcesso = range.getValue();
    var SS = SpreadsheetApp.getActiveSpreadsheet();
    var wsCaixa = SS.getSheetByName("CAIXA DE ENTRADA");

    if (numeroProcesso === "" || numeroProcesso === null) {
      sheet.getRange(5, 2).setValue("vazio");
      return;
    }

    var colE = wsCaixa.getRange("E:E").getValues(); // coluna de busca (processo)
    var colD = wsCaixa.getRange("D:D").getValues(); // coluna com CNPJ

    var cnpj = "vazio";
    for (var i = 0; i < colE.length; i++) {
      if (colE[i][0] == numeroProcesso) {
        cnpj = colD[i][0] || "vazio";
        break;
      }
    }

    sheet.getRange(5, 2).setValue(cnpj); // B5 = CNPJ
  }
}