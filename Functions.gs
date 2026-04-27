/*
******************************************
*        SISTEMA FINANCEIRO BANESTES     *
*              BANESTES DTVM             *
*              NUADM E GECIR             *
*                  2025                  *
*                                        *
*      FORMULÁRIO DE RISCO DE CRÉDITO    *
*           DESENVOLVIDO POR             *
*           IODUTRA - 030095285          *
*                                        *
******************************************
*/


function novoForm() {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var ws = SS.getSheetByName("PARECER");
    
  ws.getRange(6,2).setValue(""); //CNPJ
  ws.getRange(14,1).setValue(""); //Parecer GECIR
  ws.getRange(27,2).setValue(0); //Ajuste de pontuação
  ws.getRange(33,2).setValue(""); //Corpo do email(+)
}



/**
 * Retorna o valor de Rating da DTVM.
 * @param {number} valor - Informe a pontuação a ser analisada
 * @return Retorna a classificação de Rating DTVM
 * @customfunction
*/
function calculaRating(valor) {
  switch(true){
    case valor > 299: // MAX
      return "A1";
    case valor <= 113: // MIN
      return "SEM RATING";

    // CASE B //////////////////////
    case valor > 113 && valor <= 126:
      return "B3";
    case valor > 126 && valor <= 139:
      return "B2";
    case valor > 139 && valor <= 152:
      return "B1";

    // CASE A //////////////////////
    case valor > 152 && valor <= 177:
      return "A7";
    case valor > 177 && valor <= 201:
      return "A6";
    case valor > 201 && valor <= 226:
      return "A5";
    case valor > 226 && valor <= 250:
      return "A4";
    case valor > 250 && valor <= 275:
      return "A3";
    case valor > 275 && valor <= 299:
      return "A2";
    
    // CASE DEFAULT /////////////////
    default:
      return "vazio";
}

}

/**
 * Retorna o valor do próximo ID.
 * @param {text} cnpj - Informe o CNPJ
 * @return Retorna o valor do próximo ID
 * @customfunction
*/
function proxID(cnpj) {
  if (cnpj != ''){
    var SS = SpreadsheetApp.getActiveSpreadsheet();
    var wa = SS.getSheetByName("REPASSES");

    ultlinha = wa.getLastRow()

    if (ultlinha == 1){
      return "0001";
    }else{
      return (wa.getRange("C"+wa.getLastRow()).getValue()+1);
    }

  }else{
    return "0000";
  }


}
