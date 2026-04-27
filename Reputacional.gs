function reputacionalMacro() {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var ww = SS.getSheetByName("CAIXA DE ENTRADA");

  lr = ww.getLastRow();

  ajuste = false;

  if (lr!=1) {
    for (i=2;i<=lr;i++){
      if(ww.getRange("D"+i).getDisplayValue()!=""){
        cel = ww.getRange("F"+i).getFormula()
        if(cel == ""){
          var searchCell = "D" + i;
    
          var formula = "=VLOOKUP(QUERY(REPUTACIONAL!A2:DG;\"Select Col24 where Col4='\"&" + searchCell + "&\"' ORDER BY Col1 DESC LIMIT 1\";-1);APOIO!A:B;2;0)";
          ww.getRange("F"+i).setFormula(formula);

          var formula = "=IF(ISNUMBER(SEARCH(\"201-1\";QUERY(REPUTACIONAL!A2:DG;\"Select Col52 where Col4='\"&" + searchCell + "&\"' ORDER BY Col1 DESC LIMIT 1\";-1)));QUERY(APOIO!K1:M102;\"Select Col3 where Col1='\"&QUERY(REPUTACIONAL!A2:DG;\"Select Col52 where Col4='\"&" + searchCell + "&\"' ORDER BY Col1 DESC LIMIT 1\";-1)&\"' and Col2='\"&QUERY(REPUTACIONAL!A2:DG;\"Select Col53 where Col4='\"&" + searchCell + "&\"' ORDER BY Col1 DESC LIMIT 1\";-1)&\"'\";-1);IF(ISNUMBER(SEARCH(\"203-8\";QUERY(REPUTACIONAL!A2:DG;\"Select Col52 where Col4='\"&" + searchCell + "&\"' ORDER BY Col1 DESC LIMIT 1\";-1)));QUERY(APOIO!K1:M102;\"Select Col3 where Col1='\"&QUERY(REPUTACIONAL!A2:DG;\"Select Col52 where Col4='\"&" + searchCell + "&\"' ORDER BY Col1 DESC LIMIT 1\";-1)&\"' and Col2='\"&QUERY(REPUTACIONAL!A2:DG;\"Select Col53 where Col4='\"&" + searchCell + "&\"' ORDER BY Col1 DESC LIMIT 1\";-1)&\"'\";-1);VLOOKUP(QUERY(REPUTACIONAL!A2:DG;\"Select Col52 where Col4='\"&" + searchCell + "&\"' ORDER BY Col1 DESC LIMIT 1\";-1);APOIO!K:M;3;FALSE)))";
          ww.getRange("G"+i).setFormula(formula);

          var formula = "=IF(QUERY(REPUTACIONAL!A2:DG;\"Select Col48 where Col4='\"&" + searchCell + "&\"' ORDER BY Col1 DESC LIMIT 1\";-1)=\"NÃO\";\"NÃO\";\"\")";
          ww.getRange("Q"+i).setFormula(formula);

          var formula = "=IF(QUERY(REPUTACIONAL!A2:DG;\"Select Col48 where Col4='\"&" + searchCell + "&\"' ORDER BY Col1 DESC LIMIT 1\";-1)=\"SIM\";\"SIM\";\"\")";
          ww.getRange("R"+i).setFormula(formula);

          var formula = "=IF(QUERY(REPUTACIONAL!A2:DG;\"Select Col48 where Col4='\"&" + searchCell + "&\"' ORDER BY Col1 DESC LIMIT 1\";-1)=\"Não foi possível consultar.\";\"Não foi possível consultar.\";\"\")";
          ww.getRange("S"+i).setFormula(formula);

          ww.getRange("EX"+i).setFormula("=SUM(C"+i+";IF(Q"+i+"="+String.fromCharCode(34)+"NÃO"+String.fromCharCode(34)+";10;IF(R"+i+"="+String.fromCharCode(34)+"SIM"+String.fromCharCode(34)+";"+"2;0)))");
          ajuste = true
        }
      }
    }
    if (ajuste){
      Browser.msgBox("Ajustes feitos!")
    }else{
      Browser.msgBox("Não há novos ajustes!")
    }
  


  }else{
    Browser.msgBox("Não há novos ajustes!")
  }



}
