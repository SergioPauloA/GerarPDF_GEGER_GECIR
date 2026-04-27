/** @OnlyCurrentDoc */

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

//global
var SS = SpreadsheetApp.getActiveSpreadsheet();
var ws = SS.getSheetByName("PARECER");
var ww = SS.getSheetByName("CAIXA DE ENTRADA");
var wa = SS.getSheetByName("REPASSES");
var wb = SS.getSheetByName("CONCLUÍDAS");
var wc = SS.getSheetByName("APOIO");
function generatePdf() {
  if (ws.getRange("B5").getValue() != ''){
    var res = Browser.msgBox('ENVIAR PARECER','Deseja prosseguir? ',Browser.Buttons.YES_NO);
    if (res == 'yes'){

      var ssID = SS.getId();
      var shID = ws.getSheetId();

      //additional parameters for exporting the sheet as a pdf

      var url_ext = "/export?exportFormat=pdf&format=pdf"+
        // following parameters are optional...
        + '&size=A4'      // paper size
        + '&portrait=true'    // orientation, false for landscape
        + '&fitw=true'        // fit to width, false for actual size
        + '&sheetnames=false&printtitle=false&pagenumbers=false'  //hide optional headers and footers
        + '&gridlines=false'  // hide gridlines
        + '&fzr=false'       // do not repeat row headers (frozen rows) on each page
        + '&r1=' + 0
        + '&r2=' + 24
        + '&c1=' + 0
        + '&c2=' + 2
        + '&gid='+shID; // NEW CODE

    
      var options = {
      headers: {
        'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken()
        }
      }

      ws.autoResizeRows(1, 25); // ajusta alturas das linhas para evitar espaços em branco no PDF

      var cnpj = ws.getRange(5,2).getValue().toString().replace(".","").replace(".","").replace("-","").replace("/","");
      var cnpjo = ws.getRange(5,2).getValue();
      var pdfName = "RATING DTVM DEB - CNPJ " + cnpj;
      var response = UrlFetchApp.fetch("https://docs.google.com/spreadsheets/d/"+ ssID + url_ext, options); // NEW CODE
      var blob = response.getBlob().setName(pdfName + '.pdf');
      var arquivo = DriveApp.createFile(blob);
      var envio = envioEmail(arquivo,ws,cnpj);
      const pastaRaiz = DriveApp.getFolderById("1DZiI1fRRDJtLuGErdFXMnNcDKVnEdtOb")
      
      //ANO E MES
      var pastas = pastaRaiz.getFoldersByName(new Date().getFullYear());//BUSCA ANO
      if (pastas.hasNext()) {
        pastaEncontrada = pastas.next();
      } else {
        pastaEncontrada = pastaRaiz.createFolder(new Date().getFullYear().toString());
      }
      var pastas2 = pastaEncontrada.getFoldersByName((new Date().getMonth()+1).toString());//BUSCA MES
      if (pastas2.hasNext()) {
        pastaEncontrada2 = pastas2.next();
      } else {
        pastaEncontrada2 = pastaEncontrada.createFolder((new Date().getMonth()+1).toString());
      }
      arquivo.moveTo(pastaEncontrada2);
      if (envio){
        encaminharParaBackup(cnpjo);
      }
    }
  }else{
    Browser.msgBox('ENVIAR PARECER','Empresa não selecionada',Browser.Buttons.OK);
  }
}

function envioEmail(anexo,planilha,cnpj) {
  var destinatario = planilha.getRange("B31").getValue();
  var copiaCarbono = planilha.getRange("B32").getValue();
  const assunto = "RATING DTVM DEB - CNPJ " + cnpj;
  var mensagem = wc.getRange("G2").getValue();
  if (planilha.getRange("B33").getValue().toString() != ''){
    mensagem = mensagem + '<br><br>' + planilha.getRange("B33").getValue().toString()
  }
  const assinatura = Gmail.Users.Settings.SendAs.list('me').sendAs.filter(function(account){if(account.isDefault){return true}})[0].signature;
  const email = GmailApp.createDraft(destinatario, assunto, '', {
    cc: copiaCarbono,
    htmlBody: mensagem + '<br><br>' + assinatura,
    attachments: [DriveApp.getFileById(anexo.getId())]
  });
  
  var res = Browser.msgBox('Enviar Email', 
                            `Para: ` + destinatario + '\\n' + 
                            'CC: '+ copiaCarbono + '\\n' + 
                            'Assunto: ' + assunto + '\\n' + 
                            'Mensagem: ' + mensagem + '\\n' + '\\n' +
                            'Anexos: '+ anexo.getName() + '\\n' + '\\n' + 
                            'Deseja Enviar ?' , Browser.Buttons.YES_NO);
    if (res == 'yes'){
      email.send();
      return true;
    }else{
      return false;
    }
}

function encaminharParaBackup(cnpj){
//var ws = SS.getSheetByName("FORM");
//var ww = SS.getSheetByName("Respostas ao formulário 1");
//var wa = SS.getSheetByName("REPASSE");
//var wb = SS.getSheetByName("CONCLUÍDAS");



// ws -> wa
var vetor = []
vetor.push(ws.getRange("B5").getValue().toString()); //cnpj
vetor.push(ws.getRange("B6").getValue());            //nome
vetor.push(ws.getRange("B7").getValue());            //id

vetor.push(ws.getRange("A10").getValue());            //geger
vetor.push(ws.getRange("B11").getValue());            //recomendacao

vetor.push(ws.getRange("B14").getValue());            //gecir
vetor.push(ws.getRange("B15").getValue());            //rating
vetor.push(ws.getRange("B16").getValue());            //porte
vetor.push(ws.getRange("B17").getValue());            //natureza
vetor.push(ws.getRange("B18").getValue());            //limite exposicao
vetor.push(ws.getRange("B19").getValue());            //prazo
vetor.push(ws.getRange("B20").getValue());            //validade rating

vetor.push(ws.getRange("B22").getValue());            //disclaimer


vetor.push(ws.getRange("B26").getValue());            //pontuacao
vetor.push(ws.getRange("B27").getValue());
vetor.push(ws.getRange("B28").getValue());

vetor.push(Session.getEffectiveUser().getEmail());        //usuario
vetor.push(new Date().toLocaleDateString('pt-br'));       //data horario


wa.appendRow(vetor);


// ww -> wb

var row = ww.getRange("D:D").createTextFinder(cnpj).findNext().getRow();
wb.appendRow(ww.getRange("A" + row + ":EZ" + row).getValues()[0]);
ww.deleteRow(row);
novoForm();

Browser.msgBox("Processo finalizado!")

}
