function registroDeBacklog() {
  //EDITAR
  var planilhaOrigem = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Problemas");
  var planilhaDestino = SpreadsheetApp.openById('1saG_m3nqTvsTRT6yG46oYBVi8A5kkzdXt5baevbupqA');
  //EDITAR
  var hoje = new Date();
  var nomeAba = Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd-MM-yyyy");
  var aba = planilhaDestino.getSheetByName(nomeAba);
  if (!aba) {
    aba = planilhaDestino.insertSheet(nomeAba);  
    if(planilhaOrigem){
      var valores = planilhaOrigem.getDataRange().getValues();
      aba.getRange(1, 1, valores.length, valores[0].length).setValues(valores);
    }
  }
}
