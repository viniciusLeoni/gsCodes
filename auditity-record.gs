function registroDeAuditoria() {
  //EDITAR
  var planilhaOrigem = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Página1");
  var idDestino = SpreadsheetApp.openById('1fMc-ziiSuargpQfA4wEqrY1cBcFfkQ4Jqn3liVpDgUE');
  //EDITAR
  var planilhaDestino = idDestino.getSheetByName("Sitemap");
  if (planilhaDestino) {
    if(planilhaOrigem){
      var segundaLinhhaA = planilhaDestino.getRange("A2").getValue();
      // Guarda os valores atuais da coluna A
      if (segundaLinhhaA) {
        var colunaADestino = planilhaDestino.getRange("A2:A" + (planilhaDestino.getLastRow())).getValues();
        var colunaADestinoSize = colunaADestino.length;
      }else{
        var colunaADestinoSize = 0;
      }
      var colunaAOrigem = planilhaOrigem.getRange("A2:A" + (planilhaOrigem.getLastRow())).getValues();
      var colunaAOrigemSize = colunaAOrigem.length;

      if(colunaAOrigemSize === colunaADestinoSize){
        Logger.log("Registro já foi executado. As planilhas tem a mesmoa quantidade de linhas.");
        return;
      }

      //Limpa todos os dados
      planilhaDestino.clear();
      var valores = planilhaOrigem.getDataRange().getValues();
      planilhaDestino.getRange(1, 1, valores.length, valores[0].length).setValues(valores);
      planilhaDestino.insertColumnBefore(1);

      var ontem = new Date();
      ontem.setDate(ontem.getDate() - 1);
      var dataAtual = Utilities.formatDate(ontem, Session.getScriptTimeZone(), "dd-MM-yyyy");
      var ultimaLinha = planilhaDestino.getLastRow();
      planilhaDestino.getRange("A1").setValue("Data");
      if(colunaADestino){
        planilhaDestino.getRange("A2:A" + (ultimaLinha - 1)).setValues(colunaADestino);
      }
      planilhaDestino.getRange("A"+ultimaLinha).setValue(dataAtual);
      planilhaDestino.getRange("B2:E"+ultimaLinha).setNumberFormat("#,##0");
    }
  }
}
