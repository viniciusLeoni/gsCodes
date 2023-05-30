var planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dados GSC");

function getDataFromGSC() {
  var propertyId = "sc-domain:oiplace.com.br";
  var lastRow = planilha.getLastRow();
  var data = [];
  if (lastRow > 0) {
    var lastRowValue = planilha.getRange(lastRow, 1).getValue();
    var umDiaEmMilissegundos = 86400000; // 1 dia em milissegundos
    var novaData = new Date(lastRowValue.getTime() + umDiaEmMilissegundos);
    var startDate = novaData.toISOString().slice(0, 10);
  }else{
    var lastRow = 1;
    var startDate = "2020-01-01";
    var lineOne = ["Date", "Clicks", "Impressions", "CTR", "Position"];
    data.push(lineOne);
  }
  var today = new Date();
  endDate = today.toISOString().slice(0, 10);
  //Logger.log(startDate);
  //Logger.log(endDate);
  var searchConsoleAPIEndpoint = "https://www.googleapis.com/webmasters/v3/sites/" + propertyId + "/searchAnalytics/query";
  var headers = {
    "Authorization": "Bearer " + ScriptApp.getOAuthToken()
  };
  var payload = {
    "startDate": startDate,
    "endDate": endDate,
    "dimensions": ["date"],
    "searchType": "web",
    "rowLimit": 25000
  };
  var options = {
    "method": "post",
    "contentType": "application/json",
    "headers": headers,
    "payload": JSON.stringify(payload)
  };
  var response = UrlFetchApp.fetch(searchConsoleAPIEndpoint, options);
  var responseData = JSON.parse(response.getContentText());
  var rows = responseData.rows;
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var date = row.keys[0];
    var date = Utilities.formatDate(new Date(date), "GMT-3", "dd/MM/yyyy");
    var clicks = row.clicks.toLocaleString('pt-BR');
    var impressions = row.impressions.toLocaleString('pt-BR');
    var ctr = row.ctr * 100;
    var ctr = ctr.toLocaleString('pt-BR');
    var position = row.position.toLocaleString('pt-BR');
    data.push([date, clicks, impressions, ctr, position]);
  }
  // Limpa os dados existentes na planilha e salva os novos dados
  //planilha.clearContents();
  planilha.getRange(lastRow, 1, data.length, data[0].length).setValues(data);
  //Cria a mensagem para ser enviada pelo Telegram
  sumarioGSC(30);

}

function sumarioGSC(alertNumber) {
  // Nome da planilha e colunas a serem analisadas
  var nomePlanilha = "Dados GSC";
  
  // Abre a planilha pelo nome
  var planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomePlanilha);
  
  // Obtém o valor da última linha na coluna A
  var lastDay = planilha.getRange("A" + planilha.getLastRow()).getValue();
  var formattedDate = Utilities.formatDate(lastDay, 'GMT-03:00', 'dd/MM/yyyy');

  var coluna = planilha.getRange("A:A").getValues();
  var d = coluna.filter(String).length;
  var dmenos1 = coluna.filter(String).length-1;
  var dmenos7 = coluna.filter(String).length-7;

  // Obtém o valor da última linha na coluna B
  var cliques = planilha.getRange("B" + d).getValue();
  var cliquesD1 = planilha.getRange("B" + dmenos1).getValue();
  var cliquesD7 = planilha.getRange("B" + dmenos7).getValue();
  var delta_cliquesD1 = (cliques/cliquesD1-1)*100;
  var delta_cliquesD7 = (cliques/cliquesD7-1)*100;
  
  // Obtém o valor da última linha na coluna C
  var impressoes = planilha.getRange("C" + d).getValue();
  var impressoesD1 = planilha.getRange("C" + dmenos1).getValue();
  var impressoesD7 = planilha.getRange("C" + dmenos7).getValue();
  var delta_impressoesD1 = (impressoes/impressoesD1-1)*100;
  var delta_impressoesD7 = (impressoes/impressoesD7-1)*100;

  // Obtém o valor da última linha na coluna D
  var ctr = (planilha.getRange("D" + d).getValue()*100);
  var ctrD1 = (planilha.getRange("D" + dmenos1).getValue()*100);
  var ctrD7 = (planilha.getRange("D" + dmenos7).getValue()*100);
  var delta_ctrD1 = ctr-ctrD1;
  var delta_ctrD7 = ctr-ctrD7;

  // Obtém o valor da última linha na coluna E
  var posicao = planilha.getRange("E" + d).getValue();
  var posicaoD1 = planilha.getRange("E" + dmenos1).getValue();
  var posicaoD7 = planilha.getRange("E" + dmenos7).getValue();
  var delta_posicaoD1 = ((posicao/posicaoD1-1)*-1)*100;
  var delta_posicaoD7 = ((posicao/posicaoD7-1)*-1)*100;
  
  // Obtém o valor da última linha na coluna D
  var ultimaLinhaD = planilha.getRange("D" + planilha.getLastRow()).getValue();
  
  // Obtém o valor da última linha na coluna E
  var ultimaLinhaE = planilha.getRange("E" + planilha.getLastRow()).getValue();
  var mensagem = "";
  //Alertas máximos
  if( (delta_cliquesD1)*-1 >= alertNumber  ){
    mensagem += "\n<b>ATENÇÃO!!!</b>"
    mensagem += "\nQueda de <b>cliques</b> maior que <b>"+alertNumber+"%</b> em relação ao dia anterior\n\n"
  }else  if( (delta_cliquesD7)*-1 >= alertNumber  ){
    mensagem += "\n<b>ATENÇÃO!!!</b>"
    mensagem += "\nQueda em <b>cliques</b> maior que <b>"+alertNumber+"%</b> em relação à semana anterior\n\n"
  }else  if( (delta_impressoesD1)*-1 >= alertNumber  ){
    mensagem += "\n<b>ATENÇÃO!!!</b>"
    mensagem += "\nQueda em <b>impressões</b> maior que <b>"+alertNumber+"%</b> em relação ao dia anterior\n\n"
  }else  if( (delta_impressoesD7)*-1 >= alertNumber  ){
    mensagem += "\n<b>ATENÇÃO!!!</b>"
    mensagem += "\nQueda em <b>impressões</b> maior que <b>"+alertNumber+"%</b> em relação à semana anterior\n\n"
  }
  mensagem += "<b>Informações referentes ao dia "+formattedDate+"</b>\n";
  // Cliques
  mensagem += "\n<b>Cliques:</b> " + cliques.toLocaleString('pt-BR');
  mensagem += "\nDia anterior: "+cliquesD1.toLocaleString('pt-BR')+" ("+delta_cliquesD1.toFixed(2).replace(/\./g, ',')+"%)";
  mensagem += "\nSemana anterior: "+cliquesD7.toLocaleString('pt-BR')+" ("+delta_cliquesD7.toFixed(2).replace(/\./g, ',')+"%)\n";
  // Impressões
  mensagem += "\n<b>Impressões:</b> " + impressoes.toLocaleString('pt-BR');
  mensagem += "\nDia anterior: "+impressoesD1.toLocaleString('pt-BR')+" ("+delta_impressoesD1.toFixed(2).replace(/\./g, ',')+"%)";
  mensagem += "\nSemana anterior: "+impressoesD7.toLocaleString('pt-BR')+" ("+delta_impressoesD7.toFixed(2).replace(/\./g, ',')+"%)\n";
  // CTR
  mensagem += "\n<b>CTR:</b> " + ctr.toFixed(2).replace(/\./g, ',');
  mensagem += "\nDia anterior: "+ctrD1.toFixed(2).replace(/\./g, ',')+" ("+delta_ctrD1.toFixed(2).replace(/\./g, ',')+"pp)";
  mensagem += "\nSemana anterior: "+ctrD7.toFixed(2).replace(/\./g, ',')+" ("+delta_ctrD7.toFixed(2).replace(/\./g, ',')+"pp)\n";
  // Posição
  mensagem += "\n<b>Posição Média:</b> " + posicao.toFixed(2).replace(/\./g, ',');
  mensagem += "\nDia anterior: "+posicaoD1.toFixed(2).replace(/\./g, ',')+" ("+delta_posicaoD1.toFixed(2).replace(/\./g, ',')+"%)";
  mensagem += "\nSemana anterior: "+posicaoD7.toFixed(2).replace(/\./g, ',')+" ("+delta_posicaoD7.toFixed(2).replace(/\./g, ',')+"%)\n";

  // Obter a data e hora atual
  var dataHoraAtual = new Date();
  var dataChecker = Utilities.formatDate(dataHoraAtual, 'GMT-03:00', 'dd/MM/yyyy');
  var horaChecker = Utilities.formatDate(dataHoraAtual, 'GMT-03:00', 'HH:mm');
  mensagem += '\n<b>Checagem realizada em ' + dataChecker+' às '+horaChecker+'</b>';

  //Enviar dados pelo Telegram
  enviarGSCTelegram(mensagem);
  
}

function enviarGSCTelegram(mensagem) {
  //EDITAR
  var botToken = "5878777413:AAHTOabO4RlSJimIbp0tdFhOWaKIaCCUnfM"; // Insira o token do seu bot do Telegram
  var chatId = "-932152817"; // Insira o ID do chat do Telegram para onde a mensagem será enviada
  //EDITAR
  var url = "https://api.telegram.org/bot" + botToken + "/sendMessage";
  var payload = {
    "chat_id": chatId,
    "text": mensagem,
    "parse_mode": "HTML"
  };
  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };
  UrlFetchApp.fetch(url, options);
  //Logger.log("Dados de checagem enviados no Telegram.");
}
