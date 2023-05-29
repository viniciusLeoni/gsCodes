// ID do arquivo do Google Sheets de origem
var idPlanilhaOrigem = "10wAX0xkGS_IU8AXpNSIbTRMLUqq25zLSNFAFNjOs2Vg";
// Abre a planilha de origem
var planilhaOrigem = SpreadsheetApp.openById(idPlanilhaOrigem);
// Obtém a planilha atual
var planilhaDestino = SpreadsheetApp.getActiveSpreadsheet();
var validador = "www.oiplace";
var datasExcluidas = [];
var fakeHoje = "";
// Nome da planilha com a data de hoje na planilha de origem
if(fakeHoje !== ""){
  var nomePlanilhaHoje = fakeHoje;
}else{
  var fusoHorario = Session.getScriptTimeZone(); // Obtém o fuso horário do script
  var nomePlanilhaHoje = Utilities.formatDate(new Date(), fusoHorario, "dd-MM-yyyy");
}

function sitemapChangesChecker() {
  
  // Nome da planilha atual onde os dados serão escritos
  var nomeAbaSitemapHoje = "Sitemap Hoje";
  
  // Obtém a planilha com a data de hoje na planilha de origem
  var abaOrigem = planilhaOrigem.getSheetByName(nomePlanilhaHoje);
  
  // Verifica se a planilha de origem existe
  if (abaOrigem) {
    // Obtém os dados da coluna A da planilha de origem
    var dadosOrigem = abaOrigem.getRange("A:A").getValues().flat();
    
    // Obtém a planilha de destino
    var abaDestino = planilhaDestino.getSheetByName(nomeAbaSitemapHoje);
    
    // Limpa os dados existentes na planilha de destino
    if (abaDestino) {
      abaDestino.clear();
    } else {
      // Se a planilha de destino não existir, cria uma nova planilha com o nome especificado
      abaDestino = planilhaDestino.insertSheet(nomeAbaSitemapHoje);
    }
    
    // Escreve os dados na planilha de destino
    abaDestino.getRange(1, 1, dadosOrigem.length, 1).setValues(dadosOrigem.map(function(valor) {
      return [valor];
    }));

    
    // Chama função para escrever os dados dos sitemaps antigos
    escreverSitemapsAntigos();
    // Chama função para comparar os sitemaps
    compararSitemaps();
    //Chama a funlção que monta a mensagem para o Telegram
    sumarioComparador("URLs Removidas","Novas URLs","A",nomePlanilhaHoje,);
    
  } else {
    Logger.log("A planilha de origem com a data de hoje não foi encontrada.");
  }
}

function escreverSitemapsAntigos() {
  
  // Nome da aba na planilha atual
  var nomeAbaSitemapsAntigos = "Sitemaps Antigos";

  // Abre a planilha de origem

  // Obtém a aba "Sitemaps Antigos" na planilha atual
  var abaSitemapsAntigos = planilhaDestino.getSheetByName(nomeAbaSitemapsAntigos);

  // Limpa os dados existentes na aba "Sitemaps Antigos" na planilha atual
  abaSitemapsAntigos.clearContents();

  // Obtém todas as abas da planilha de origem
  var abasOrigem = planilhaOrigem.getSheets();

  //Escreve o cabeçalho em Sitemaps Antigos
  abaSitemapsAntigos.getRange(1,1).setValue("URLs");

  //Salva o nome de todas as abas pra pegar a mais recente
  var nomesAbasOrigem = [];

  // Loop pelas abas de origem
  for (var i = 0; i < abasOrigem.length; i++) {
    var abaOrigem = abasOrigem[i];
    var nomeAbaOrigem = abaOrigem.getName();
    // Verifica se o nome da aba não é "Sumário" nem a data de hoje
    if (!datasExcluidas.includes(nomeAbaOrigem) && nomeAbaOrigem !== "Sumário" && nomeAbaOrigem !== nomePlanilhaHoje) {
      nomesAbasOrigem.push(nomeAbaOrigem);
    }
  }
  
  var nomePlanihaAnterior = nomesAbasOrigem[0];
  var planilhaAnterior = planilhaOrigem.getSheetByName(nomePlanihaAnterior);

  var coluna = planilhaAnterior.getRange("A:A").getValues();
  var ultimaLinhaColuna = coluna.filter(String).length-1;
  var valoresOrigem = planilhaAnterior.getRange("A2:A"+ultimaLinhaColuna).getValues();

  // Escreve os valores únicos na planilha atual, na aba "Sitemaps Antigos"
  var ultimaLinha = abaSitemapsAntigos.getLastRow();
  abaSitemapsAntigos.getRange(ultimaLinha + 1, 1, valoresOrigem.length, 1).setValues(valoresOrigem);

  // Escreve a data do sitemap da data mais recente, que não hoje, na última linha de Sitemaps Antigos
  abaSitemapsAntigos.getRange(valoresOrigem.length+2, 1).setValue(nomePlanihaAnterior);

}

function compararSitemaps() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var planilhaAtual = spreadsheet.getSheetByName("Sitemap Hoje");
  var planilhaAntigo = spreadsheet.getSheetByName("Sitemaps Antigos");
  var planilhaRemoviveis = spreadsheet.getSheetByName("URLs Removidas");
  var planilhaNovas = spreadsheet.getSheetByName("Novas URLs");

  var colunaAntiga = planilhaAntigo.getRange("A:A").getValues();
  var ultimaLinhaAntiga = (colunaAntiga.filter(String).length)-1;

  var colunaAtual = planilhaAtual.getRange("A:A").getValues();
  var ultimaLinhaAtual = (colunaAtual.filter(String).length)-1;

  var urlsAntigas = planilhaAntigo.getRange("A2:A"+ultimaLinhaAntiga).getValues();
  var urlsAtuais = planilhaAtual.getRange("A2:A"+ultimaLinhaAtual).getValues();

  var urlsRemoviveis = [];
  var urlsNovas = [];

  for (var i = 0; i < urlsAntigas.length; i++) {
    var url = urlsAntigas[i][0];
    if (url.endsWith('/') && url !== "" && !(url instanceof Date)) {
      var normalizedUrl = removerBarraFinal(url);
    }
    if (!urlsAtuais.some(function(item) {
      return item[0] === url || removerBarraFinal(item[0]) === normalizedUrl;
    })) {
      urlsRemoviveis.push([url]);
    }
  }

  for (var j = 0; j < urlsAtuais.length; j++) {
    var url = urlsAtuais[j][0];
      var normalizedUrl = removerBarraFinal(url);
    if (!urlsAntigas.some(function(item) {
      return item[0] === url || removerBarraFinal(item[0]) === normalizedUrl;
    })) {
      urlsNovas.push([url]);
    }
  }
  

  planilhaRemoviveis.clearContents();
  planilhaRemoviveis.getRange(1, 1).setValue("URLs");
  if(urlsRemoviveis.length > 0){
    planilhaRemoviveis.getRange(2, 1, urlsRemoviveis.length, 1).setValues(urlsRemoviveis);
  }

  planilhaNovas.clearContents();
  planilhaNovas.getRange(1, 1).setValue("URLs");
  if(urlsNovas.length > 0){
    planilhaNovas.getRange(2, 1, urlsNovas.length, 1).setValues(urlsNovas);
  }
  
  registroDeComparador("URLs Removidas","1ED_nwKbV2GCUT4kgv8l6CvW4GAQc_TzibVZsF8U2em4");
  registroDeComparador("Novas URLs","13_YreCERsHusNXuXx27wnLQ35GiLvwZ4Tc6V3StcSx4");

}

function removerBarraFinal(url) {
  return url.replace(/\/$/, '');
}

function sumarioComparador(planilhaAlvo1,planilhaAlvo2,letraAlvo) {
  
  //Planilha URLs Removidas
  var planilha1 = planilhaDestino.getSheetByName(planilhaAlvo1);
  // Definir a planilha e a coluna a serem verificadas
  var colunaAlvo1 = planilha1.getRange(letraAlvo+":"+letraAlvo).getValues();
  var ultimaLinha1 = colunaAlvo1.filter(String).length;
  if(ultimaLinha1 === 1){
    ultimaLinha1 = 2;
  }
  var dados1 = planilha1.getRange(letraAlvo+"2:"+letraAlvo+ultimaLinha1).getValues();
  // Obter o total de linhas preenchidas
  var totalLinhasPreenchidas1 = dados1.filter(function (valor) {
    return valor[0] !== '';
  }).length;
  var totalFormatado1 = totalLinhasPreenchidas1.toLocaleString('pt-BR');

  //Planilha NovasURLs
  var planilha2 = planilhaDestino.getSheetByName(planilhaAlvo2);
  // Definir a planilha e a coluna a serem verificadas
  var colunaAlvo2 = planilha2.getRange(letraAlvo+":"+letraAlvo).getValues();
  var ultimaLinha2 = colunaAlvo2.filter(String).length;
  if(ultimaLinha2 === 1){
    ultimaLinha2 = 2;
  }
  var dados2 = planilha2.getRange(letraAlvo+"2:"+letraAlvo+ultimaLinha2).getValues();
  // Obter o total de linhas preenchidas
  var totalLinhasPreenchidas2 = dados2.filter(function (valor) {
    return valor[0] !== '';
  }).length;
  var totalFormatado2 = totalLinhasPreenchidas2.toLocaleString('pt-BR');

  //Planilha Sitemaps Antigos
  var planilha3 = planilhaDestino.getSheetByName("Sitemaps Antigos");
  // Definir a planilha e a coluna a serem verificadas
  var colunaAlvo3 = planilha3.getRange(letraAlvo+":"+letraAlvo).getValues();
  // Encontrar o valor na última linha
  var ultimaLinha3 = colunaAlvo3.filter(String).length;
  if(ultimaLinha3 === 1){
    ultimaLinha3 = 2;
  }
  var dados3 = planilha3.getRange(letraAlvo+"2:"+letraAlvo+ultimaLinha3).getValues();
  // Obter o total de linhas preenchidas
  var totalLinhasPreenchidas3 = dados3.filter(function (valor) {
    return valor[0] !== '';
  }).length;
  var valorUltimaLinha3 = planilha3.getRange(letraAlvo+ultimaLinha3+":"+letraAlvo+ultimaLinha3).getValues();
  var fusoHorario = Session.getScriptTimeZone(); // Obtém o fuso horário do script
  var dataCell = new Date(valorUltimaLinha3);
  var dataOntem = Utilities.formatDate(dataCell, fusoHorario, 'dd/MM/yyyy');
  var dataHoje = nomePlanilhaHoje.replace(/-/g, "/");

  //Planilha Sitemap Hoje
  var planilha4 = planilhaDestino.getSheetByName("Sitemap Hoje");
  // Definir a planilha e a coluna a serem verificadas
  var colunaAlvo4 = planilha4.getRange(letraAlvo+":"+letraAlvo).getValues();
  // Encontrar o valor na última linha
  var ultimaLinha4 = colunaAlvo4.filter(String).length;
  if(ultimaLinha4 === 1){
    ultimaLinha4 = 2;
  }
  var dados4 = planilha4.getRange(letraAlvo+"2:"+letraAlvo+ultimaLinha4).getValues();
  // Obter o total de linhas preenchidas
  var totalLinhasPreenchidas4 = dados4.filter(function (valor) {
    return valor[0] !== '';
  }).length;

  // Obter a data e hora atual
  var dataHoraAtual = new Date();
  var dataChecker = Utilities.formatDate(dataHoraAtual, fusoHorario, 'dd/MM/yyyy');
  var horaChecker = Utilities.formatDate(dataHoraAtual, fusoHorario, 'HH:mm');
  // Construir a mensagem com os resultados
  var inicioMensagem = 'Dados da comparação de Sitemaps<b>\n'+dataHoje+' vs '+dataOntem+'</b>\n\n';
  var percentualRemovidas = ((totalLinhasPreenchidas1/totalLinhasPreenchidas3)*100).toFixed(2);
  var percentualNovas = ((totalLinhasPreenchidas2/totalLinhasPreenchidas4)*100).toFixed(2);
  var mensagem = inicioMensagem+
    '<b>'+planilhaAlvo1+':</b> ' + totalFormatado1 + ' ('+percentualRemovidas+'%)\n' +
    '<b>'+planilhaAlvo2+':</b> ' + totalFormatado2 + ' ('+percentualNovas+'%)\n';
    mensagem += '\nInformações salvas nos arquivos <b>Registro de '+planilhaAlvo1+'</b> e <b>Registro de '+planilhaAlvo2+'</b>\n';
    mensagem += '\n<b>Checagem realizada em ' + dataChecker+' às '+horaChecker+'</b>';
  // Enviar mensagem pelo Telegram
  alertaComparadorTelegram(mensagem);
}

function alertaComparadorTelegram(mensagem) {
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
}

function registroDeComparador(nomePlanilhaOrigem,idDestino) {
  //EDITAR
  var abaOrigem = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomePlanilhaOrigem);
  var planilhaRegistro = SpreadsheetApp.openById(idDestino);
  //EDITAR
  // var hoje = new Date();
  // var nomeAba = Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd-MM-yyyy");
  var nomeAba = nomePlanilhaHoje;
  var aba = planilhaRegistro.getSheetByName(nomeAba);
  if (!aba) {
    aba = planilhaRegistro.insertSheet(nomeAba);  
    if(abaOrigem){
      var valores = abaOrigem.getDataRange().getValues();
      aba.getRange(1, 1, valores.length, valores[0].length).setValues(valores);
    }
  }
}
