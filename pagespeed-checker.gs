var nomePlanilha = "Pagespeed"; // Insira o nome da planilha (aba) que você deseja escrever os dados de cobertura
var nomeSitemapPlanilha = "Sitemap"; // Insira o nome da planilha (aba) onde estão os dados das suas URLs
var planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomePlanilha); // Obtém a planilha pelo nome
var sitemapPlanilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeSitemapPlanilha); // Obtém a planilha pelo nome
var dataAtual = new Date();

function pagespeedChecker(){
  var colunaSitemap = sitemapPlanilha.getRange("A:A").getValues();
  var ultimaLinhaSitemap = colunaSitemap.filter(String).length;
  if(ultimaLinhaSitemap === 0){
    var ultimaLinhaSitemap = 1;
  }
  var dataSitemap = sitemapPlanilha.getRange(ultimaLinhaSitemap, 1).getValue();
  //
  var colunaURLs = planilha.getRange("A:A").getValues();
  var ultimaLinhaURLs = colunaURLs.filter(String).length;
  if(ultimaLinhaURLs === 0){
    var ultimaLinhaURLs = 1;
  }
  var dataURLs = planilha.getRange(ultimaLinhaURLs, 1).getValue();
  //
  var colunaPagespeed = planilha.getRange("B:B").getValues();
  var ultimaLinhaPagespeed = colunaPagespeed.filter(String).length;
  if(ultimaLinhaPagespeed === 0){
    var ultimaLinhaPagespeed = 1;
  }
  var dataPagespeed = planilha.getRange(ultimaLinhaPagespeed, 2).getValue();
  //
  // Compara as da verificação do sitemap com a data atual
  if (!(dataSitemap instanceof Date && dataSitemap.toDateString() === dataAtual.toDateString())) {
    //Se ainda não foi feita a atualização do sitemap hoje, encerra o script
    Logger.log("A planilha '"+nomeSitemapPlanilha+"' ainda não foi atualizada hoje.");
    return
  }else if ( dataPagespeed instanceof Date && dataPagespeed.toDateString() === dataAtual.toDateString() ) {
    //Se a atualização da planilha do PageSpeed já foi feita hoje, encerra o script
    Logger.log("A atualização da planilha '"+nomePlanilha+"' foi concluída hoje.");
    return
  }else if(ultimaLinhaURLs === ultimaLinhaPagespeed || ultimaLinhaURLs === 1){
    Logger.log("Limpando todos os dados de '"+nomePlanilha+"' antes da nova importação.");
    planilha.clearContents();
    Logger.log("Importando os dados de '"+nomeSitemapPlanilha+"' para '"+nomePlanilha+"'.");
    var valoresColunaA = sitemapPlanilha.getRange("A:A").getValues();
    for (var i = 0; i < valoresColunaA.length; i++) {
      var celula = valoresColunaA[i][0];
      if (celula.toString().indexOf("smplaces.") != -1) {
        celula = celula.toString().replace("smplaces.", "www.");
        valoresColunaA[i][0] = celula;
      }
    }
    planilha.getRange(1, 1, valoresColunaA.length, 1).setValues(valoresColunaA);
    Logger.log("Importação de '"+nomeSitemapPlanilha+"' para '"+nomePlanilha+"' finalizada.");
    writeMetrics(planilha);
    sumarioPagespeed();
  }else if(ultimaLinhaURLs !== ultimaLinhaPagespeed){
    Logger.log("Iniciando o processo de conexão com o PageSpeed Insights.")
    writeMetrics(planilha);
    sumarioPagespeed();
  }
}

function writeMetrics(planilha) {
  var colunaA = planilha.getRange("A:A").getValues();
  var ultimaLinhaA = colunaA.filter(String).length-1;
  var colunaB = planilha.getRange("B:B").getValues();
  var ultimaLinhaB = colunaB.filter(String).length;
  if(ultimaLinhaB===0){
    //Escreve os cabeçalhos na planilha
    Logger.log("Escrevendo os cabeçalhos em '"+nomePlanilha+"'.");
    planilha.getRange('B1').setValue('Score');
    planilha.getRange('C1').setValue('Largest Contentful Paint');
    planilha.getRange('D1').setValue('Cumulative Layout Shift');
    planilha.getRange('E1').setValue('First Contentful Paint');
    planilha.getRange('F1').setValue('Total Blocking Time');
    planilha.getRange('G1').setValue('Speed Index');
    var linhaInicial = 2;
  }else{
    var linhaInicial = ultimaLinhaB+1;
  }
  var urls = planilha.getRange("A"+linhaInicial+":A" + ultimaLinhaA).getValues();
  Logger.log("Iniciando captura de métricas das URLs");
  Logger.log("Começando na linha: "+linhaInicial);
  Logger.log("Total de urls: "+urls.length);
  for (var i = 0; i < urls.length; i++) {
    var linhaAtual = linhaInicial+i;
    var url = urls[i][0];
    var url_trimmed = url.trim();
    psConect(url_trimmed,linhaAtual);
  }
  var dataAtual = new Date();
  planilha.getRange(ultimaLinhaA+1,2).clearFormat();
  planilha.getRange(ultimaLinhaA+1,2).setValue(dataAtual); //escreve a data na primeira linha depois do último status
  //Faz o backup das informação eu uma outra planilha
  registroDePagespeed();
  Logger.log("Todos as métricas foram escritos na planilha.");
}

function psConect(url,linha) {
  Logger.log("Iniciando o processo de escrita ds métricas para a url "+url+ " na planilha "+nomePlanilha+" na linha "+linha);
  var apiKey = "AIzaSyAsxXOKXK75NzmnWXosnHij2UsOoXbPCQY";
  var pageSpeedEndpointUrl = "https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url=" + encodeURIComponent(url) + "&key=" + apiKey + "&strategy=mobile";
  var response = UrlFetchApp.fetch(pageSpeedEndpointUrl);
  var json = response.getContentText();
  var parsedJson = JSON.parse(json);
  var lighthouse = parsedJson['lighthouseResult'];
  var score = lighthouse['categories']['performance']['score']*100;
  var largestContentfulPaint = lighthouse['audits']['largest-contentful-paint']['numericValue']/1000;
  var cumulativeLayoutShift = lighthouse['audits']['cumulative-layout-shift']['numericValue'];
  var firstContentfulPaint = lighthouse['audits']['first-contentful-paint']['numericValue']/1000;
  var totalBlockingTime = lighthouse['audits']['total-blocking-time']['numericValue']/1000;
  var speedIndex = lighthouse['audits']['speed-index']['numericValue']/1000;
  planilha.getRange('B'+linha).clearFormat();
  planilha.getRange('B'+linha).setValue(Math.round(score));
  planilha.getRange('C'+linha).setValue(largestContentfulPaint);
  planilha.getRange('D'+linha).setValue(cumulativeLayoutShift);
  planilha.getRange('E'+linha).setValue(firstContentfulPaint);
  planilha.getRange('F'+linha).setValue(totalBlockingTime);
  planilha.getRange('G'+linha).setValue(speedIndex);

  var result = {
    'score': score,
    'largestContentfulPaint': largestContentfulPaint,
    'cumulativeLayoutShift': cumulativeLayoutShift,
    'firstContentfulPaint': firstContentfulPaint,
    'totalBlockingTime': totalBlockingTime,
    'speedIndex': speedIndex,
  }
  Logger.log(JSON.stringify(result, null, 2));
  Logger.log("As métricas para a url "+url+ "foram escritas na planilha "+nomePlanilha);
}

function sumarioPagespeed() {
  // Definir a planilha e a coluna a serem verificadas
  var colunaA = planilha.getRange("A:A").getValues();
  var ultimaLinhaA = colunaA.filter(String).length;
  var colunaB = planilha.getRange("B:B").getValues();
  var ultimaLinhaB = colunaB.filter(String).length;
  var penultimaLinhaB = colunaB.filter(String).length-1;
  var dados = planilha.getRange("B2:B"+penultimaLinhaB).getValues();
  // Encontrar o valor na última linha
  var valorUltimaLinhaA = planilha.getRange("A"+ultimaLinhaA+":A"+ultimaLinhaA).getValues();
  var valorUltimaLinhaB = planilha.getRange("B"+ultimaLinhaB+":B"+ultimaLinhaB).getValues();
  var fusoHorario = Session.getScriptTimeZone(); // Obtém o fuso horário do script
  var dataCellA = new Date(valorUltimaLinhaA); // Obtém a data atual
  var dataCellB = new Date(valorUltimaLinhaB); // Obtém a data atual
  var dataExtractionStart = Utilities.formatDate(dataCellA, fusoHorario, 'dd/MM/yyyy');
  var dataExtractionEnd = Utilities.formatDate(dataCellB, fusoHorario, 'dd/MM/yyyy');
  // Obter o total de linhas preenchidas
  var totalLinhasPreenchidas = dados.filter(function (valor) {
    return valor[0] !== '';
  }).length;
  var totalFormatado = totalLinhasPreenchidas.toLocaleString('pt-BR');
  // Obter a data e hora atual
  var dataHoraAtual = new Date();
  var dataChecker = Utilities.formatDate(dataHoraAtual, fusoHorario, 'dd/MM/yyyy');
  var horaChecker = Utilities.formatDate(dataHoraAtual, fusoHorario, 'HH:mm');
  // Construir a mensagem com os resultados
  var mensagem = 'Verificação iniciada em '+ dataExtractionStart + ' e finalizada em ' + dataExtractionEnd + '\n\n';
  mensagem += '<b>URLs checadas: ' + totalFormatado + '</b>\n\n';
  var totalURLs = 0;
  var scores = 0;
  var totalHome = 0;
  var scoreHome = 0;
  var totalLp = 0;
  var scoreLp = 0;
  var totalCat = 0;
  var scoreCat = 0;
  var totalProd = 0;
  var scoreProd = 0;
  for (var i = 0; i < dados.length; i++) {
    var score = dados[i][0];
    var linha = i+2;
    var url = planilha.getRange("A"+linha).getValue();
    totalURLs++;
    scores += score;

    if(url.endsWith('/')){
      var indexUltimaBarra = url.lastIndexOf('/');
      var urlFinal = url.substring(0, indexUltimaBarra);
    }else{
      var urlFinal = url;
    }
    
    var numeroBarras = (urlFinal.match(/\//g) || []).length;
    var diretorios = numeroBarras-2;

    if(url === "https://www.oiplace.com.br/") {
      totalHome++;
      scoreHome += score;
      Logger.log("Home | "+url);

    } else if (url.endsWith(".html")) {
      if (diretorios === 1) {
        totalLp++;
        scoreLp += score;
        Logger.log("LP | "+url);
      }else if (diretorios === 2) {
        totalProd++;
        scoreProd += score;
        Logger.log("Prod | "+url);
      }
    } else if (url.indexOf("/perguntas-frequentes/") !== -1 || url.indexOf("/minha-conta/") !== -1 ) {
      totalLp++;
      scoreLp += score;
      Logger.log("LP | "+url);
    }else{
      totalCat++;
      scoreCat += score;
      Logger.log("Cat | "+url);
    }
  }

  mensagem += '<b>Média geral de score:</b> '+(scores/totalURLs).toFixed(0)+'\n\n';
  mensagem += '<b>Score da home:</b> '+(scoreHome/totalHome).toFixed(0)+'\n';
  mensagem += '<b>Score médio das páginas de produto:</b> '+(scoreProd/totalProd).toFixed(0)+'\n';
  mensagem += '<b>Score médio das páginas de categoria:</b> '+(scoreCat/totalCat).toFixed(0)+'\n';
  mensagem += '<b>Score médio das landing pages:</b> '+(scoreLp/totalLp).toFixed(0)+'\n';
  mensagem += '\nDados salvos na planilha <b>Registro de PageSpeed</b> em ' + dataChecker+' às '+horaChecker;
  Logger.log(mensagem);
  // Enviar mensagem pelo Telegram
  alertaPagespeedTelegram(mensagem);
}

function alertaPagespeedTelegram(mensagem) {
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
  Logger.log("Dados de checagem enviados no Telegram.");
}

function registroDePagespeed() {
  var planilhaOrigem = planilha;
  // EDITAR
  var planilhaDestino = SpreadsheetApp.openById('1kab8sJY_WDHlV39UTmd8Z8bj2ofyBC0QBcBUJ6q_NL4');
  // EDITAR
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
