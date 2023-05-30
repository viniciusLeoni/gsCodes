var nomePlanilha = "Cobertura"; // Insira o nome da planilha (aba) que você deseja escrever os dados de cobertura
var nomeSitemapPlanilha = "Sitemap"; // Insira o nome da planilha (aba) onde estão os dados das suas URLs
var siteUrl = "sc-domain:oiplace.com.br";
var siteUrlAlt = "https://www.oiplace.com.br/";

var planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomePlanilha); // Obtém a planilha pelo nome
var sitemapPlanilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeSitemapPlanilha); // Obtém a planilha pelo nome
var dataAtual = new Date();

function coverageChecker(){
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
  var colunaCoverage = planilha.getRange("B:B").getValues();
  var ultimaLinhaCoverage = colunaCoverage.filter(String).length;
  if(ultimaLinhaCoverage === 0){
    var ultimaLinhaCoverage = 1;
  }
  var dataCoverage = planilha.getRange(ultimaLinhaCoverage, 2).getValue();
  //
  // Compara as da verificação do sitemap com a data atual
  if (!(dataSitemap instanceof Date && dataSitemap.toDateString() === dataAtual.toDateString())) {
    //Se ainda não foi feita a atualização do sitemap hoje, encerra o script
    // Logger.log("A planilha '"+nomeSitemapPlanilha+"' ainda não foi atualizada hoje.");
    return
  }else if ( dataCoverage instanceof Date && dataCoverage.toDateString() === dataAtual.toDateString() ) {
    //Se a atualização da planilha de Cobertura já foi feita hoje, encerra o script
    // Logger.log("A atualização da planilha '"+nomePlanilha+"' foi concluída hoje.");
    return
  }else if(ultimaLinhaURLs === ultimaLinhaCoverage || ultimaLinhaURLs === 1){
    // Logger.log("Limpando todos os dados de '"+nomePlanilha+"' antes da nova importação.");
    planilha.clearContents();
    // Logger.log("Importando os dados de '"+nomeSitemapPlanilha+"' para '"+nomePlanilha+"'.");
    var valoresColunaA = sitemapPlanilha.getRange("A:A").getValues();
    for (var i = 0; i < valoresColunaA.length; i++) {
      var celula = valoresColunaA[i][0];
      if (celula.toString().indexOf("smplaces.") != -1) {
        celula = celula.toString().replace("smplaces.", "www.");
        valoresColunaA[i][0] = celula;
      }
    }
    planilha.getRange(1, 1, valoresColunaA.length, 1).setValues(valoresColunaA);
    // Logger.log("Importação de '"+nomeSitemapPlanilha+"' para '"+nomePlanilha+"' finalizada.");
    //Escreve os cabeçalhos na planilha
    // Logger.log("Escrevendo os cabeçalhos em '"+nomePlanilha+"'.");
    planilha.getRange('B1').setValue('coverageState');
    planilha.getRange('C1').setValue('pageFetchState');
    planilha.getRange('D1').setValue('verdict');
    planilha.getRange('E1').setValue('googleCanonical');
    planilha.getRange('F1').setValue('lastCrawlTime');
    gscConect();
  }else if(ultimaLinhaURLs !== ultimaLinhaCoverage){
    //Verifica se estamos no horário determinado e, se sim, envia mensagem com o status atual no Telegram
    var horaAtual = dataAtual.getHours();
    if(planilha.getRange('B1').getValue() !== "coverage" && horaAtual >= 15 && horaAtual < 16 ) {
      // Logger.log("Trocando o cabeçalho da coluna B.");
      planilha.getRange('B1').setValue('coverage');
      //Cria a mensagem para ser enviada para o Telegram
      sumarioCobertura("parcial","D");
    }else if(planilha.getRange('B1').getValue() === "coverage" && horaAtual >= 16 && horaAtual < 17 ){
      //Reescreve o cabeçalho para o padrão após envio de mensagem
      planilha.getRange('B1').setValue('coverageState');
    }
    // Conecta com a API do GSC
    gscConect();
    // Envia sumário no Telegram
    sumarioCobertura("final","D");
  }
}

function gscConect(){
  // Logger.log("Conectando com a API do GSC.");
  var colunaA = planilha.getRange("A:A").getValues();
  var ultimaLinhaA = (colunaA.filter(String).length)-1;
  if(ultimaLinhaA === 0){
    var ultimaLinhaA = 1;
  }
  var colunaB = planilha.getRange("B:B").getValues();
  var ultimaLinhaB = colunaB.filter(String).length;
  if(ultimaLinhaB === 0){
    var ultimaLinhaB = 1;
  }
  //Verifica a última linha da coluna B pra definir a linha inicial do intervalo;
  if(ultimaLinhaB === 0 || ultimaLinhaA === ultimaLinhaB){
    var linhaInicial = 2;
  }else{
    var linhaInicial = ultimaLinhaB+1;
  }
  //Inicia o processo de conexão com o GSC
  // Logger.log("Linha inicial:"+linhaInicial);
  // Logger.log("Linha final:"+ultimaLinhaA);
  var urls = planilha.getRange("A"+linhaInicial+":A"+ultimaLinhaA).getValues();
  var valoresURLs = urls.map(function(row) {
    return row[0]; // Extrai os valores da primeira coluna (coluna A)
  });
  // Logger.log("Escrevendo os dados do GSC na planilha '"+nomePlanilha+"'.");
  for (var i = 0; i < valoresURLs.length; i++) {
    var url = urls[i].toString();
    try {
      var request = {
        "inspectionUrl": url,
        "siteUrl": siteUrl,
        "languageCode": "pt-BR"
      };
      var response = UrlFetchApp.fetch('https://searchconsole.googleapis.com/v1/urlInspection/index:inspect', {
        headers: {
          Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
        },
        'method' : 'post',
        'contentType' : 'application/json',
        'payload' : JSON.stringify(request, null, 2)
      });
    } catch (e) {
      var request = {
        "inspectionUrl": url,
        "siteUrl": siteUrlAlt,
        "languageCode": "pt-BR"
      };
      var response = UrlFetchApp.fetch('https://searchconsole.googleapis.com/v1/urlInspection/index:inspect', {
        headers: {
          Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
        },
        'method' : 'post',
        'contentType' : 'application/json',
        'payload' : JSON.stringify(request, null, 2)
      });
    }
    var json = JSON.parse(response.getContentText());
    //Escreve linha por linha as informações da inspeção do GSC na planilha
    gscWriteInspect(linhaInicial+i,url,json);
    //return //apagar
  }
  //Inclui a data na última linha da coluna A
  planilha.getRange('B' + (ultimaLinhaA+1)).setValue(dataAtual);
}

function gscWriteInspect(writeLine,url,json){
  // Logger.log("Escrevendo a linha referente a '"+url+"'.");
  var coverageState = json.inspectionResult.indexStatusResult.coverageState;
  var pageFetchState = json.inspectionResult.indexStatusResult.pageFetchState;
  var pageFetchState;
  //Traduz o resultado de pageFetchState
  switch (pageFetchState) {
    case "PAGE_FETCH_STATE_UNSPECIFIED":
      pageFetchState = "Desconhecido";
      break;
    case "SUCCESSFUL":
      pageFetchState = "Bem-sucedida";
      break;
    case "SOFT_404":
      pageFetchState = "Erro soft 404";
      break;
    case "BLOCKED_ROBOTS_TXT":
      pageFetchState = "	Bloqueado pelo robots.txt";
      break;
    case "NOT_FOUND":
      pageFetchState = "Não encontrado (404)";
      break;
    case "ACCESS_DENIED":
      pageFetchState = "	Bloqueado devido a solicitação não autorizada (401)";
      break;
    case "SERVER_ERROR":
      pageFetchState = "Erro no servidor (5xx)";
      break;
    case "REDIRECT_ERROR":
      pageFetchState = "Erro de redirecionamento";
      break;
    case "ACCESS_FORBIDDEN":
      pageFetchState = "Bloqueado devido a acesso proibido (403)";
      break;
    case "BLOCKED_4XX":
      pageFetchState = "Bloqueado devido a outro problema 4xx (não 403, 404)";
      break;
    case "INTERNAL_CRAWL_ERROR":
      pageFetchState = "Erro interno";
      break;
    case "INVALID_URL":
      pageFetchState = "Erro soft 404";
      break;
    default:
      pageFetchState = "URL inválido";
      break;
  }
  var verdict = json.inspectionResult.indexStatusResult.verdict;
  var verdict;
  //Traduz o resultado de verdict
  switch (verdict) {
    case "VERDICT_UNSPECIFIED":
      verdict = "Não especificado";
      break;
    case "PASS":
      verdict = "Válido";
      break;
    case "PARTIAL":
      verdict = "Aprovado/Não utilizado";
      break;
    case "FAIL":
      verdict = "Inválido";
      break;
    case "NEUTRAL":
      verdict = "Excluído";
      break;
    default:
      verdict = "Resultado desconhecido";
      break;
  }
  var googleCanonical = json.inspectionResult.indexStatusResult.googleCanonical;
  var lastCrawlTime = json.inspectionResult.indexStatusResult.lastCrawlTime;
  // Logger.log("Escrevendo na linha'"+writeLine+"'.");
  planilha.getRange('B' + (writeLine)).setValue(coverageState);
  planilha.getRange('C' + (writeLine)).setValue(pageFetchState);
  if(verdict !== ""){
    planilha.getRange('D' + (writeLine)).setValue(verdict);
  }else{
    planilha.getRange('D' + (writeLine)).setValue("Desconhecido");
  }
  planilha.getRange('E' + (writeLine)).setValue(googleCanonical);
  planilha.getRange('F' + (writeLine)).setValue(lastCrawlTime);
}

function sumarioCobertura(status,letraAlvo) {
  // Definir a planilha e a coluna a serem verificadas
  var colunaAlvo = planilha.getRange(letraAlvo+":"+letraAlvo).getValues();
  var ultimaLinha = colunaAlvo.filter(String).length;
  var penultimaLinha = colunaAlvo.filter(String).length-1;
  var dados = planilha.getRange(letraAlvo+"2:"+letraAlvo+penultimaLinha).getValues();
  // Encontrar o valor na última linha
  var valorUltimaLinha = planilha.getRange(letraAlvo+ultimaLinha+":"+letraAlvo+ultimaLinha).getValues();
  var fusoHorario = Session.getScriptTimeZone(); // Obtém o fuso horário do script
  var dataCell = new Date(valorUltimaLinha); // Obtém a data atual
  var dataExtraction = Utilities.formatDate(dataCell, fusoHorario, 'dd/MM/yyyy');
  // Obter o total de linhas preenchidas
  var totalLinhasPreenchidas = dados.filter(function (valor) {
    return valor[0] !== '';
  }).length;
  var totalFormatado = totalLinhasPreenchidas.toLocaleString('pt-BR');
  // Obter os valores únicos encontrados
  var valoresUnicos = [...new Set(dados.flat())];
  // Calcular a porcentagem de cada valor único referente ao total de linhas preenchidas
  var resultados = valoresUnicos.map(function (valorUnico) {
    var totalOcorrencias = dados.filter(function (valor) {
      return valor[0] === valorUnico;
    }).length;
    var porcentagem = (totalOcorrencias / totalLinhasPreenchidas) * 100;
    return { valor: valorUnico, total: totalOcorrencias, porcentagem: porcentagem };
  });;
  // Obter a data e hora atual
  var dataHoraAtual = new Date();
  var dataChecker = Utilities.formatDate(dataHoraAtual, fusoHorario, 'dd/MM/yyyy');
  var horaChecker = Utilities.formatDate(dataHoraAtual, fusoHorario, 'HH:mm');
  // Construir a mensagem com os resultados
  if(status === "final"){
    var inicioMensagem = 'Dados <b>finais</b> da inspeção de cobertura.\nInformações salvas no arquivo "Registros de cobertura"\n\n';
    registroDeCobertura();
  }else{
    var inicioMensagem = 'Dados <b>parciais</b> da inspeção de cobertura.\n\n';
  }
  var mensagem = inicioMensagem+
    'URLs inspecionadas: ' + totalFormatado + '\n' +
    '\n<b>Status de cobertura encontrados</b>\n';
  resultados.forEach(function (resultado) {
    mensagem += resultado.valor + ': ' + resultado.porcentagem.toFixed(2).replace(/\./g, ',') + '% (' + resultado.total.toLocaleString('pt-BR') + ')\n';
  });
  mensagem += '\n<b>Checagem realizada em ' + dataChecker+' às '+horaChecker+'</b>';
  // Enviar mensagem pelo Telegram
  Logger.log(mensagem);
  alertaCoverageTelegram(mensagem);
}

function alertaCoverageTelegram(mensagem) {
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
  // Logger.log("Dados de checagem enviados no Telegram.");
}

function registroDeCobertura() {
  var planilhaOrigem = planilha;
  // EDITAR
  var planilhaDestino = SpreadsheetApp.openById('18by-hVTXKUMUV1__t6dD_McndhtP0sm9c4_-VPby3k8');
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
