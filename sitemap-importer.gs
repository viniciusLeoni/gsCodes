function sitemapImporter() {
  //EDITAR
  var sitemap_url = "https://www.oiplace.com.br/sitemap_index.xml"; // Insira a URL do arquivo XML que você deseja importar
  var nomePlanilha = "Sitemap"; // Insira o nome da planilha (aba) que você deseja escrever os dados do XML
  //EDITAR
  var response = UrlFetchApp.fetch(sitemap_url);
  var planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomePlanilha); // Obtém a planilha pelo nome
  //Checa se a planilha existe
  if (!planilha) {
    // Verifica se a planilha foi encontrada
    // Logger.log("A planilha com o nome '" + nomePlanilha + "' não foi encontrada.");
    return;
  }
  //Checa se a planilha já foi completamente preenchida hoje
  var ultimaLinha = planilha.getLastRow();
  var dataAtual = new Date(); // Obtém a data atual
  if (ultimaLinha !== 0) {
    var valorUltimaCelulaA = planilha.getRange(ultimaLinha, 1).getValue();
    var valorUltimaCelulaB = planilha.getRange(ultimaLinha, 2).getValue();
    if (valorUltimaCelulaA instanceof Date && valorUltimaCelulaA.toDateString() === dataAtual.toDateString() && valorUltimaCelulaB instanceof Date && valorUltimaCelulaB.toDateString() === dataAtual.toDateString()) {
      // Logger.log("A planilha com o nome '" + nomePlanilha + "' já foi preenchida hoje.");
      var cellB1 = planilha.getRange(1,2).getValue();
      if(cellB1 !== "Códigos"){
        planilha.getRange(1,2).setValue("Códigos")
        sumarioSitemap(planilha);
      }
      return;
    }
  }
  //Checa se o sitemap existe e retorna 200
  if (response.getResponseCode() == 200) {
    //Importa as urls do sitemap para a planilha 
    var xml = UrlFetchApp.fetch(sitemap_url).getContentText();
    var document = XmlService.parse(xml);
    var root = document.getRootElement();
    //Checa se tem a data atual está preenchida na última célula da Coluna A, se não, importa os dados
    if (!(valorUltimaCelulaA instanceof Date && valorUltimaCelulaA.toDateString() === dataAtual.toDateString())) {
      // Limpa os dados existentes na planilha
      planilha.clearContents();
      // Logger.log("Planilha foi completamente limpa. Iniciando os trabalhos.");
      //Insere os dados na planilha
      importarXML(sitemap_url,planilha,root,xml);
    }
  }else{
    // Logger.log("Falha ao obter o sitemap. Status code: " + response.getResponseCode());
  }
}

function importarXML(sitemap_url,planilha,root,xml) {
  var limitador = 5;
  var ns = XmlService.getNamespace("http://www.sitemaps.org/schemas/sitemap/0.9"); // Define o namespace do XML
  var locs = []; // Array para armazenar os valores de loc
  if (xml.indexOf("<sitemap>") !== -1) { 
    var sitemaps = root.getChildren("sitemap", ns); // Obtém todos os elementos filho "url" do elemento raiz, usando o namespace
    // Logger.log("Iniciando sitemap índice - "+sitemap_url);
    // Loop para percorrer os elementos url e obter os valores de loc
    for (var j = 0; j < sitemaps.length; j++) {
      var urlElementSitemap = sitemaps[j]; // Obtém o elemento url atual
      var locElementSitemap = urlElementSitemap.getChild("loc", ns); // Obtém o elemento loc dentro do elemento url, usando o namespace
      var locValueSitemap = locElementSitemap.getValue(); // Obtém o valor de texto do elemento loc
      // Logger.log("Iniciando o sitemap único - "+locValueSitemap);
      var xmlFilho = UrlFetchApp.fetch(locValueSitemap).getContentText(); // Faz a solicitação do URL e obtém o conteúdo em formato de texto
      var documentFilho = XmlService.parse(xmlFilho); // Analisa o XML obtido em um objeto Document
      var rootFilho = documentFilho.getRootElement(); // Obtém o elemento raiz do XML
      var urls = rootFilho.getChildren("url", ns); // Obtém todos os elementos filho "url" do elemento raiz, usando o namespace
      // Logger.log(urls.length);
      // Loop para percorrer os elementos url e obter os valores de loc
      for (var i = 0; i < urls.length; i++) {
        var urlElement = urls[i]; // Obtém o elemento url atual
        var locElement = urlElement.getChild("loc", ns); // Obtém o elemento loc dentro do elemento url, usando o namespace
        var locValue = locElement.getValue(); // Obtém o valor de texto do elemento loc
        if (locValue.toString().indexOf("smplaces.") != -1) {
          locValue = locValue.toString().replace("smplaces.", "www.");
        }
        if (!locValue.includes("/demandware.store/")) {
          // Logger.log(locValue);
          locs.push([locValue]); // Adiciona o valor de loc ao array locs
        }
        if(i >= 100){
          break;
        }
      }
      // Logger.log("Finalizando o sitemap único - "+locValueSitemap);
    }
    // Logger.log("Finalizando sitemap índice - "+sitemap_url);
  }else{
    // Logger.log("Iniciando sitemap único - "+sitemap_url);
    var urls = root.getChildren("url", ns); // Obtém todos os elementos filho "url" do elemento raiz, usando o namespace
    // Loop para percorrer os elementos url e obter os valores de loc
    for (var i = 0; i < urls.length; i++) {
      var urlElement = urls[i]; // Obtém o elemento url atual
      var locElement = urlElement.getChild("loc", ns); // Obtém o elemento loc dentro do elemento url, usando o namespace
      var locValue = locElement.getValue(); // Obtém o valor de texto do elemento loc
      if (locValue.toString().indexOf("smplaces.") != -1) {
        locValue = locValue.toString().replace("smplaces.", "www.");
      }
      // Logger.log(locValue);
      if (!locValue.includes("/demandware.store/")) {
        locs.push([locValue]); // Adiciona o valor de loc ao array locs
      }
      if(i >= 100){
          break;
        }
    }
    // Logger.log("Finalizando sitemap único - "+sitemap_url);
  }
  planilha.getRange(1,1).setValue("URLs");
  planilha.getRange(2, 1, locs.length, locs[0].length).setValues(locs);
  var dataAtual = new Date();
  planilha.getRange(locs.length+2,1).setValue(dataAtual); //escreve a data na primeira linha depois do último status
}
