function sendUrlsToGoogleIndex() {
  // Define a URL base da API de Indexação do Google
  var API_ENDPOINT = 'https://indexing.googleapis.com/v3/urlNotifications:publish';
  
  // Obtém a planilha ativa
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Obtém os dados da planilha
  var data = sheet.getDataRange().getValues();
  
  // Loop através dos dados da planilha e envia as URLs para indexação
  for (var i = 1; i < data.length; i++) {
    var url = data[i][0];
    var payload = {
      'url': url,
      'type': 'URL_UPDATED'
    };
    var options = {
      'method': 'POST',
      'muteHttpExceptions': true,
      'headers': {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      },
      'payload': JSON.stringify(payload)
    };
    var response = UrlFetchApp.fetch(API_ENDPOINT, options);
    var json = JSON.parse(response.getContentText());
    Logger.log(json);
  }
}
