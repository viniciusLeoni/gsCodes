var arquivo = SpreadsheetApp.getActiveSpreadsheet();
var planilhas = arquivo.getSheets();
var sumario = arquivo.getSheetByName("Sumário");
  
function criarSumario() {
  
  // Limpa os dados existentes na coluna A da planilha "Sumário"
  sumario.clearContents();
  
  var nomesPlanilhas = [];
  for (var i = 0; i < planilhas.length; i++) {
    var planilha = planilhas[i];
    if (planilha.getName() !== "Sumário") {
      nomesPlanilhas.push(planilha.getName());
    }
  }
  
  // Ordena os nomes das planilhas por data (da mais antiga para a mais nova)
  nomesPlanilhas.sort(function(a, b) {
    var dataA = getDataDaPlanilha(a);
    var dataB = getDataDaPlanilha(b);
    if (dataA < dataB) {
      return -1;
    }
    if (dataA > dataB) {
      return 1;
    }
    return 0;
  });

  // Escreve "Data" em A1
  sumario.getRange("A1").setValue("Data");  
  
  // Grava os nomes das planilhas ordenados na coluna A da planilha "Sumário"
  var range = sumario.getRange(2, 1, nomesPlanilhas.length, 1);
  range.setValues(nomesPlanilhas.map(function(nome) { return [nome]; }));

  //Modificar esse trecho de acordo com a métrica desejada
  preencherPlanilha("A");
    
}

function getDataDaPlanilha(nome) {
  // Extrai a data da string do nome da planilha
  var partes = nome.split("-");
  var dia = partes[0];
  var mes = partes[1];
  var ano = partes[2];

  // Retorna um objeto Date com a data extraída
  return new Date(ano, mes - 1, dia);
}

function preencherPlanilha(colunaAlvo) {
  var arquivo = SpreadsheetApp.getActiveSpreadsheet();
  var sumario = arquivo.getSheetByName("Sumário");
  var dataRange = sumario.getRange("A2:A" + sumario.getLastRow());
  var datas = dataRange.getValues();
  
  // Limpa os dados existentes nas colunas B a Z da planilha "Sumário"
  sumario.getRange("B:Z").clearContent();

  sumario.getRange(1, 2).setValue("URLs Removidas");

  for (var i = 0; i < datas.length; i++) {
    var data = Utilities.formatDate(datas[i][0], Session.getScriptTimeZone(), "dd-MM-yyyy");
    var planilha = arquivo.getSheetByName(data);
    
    if (planilha) {

      // Obtém os valores da coluna D da planilha correspondente à data
      var valoresColuna = planilha.getRange(colunaAlvo+"2:"+colunaAlvo).getValues();
      // Obter o total de linhas preenchidas
      var totalLinhasPreenchidas = valoresColuna.filter(function (valor) {
        return valor[0] !== '';
      }).length;
      var linhaEscrita = i + 2;
      sumario.getRange(linhaEscrita, 2).setValue(totalLinhasPreenchidas);

    }

  }

}
