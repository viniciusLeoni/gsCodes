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

  getMediumScore("B");
  getMediumScore("C");
  getMediumScore("D");
  getMediumScore("E");
  getMediumScore("F");
  getMediumScore("G");
    
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

function getMediumScore(colunaAlvo) {

  var arquivo = SpreadsheetApp.getActiveSpreadsheet();
  var sumario = arquivo.getSheetByName("Sumário");
  var dataRange = sumario.getRange("A2:A" + sumario.getLastRow());
  var datas = dataRange.getValues();
  
  // Limpa os dados existentes nas colunas B e C da planilha "Sumário"
  sumario.getRange(colunaAlvo+":"+colunaAlvo).clearContent();
  
  // Escreve os índices das colunas em B1 e C1, respectivamente
  sumario.getRange("B1").setValue("Score Médio");
  sumario.getRange("C1").setValue("LCP Médio");
  sumario.getRange("D1").setValue("CLS Médio");
  sumario.getRange("E1").setValue("FCP Médio");
  sumario.getRange("F1").setValue("TBT Médio");
  sumario.getRange("G1").setValue("SI Médio");
  
  // Itera sobre as datas na coluna A do "Sumário"
  for (var i = 0; i < datas.length; i++) {
    var data = Utilities.formatDate(datas[i][0], Session.getScriptTimeZone(), "dd-MM-yyyy");
    var planilha = arquivo.getSheetByName(data);
    
    if (planilha) {
      // Obtém os valores da coluna D da planilha correspondente à data
     
      var valoresColuna = planilha.getRange(colunaAlvo+":"+colunaAlvo).getValues();

      var ultimaLinhaAlvo = (valoresColuna.filter(String).length)-1;
      var valoresColunaAlvo = planilha.getRange(colunaAlvo+":"+colunaAlvo+ultimaLinhaAlvo).getValues();
    
      // Conta todas os valores da coluna alvo
      var total = -1; //começando negativo para descontar a primeira linha que é um cabeçalho    
      // Soma todas os valores da coluna alvo
      var soma = 0;
      for (var j = 0; j < valoresColuna.length; j++) {
        var valor = valoresColuna[j][0];
        if (j > 0 && valor !== "" && !(valor instanceof Date) ) {
          total++;
          soma += valor;
        }
      }
      
      // Escreve o score médio referente a data
      var linha = i + 2;
      var scoreMedio = soma/total;
      if(colunaAlvo === "B"){
        sumario.getRange(colunaAlvo+linha).setValue(Math.round(scoreMedio));
      }else{
        sumario.getRange(colunaAlvo+linha).setValue(scoreMedio);
      }
      
    }
  }
}
