var arquivo = SpreadsheetApp.getActiveSpreadsheet();
var planilhas = arquivo.getSheets();
var sumario = arquivo.getSheetByName("Sumário");
var sumarioPrioridades = arquivo.getSheetByName("Sumário Prioridades");
  
function criarSumario() {
  
  // Limpa os dados existentes na coluna A da planilha "Sumário"
  sumario.clearContents();
  
  var nomesPlanilhas = [];
  for (var i = 0; i < planilhas.length; i++) {
    var planilha = planilhas[i];
    if (planilha.getName() !== "Sumário" && planilha.getName() !== "Sumário Prioridades") {
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
  criarSumarioPrioridades();
    
}

function criarSumarioPrioridades() {
  
  // Limpa os dados existentes na coluna A da planilha "Sumário"
  sumarioPrioridades.clearContents();
  
  var nomesPlanilhas = [];
  for (var i = 0; i < planilhas.length; i++) {
    var planilha = planilhas[i];
    if (planilha.getName() !== "Sumário" && planilha.getName() !== "Sumário Prioridades") {
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
  sumarioPrioridades.getRange("A1").setValue("Data");  
  
  // Grava os nomes das planilhas ordenados na coluna A da planilha "Sumário"
  var range = sumarioPrioridades.getRange(2, 1, nomesPlanilhas.length, 1);
  range.setValues(nomesPlanilhas.map(function(nome) { return [nome]; }));

  //Modificar esse trecho de acordo com a métrica desejada
  preencherPlanilhaPrioridades("A","B");
    
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
  var dataRange = sumario.getRange("A2:A" + sumario.getLastRow());
  var datas = dataRange.getValues();
  
  // Limpa os dados existentes nas colunas B a Z da planilha "Sumário"
  sumario.getRange("B:Z").clearContent();

  var valoresUnicos = [];

  for (var i = 0; i < datas.length; i++) {
    var data = Utilities.formatDate(datas[i][0], Session.getScriptTimeZone(), "dd-MM-yyyy");
    var planilha = arquivo.getSheetByName(data);
    
    if (planilha) {

      // Obtém os valores da coluna D da planilha correspondente à data
      var valoresColuna = planilha.getRange(colunaAlvo+"2:"+colunaAlvo).getValues();
          
      for (var j = 0; j < valoresColuna.length; j++) {

        var valor = valoresColuna[j][0];
        if (valor !== "" && valoresUnicos.indexOf(valor) === -1 && !(valor instanceof Date) ) {
          valoresUnicos.push(valor);
        }
        
      }

    }

  }

  for (var k = 0; k < valoresUnicos.length; k++) {

    var valorUnico = valoresUnicos[k];
    var colunaEscrita = k+2;
    sumario.getRange(1,colunaEscrita).setValue(valorUnico);

    for (var i = 0; i < datas.length; i++) {
      var data = Utilities.formatDate(datas[i][0], Session.getScriptTimeZone(), "dd-MM-yyyy");
      var planilha = arquivo.getSheetByName(data);
      
      if (planilha) {
        // Obtém os valores da coluna D da planilha correspondente à data
        var valoresColuna = planilha.getRange(colunaAlvo+"2:"+colunaAlvo).getValues();
        
        // Conta as ocorrências na coluna        
        var contador = [];
        contador[k] = 0;

        for (var j = 0; j < valoresColuna.length; j++) {
          
          var valor = valoresColuna[j][0];
          

          if (valor === valorUnico) {
            contador[k]++;
          }
        
        }
        
        // Escreve o contador de na linha correspondente a partir da coluna B
        var linhaEscrita = i + 2;
        sumario.getRange(linhaEscrita, colunaEscrita).setValue(contador[k]);

      
      }
    }
    
  }

}

function preencherPlanilhaPrioridades(colunaAlvo,colunaFiltro) {
  var arquivo = SpreadsheetApp.getActiveSpreadsheet();
  var dataRange = sumarioPrioridades.getRange("A2:A" + sumarioPrioridades.getLastRow());
  var datas = dataRange.getValues();
  
  // Limpa os dados existentes nas colunas B a Z da planilha "Sumário"
  sumarioPrioridades.getRange("B:Z").clearContent();

  var valoresAlvoUnicos = [];
  var valoresFiltroUnicos = [];

  for (var i = 0; i < datas.length; i++) {
    var data = Utilities.formatDate(datas[i][0], Session.getScriptTimeZone(), "dd-MM-yyyy");
    var planilha = arquivo.getSheetByName(data);
    
    if (planilha) {

      // Obtém os valores da coluna da planilha correspondente à data
      var valoresColunaAlvo = planilha.getRange(colunaAlvo+"2:"+colunaAlvo).getValues();
      var valoresColunaFiltro = planilha.getRange(colunaFiltro+"2:"+colunaFiltro).getValues();
          
      for (var j = 0; j < valoresColunaAlvo.length; j++) {

        var valorAlvo = valoresColunaAlvo[j][0];
        var valorFiltro = valoresColunaFiltro[j][0];
        
        if (valorAlvo !== "" && valoresAlvoUnicos.indexOf(valorAlvo) === -1 && !(valorAlvo instanceof Date) && valorFiltro === 1) {
          valoresAlvoUnicos.push(valorAlvo);
        }
        
      }

    }

  }

  for (var k = 0; k < valoresAlvoUnicos.length; k++) {

    var valorAlvoUnico = valoresAlvoUnicos[k];
    var colunaEscrita = k+2;
    sumarioPrioridades.getRange(1,colunaEscrita).setValue(valorAlvoUnico);

    for (var i = 0; i < datas.length; i++) {
      var data = Utilities.formatDate(datas[i][0], Session.getScriptTimeZone(), "dd-MM-yyyy");
      var planilha = arquivo.getSheetByName(data);
      
      if (planilha) {
        // Obtém os valores da coluna da planilha correspondente à data
        var valoresColunaAlvo = planilha.getRange(colunaAlvo+"2:"+colunaAlvo).getValues();
        
        // Conta as ocorrências na coluna        
        var contador = [];
        contador[k] = 0;

        for (var j = 0; j < valoresColunaAlvo.length; j++) {
          
          var valorAlvo = valoresColunaAlvo[j][0];
          var valorFiltro = valoresColunaFiltro[j][0];

          if (valorAlvo === valorAlvoUnico && valorFiltro === 1) {
            contador[k]++;
          }
        
        }
        
        // Escreve o contador de na linha correspondente a partir da coluna B
        var linhaEscrita = i + 2;
        sumarioPrioridades.getRange(linhaEscrita, colunaEscrita).setValue(contador[k]);

      
      }
    }
    
  }

}
