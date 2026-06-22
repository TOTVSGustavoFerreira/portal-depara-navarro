/**
 * Portal De-Para TOTVS - Backend (Google Apps Script)
 * Desenvolvido por Antigravity
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Portal De-Para - TOTVS RM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Configurações Globais
var SHEET_ZDEPARA = "ZDEPARA_EVENTOS";
var SHEET_DADOS_RM = "DADOS_RM_EVENTOS";

// Mapeamento das colunas esperadas na ZDEPARA_EVENTOS
var EXPECTED_COLUMNS_DEPARA = [
  "EMPRESA_DE",
  "CODIGO_DE",
  "NOME_DE",
  "TIPO_EVENTO",
  "COLIGADA_PARA",
  "CODIGO_PARA",
  "NOME_RM"
];

/**
 * Realiza o diagnóstico da planilha e verifica se a estrutura necessária está correta.
 */
function getSystemStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var status = {
    valid: true,
    errors: [],
    warnings: [],
    sheets: {
      deparaExists: false,
      dadosRmExists: false
    },
    columns: {
      deparaMissing: [],
      hasObservacao: false
    }
  };
  
  var sheetDepara = ss.getSheetByName(SHEET_ZDEPARA);
  var sheetDadosRM = ss.getSheetByName(SHEET_DADOS_RM);
  
  if (sheetDepara) {
    status.sheets.deparaExists = true;
  } else {
    status.valid = false;
    status.errors.push("Aba '" + SHEET_ZDEPARA + "' não encontrada na planilha.");
  }
  
  if (sheetDadosRM) {
    status.sheets.dadosRmExists = true;
  } else {
    status.valid = false;
    status.errors.push("Aba '" + SHEET_DADOS_RM + "' não encontrada na planilha.");
  }
  
  if (sheetDepara) {
    var headers = sheetDepara.getRange(1, 1, 1, sheetDepara.getLastColumn()).getValues()[0];
    
    EXPECTED_COLUMNS_DEPARA.forEach(function(col) {
      if (headers.indexOf(col) === -1) {
        if (col === "NOME_RM") {
          var countNomeRm = headers.filter(function(h) { return h === "NOME_RM"; }).length;
          if (countNomeRm === 0) {
            status.valid = false;
            status.errors.push("Coluna obrigatória '" + col + "' não encontrada na aba " + SHEET_ZDEPARA);
            status.columns.deparaMissing.push(col);
          }
        } else {
          status.valid = false;
          status.errors.push("Coluna obrigatória '" + col + "' não encontrada na aba " + SHEET_ZDEPARA);
          status.columns.deparaMissing.push(col);
        }
      }
    });
    
    var obsIndex = headers.indexOf("OBSERVACAO");
    if (obsIndex !== -1) {
      status.columns.hasObservacao = true;
    } else {
      status.warnings.push("Coluna 'OBSERVACAO' não encontrada. Ela será criada automaticamente.");
    }
  }
  
  return status;
}

/**
 * Garante que a coluna OBSERVACAO exista na planilha ZDEPARA_EVENTOS.
 */
function checkAndCreateObservationColumn(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var obsIndex = headers.indexOf("OBSERVACAO");
  if (obsIndex === -1) {
    var lastCol = sheet.getLastColumn();
    sheet.getRange(1, lastCol + 1).setValue("OBSERVACAO");
    SpreadsheetApp.flush();
    return lastCol + 1;
  }
  return obsIndex + 1;
}

/**
 * Retorna a descrição de um evento com base no código lido da aba DADOS_RM_EVENTOS.
 * Retorna texto estático para evitar uso de fórmulas complexas que geram erros na planilha do usuário.
 */
function lookupEventDescription(code, rmEventsList) {
  if (!code) return "";
  if (code === "NAO IMPORTAR" || code === "P/ ANALISE") return code;
  
  for (var i = 0; i < rmEventsList.length; i++) {
    if (rmEventsList[i].codigo === code) {
      return rmEventsList[i].descricao;
    }
  }
  return "";
}

/**
 * Retorna todos os dados para alimentar o Portal (status, dados da tabela, autocomplete e estatísticas).
 */
function getPortalData() {
  var status = getSystemStatus();
  if (!status.valid) {
    return { status: status, data: [], rmEvents: [], stats: {} };
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDepara = ss.getSheetByName(SHEET_ZDEPARA);
  var sheetDadosRM = ss.getSheetByName(SHEET_DADOS_RM);
  
  checkAndCreateObservationColumn(sheetDepara);
  
  var headers = sheetDepara.getRange(1, 1, 1, sheetDepara.getLastColumn()).getValues()[0];
  
  var colIndexes = {};
  headers.forEach(function(h, idx) {
    if (!colIndexes[h]) {
      colIndexes[h] = [];
    }
    colIndexes[h].push(idx);
  });
  
  var lastRow = sheetDepara.getLastRow();
  var rawData = [];
  if (lastRow > 1) {
    rawData = sheetDepara.getRange(2, 1, lastRow - 1, headers.length).getValues();
  }
  
  var deparaData = rawData.map(function(row, rIdx) {
    var rowNum = rIdx + 2;
    
    var getValueByName = function(name, occurrenceIndex) {
      occurrenceIndex = occurrenceIndex || 0;
      var indexes = colIndexes[name];
      if (indexes && indexes[occurrenceIndex] !== undefined) {
        return row[indexes[occurrenceIndex]];
      }
      return "";
    };
    
    return {
      rowNum: rowNum,
      empresaDe: String(getValueByName("EMPRESA_DE")),
      codigoDe: String(getValueByName("CODIGO_DE")),
      nomeDe: String(getValueByName("NOME_DE")),
      tipoEvento: String(getValueByName("TIPO_EVENTO")),
      coligadaPara: String(getValueByName("COLIGADA_PARA")),
      codigoPara: String(getValueByName("CODIGO_PARA")),
      nomeRm: String(getValueByName("NOME_RM", 0)),
      
      codigoParaFichaMes1: String(getValueByName("CODIGO_PARA_FICHA_MES1")),
      nomeRmFichaMes1: String(getValueByName("NOME_RM", 1)),
      
      codigoParaFichaMes2: String(getValueByName("CODIGO_PARA_FICHA_MES2")),
      nomeRmFichaMes2: String(getValueByName("NOME_RM", 2)),
      
      codigoParaVerbasFerias: String(getValueByName("CODIGO_PARA_VERBAS_FERIAS")),
      nomeRmVerbasFerias: String(getValueByName("NOME_RM", 3)),
      
      observacao: String(getValueByName("OBSERVACAO"))
    };
  });
  
  var rmEvents = [];
  var lastRowRM = sheetDadosRM.getLastRow();
  if (lastRowRM > 1) {
    var rawRM = sheetDadosRM.getRange(2, 1, lastRowRM - 1, 3).getValues();
    rmEvents = rawRM.map(function(row) {
      return {
        codigo: String(row[0]),
        descricao: String(row[1]),
        tipo: String(row[2])
      };
    });
  }
  
  // Estatísticas e Divergências
  var total = deparaData.length;
  var preenchidos = 0;
  var naoPreenchidos = 0;
  var pAnalise = 0;
  
  var keyMap = {};
  var duplicateKeys = {};
  
  deparaData.forEach(function(item) {
    var cod = item.codigoPara;
    if (cod === "P/ ANALISE") {
      pAnalise++;
    } else if (!cod) {
      naoPreenchidos++;
    } else {
      preenchidos++;
    }
    
    if (item.nomeDe && item.tipoEvento && cod && cod !== "P/ ANALISE") {
      var key = item.nomeDe.toLowerCase() + "|||" + item.tipoEvento;
      if (!keyMap[key]) {
        keyMap[key] = [];
      }
      if (keyMap[key].indexOf(cod) === -1) {
        keyMap[key].push(cod);
      }
      if (keyMap[key].length > 1) {
        duplicateKeys[key] = true;
      }
    }
  });
  
  var divergenciasCount = 0;
  deparaData.forEach(function(item) {
    if (item.nomeDe && item.tipoEvento) {
      var key = item.nomeDe.toLowerCase() + "|||" + item.tipoEvento;
      if (duplicateKeys[key]) {
        item.hasDivergencia = true;
        divergenciasCount++;
      }
    }
  });
  
  var stats = {
    total: total,
    preenchidos: preenchidos,
    naoPreenchidos: naoPreenchidos,
    pAnalise: pAnalise,
    divergencias: divergenciasCount
  };
  
  return {
    status: status,
    data: deparaData,
    rmEvents: rmEvents,
    stats: stats
  };
}

/**
 * Salva o mapeamento de um ou múltiplos eventos (em lote).
 */
function saveEventMapping(rowNum, data, applyToAllMatches) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDepara = ss.getSheetByName(SHEET_ZDEPARA);
  var sheetDadosRM = ss.getSheetByName(SHEET_DADOS_RM);
  var headers = sheetDepara.getRange(1, 1, 1, sheetDepara.getLastColumn()).getValues()[0];
  
  var colIndexes = {};
  headers.forEach(function(h, idx) {
    if (!colIndexes[h]) {
      colIndexes[h] = [];
    }
    colIndexes[h].push(idx);
  });
  
  var getColIndex = function(name, occurrence) {
    occurrence = occurrence || 0;
    return (colIndexes[name] && colIndexes[name][occurrence] !== undefined) ? colIndexes[name][occurrence] + 1 : null;
  };
  
  // Obter lista atualizada de eventos RM para fazer o Procv/Lookup estático no backend
  var rmEventsList = [];
  var lastRowRM = sheetDadosRM.getLastRow();
  if (lastRowRM > 1) {
    var rawRM = sheetDadosRM.getRange(2, 1, lastRowRM - 1, 2).getValues();
    rmEventsList = rawRM.map(function(row) {
      return { codigo: String(row[0]), descricao: String(row[1]) };
    });
  }
  
  var saveRow = function(r) {
    var colColigadaPara = getColIndex("COLIGADA_PARA");
    if (colColigadaPara && data.coligadaPara) {
      sheetDepara.getRange(r, colColigadaPara).setValue(data.coligadaPara);
    }
    
    var setFieldStaticValue = function(codColName, nomeColName, occNum, codValue) {
      var colCod = getColIndex(codColName);
      var colNome = getColIndex(nomeColName, occNum);
      
      if (colCod) {
        sheetDepara.getRange(r, colCod).setValue(codValue);
      }
      
      if (colNome) {
        // Gravar descrição estática em formato texto limpo, resolvendo 100% dos problemas de fórmulas
        var descValue = lookupEventDescription(codValue, rmEventsList);
        sheetDepara.getRange(r, colNome).setValue(descValue);
      }
    };
    
    setFieldStaticValue("CODIGO_PARA", "NOME_RM", 0, data.codigoPara);
    setFieldStaticValue("CODIGO_PARA_FICHA_MES1", "NOME_RM", 1, data.codigoParaFichaMes1);
    setFieldStaticValue("CODIGO_PARA_FICHA_MES2", "NOME_RM", 2, data.codigoParaFichaMes2);
    setFieldStaticValue("CODIGO_PARA_VERBAS_FERIAS", "NOME_RM", 3, data.codigoParaVerbasFerias);
    
    var colObs = getColIndex("OBSERVACAO");
    if (colObs) {
      sheetDepara.getRange(r, colObs).setValue(data.observacao || "");
    }
  };
  
  saveRow(rowNum);
  
  var affectedRows = 1;
  if (applyToAllMatches && data.nomeDe && data.tipoEvento) {
    var lastRow = sheetDepara.getLastRow();
    if (lastRow > 1) {
      var colNomeDeIdx = headers.indexOf("NOME_DE");
      var colTipoEventoIdx = headers.indexOf("TIPO_EVENTO");
      
      if (colNomeDeIdx !== -1 && colTipoEventoIdx !== -1) {
        var allRows = sheetDepara.getRange(2, 1, lastRow - 1, headers.length).getValues();
        allRows.forEach(function(rowValues, idx) {
          var currentRowNum = idx + 2;
          if (currentRowNum !== rowNum) {
            var cellNomeDe = String(rowValues[colNomeDeIdx]);
            var cellTipoEvento = String(rowValues[colTipoEventoIdx]);
            
            if (cellNomeDe === data.nomeDe && cellTipoEvento === data.tipoEvento) {
              saveRow(currentRowNum);
              affectedRows++;
            }
          }
        });
      }
    }
  }
  
  SpreadsheetApp.flush();
  return { success: true, affectedRows: affectedRows };
}

/**
 * Cria um novo evento RM na tabela DADOS_RM_EVENTOS
 */
function createNewRMEvent(eventData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDadosRM = ss.getSheetByName(SHEET_DADOS_RM);
  
  if (!sheetDadosRM) {
    return { success: false, error: "Aba de dados de referência RM não encontrada." };
  }
  
  var lastRow = sheetDadosRM.getLastRow();
  
  var nextCodeStr = "0001";
  if (lastRow > 1) {
    var allCodes = sheetDadosRM.getRange(2, 1, lastRow - 1, 1).getValues().map(function(r) {
      return parseInt(r[0], 10);
    }).filter(function(v) {
      return !isNaN(v);
    });
    
    if (allCodes.length > 0) {
      var maxCode = Math.max.apply(null, allCodes);
      var nextCodeNum = maxCode + 1;
      nextCodeStr = String(nextCodeNum);
      while (nextCodeStr.length < 4) {
        nextCodeStr = "0" + nextCodeStr;
      }
    }
  }
  
  var desc = "[PENDENTE RM] " + eventData.nomeDe;
  var tipo = "PROVENTO";
  if (eventData.tipoEvento === "D-DESCONTO" || eventData.tipoEvento === "DESCONTO") {
    tipo = "DESCONTO";
  } else if (eventData.tipoEvento === "B-BASE" || eventData.tipoEvento === "BASE") {
    tipo = "BASE DE CALCULO";
  }
  
  var valHorDiaRef = "VALOR";
  var natEsocial = "9999-Outros";
  
  var newRow = lastRow + 1;
  sheetDadosRM.getRange(newRow, 1, 1, 5).setValues([[
    nextCodeStr,
    desc,
    tipo,
    valHorDiaRef,
    natEsocial
  ]]);
  
  sheetDadosRM.getRange(newRow, 1, 1, 5).setBackground("#FFF2CC");
  
  SpreadsheetApp.flush();
  
  return {
    success: true,
    code: nextCodeStr,
    description: desc,
    type: tipo
  };
}

function getColumnLetter(colIndex) {
  var letter = "";
  while (colIndex > 0) {
    var temp = (colIndex - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    colIndex = (colIndex - temp - 1) / 26;
  }
  return letter;
}

/**
 * Cria o menu customizado quando a planilha é aberta.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Portal De-Para')
    .addItem('Abrir Portal (Tela Inteira)', 'openPortalDialog')
    .addItem('Abrir Portal (Menu Lateral)', 'openPortalSidebar')
    .addToUi();
}

/**
 * Abre o portal como uma nova aba cheia no navegador (ideal para usar 100% da tela).
 */
function openPortalDialog() {
  var url = ScriptApp.getService().getUrl();
  if (!url) {
    // Fallback usando a URL do Web App informada
    url = "https://script.google.com/a/macros/totvs.com.br/s/AKfycbU5KymU1hENYrX5w8q5WGtisAtnipmPC0rNIGOZB0xgd4EEomXdlHtzG_6nRHAs3Sy/exec";
  }
  
  var htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@400;600&display=swap" rel="stylesheet">
      <style>
        body { font-family: 'Outfit', sans-serif; text-align: center; padding: 24px; background-color: #f4f7f9; margin: 0; }
        .btn { display: inline-block; padding: 12px 24px; background: #002233; color: white; text-decoration: none; border-radius: 8px; font-weight: 600; font-size: 0.95rem; margin-top: 15px; border: 1px solid #002233; transition: all 0.3s; cursor: pointer; }
        .btn:hover { background: #00DBFF; color: #002233; border-color: #00DBFF; }
        p { color: #64748b; font-size: 0.9rem; margin: 0 0 10px 0; }
      </style>
    </head>
    <body>
      <p style="font-weight: 600; color: #002233; font-size: 1.05rem;">Abrindo Portal De-Para em Tela Cheia...</p>
      <p>Se a nova aba não abrir automaticamente devido ao bloqueador de popups do navegador, clique no botão abaixo:</p>
      <a href="${url}" target="_blank" id="openLink" class="btn" onclick="setTimeout(function(){ google.script.host.close(); }, 500);">Acessar Portal em Tela Cheia</a>
      <script>
        // Tenta abrir automaticamente em nova aba
        var win = window.open("${url}", "_blank");
        if (win) {
          setTimeout(function() { google.script.host.close(); }, 1200);
        }
      </script>
    </body>
    </html>
  `;
  
  var html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(450)
    .setHeight(220)
    .setTitle('Redirecionando para o Portal');
  SpreadsheetApp.getUi().showModalDialog(html, 'Portal De-Para - TOTVS RM');
}

/**
 * Abre o portal como um painel acoplado na lateral direita da tela.
 */
function openPortalSidebar() {
  var html = HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Portal De-Para - TOTVS RM');
  SpreadsheetApp.getUi().showSidebar(html);
}

// ID da planilha oficial da empresa (TOTVS)
var COMPANY_SPREADSHEET_ID = "1tEYVPT9ulbQMq34qdA1zItgi1gGdsfxlLgF90Ld0YXg";

/**
 * Sincroniza a planilha de trabalho pessoal para a planilha oficial da empresa.
 * Copia os dados das abas ZDEPARA_EVENTOS e DADOS_RM_EVENTOS.
 */
function syncToCompanySpreadsheet() {
  try {
    var personalSS = SpreadsheetApp.getActiveSpreadsheet();
    var companySS = SpreadsheetApp.openById(COMPANY_SPREADSHEET_ID);
    
    // Lista de abas a serem sincronizadas
    var sheetsToSync = [SHEET_ZDEPARA, SHEET_DADOS_RM];
    
    sheetsToSync.forEach(function(sheetName) {
      var personalSheet = personalSS.getSheetByName(sheetName);
      var companySheet = companySS.getSheetByName(sheetName);
      
      if (personalSheet && companySheet) {
        // Obter todos os valores e formatações da planilha pessoal
        var dataRange = personalSheet.getDataRange();
        var values = dataRange.getValues();
        var backgrounds = dataRange.getBackgrounds();
        
        // Limpar a aba correspondente na planilha da empresa
        companySheet.clearContents();
        companySheet.clearFormats();
        
        // Inserir os novos valores na planilha da empresa
        companySheet.getRange(1, 1, values.length, values[0].length).setValues(values);
        companySheet.getRange(1, 1, backgrounds.length, backgrounds[0].length).setBackgrounds(backgrounds);
      }
    });
    
    return { success: true, message: "Sincronização concluída com sucesso!" };
  } catch (error) {
    return { success: false, error: "Erro ao sincronizar: " + error.message };
  }
}
