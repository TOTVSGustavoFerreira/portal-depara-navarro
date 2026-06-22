/**
 * Portal De-Para TOTVS - Backend Version 1.3 (Google Apps Script)
 * Desenvolvido por Antigravity
 */

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  template.spreadsheetId = (e && e.parameter && e.parameter.id) ? e.parameter.id : "";
  return template.evaluate()
    .setTitle('Portal De-Para - TOTVS RM v1.3')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getActiveSS(spreadsheetId) {
  if (spreadsheetId && spreadsheetId !== "undefined" && spreadsheetId !== "null" && spreadsheetId !== "") {
    try {
      return SpreadsheetApp.openById(spreadsheetId);
    } catch (e) {
      Logger.log("Erro ao abrir planilha por ID (" + spreadsheetId + "): " + e.message);
    }
  }
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    Logger.log("Nenhuma planilha ativa encontrada: " + e.message);
    return null;
  }
}

/**
 * Busca uma aba na planilha de forma dinâmica (suporta singular/plural).
 */
function getSheetDynamic(ss, name) {
  if (!ss) return null;
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  
  // Tenta com 'S' no final (Plural)
  if (name.slice(-1).toUpperCase() !== "S") {
    sheet = ss.getSheetByName(name + "S") || ss.getSheetByName(name + "s");
    if (sheet) return sheet;
  }
  
  // Tenta remover 'S' no final (Singular)
  if (name.slice(-1).toUpperCase() === "S") {
    sheet = ss.getSheetByName(name.slice(0, -1));
    if (sheet) return sheet;
  }
  
  return null;
}

// Configurações Globais de Abas
var SHEET_ZDEPARA = "ZDEPARA_EVENTOS";
var SHEET_DADOS_RM = "DADOS_RM_EVENTOS";
var SHEET_CONFIG = "CONFIG_CONEXAO";
var SHEET_LOG = "LOG_SINCRONIZACAO";

// Lista oficial de 9 abas que devem ser sempre sincronizadas por padrão
var DEFAULT_SHEETS_TO_SYNC = [
  "ZDEPARA_COLIGADAS",
  "ZDEPARA_FUNCOES",
  "ZDEPARA_SINDICATOS",
  "ZDEPARA_SECOES",
  "ZDEPARA_EVENTOS",
  "ZDEPARA_SITUACAO",
  "DADOS_RM_MOTIVOS",
  "DADOS_RM_SITUACAO",
  "DADOS_RM_EVENTOS"
];

// Colunas obrigatórias da aba ZDEPARA_EVENTOS
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
 * Retorna as configurações de conexão e informações das planilhas.
 */
function getConnectionInfo(spreadsheetId) {
  var ss = getActiveSS(spreadsheetId);
  if (!ss) {
    return { isInstalled: false, extraSheets: [] };
  }
  var sheetConfig = ss.getSheetByName(SHEET_CONFIG);
  
  var info = {
    originName: ss.getName(),
    destinationUrl: "",
    destinationName: "Não conectada",
    destinationPath: "Não disponível",
    extraSheets: [],
    isInstalled: false
  };
  
  // Verificar se as abas essenciais existem
  var sheetDepara = ss.getSheetByName(SHEET_ZDEPARA);
  var sheetDadosRM = ss.getSheetByName(SHEET_DADOS_RM);
  
  if (sheetConfig && sheetDepara && sheetDadosRM) {
    info.isInstalled = true;
    var destUrl = sheetConfig.getRange("B1").getValue().toString().trim();
    info.destinationUrl = destUrl;
    
    if (destUrl) {
      var destId = extractSpreadsheetId(destUrl);
      if (destId) {
        try {
          var destSS = SpreadsheetApp.openById(destId);
          info.destinationName = destSS.getName();
          info.destinationPath = getDriveFilePath(destId);
        } catch (e) {
          info.destinationName = "Erro de acesso (verifique compartilhamento)";
        }
      }
    }
    
    // Ler abas extras adicionais na aba de configuração (a partir da linha 3)
    var lastRow = sheetConfig.getLastRow();
    if (lastRow >= 3) {
      var values = sheetConfig.getRange(3, 1, lastRow - 2, 1).getValues();
      values.forEach(function(row) {
        var val = row[0].toString().trim();
        if (val && DEFAULT_SHEETS_TO_SYNC.indexOf(val) === -1 && val !== SHEET_CONFIG && val !== SHEET_LOG) {
          info.extraSheets.push(val);
        }
      });
    }
  }
  
  return info;
}

/**
 * Extrai o ID da planilha a partir de uma URL ou retorna o próprio ID se já for um ID.
 */
function extractSpreadsheetId(urlOrId) {
  if (!urlOrId) return null;
  if (urlOrId.indexOf("docs.google.com/spreadsheets") !== -1) {
    var matches = urlOrId.match(/\/d\/([a-zA-Z0-9-_]+)/);
    return matches ? matches[1] : null;
  }
  return urlOrId;
}

/**
 * Obtém o caminho de pastas no Google Drive de um arquivo.
 */
function getDriveFilePath(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    var parents = file.getParents();
    var path = [];
    while (parents.hasNext()) {
      var folder = parents.next();
      path.unshift(folder.getName());
      parents = folder.getParents();
    }
    return path.length > 0 ? path.join(" > ") : "Raiz do Drive";
  } catch (e) {
    return "Drive Virtual (Acesso corporativo)";
  }
}

/**
 * Executa a auto-instalação das 9 tabelas e abas de configuração necessárias.
 */
function autoInstallStructure(spreadsheetId) {
  var ss = getActiveSS(spreadsheetId);
  if (!ss) {
    return { success: false, error: "Planilha não encontrada." };
  }
  
  try {
    // 1. Criar aba CONFIG_CONEXAO
    var sheetConfig = ss.getSheetByName(SHEET_CONFIG);
    if (!sheetConfig) {
      sheetConfig = ss.insertSheet(SHEET_CONFIG);
      sheetConfig.getRange("A1").setValue("Link Planilha Destino (Empresa):").setFontWeight("bold");
      sheetConfig.getRange("B1").setValue("https://docs.google.com/spreadsheets/d/1zEMXK--jTyXQKxFHpDbKaBtudXdi5urcAu7juaWJeuM/edit?usp=sharing");
      sheetConfig.getRange("A2").setValue("Abas Extras Personalizadas a Sincronizar:").setFontWeight("bold");
      sheetConfig.setColumnWidth(1, 280);
      sheetConfig.setColumnWidth(2, 450);
    } else {
      // Caso a aba já exista mas B1 esteja vazio, preencher com o padrão automaticamente
      var currentLink = sheetConfig.getRange("B1").getValue().toString().trim();
      if (!currentLink) {
        sheetConfig.getRange("B1").setValue("https://docs.google.com/spreadsheets/d/1zEMXK--jTyXQKxFHpDbKaBtudXdi5urcAu7juaWJeuM/edit?usp=sharing");
      }
    }
    
    // 2. Garantir a existência das 9 abas padrão
    DEFAULT_SHEETS_TO_SYNC.forEach(function(sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        
        // Estruturação inicial baseada na aba
        if (sheetName === "ZDEPARA_EVENTOS") {
          var deparaHeaders = [
            "EMPRESA_DE", "CODIGO_DE", "NOME_DE", "TIPO_EVENTO", 
            "COLIGADA_PARA", "CODIGO_PARA", "NOME_RM", 
            "CODIGO_PARA_FICHA_MES1", "CODIGO_PARA_FICHA_MES2", "CODIGO_PARA_VERBAS_FERIAS",
            "OBSERVACAO"
          ];
          sheet.getRange(1, 1, 1, deparaHeaders.length).setValues([deparaHeaders]).setFontWeight("bold").setBackground("#d9e1f2");
          // Formatar colunas de códigos como texto para manter os zeros à esquerda
          sheet.getRange("B2:B").setNumberFormat("@"); // CODIGO_DE
          sheet.getRange("F2:F").setNumberFormat("@"); // CODIGO_PARA
          sheet.getRange("H2:J").setNumberFormat("@"); // CODIGO_PARA_FICHA_MES1, CODIGO_PARA_FICHA_MES2, CODIGO_PARA_VERBAS_FERIAS
        } else if (sheetName === "DADOS_RM_EVENTOS") {
          var rmHeaders = ["CÓDIGO", "DESCRIÇÃO", "TIPO", "VALHORDIAREF", "NATE_ESOCIAL"];
          sheet.getRange(1, 1, 1, rmHeaders.length).setValues([rmHeaders]).setFontWeight("bold").setBackground("#e2efda");
          sheet.getRange("A2:A").setNumberFormat("@"); // CÓDIGO
        } else {
          // Outras abas apenas com cabeçalho genérico se criadas vazias
          sheet.getRange("A1").setValue("CODIGO").setFontWeight("bold");
          sheet.getRange("B1").setValue("DESCRICAO").setFontWeight("bold");
        }
      }
    });
    
    // 3. Criar aba LOG_SINCRONIZACAO
    var sheetLog = ss.getSheetByName(SHEET_LOG);
    if (!sheetLog) {
      sheetLog = ss.insertSheet(SHEET_LOG);
      var logHeaders = ["DATA/HORA", "USUÁRIO", "TIPO AÇÃO", "DETALHES DO PROCESSO", "STATUS"];
      sheetLog.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]).setFontWeight("bold").setBackground("#fff2cc");
    }
    
    SpreadsheetApp.flush();
    return { success: true, message: "Estrutura do Portal v1.3 instalada com sucesso! As 9 abas obrigatórias foram mapeadas. Link da planilha preenchido em CONFIG_CONEXAO." };
  } catch (e) {
    return { success: false, error: "Erro ao instalar: " + e.message };
  }
}

/**
 * Diagnóstico de Gaps de Criação Manual (restrito à faixa de Inclusões Manuais) e Duplicidades.
 */
function getReferenceDiagnostics(rmEvents) {
  var diag = {
    gaps: [],
    duplicates: []
  };
  
  if (!rmEvents || rmEvents.length === 0) return diag;
  
  // 1. Verificar Duplicidades no RM (apenas para novas inclusões manuais)
  var codeCounts = {};
  rmEvents.forEach(function(ev) {
    if (ev.codigo && ev.descricao && ev.descricao.indexOf("[INCLUSAO MANUAL]") !== -1) {
      codeCounts[ev.codigo] = (codeCounts[ev.codigo] || 0) + 1;
    }
  });
  
  for (var code in codeCounts) {
    if (codeCounts[code] > 1) {
      diag.duplicates.push(code);
    }
  }
  
  // 2. Verificar Gaps apenas nas Inclusões Manuais (que contêm [INCLUSAO MANUAL])
  var manualCodes = rmEvents.filter(function(ev) {
    return ev.descricao.indexOf("[INCLUSAO MANUAL]") !== -1;
  }).map(function(ev) {
    return parseInt(ev.codigo, 10);
  }).filter(function(num) {
    return !isNaN(num);
  }).sort(function(a, b) {
    return a - b;
  });
  
  if (manualCodes.length > 1) {
    var min = manualCodes[0];
    var max = manualCodes[manualCodes.length - 1];
    
    for (var i = min; i <= max; i++) {
      if (manualCodes.indexOf(i) === -1) {
        var gapStr = String(i);
        while (gapStr.length < 4) {
          gapStr = "0" + gapStr;
        }
        diag.gaps.push(gapStr);
      }
    }
  }
  
  return diag;
}

/**
 * Retorna todos os dados para alimentar a interface do Portal.
 */
function getPortalData(spreadsheetId) {
  var connection = getConnectionInfo(spreadsheetId);
  if (!connection.isInstalled) {
    return { status: { valid: false, errors: ["A estrutura v1.3 do portal não está configurada nesta planilha. Execute a instalação."] }, data: [], rmEvents: [], stats: {}, connection: connection };
  }
  
  var ss = getActiveSS(spreadsheetId);
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
  
  // Diagnóstico de gaps e duplicidades
  var diagnostics = getReferenceDiagnostics(rmEvents);
  
  // Calcular Estatísticas
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
  
  // Filtrar Novos Eventos Criados que Não Foram Utilizados
  var utilizedCodes = new Set();
  deparaData.forEach(function(item) {
    if (item.codigoPara) utilizedCodes.add(item.codigoPara);
    if (item.codigoParaFichaMes1) utilizedCodes.add(item.codigoParaFichaMes1);
    if (item.codigoParaFichaMes2) utilizedCodes.add(item.codigoParaFichaMes2);
    if (item.codigoParaVerbasFerias) utilizedCodes.add(item.codigoParaVerbasFerias);
  });
  
  var unusedCreatedEvents = rmEvents.filter(function(ev) {
    return ev.descricao.indexOf("[INCLUSAO MANUAL]") !== -1 && !utilizedCodes.has(ev.codigo);
  });
  
  var stats = {
    total: total,
    preenchidos: preenchidos,
    naoPreenchidos: naoPreenchidos,
    pAnalise: pAnalise,
    divergencias: divergenciasCount,
    gapsCount: diagnostics.gaps.length,
    duplicatesCount: diagnostics.duplicates.length,
    unusedCreatedCount: unusedCreatedEvents.length
  };
  
  return {
    status: { valid: true, errors: [] },
    data: deparaData,
    rmEvents: rmEvents,
    unusedCreatedEvents: unusedCreatedEvents,
    diagnostics: diagnostics,
    stats: stats,
    connection: connection
  };
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
 * Faz a busca da descrição de um evento.
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
 * Cria um novo evento RM na tabela DADOS_RM_EVENTOS
 */
function createNewRMEvent(spreadsheetId, eventData) {
  var ss = getActiveSS(spreadsheetId);
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
  
  var desc = "[INCLUSAO MANUAL] " + eventData.nomeDe;
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

/**
 * Exclui um Novo Evento Não Utilizado de forma segura.
 */
function deleteUnusedCreatedEvent(spreadsheetId, code) {
  try {
    var ss = getActiveSS(spreadsheetId);
    var sheetDadosRM = ss.getSheetByName(SHEET_DADOS_RM);
    
    var lastRowRM = sheetDadosRM.getLastRow();
    if (lastRowRM <= 1) return { success: false, error: "Nenhum evento cadastrado." };
    
    var rmDataRange = sheetDadosRM.getRange(2, 1, lastRowRM - 1, 2);
    var rmValues = rmDataRange.getValues();
    
    var targetRowIndex = -1;
    var isManual = false;
    
    for (var i = 0; i < rmValues.length; i++) {
      if (String(rmValues[i][0]) === code) {
        targetRowIndex = i + 2;
        isManual = rmValues[i][1].indexOf("[INCLUSAO MANUAL]") !== -1;
        break;
      }
    }
    
    if (targetRowIndex === -1) {
      return { success: false, error: "Código do evento não encontrado no RM." };
    }
    
    if (!isManual) {
      return { success: false, error: "Operação proibida: Não é possível deletar eventos padrão do RM." };
    }
    
    sheetDadosRM.deleteRow(targetRowIndex);
    SpreadsheetApp.flush();
    
    return { success: true, message: "Evento manual código " + code + " excluído. Uma lacuna (gap) foi gerada e poderá ser revisada no diagnóstico." };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Salva as edições locais da linha na Planilha Pessoal.
 */
function saveEventMapping(spreadsheetId, rowNum, data, applyToAllMatches) {
  var ss = getActiveSS(spreadsheetId);
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
    if (colColigadaPara) {
      if (data.coligadaPara === "" || data.coligadaPara === null || data.coligadaPara === undefined || data.coligadaPara === "-") {
        sheetDepara.getRange(r, colColigadaPara).clearContent();
      } else {
        sheetDepara.getRange(r, colColigadaPara).setValue(data.coligadaPara);
      }
    }
    
    var setFieldStaticValue = function(codColName, nomeColName, occNum, codValue) {
      var colCod = getColIndex(codColName);
      var colNome = getColIndex(nomeColName, occNum);
      
      if (colCod) {
        if (codValue === "" || codValue === null || codValue === undefined || codValue === "-") {
          sheetDepara.getRange(r, colCod).clearContent();
        } else {
          sheetDepara.getRange(r, colCod).setNumberFormat("@").setValue(String(codValue));
        }
      }
      
      if (colNome) {
        var descValue = lookupEventDescription(codValue, rmEventsList);
        if (descValue === "" || descValue === null || descValue === undefined) {
          sheetDepara.getRange(r, colNome).clearContent();
        } else {
          sheetDepara.getRange(r, colNome).setValue(descValue);
        }
      }
    };
    
    setFieldStaticValue("CODIGO_PARA", "NOME_RM", 0, data.codigoPara);
    setFieldStaticValue("CODIGO_PARA_FICHA_MES1", "NOME_RM", 1, data.codigoParaFichaMes1);
    setFieldStaticValue("CODIGO_PARA_FICHA_MES2", "NOME_RM", 2, data.codigoParaFichaMes2);
    setFieldStaticValue("CODIGO_PARA_VERBAS_FERIAS", "NOME_RM", 3, data.codigoParaVerbasFerias);
    
    var colObs = getColIndex("OBSERVACAO");
    if (colObs) {
      if (data.observacao === "" || data.observacao === null || data.observacao === undefined) {
        sheetDepara.getRange(r, colObs).clearContent();
      } else {
        sheetDepara.getRange(r, colObs).setValue(data.observacao);
      }
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
 * IMPORTAR: Copia as 9 abas obrigatórias da planilha da empresa para a pessoal.
 */
function importFromCompanySpreadsheet(spreadsheetId) {
  var connection = getConnectionInfo(spreadsheetId);
  if (!connection.destinationUrl) {
    return { success: false, error: "Nenhuma URL de planilha corporativa configurada na aba CONFIG_CONEXAO." };
  }
  
  var destId = extractSpreadsheetId(connection.destinationUrl);
  if (!destId) {
    return { success: false, error: "ID da planilha de destino inválido." };
  }
  
  try {
    var personalSS = getActiveSS(spreadsheetId);
    var companySS = SpreadsheetApp.openById(destId);
    
    // Unificar lista padrão de 9 abas + extras do usuário
    var sheetsToSync = DEFAULT_SHEETS_TO_SYNC.concat(connection.extraSheets);
    var details = [];
    
    sheetsToSync.forEach(function(sheetName) {
      var personalSheet = personalSS.getSheetByName(sheetName);
      var companySheet = companySS.getSheetByName(sheetName);
      
      if (personalSheet && companySheet) {
        var dataRange = companySheet.getDataRange();
        var values = dataRange.getValues();
        var backgrounds = dataRange.getBackgrounds();
        var numberFormats = dataRange.getNumberFormats();
        
        personalSheet.clearContents();
        personalSheet.clearFormats();
        
        personalSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
        personalSheet.getRange(1, 1, backgrounds.length, backgrounds[0].length).setBackgrounds(backgrounds);
        personalSheet.getRange(1, 1, numberFormats.length, numberFormats[0].length).setNumberFormats(numberFormats);
        
        details.push(sheetName + " (" + values.length + " linhas)");
      }
    });
    
    writeLog(spreadsheetId, "IMPORTAÇÃO", "Importados dados da Planilha Destino: " + details.join(", "), "SUCESSO");
    return { success: true, message: "Dados importados da TOTVS com sucesso!" };
  } catch (e) {
    writeLog(spreadsheetId, "IMPORTAÇÃO", "Erro na importação: " + e.message, "FALHA");
    return { success: false, error: e.message };
  }
}

/**
 * EXPORTAR: Copia os dados da planilha pessoal para a oficial da empresa.
 */
function syncToCompanySpreadsheet(spreadsheetId) {
  var connection = getConnectionInfo(spreadsheetId);
  if (!connection.destinationUrl) {
    return { success: false, error: "Nenhuma URL de planilha corporativa configurada na aba CONFIG_CONEXAO." };
  }
  
  var destId = extractSpreadsheetId(connection.destinationUrl);
  if (!destId) {
    return { success: false, error: "ID da planilha de destino inválido." };
  }
  
  try {
    var personalSS = getActiveSS(spreadsheetId);
    var companySS = SpreadsheetApp.openById(destId);
    
    // Unificar lista padrão de 9 abas + extras do usuário
    var sheetsToSync = DEFAULT_SHEETS_TO_SYNC.concat(connection.extraSheets);
    var details = [];
    
    sheetsToSync.forEach(function(sheetName) {
      var personalSheet = personalSS.getSheetByName(sheetName);
      var companySheet = companySS.getSheetByName(sheetName);
      
      if (personalSheet && companySheet) {
        var dataRange = personalSheet.getDataRange();
        var values = dataRange.getValues();
        var backgrounds = dataRange.getBackgrounds();
        var numberFormats = dataRange.getNumberFormats();
        
        companySheet.clearContents();
        companySheet.clearFormats();
        
        companySheet.getRange(1, 1, values.length, values[0].length).setValues(values);
        companySheet.getRange(1, 1, backgrounds.length, backgrounds[0].length).setBackgrounds(backgrounds);
        companySheet.getRange(1, 1, numberFormats.length, numberFormats[0].length).setNumberFormats(numberFormats);
        
        details.push(sheetName + " (" + values.length + " linhas)");
      }
    });
    
    writeLog(spreadsheetId, "EXPORTAÇÃO", "Exportados dados para a Planilha Destino: " + details.join(", "), "SUCESSO");
    return { success: true, message: "Sincronização concluída com sucesso! Verifique a planilha da TOTVS e os logs de auditoria." };
  } catch (e) {
    writeLog(spreadsheetId, "EXPORTAÇÃO", "Erro na exportação: " + e.message, "FALHA");
    return { success: false, error: e.message };
  }
}

/**
 * Grava uma linha na aba de LOG_SINCRONIZACAO.
 */
function writeLog(spreadsheetId, actionType, detail, status) {
  var userEmail = Session.getActiveUser().getEmail() || "Usuário do Portal";
  var timestamp = new Date();
  
  var writeToSheet = function(ss) {
    var sheetLog = ss.getSheetByName(SHEET_LOG);
    if (!sheetLog) {
      sheetLog = ss.insertSheet(SHEET_LOG);
      var logHeaders = ["DATA/HORA", "USUÁRIO", "TIPO AÇÃO", "DETALHES DO PROCESSO", "STATUS"];
      sheetLog.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]).setFontWeight("bold").setBackground("#fff2cc");
    }
    sheetLog.appendRow([timestamp, userEmail, actionType, detail, status]);
  };
  
  try {
    writeToSheet(getActiveSS(spreadsheetId));
  } catch(e) {}
  
  try {
    var connection = getConnectionInfo(spreadsheetId);
    var destId = extractSpreadsheetId(connection.destinationUrl);
    if (destId) {
      writeToSheet(SpreadsheetApp.openById(destId));
    }
  } catch(e) {}
}

/**
 * Cria o menu customizado quando a planilha é aberta.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Portal De-Para')
    .addItem('Abrir Portal (Tela Inteira)', 'openPortalDialog')
    .addItem('Abrir Portal (Menu Lateral)', 'openPortalSidebar')
    .addItem('Instalar Estrutura do Portal', 'autoInstallStructure')
    .addToUi();
}

function openPortalDialog() {
  var url = ScriptApp.getService().getUrl();
  if (!url) {
    url = "https://script.google.com/a/macros/totvs.com.br/s/AKfycbU5KymU1hENYrX5w8q5WGtisAtnipmPC0rNIGOZB0xgd4EEomXdlHtzG_6nRHAs3Sy/exec";
  }
  
  // Anexa o ID da planilha atual como parâmetro na URL
  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  url += (url.indexOf("?") === -1 ? "?" : "&") + "id=" + ssId;
  
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
      <p>Se a nova aba não abrir automaticamente devido ao bloqueador de popups do seu navegador, clique no botão abaixo:</p>
      <a href="${url}" target="_blank" class="btn" onclick="setTimeout(function(){ google.script.host.close(); }, 500);">Acessar Portal em Tela Cheia</a>
      <script>
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

function openPortalSidebar() {
  var template = HtmlService.createTemplateFromFile('index');
  try {
    template.spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  } catch (e) {
    template.spreadsheetId = "";
  }
  var html = template.evaluate()
    .setTitle('Portal De-Para - TOTVS RM');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * IMPORTAR UMA ABA: Copia uma única aba da planilha da empresa para a pessoal.
 */
function importSingleSheet(spreadsheetId, sheetName) {
  try {
    var connection = getConnectionInfo(spreadsheetId);
    var destId = extractSpreadsheetId(connection.destinationUrl);
    if (!destId) return { success: false, error: "Planilha de destino inválida." };
    
    var personalSS = getActiveSS(spreadsheetId);
    var companySS;
    try {
      companySS = SpreadsheetApp.openById(destId);
    } catch(err) {
      // Se destId não for um ID de planilha válido (ex: for o nome do arquivo "00-DE-PARA-NAVARRO -TESTE")
      if (destId.indexOf("/") === -1 && destId.length < 25) {
        return { 
          success: false, 
          error: "O link configurado na aba 'CONFIG_CONEXAO' (célula B1) é inválido (parece ser apenas o nome '" + destId + "'). Certifique-se de preencher a célula B1 com a URL completa da planilha oficial corporativa da TOTVS." 
        };
      }
      return { 
        success: false, 
        error: "Não foi possível abrir a Planilha Destino (ID: " + destId + "). Verifique se a célula B1 da aba 'CONFIG_CONEXAO' possui a URL correta e se você possui permissão de acesso a ela." 
      };
    }
    
    var companySheet = getSheetDynamic(companySS, sheetName);
    
    // Se a aba não existe na empresa (TOTVS), pulamos silenciosamente para manter o comportamento anterior
    if (!companySheet) {
      writeLog(spreadsheetId, "IMPORTAÇÃO PARCIAL", "Aba " + sheetName + " não encontrada na Planilha Destino (pulada).", "SUCESSO");
      return { success: true, message: sheetName + " pulado (não existe no destino)." };
    }
    
    // Pegar o nome exato da aba que foi encontrada na empresa
    var realSheetName = companySheet.getName();
    var personalSheet = getSheetDynamic(personalSS, realSheetName);
    
    // Se a aba não existe na pessoal, criamos automaticamente com o nome exato do destino
    if (!personalSheet) {
      personalSheet = personalSS.insertSheet(realSheetName);
    }
    
    var dataRange = companySheet.getDataRange();
    var values = dataRange.getValues();
    var backgrounds = dataRange.getBackgrounds();
    var numberFormats = dataRange.getNumberFormats();
    
    personalSheet.clearContents();
    personalSheet.clearFormats();
    
    personalSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    personalSheet.getRange(1, 1, backgrounds.length, backgrounds[0].length).setBackgrounds(backgrounds);
    personalSheet.getRange(1, 1, numberFormats.length, numberFormats[0].length).setNumberFormats(numberFormats);
    
    // Tratamento dos códigos dos eventos para manter zeros à esquerda (Evitar conversão em inteiro)
    if (realSheetName === "ZDEPARA_EVENTOS") {
      var personalHeaders = values[0];
      
      var colCodigoDeIdx = personalHeaders.indexOf("CODIGO_DE");
      var colCodigoParaIdx = personalHeaders.indexOf("CODIGO_PARA");
      var colFicha1Idx = personalHeaders.indexOf("CODIGO_PARA_FICHA_MES1");
      var colFicha2Idx = personalHeaders.indexOf("CODIGO_PARA_FICHA_MES2");
      var colFeriasIdx = personalHeaders.indexOf("CODIGO_PARA_VERBAS_FERIAS");
      
      var treatColumn = function(idx) {
        if (idx === -1) return;
        var rangeCol = personalSheet.getRange(2, idx + 1, values.length - 1, 1);
        rangeCol.setNumberFormat("@");
        var cellValues = rangeCol.getValues();
        var fixedValues = cellValues.map(function(row) {
          var val = row[0];
          if (val === "" || val === null || val === undefined) return [""];
          // Se for número, converter para string e preencher com zeros se necessário (ex: 76 -> "0076")
          if (typeof val === "number") {
            var strVal = String(val);
            while (strVal.length < 4) {
              strVal = "0" + strVal;
            }
            return [strVal];
          }
          return [String(val).trim()];
        });
        rangeCol.setValues(fixedValues);
      };
      
      if (values.length > 1) {
        treatColumn(colCodigoDeIdx);
        treatColumn(colCodigoParaIdx);
        treatColumn(colFicha1Idx);
        treatColumn(colFicha2Idx);
        treatColumn(colFeriasIdx);
      }
    } 
    else if (realSheetName === "DADOS_RM_EVENTOS") {
      var personalHeaders = values[0];
      var colCodigoIdx = personalHeaders.indexOf("CÓDIGO");
      if (colCodigoIdx === -1) colCodigoIdx = personalHeaders.indexOf("CODIGO");
      
      if (colCodigoIdx !== -1 && values.length > 1) {
        var rangeCol = personalSheet.getRange(2, colCodigoIdx + 1, values.length - 1, 1);
        rangeCol.setNumberFormat("@");
        var cellValues = rangeCol.getValues();
        var fixedValues = cellValues.map(function(row) {
          var val = row[0];
          if (val === "" || val === null || val === undefined) return [""];
          if (typeof val === "number") {
            var strVal = String(val);
            while (strVal.length < 4) {
              strVal = "0" + strVal;
            }
            return [strVal];
          }
          return [String(val).trim()];
        });
        rangeCol.setValues(fixedValues);
      }
    }
    
    writeLog(spreadsheetId, "IMPORTAÇÃO PARCIAL", "Importada aba: " + realSheetName + " (" + values.length + " linhas)", "SUCESSO");
    return { success: true, message: realSheetName + " importado." };
  } catch (e) {
    writeLog(spreadsheetId, "IMPORTAÇÃO PARCIAL", "Erro ao importar aba " + sheetName + ": " + e.message, "FALHA");
    return { success: false, error: e.message };
  }
}

/**
 * EXPORTAR UMA ABA: Copia os dados de uma única aba pessoal para a oficial da empresa.
 */
function syncSingleSheet(spreadsheetId, sheetName) {
  try {
    var connection = getConnectionInfo(spreadsheetId);
    var destId = extractSpreadsheetId(connection.destinationUrl);
    if (!destId) return { success: false, error: "Planilha de destino inválida." };
    
    var personalSS = getActiveSS(spreadsheetId);
    var companySS;
    try {
      companySS = SpreadsheetApp.openById(destId);
    } catch(err) {
      if (destId.indexOf("/") === -1 && destId.length < 25) {
        return { 
          success: false, 
          error: "O link configurado na aba 'CONFIG_CONEXAO' (célula B1) é inválido (parece ser apenas o nome '" + destId + "'). Certifique-se de preencher a célula B1 com a URL completa da planilha oficial corporativa da TOTVS." 
        };
      }
      return { 
        success: false, 
        error: "Não foi possível abrir a Planilha Destino (ID: " + destId + "). Verifique se a célula B1 da aba 'CONFIG_CONEXAO' possui a URL correta e se você possui permissão de acesso a ela." 
      };
    }
    
    var personalSheet = getSheetDynamic(personalSS, sheetName);
    
    // Se a aba não existe na pessoal, não temos o que exportar, pulamos
    if (!personalSheet) {
      return { success: true, message: sheetName + " pulado (não existe na origem)." };
    }
    
    var realSheetName = personalSheet.getName();
    var companySheet = getSheetDynamic(companySS, realSheetName);
    
    // Se não existe na de destino (empresa), criamos automaticamente lá com o mesmo nome exato
    if (!companySheet) {
      companySheet = companySS.insertSheet(realSheetName);
    }
    
    var personalRange = personalSheet.getDataRange();
    var personalValues = personalRange.getValues();
    var personalBackgrounds = personalRange.getBackgrounds();
    var personalNumberFormats = personalRange.getNumberFormats();
    
    if (personalValues.length === 0) {
      return { success: true, message: realSheetName + " está vazia na origem." };
    }
    
    var personalHeaders = personalValues[0];
    
    // Para abas de referência como DADOS_RM_EVENTOS e DADOS_RM_MOTIVOS (Tabelas de apoio do RM), 
    // nunca devemos apagar linhas nativas existentes ou colunas extras do cliente.
    // Faremos sincronização incremental baseada no Código/Chave Primária na primeira coluna.
    if (realSheetName.startsWith("DADOS_RM_")) {
      var companyRange = companySheet.getDataRange();
      var companyValues = companyRange.getValues();
      
      // Se a planilha de destino estiver vazia, podemos copiar tudo diretamente
      if (companyValues.length <= 1 || companyValues[0].length === 0) {
        companySheet.clearContents();
        companySheet.clearFormats();
        companySheet.getRange(1, 1, personalValues.length, personalValues[0].length).setValues(personalValues);
        companySheet.getRange(1, 1, personalBackgrounds.length, personalBackgrounds[0].length).setBackgrounds(personalBackgrounds);
        companySheet.getRange(1, 1, personalNumberFormats.length, personalNumberFormats[0].length).setNumberFormats(personalNumberFormats);
      } else {
        var companyHeaders = companyValues[0];
        
        // Mapear linhas do destino por chave primaria (coluna A)
        var companyKeysMap = {};
        for (var i = 1; i < companyValues.length; i++) {
          var key = String(companyValues[i][0]).trim();
          if (key) {
            companyKeysMap[key] = i + 1; // Guarda o número da linha correspondente (1-based)
          }
        }

        // Mapear índices das colunas correspondentes
        var colMap = []; // Array de { personalColIdx: num, companyColIdx: num }
        personalHeaders.forEach(function(pHeader, pIdx) {
          var cIdx = companyHeaders.indexOf(pHeader);
          if (cIdx !== -1) {
            colMap.push({ personalColIdx: pIdx, companyColIdx: cIdx });
          }
        });
        
        // Percorrer a planilha pessoal (origem) e atualizar/inserir
        var rowsAdded = 0;
        var rowsUpdated = 0;
        
        for (var j = 1; j < personalValues.length; j++) {
          var origKey = String(personalValues[j][0]).trim();
          if (!origKey) continue;
          
          var targetRow;
          if (companyKeysMap[origKey]) {
            targetRow = companyKeysMap[origKey];
            rowsUpdated++;
          } else {
            // Se for novo, inserimos uma nova linha no final e mapeamos
            companySheet.appendRow(new Array(companyHeaders.length).fill(""));
            targetRow = companySheet.getLastRow();
            companyKeysMap[origKey] = targetRow;
            rowsAdded++;
          }
          
          // Sobrescrever formatos e valores APENAS nas colunas mapeadas de forma segura e não destrutiva
          colMap.forEach(function(mapping) {
            var pColIdx = mapping.personalColIdx;
            var cColIdx = mapping.companyColIdx;
            
            var cellVal = personalValues[j][pColIdx];
            var cellBg = personalBackgrounds[j][pColIdx];
            var cellFmt = personalNumberFormats[j][pColIdx];
            
            var targetCell = companySheet.getRange(targetRow, cColIdx + 1);
            targetCell.setValue(cellVal);
            targetCell.setBackground(cellBg);
            targetCell.setNumberFormat(cellFmt);
          });
        }
        writeLog(spreadsheetId, "EXPORTAÇÃO INCREMENTAL", "Aba: " + realSheetName + " | Inseridos: " + rowsAdded + " | Atualizados: " + rowsUpdated, "SUCESSO");
      }
    } else {
      // Para abas ZDEPARA_COLIGADAS, ZDEPARA_EVENTOS, etc.
      // Em vez de limpar formatações inteiras ou apagar colunas que o usuário possa ter adicionado no destino (ex: IDs no início),
      // faremos um mapeamento por cabeçalho, escrevendo somente nas colunas correspondentes que existem em comum.
      var companyRange = companySheet.getDataRange();
      var companyValues = companyRange.getValues();
      
      if (companyValues.length <= 1 || companyValues[0].length === 0) {
        // Se vazia no destino, copia integralmente
        companySheet.clearContents();
        companySheet.clearFormats();
        companySheet.getRange(1, 1, personalValues.length, personalValues[0].length).setValues(personalValues);
        companySheet.getRange(1, 1, personalBackgrounds.length, personalBackgrounds[0].length).setBackgrounds(personalBackgrounds);
        companySheet.getRange(1, 1, personalNumberFormats.length, personalNumberFormats[0].length).setNumberFormats(personalNumberFormats);
      } else {
        var companyHeaders = companyValues[0];
        
        // Mapear índices das colunas correspondentes
        var colMap = []; // Array de { personalColIdx: num, companyColIdx: num }
        personalHeaders.forEach(function(pHeader, pIdx) {
          var cIdx = companyHeaders.indexOf(pHeader);
          if (cIdx !== -1) {
            colMap.push({ personalColIdx: pIdx, companyColIdx: cIdx });
          }
        });
        
        // Garantir o número de linhas no destino
        var diffRows = personalValues.length - companyValues.length;
        if (diffRows > 0) {
          companySheet.insertRowsAfter(companyValues.length, diffRows);
        } else if (diffRows < 0) {
          var startDelete = personalValues.length + 1;
          var numDelete = companyValues.length - personalValues.length;
          // IMPORTANTE: Limpar apenas as colunas que nós efetivamente gerenciamos (mapeadas), 
          // preservando colunas adicionais de terceiros/IDs no destino
          colMap.forEach(function(mapping) {
            var cColIdx = mapping.companyColIdx;
            companySheet.getRange(startDelete, cColIdx + 1, numDelete, 1).clearContent().clearFormat();
          });
        }
        
        // Atualizar células correspondentes
        colMap.forEach(function(mapping) {
          var pColIdx = mapping.personalColIdx;
          var cColIdx = mapping.companyColIdx;
          
          // Obter dados da coluna inteira na origem
          var colValues = [];
          var colBackgrounds = [];
          var colFormats = [];
          
          for (var r = 0; r < personalValues.length; r++) {
            colValues.push([personalValues[r][pColIdx]]);
            colBackgrounds.push([personalBackgrounds[r][pColIdx]]);
            colFormats.push([personalNumberFormats[r][pColIdx]]);
          }
          
          // Gravar na coluna de destino correspondente
          var targetRange = companySheet.getRange(1, cColIdx + 1, personalValues.length, 1);
          targetRange.setValues(colValues);
          targetRange.setBackgrounds(colBackgrounds);
          targetRange.setNumberFormats(colFormats);
        });
        
        writeLog(spreadsheetId, "EXPORTAÇÃO MAPEADA", "Aba: " + realSheetName + " (" + personalValues.length + " linhas atualizadas por coluna)", "SUCESSO");
      }
    }
    
    return { success: true, message: realSheetName + " sincronizado com segurança." };
  } catch (e) {
    writeLog(spreadsheetId, "EXPORTAÇÃO PARCIAL", "Erro ao exportar aba " + sheetName + ": " + e.message, "FALHA");
    return { success: false, error: e.message };
  }
}
