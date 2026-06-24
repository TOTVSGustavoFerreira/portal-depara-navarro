/**
 * Portal De-Para TOTVS - Backend API Version 1.3 (Google Apps Script)
 * Adaptado para arquitetura de API REST com suporte a CORS (hospedagem externa no GitHub Pages)
 * Desenvolvido por Antigravity
 */

// Permite requisições de outros domínios como o GitHub Pages
function isCodeIgnored(cod) {
  if (!cod) return true;
  var normalized = String(cod)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") // Remove accents
    .trim()
    .toUpperCase();
  return normalized === "NAO IMPORTAR" || normalized === "P/ ANALISE";
}

function doGet(e) {
  var id = (e && e.parameter && e.parameter.id) ? e.parameter.id : "";
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "";
  
  if (!id) {
    return createJsonResponse({ success: false, error: "ID da planilha nao fornecido (?id=...)" });
  }
  
  try {
    if (action === "getPortalData") {
      var data = getPortalData(id);
      return createJsonResponse(data);
    } 
    else if (action === "getConnectionInfo") {
      var conn = getConnectionInfo(id);
      return createJsonResponse(conn);
    }
    else if (action === "importSingleSheet") {
      var sheetName = e.parameter.sheetName;
      var res = importSingleSheet(id, sheetName);
      return createJsonResponse(res);
    }
    else if (action === "syncSingleSheet") {
      var sheetName = e.parameter.sheetName;
      var res = syncSingleSheet(id, sheetName);
      return createJsonResponse(res);
    }
    else if (action === "autoInstallStructure") {
      var res = autoInstallStructure(id);
      return createJsonResponse(res);
    }
    
    return createJsonResponse({ success: false, error: "Acao GET desconhecida ou nao informada." });
  } catch (err) {
    return createJsonResponse({ success: false, error: "Erro de execucao (GET): " + err.message });
  }
}

function doPost(e) {
  var params;
  try {
    params = JSON.parse(e.postData.contents);
  } catch (err) {
    return createJsonResponse({ success: false, error: "JSON invalido no corpo do POST." });
  }
  
  var id = params.id;
  var action = params.action;
  
  if (!id) {
    return createJsonResponse({ success: false, error: "ID da planilha nao fornecido no payload." });
  }
  
  try {
    if (action === "saveEventMapping") {
      var res = saveEventMapping(id, params.rowNum, params.data, params.applyToAllMatches);
      return createJsonResponse(res);
    }
    else if (action === "createNewRMEvent") {
      var res = createNewRMEvent(id, params.eventData);
      return createJsonResponse(res);
    }
    else if (action === "deleteUnusedCreatedEvent") {
      var res = deleteUnusedCreatedEvent(id, params.code);
      return createJsonResponse(res);
    }
    
    return createJsonResponse({ success: false, error: "Acao POST desconhecida." });
  } catch (err) {
    return createJsonResponse({ success: false, error: "Erro de execucao (POST): " + err.message });
  }
}

function createJsonResponse(data) {
  var output = ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ==========================================
// FUNÇÕES DE MANIPULAÇÃO DO GOOGLE SHEETS
// ==========================================

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

function getSheetDynamic(ss, name) {
  if (!ss) return null;
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  
  if (name.slice(-1).toUpperCase() !== "S") {
    sheet = ss.getSheetByName(name + "S") || ss.getSheetByName(name + "s");
    if (sheet) return sheet;
  }
  
  if (name.slice(-1).toUpperCase() === "S") {
    sheet = ss.getSheetByName(name.slice(0, -1));
    if (sheet) return sheet;
  }
  
  return null;
}

var SHEET_ZDEPARA = "ZDEPARA_EVENTOS";
var SHEET_DADOS_RM = "DADOS_RM_EVENTOS";
var SHEET_CONFIG = "CONFIG_CONEXAO";
var SHEET_LOG = "LOG_SINCRONIZACAO";

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

var EXPECTED_COLUMNS_DEPARA = [
  "EMPRESA_DE",
  "CODIGO_DE",
  "NOME_DE",
  "TIPO_EVENTO",
  "COLIGADA_PARA",
  "CODIGO_PARA",
  "NOME_RM"
];

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

function extractSpreadsheetId(urlOrId) {
  if (!urlOrId) return null;
  if (urlOrId.indexOf("docs.google.com/spreadsheets") !== -1) {
    var matches = urlOrId.match(/\/d\/([a-zA-Z0-9-_]+)/);
    return matches ? matches[1] : null;
  }
  return urlOrId;
}

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

function autoInstallStructure(spreadsheetId) {
  var ss = getActiveSS(spreadsheetId);
  if (!ss) {
    return { success: false, error: "Planilha não encontrada." };
  }
  
  try {
    var sheetConfig = ss.getSheetByName(SHEET_CONFIG);
    if (!sheetConfig) {
      sheetConfig = ss.insertSheet(SHEET_CONFIG);
      sheetConfig.getRange("A1").setValue("Link Planilha Destino (Empresa):").setFontWeight("bold");
      sheetConfig.getRange("B1").setValue("https://docs.google.com/spreadsheets/d/1zEMXK--jTyXQKxFHpDbKaBtudXdi5urcAu7juaWJeuM/edit?usp=sharing");
      sheetConfig.getRange("A2").setValue("Abas Extras Personalizadas a Sincronizar:").setFontWeight("bold");
      sheetConfig.setColumnWidth(1, 280);
      sheetConfig.setColumnWidth(2, 450);
    } else {
      var currentLink = sheetConfig.getRange("B1").getValue().toString().trim();
      if (!currentLink) {
        sheetConfig.getRange("B1").setValue("https://docs.google.com/spreadsheets/d/1zEMXK--jTyXQKxFHpDbKaBtudXdi5urcAu7juaWJeuM/edit?usp=sharing");
      }
    }
    
    DEFAULT_SHEETS_TO_SYNC.forEach(function(sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        
        if (sheetName === "ZDEPARA_EVENTOS") {
          var deparaHeaders = [
            "EMPRESA_DE", "CODIGO_DE", "NOME_DE", "TIPO_EVENTO", 
            "COLIGADA_PARA", "CODIGO_PARA", "NOME_RM", 
            "CODIGO_PARA_FICHA_MES1", "CODIGO_PARA_FICHA_MES2", "CODIGO_PARA_VERBAS_FERIAS",
            "OBSERVACAO"
          ];
          sheet.getRange(1, 1, 1, deparaHeaders.length).setValues([deparaHeaders]).setFontWeight("bold").setBackground("#d9e1f2");
          sheet.getRange("B2:B").setNumberFormat("@");
          sheet.getRange("F2:F").setNumberFormat("@");
          sheet.getRange("H2:J").setNumberFormat("@");
        } else if (sheetName === "DADOS_RM_EVENTOS") {
          var rmHeaders = ["CÓDIGO", "DESCRIÇÃO", "TIPO", "VALHORDIAREF", "NATE_ESOCIAL"];
          sheet.getRange(1, 1, 1, rmHeaders.length).setValues([rmHeaders]).setFontWeight("bold").setBackground("#e2efda");
          sheet.getRange("A2:A").setNumberFormat("@");
        } else {
          sheet.getRange("A1").setValue("CODIGO").setFontWeight("bold");
          sheet.getRange("B1").setValue("DESCRICAO").setFontWeight("bold");
        }
      }
    });
    
    var sheetLog = ss.getSheetByName(SHEET_LOG);
    if (!sheetLog) {
      sheetLog = ss.insertSheet(SHEET_LOG);
      var logHeaders = ["DATA/HORA", "USUÁRIO", "TIPO AÇÃO", "DETALHES DO PROCESSO", "STATUS"];
      sheetLog.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]).setFontWeight("bold").setBackground("#fff2cc");
    }
    
    SpreadsheetApp.flush();
    return { success: true, message: "Estrutura v1.3 instalada com sucesso!" };
  } catch (e) {
    return { success: false, error: "Erro ao instalar: " + e.message };
  }
}

function getReferenceDiagnostics(rmEvents) {
  var diag = { gaps: [], duplicates: [] };
  
  var manualCodes = rmEvents
    .filter(function(e) { return e.descricao.indexOf("[INCLUSAO MANUAL]") !== -1; })
    .map(function(e) { return parseInt(e.codigo, 10); })
    .filter(function(code) { return !isNaN(code); })
    .sort(function(a, b) { return a - b; });
  
  if (manualCodes.length > 1) {
    for (var i = 0; i < manualCodes.length - 1; i++) {
      var current = manualCodes[i];
      var next = manualCodes[i + 1];
      if (next - current > 1) {
        for (var g = current + 1; g < next; g++) {
          var gapStr = String(g);
          while (gapStr.length < 4) { gapStr = "0" + gapStr; }
          diag.gaps.push(gapStr);
        }
      }
    }
  }
  
  var occurrences = {};
  rmEvents.forEach(function(e) {
    if (e.descricao.indexOf("[INCLUSAO MANUAL]") !== -1) {
      occurrences[e.codigo] = (occurrences[e.codigo] || 0) + 1;
    }
  });
  
  for (var code in occurrences) {
    if (occurrences[code] > 1) {
      diag.duplicates.push(code);
    }
  }
  
  return diag;
}

function checkAndCreateObservationColumn(sheetDepara) {
  if (!sheetDepara) return;
  var lastCol = sheetDepara.getLastColumn();
  if (lastCol === 0) return;
  var headers = sheetDepara.getRange(1, 1, 1, lastCol).getValues()[0];
  var colIdx = headers.indexOf("OBSERVACAO");
  
  if (colIdx === -1) {
    sheetDepara.getRange(1, lastCol + 1).setValue("OBSERVACAO").setFontWeight("bold");
    SpreadsheetApp.flush();
  }
}

function getPortalData(spreadsheetId) {
  var connection = getConnectionInfo(spreadsheetId);
  if (!connection.isInstalled) {
    return { status: { valid: false, errors: ["A estrutura v1.3 nao esta instalada."] }, data: [], rmEvents: [], stats: {}, connection: connection };
  }
  
  var ss = getActiveSS(spreadsheetId);
  var sheetDepara = ss.getSheetByName(SHEET_ZDEPARA);
  var sheetDadosRM = ss.getSheetByName(SHEET_DADOS_RM);
  
  checkAndCreateObservationColumn(sheetDepara);
  
  var headers = sheetDepara.getRange(1, 1, 1, sheetDepara.getLastColumn()).getValues()[0];
  
  var colIndexes = {};
  headers.forEach(function(h, idx) {
    if (!colIndexes[h]) colIndexes[h] = [];
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
  
  var diagnostics = getReferenceDiagnostics(rmEvents);
  
  var total = deparaData.length;
  var preenchidos = 0;
  var naoPreenchidos = 0;
  var pAnalise = 0;
  
  var keyMap = {};
  var duplicateKeys = {};
  
  deparaData.forEach(function(item) {
    var cod = item.codigoPara;
    if (cod) {
      var isIgnored = isCodeIgnored(cod);
      if (isIgnored) {
        // Let's count them appropriately or treat them as filled but ignored.
        // For stats purposes, they are mapped, but we handle their status
        if (String(cod).toUpperCase().indexOf("ANALISE") !== -1) {
          pAnalise++;
        } else {
          preenchidos++; // NAO IMPORTAR counts as filled
        }
      } else {
        preenchidos++;
      }
    } else {
      naoPreenchidos++;
    }
    
    if (item.nomeDe && item.tipoEvento && cod && !isCodeIgnored(cod)) {
      var key = item.nomeDe.toLowerCase() + "|||" + item.tipoEvento;
      if (!keyMap[key]) keyMap[key] = [];
      if (keyMap[key].indexOf(cod) === -1) keyMap[key].push(cod);
      if (keyMap[key].length > 1) duplicateKeys[key] = true;
    }
  });
  
  // Helper local no backend para normalizar o tipo de evento (tolerância de singular/plural, P/D/B)
  function normalizeEventType(type) {
    if (!type) return "";
    var t = String(type).toUpperCase().replace(/\s/g, "");
    if (t.indexOf("PROV") !== -1 || t === "P") return "PROVENTO";
    if (t.indexOf("DESC") !== -1 || t === "D") return "DESCONTO";
    if (t.indexOf("BASE") !== -1 || t.indexOf("CALC") !== -1 || t === "B") return "BASE";
    return t;
  }

  // Mapa de eventos RM para validação de tipo rápida
  var rmEventsMap = {};
  rmEvents.forEach(function(ev) {
    rmEventsMap[ev.codigo] = ev;
  });

  var divergenciasCount = 0;
  deparaData.forEach(function(item) {
    var cod = item.codigoPara;
    var isIgnored = isCodeIgnored(cod);
    
    var hasDuplicidade = false;
    var hasTipoMismatch = false;

    // 1. Verificar duplicidade de chave (nome legado + tipo mapeado para diferentes RMs)
    if (item.nomeDe && item.tipoEvento && cod && !isIgnored) {
      var key = item.nomeDe.toLowerCase() + "|||" + item.tipoEvento;
      if (duplicateKeys[key]) {
        hasDuplicidade = true;
      }
    }

    // 2. Verificar incompatibilidade de tipo
    if (cod && !isIgnored && rmEventsMap[cod] && item.tipoEvento) {
      var rmEvent = rmEventsMap[cod];
      var normLegacy = normalizeEventType(item.tipoEvento);
      var normRM = normalizeEventType(rmEvent.tipo);
      if (normLegacy && normRM && normLegacy !== normRM) {
        hasTipoMismatch = true;
      }
    }

    if (hasDuplicidade || hasTipoMismatch) {
      item.hasDivergencia = true;
      item.hasDivergenciaDuplicidade = hasDuplicidade;
      item.hasDivergenciaTipo = hasTipoMismatch;
      divergenciasCount++;
    }
  });
  
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

  // Detectar Mapeamentos Órfãos (Código RM no De-Para que não existe em DADOS_RM_EVENTOS)
  var rmCodesSet = {};
  rmEvents.forEach(function(ev) {
    rmCodesSet[ev.codigo] = true;
  });

  var orphans = [];
  deparaData.forEach(function(item) {
    var cod = item.codigoPara;
    if (cod && !isCodeIgnored(cod)) {
      if (!rmCodesSet[cod]) {
        orphans.push({
          rowNum: item.rowNum,
          codigoDe: item.codigoDe,
          nomeDe: item.nomeDe,
          codigoPara: cod,
          nomeRm: item.nomeRm
        });
      }
    }
  });
  diagnostics.orphans = orphans;
  
  var stats = {
    total: total,
    preenchidos: preenchidos,
    naoPreenchidos: naoPreenchidos,
    pAnalise: pAnalise,
    divergencias: divergenciasCount,
    gapsCount: diagnostics.gaps.length,
    duplicatesCount: diagnostics.duplicates.length,
    unusedCreatedCount: unusedCreatedEvents.length,
    orphansCount: orphans.length
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

function lookupEventDescription(code, eventsList) {
  if (!code) return "";
  for (var i = 0; i < eventsList.length; i++) {
    if (eventsList[i].codigo === code) {
      return eventsList[i].descricao;
    }
  }
  return "";
}

function saveEventMapping(spreadsheetId, rowNum, data, applyToAllMatches) {
  var ss = getActiveSS(spreadsheetId);
  var sheetDepara = ss.getSheetByName(SHEET_ZDEPARA);
  var sheetDadosRM = ss.getSheetByName(SHEET_DADOS_RM);
  var headers = sheetDepara.getRange(1, 1, 1, sheetDepara.getLastColumn()).getValues()[0];
  
  var colIndexes = {};
  headers.forEach(function(h, idx) {
    if (!colIndexes[h]) colIndexes[h] = [];
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

function createNewRMEvent(spreadsheetId, eventData) {
  var ss = getActiveSS(spreadsheetId);
  var sheet = ss.getSheetByName(SHEET_DADOS_RM);
  if (!sheet) return { success: false, error: "Aba DADOS_RM_EVENTOS nao encontrada." };
  
  try {
    var lastRow = sheet.getLastRow();
    var lastCode = 0;
    
    if (lastRow > 1) {
      var codes = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      var numericCodes = codes.map(function(row) {
        return parseInt(row[0], 10);
      }).filter(function(code) {
        return !isNaN(code);
      });
      
      if (numericCodes.length > 0) {
        lastCode = Math.max.apply(null, numericCodes);
      }
    }
    
    var nextCodeVal = lastCode + 1;
    var nextCodeStr = String(nextCodeVal);
    while (nextCodeStr.length < 4) {
      nextCodeStr = "0" + nextCodeStr;
    }
    
    var descCompleta = "[INCLUSAO MANUAL] " + eventData.nomeDe.toUpperCase();
    var tipoFinal = eventData.tipoEvento || "PROVENTO";
    
    sheet.appendRow(["", "", "", "", ""]);
    var targetRow = sheet.getLastRow();
    
    sheet.getRange(targetRow, 1).setNumberFormat("@").setValue(nextCodeStr);
    sheet.getRange(targetRow, 2).setValue(descCompleta);
    sheet.getRange(targetRow, 3).setValue(tipoFinal);
    sheet.getRange(targetRow, 4).setValue("");
    sheet.getRange(targetRow, 5).setValue("");
    
    SpreadsheetApp.flush();
    return { success: true, code: nextCodeStr, description: descCompleta, type: tipoFinal };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function deleteUnusedCreatedEvent(spreadsheetId, code) {
  var ss = getActiveSS(spreadsheetId);
  var sheet = ss.getSheetByName(SHEET_DADOS_RM);
  if (!sheet) return { success: false, error: "Aba DADOS_RM_EVENTOS nao encontrada." };
  
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < values.length; i++) {
        if (String(values[i][0]) === String(code)) {
          sheet.deleteRow(i + 2);
          SpreadsheetApp.flush();
          return { success: true, message: "Evento manual código " + code + " excluido com sucesso." };
        }
      }
    }
    return { success: false, error: "Evento nao encontrado." };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

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
      if (destId.indexOf("/") === -1 && destId.length < 25) {
        return { 
          success: false, 
          error: "O link configurado na aba 'CONFIG_CONEXAO' (célula B1) é inválido. Certifique-se de preencher a célula B1 com a URL completa da planilha oficial corporativa da TOTVS." 
        };
      }
      return { 
        success: false, 
        error: "Não foi possível abrir a Planilha Destino (ID: " + destId + "). Verifique a URL e as permissões de acesso." 
      };
    }
    
    var companySheet = getSheetDynamic(companySS, sheetName);
    if (!companySheet) {
      writeLog(spreadsheetId, "IMPORTAÇÃO PARCIAL", "Aba " + sheetName + " não encontrada na Planilha Destino (pulada).", "SUCESSO");
      return { success: true, message: sheetName + " pulado (não existe no destino)." };
    }
    
    var realSheetName = companySheet.getName();
    var personalSheet = getSheetDynamic(personalSS, realSheetName);
    
    if (!personalSheet) {
      personalSheet = personalSS.insertSheet(realSheetName);
    }
    
    var dataRange = companySheet.getDataRange();
    var values = dataRange.getValues();
    var backgrounds = dataRange.getBackgrounds();
    var numberFormats = dataRange.getNumberFormats();
    
    if (values.length === 0) {
      return { success: true, message: realSheetName + " está vazia no destino." };
    }
    
    var headers = values[0];
    
    // Identifica e formata as colunas em memória
    var colsToFormat = [];
    var normHeaders = headers.map(function(h) { 
      return String(h).normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toUpperCase(); 
    });
    
    if (realSheetName === "ZDEPARA_EVENTOS") {
      ["CODIGO_DE", "CODIGO_PARA", "CODIGO_PARA_FICHA_MES1", "CODIGO_PARA_FICHA_MES2", "CODIGO_PARA_VERBAS_FERIAS"].forEach(function(colName) {
        var idx = normHeaders.indexOf(colName);
        if (idx !== -1) colsToFormat.push(idx);
      });
    } else if (realSheetName === "DADOS_RM_EVENTOS") {
      var idx = normHeaders.indexOf("CODIGO");
      if (idx === -1) idx = normHeaders.indexOf("COD");
      if (idx !== -1) colsToFormat.push(idx);
    }
    
    // Processamento de formatação 100% em memória
    for (var r = 1; r < values.length; r++) {
      colsToFormat.forEach(function(colIdx) {
        var val = values[r][colIdx];
        if (val === "" || val === null || val === undefined) {
          values[r][colIdx] = "";
        } else if (typeof val === "number") {
          var strVal = String(val);
          while (strVal.length < 4) strVal = "0" + strVal;
          values[r][colIdx] = strVal;
        } else {
          values[r][colIdx] = String(val).trim();
        }
        numberFormats[r][colIdx] = "@";
      });
    }
    
    personalSheet.clearContents();
    personalSheet.clearFormats();
    
    personalSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    personalSheet.getRange(1, 1, backgrounds.length, backgrounds[0].length).setBackgrounds(backgrounds);
    personalSheet.getRange(1, 1, numberFormats.length, numberFormats[0].length).setNumberFormats(numberFormats);
    
    writeLog(spreadsheetId, "IMPORTAÇÃO PARCIAL", "Importada aba: " + realSheetName + " (" + values.length + " linhas)", "SUCESSO");
    return { success: true, message: realSheetName + " importado com sucesso." };
  } catch (e) {
    writeLog(spreadsheetId, "IMPORTAÇÃO PARCIAL", "Erro ao importar aba " + sheetName + ": " + e.message, "FALHA");
    return { success: false, error: e.message };
  }
}

function normalizeTipoEvento(tipo) {
  var t = String(tipo).trim().toUpperCase();
  if (t.indexOf("PROV") !== -1 || t === "P" || t.startsWith("P-")) return "P";
  if (t.indexOf("DESC") !== -1 || t === "D" || t.startsWith("D-")) return "D";
  if (t.indexOf("BASE") !== -1 || t === "B" || t.startsWith("B-")) return "B";
  return t;
}

function getRowKey(headers, rowValues) {
  var normHeaders = headers.map(function(h) { 
    return String(h)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .toUpperCase(); 
  });
  
  var nomeDeIdx = normHeaders.indexOf("NOME_DE");
  var tipoDeIdx = normHeaders.indexOf("TIPO_EVENTO");
  var codDeIdx = normHeaders.indexOf("CODIGO_DE");
  
  if (nomeDeIdx !== -1 && tipoDeIdx !== -1) {
    var nomeVal = String(rowValues[nomeDeIdx]).trim().toUpperCase();
    var tipoVal = normalizeTipoEvento(rowValues[tipoDeIdx]);
    if (nomeVal) {
      return nomeVal + "||" + tipoVal;
    }
  }
  
  if (nomeDeIdx !== -1) {
    var nomeVal = String(rowValues[nomeDeIdx]).trim().toUpperCase();
    if (nomeVal) return nomeVal;
  }
  
  if (codDeIdx !== -1) {
    var codVal = String(rowValues[codDeIdx]).trim().toUpperCase();
    if (codVal) return codVal;
  }
  
  return String(rowValues[0]).trim().toUpperCase();
}

function headersMatch(h1, h2) {
  var norm = function(str) {
    return String(str)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .toUpperCase()
      .replace(/[^A-Z0-9]/g, "");
  };
  
  var norm1 = norm(h1);
  var norm2 = norm(h2);
  
  if (norm1 === norm2) return true;
  
  // Trata chaves comuns com erros de digitação / abreviações comuns
  if ((norm1 === "CODIGODE" || norm1 === "CODIGOD") && (norm2 === "CODIGODE" || norm2 === "CODIGOD")) return true;
  if ((norm1 === "CODIGOPARA" || norm1 === "CODIGOPAR") && (norm2 === "CODIGOPARA" || norm2 === "CODIGOPAR")) return true;
  
  return false;
}

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
          error: "O link configurado na aba 'CONFIG_CONEXAO' (célula B1) é inválido. Certifique-se de preencher a célula B1 com a URL completa da planilha oficial corporativa da TOTVS." 
        };
      }
      return { 
        success: false, 
        error: "Não foi possível abrir a Planilha Destino (ID: " + destId + "). Verifique a URL e as permissões de acesso." 
      };
    }
    
    var personalSheet = getSheetDynamic(personalSS, sheetName);
    if (!personalSheet) {
      return { success: true, message: sheetName + " pulado (não existe na origem)." };
    }
    
    var realSheetName = personalSheet.getName();
    
    // Obter todos os dados, cores e formatos da origem em memória
    var personalRange = personalSheet.getDataRange();
    var values = personalRange.getValues();
    var backgrounds = personalRange.getBackgrounds();
    var numberFormats = personalRange.getNumberFormats();
    
    if (values.length === 0) {
      return { success: true, message: realSheetName + " está vazia na origem." };
    }
    
    // 1. Criar uma nova aba temporária no destino
    var tempSheetName = realSheetName + "_TEMP_SYNC";
    var tempSheet = companySS.getSheetByName(tempSheetName);
    if (tempSheet) {
      companySS.deleteSheet(tempSheet);
    }
    tempSheet = companySS.insertSheet(tempSheetName);
    
    // 2. Gravar os dados completos na aba temporária (se falhar aqui por timeout, a original continua intacta)
    tempSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    tempSheet.getRange(1, 1, backgrounds.length, backgrounds[0].length).setBackgrounds(backgrounds);
    tempSheet.getRange(1, 1, numberFormats.length, numberFormats[0].length).setNumberFormats(numberFormats);
    
    // 3. Deletar a aba antiga oficial
    var oldCompanySheet = getSheetDynamic(companySS, realSheetName);
    if (oldCompanySheet) {
      companySS.deleteSheet(oldCompanySheet);
    }
    
    // 4. Renomear a temporária para a oficial
    tempSheet.setName(realSheetName);
    
    // Forçar recálculo e gravação imediata
    SpreadsheetApp.flush();
    
    writeLog(spreadsheetId, "SINCRONIZAÇÃO COMPLETA", "Aba: " + realSheetName + " copiada integralmente de forma segura (Transacional)", "SUCESSO");
    return { 
      success: true, 
      message: realSheetName + " sincronizado com segurança (Cópia integral transacional).",
      details: { mode: "direta", rows: values.length - 1 }
    };
  } catch (e) {
    writeLog(spreadsheetId, "EXPORTAÇÃO PARCIAL", "Erro ao sincronizar aba " + sheetName + ": " + e.message, "FALHA");
    return { success: false, error: e.message };
  }
}

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
