/**
 * Portal De-Para - Proof of Concept (PoC) Backend API
 * Desenvolvido para teste de conexão com Planilha Corporativa TOTVS via GitHub Pages
 */

function doGet(e) {
  var spreadsheetId = (e && e.parameter && e.parameter.id) ? e.parameter.id : "";
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "read";
  
  if (!spreadsheetId) {
    return createJsonResponse({ success: false, error: "ID da planilha nao fornecido na URL (?id=...)" });
  }

  try {
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheetByName("TESTE_CONEXAO");
    
    // Se a aba de teste não existir, cria automaticamente para o usuário
    if (!sheet) {
      sheet = ss.insertSheet("TESTE_CONEXAO");
      sheet.getRange("A1").setValue("CHAVE").setFontWeight("bold");
      sheet.getRange("B1").setValue("VALOR").setFontWeight("bold");
      sheet.getRange("A2").setValue("Mensagem");
      sheet.getRange("B2").setValue("Conexao inicial bem-sucedida!");
    }
    
    if (action === "read") {
      var chave = sheet.getRange("A2").getValue().toString();
      var valor = sheet.getRange("B2").getValue().toString();
      return createJsonResponse({ 
        success: true, 
        spreadsheetName: ss.getName(),
        chave: chave, 
        valor: valor 
      });
    }
    
    return createJsonResponse({ success: false, error: "Acao GET desconhecida." });
  } catch (err) {
    return createJsonResponse({ 
      success: false, 
      error: "Erro de Acesso: " + err.message + ". Certifique-se de que a planilha possui permissao de Editor para a conta que executou este script." 
    });
  }
}

function doPost(e) {
  var params;
  try {
    params = JSON.parse(e.postData.contents);
  } catch (err) {
    return createJsonResponse({ success: false, error: "JSON invalido no corpo da requisicao." });
  }
  
  var spreadsheetId = params.id;
  var action = params.action;
  
  if (!spreadsheetId) {
    return createJsonResponse({ success: false, error: "ID da planilha nao fornecido." });
  }
  
  try {
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheetByName("TESTE_CONEXAO");
    
    if (!sheet) {
      return createJsonResponse({ success: false, error: "Aba TESTE_CONEXAO nao encontrada." });
    }
    
    if (action === "write") {
      var novoValor = params.valor || "";
      sheet.getRange("B2").setValue(novoValor);
      SpreadsheetApp.flush();
      
      return createJsonResponse({ 
        success: true, 
        message: "Valor atualizado na TOTVS com sucesso!",
        valorSalvo: novoValor
      });
    }
    
    return createJsonResponse({ success: false, error: "Acao POST desconhecida." });
  } catch (err) {
    return createJsonResponse({ success: false, error: "Erro ao gravar dados: " + err.message });
  }
}

/**
 * Cria a saida formatada em JSON com tratamento de CORS para navegadores
 */
function createJsonResponse(data) {
  var output = ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}
