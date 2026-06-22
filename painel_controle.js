/**
 * Painel de Controle - Gerenciador de Projetos De-Para (Google Apps Script)
 * 
 * Este script deve ser colado no Apps Script da sua planilha pessoal "Painel de Controle".
 * Ele automatiza a criação de novos projetos clonando o seu modelo com os scripts do portal
 * e importando os dados crús da planilha da TOTVS de forma isolada e segura.
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Gerenciador De-Para')
    .addItem('📁 Inicializar Novo Projeto com Portal', 'criarNovoProjetoComDados')
    .addToUi();
}

function criarNovoProjetoComDados() {
  var ui = SpreadsheetApp.getUi();
  
  // 1. Mensagem de Boas-Vindas e Explicação do Passo 1
  ui.alert('Passo 1 de 4: Seleção do Modelo', 
           'Olá! Vamos inicializar a estrutura de um novo projeto.\n\n' +
           'Primeiro, você precisará fornecer o link da sua Planilha MODELO pessoal (aquela que contém o portal visual v1.3 e os scripts).\n\n' +
           'Clique em OK para inserir o link.', 
           ui.ButtonSet.OK);
           
  var respostaModelo = ui.prompt('Link da Planilha Modelo', 'Insira o Link ou ID da sua Planilha MODELO:', ui.ButtonSet.OK_CANCEL);
  if (respostaModelo.getSelectedButton() !== ui.Button.OK) return;
  var linkModelo = respostaModelo.getResponseText();
  
  // 2. Explicação do Passo 2
  ui.alert('Passo 2 de 4: Seleção do Arquivo de Origem (TOTVS)', 
           'Agora, precisamos da Planilha Original gerada pela equipe que está na pasta da TOTVS.\n\n' +
           'Ela é a planilha crua que contém os dados originais/legados do cliente que serão migrados.\n\n' +
           'Clique em OK para inserir o link.', 
           ui.ButtonSet.OK);
           
  var respostaTotvs = ui.prompt('Link da Planilha da TOTVS (Crua)', 'Insira o Link ou ID da Planilha que está na pasta da TOTVS:', ui.ButtonSet.OK_CANCEL);
  if (respostaTotvs.getSelectedButton() !== ui.Button.OK) return;
  var linkTotvs = respostaTotvs.getResponseText();
  
  // 3. Explicação do Passo 3
  ui.alert('Passo 3 de 3: Pasta de Destino no Google Drive', 
           'Por fim, precisamos escolher onde a planilha com o portal visual do cliente será salva.\n\n' +
           'Como obter o ID da Pasta:\n' +
           '1. Abra a pasta do cliente no Google Drive.\n' +
           '2. Olhe a barra de endereços do seu navegador.\n' +
           '3. Copie o código longo de letras e números que fica após "/folders/".\n\n' +
           'Clique em OK para colar o ID.', 
           ui.ButtonSet.OK);
           
  var respostaPasta = ui.prompt('ID da Pasta de Destino', 'Insira o ID da Pasta do Google Drive de destino:', ui.ButtonSet.OK_CANCEL);
  if (respostaPasta.getSelectedButton() !== ui.Button.OK) return;
  var idPastaDestino = respostaPasta.getResponseText();
  
  try {
    var idModelo = extrairIdDoLink(linkModelo);
    var idTotvs = extrairIdDoLink(linkTotvs);
    
    // Obter arquivos e pastas do Google Drive
    var ssPainel = SpreadsheetApp.getActiveSpreadsheet();
    ssPainel.toast('Iniciando clonagem do projeto...', 'Gerenciador De-Para ⚙️', 5);
    
    var idDestino = extrairIdDaPasta(idPastaDestino);
    var arquivoModelo = DriveApp.getFileById(idModelo);
    var pastaDestino = DriveApp.getFolderById(idDestino);
    
    // Passo A: Copiar a Planilha Modelo inteira (copiando junto cores, formatação e scripts do portal)
    ssPainel.toast('Copiando planilha modelo para a pasta de destino...', 'Gerenciador De-Para ⚙️', 5);
    var arquivoTotvsFile = DriveApp.getFileById(idTotvs);
    var nomeNovoArquivo = arquivoTotvsFile.getName();
    var novoArquivo = arquivoModelo.makeCopy(nomeNovoArquivo, pastaDestino);
    var idNovaPlanilha = novoArquivo.getId();
    
    // Passo B: Abrir a planilha original da TOTVS e a Nova Planilha clonada
    ssPainel.toast('Abrindo planilhas para cópia dos dados...', 'Gerenciador De-Para ⚙️', 5);
    var ssTotvs = SpreadsheetApp.openById(idTotvs);
    var ssNova = SpreadsheetApp.openById(idNovaPlanilha);
    
    // Passo C: Iterar por todas as abas da TOTVS e, se existirem no modelo, copiar seus dados
    var abasTotvs = ssTotvs.getSheets();
    var abasCopiadas = [];
    var totalAbas = abasTotvs.length;
    
    abasTotvs.forEach(function(abaOrigem, index) {
      var nomeAba = abaOrigem.getName();
      var abaDestino = ssNova.getSheetByName(nomeAba);
      
      if (abaDestino) {
        ssPainel.toast('Copiando aba: ' + nomeAba + ' (' + (index + 1) + ' de ' + totalAbas + ')', 'Gerenciador De-Para ⚙️', 5);
        var ultLinha = abaOrigem.getLastRow();
        var ultColuna = abaOrigem.getLastColumn();
        
        if (ultLinha > 0 && ultColuna > 0) {
          // Limpar dados antigos do modelo na aba correspondente antes de copiar
          abaDestino.clearContents();
          
          var valores = abaOrigem.getRange(1, 1, ultLinha, ultColuna).getValues();
          abaDestino.getRange(1, 1, ultLinha, ultColuna).setValues(valores);
          abasCopiadas.push(nomeAba);
        }
      }
    });
    
    // Passo D: Compartilhar a planilha com a Conta de Serviço usada pelo Streamlit
    var emailContaServico = "depara-automacao@depara-automacao-4444.iam.gserviceaccount.com"; 
    
    try {
      novoArquivo.addEditor(emailContaServico);
      ui.alert('🎉 Sucesso!', 
               'Cópia do Projeto Inicializada com sucesso!\n\n' +
               '• Arquivo Gerado: ' + nomeNovoArquivo + '\n' +
               '• Dados importados da planilha da TOTVS.\n' +
               '• Compartilhado com a Conta de Serviço para o portal Streamlit.\n\n' +
               'Você já pode acessar a planilha ou enviar o link dela ao portal Streamlit.\n\n' +
               'Link do Arquivo:\n' + novoArquivo.getUrl(), 
               ui.ButtonSet.OK);
    } catch(eShare) {
      ui.alert('⚠️ Aviso', 
               'A planilha foi criada e os dados importados, mas ocorreu um problema ao compartilhar com o robô automaticamente:\n' + eShare.message + '\n\n' +
               'Por favor, compartilhe manualmente no Drive com o e-mail: ' + emailContaServico, 
               ui.ButtonSet.OK);
    }
    
  } catch(erro) {
    ui.alert('❌ Erro no Processo', 
             'Não foi possível inicializar o projeto:\n' + erro.toString(), 
             ui.ButtonSet.OK);
  }
}

// Função auxiliar para extrair o ID de links do Google Sheets
function extrairIdDoLink(link) {
  if (link.indexOf("https://") === 0) {
    var partes = link.split("/d/");
    if (partes.length > 1) {
      return partes[1].split("/")[0];
    }
  }
  return link.trim();
}

// Função auxiliar para extrair o ID de links de pastas do Google Drive
function extrairIdDaPasta(link) {
  if (!link) return "";
  if (link.indexOf("https://") === 0) {
    var matches = link.match(/\/folders\/([a-zA-Z0-9-_]+)/);
    if (matches && matches[1]) {
      return matches[1];
    }
  }
  return link.trim();
}
