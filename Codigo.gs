// IDs e constantes necessários para acessar a planilha e o documento modelo
const SHEETID = 'SHEETID'; // ID da planilha
const DADOABANOME = 'DADOABANOME'; // Nome da aba de dados
const CHAVE = "B:B"; // Intervalo da coluna B, onde os IDs estão armazenados
const CABECALHO = 1; // O cabeçalho começa na linha 2 (índice 1)
const ABAHISTORICO = 'ABAHISTORICO'; // aba de historico


// IDs dos templates e destinos
const TEMPLATE1_ID = 'TEMPLATE1_ID';
const DESTINO1_ID = 'DESTINO1_ID';
const TEMPLATE1_NOME = 'Documento Template 1 - Linha';


const TEMPLATE2_ID = 'TEMPLATE2_ID';
const DESTINO2_ID = 'DESTINO2_ID';
const TEMPLATE2_NOME = 'Documento Template 2 - Linha';


// Função para encontrar o índice da linha pelo ID na coluna B
function findRowIndexById(id) {
  const sheet = SpreadsheetApp.openById(SHEETID).getSheetByName(DADOABANOME);
  const data = sheet.getRange(CHAVE).getValues();


  // Loop para percorrer cada linha da coluna B
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && String(data[i][0]).trim() === String(id).trim()) {
      return i + 1; // Retorna o número da linha (começando em 1)
    }
  }


  return -1; // Retorna -1 se o ID não for encontrado
}


// Função genérica para gerar um documento com base no índice da linha e IDs do template e destino
function gerarDoc(rowIndex, templateId, destinationId, nomeDocumento) {
  try {
    const sheet = SpreadsheetApp.openById(SHEETID).getSheetByName(DADOABANOME);
    const dataRange = sheet.getDataRange();
    const lastRow = dataRange.getLastRow();
    const headers = dataRange.offset(1, 0, 1).getValues()[0]; // Cabeçalho está na segunda linha (índice 1)


    if (rowIndex < 1 || rowIndex > lastRow) {
      throw new Error('Índice de linha fora do intervalo');
    }


    // Abre o documento modelo e cria uma cópia na pasta destino
    const folder = DriveApp.getFolderById(destinationId);
    const templateFile = DriveApp.getFileById(templateId);
    const temp = templateFile.makeCopy(`${nomeDocumento} ${rowIndex}`, folder);
    const doc = DocumentApp.openById(temp.getId());
    const body = doc.getBody();


    // Substitui os placeholders no documento
    headers.forEach((header, i) => {
      const placeholder = `{${header}}`; // Placeholder no formato {Placeholder}
      const value = sheet.getRange(rowIndex, i + 1).getValue();
      body.replaceText(placeholder, value.toString()); // Substitui o placeholder pelo valor correspondente
    });


    // Salva e fecha o documento
    doc.saveAndClose();
    const documentUrl = temp.getUrl(); // Obtém o URL do documento gerado
    Logger.log(`Documento gerado: ${documentUrl}`);


    // Salva o nome, data e link do documento na aba "Histórico" da planilha
    const historicoSheet = sheet.getParent().getSheetByName(ABAHISTORICO); // Supondo que o nome da aba seja "Histórico"
    if (historicoSheet) {
      const nome = sheet.getRange(rowIndex, 2).getValue(); // Substitua pelo índice da coluna do nome
      const dataAtual = new Date().toLocaleDateString('pt-BR'); // Data atual formatada
      historicoSheet.appendRow([nome, dataAtual, documentUrl]); // Adiciona uma nova linha com nome, data e URL do documento
      Logger.log('Dados salvos no histórico.');
    } else {
      Logger.log('Aba "Histórico" não encontrada.');
    }


    // Exibe um alerta de sucesso para o usuário
    SpreadsheetApp.getUi().alert('Documento gerado com sucesso!');
  } catch (e) {
    Logger.log('Erro ao acessar o documento: ' + e.message);
    // Exibe um alerta de erro para o usuário
    SpreadsheetApp.getUi().alert(`Erro ao gerar o documento: ${e.message}`);
  }
}


// Função para exibir um diálogo e chamar a função de busca e geração de documento
function showDialog(templateId, destinationId, nomeDocumento) {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Informe o ID que deseja buscar:', ui.ButtonSet.OK_CANCEL);


  if (result.getSelectedButton() === ui.Button.OK) {
    const id = result.getResponseText();
    const rowIndex = findRowIndexById(id);


    if (rowIndex !== -1) {
      gerarDoc(rowIndex, templateId, destinationId, nomeDocumento); // Chama a função para gerar o documento com o índice da linha encontrada
    } else {
      ui.alert(`ID "${id}" não encontrado.`);
    }
  }
}


// Função para adicionar o menu personalizado "Menu Avançado" e vincular à função showSidebar
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Menu Avançado')
    .addItem('Abrir Menu', 'showSidebar')
    .addToUi();
}


// Função para exibir a barra lateral com a navegação personalizada
function showSidebar() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Menu Avançado')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}


// Função para exibir um diálogo específico para um template
function showDialogTemplate1() {
  showDialog(TEMPLATE1_ID, DESTINO1_ID, TEMPLATE1_NOME);
}


// Função para exibir um diálogo específico para outro template
function showDialogTemplate2() {
  showDialog(TEMPLATE2_ID, DESTINO2_ID, TEMPLATE2_NOME);
}
