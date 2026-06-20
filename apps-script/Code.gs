// ============================================================
// Painel de Controle da Líder — Backend Google Apps Script
// Salva relatórios na planilha e cria Google Docs na pasta
// ============================================================

// ── CONFIGURE AQUI ──────────────────────────────────────────
// 1. ID da planilha (está na URL do Google Sheets após /d/)
var SHEET_ID = '1UFb2ZGLAWBwN0XUKyXakoKZPYiouTWlVTSyWwhWVuNg';

// 2. ID da pasta do Drive (está na URL da pasta após /folders/)
var FOLDER_ID = '1t7Dw49g74DmKBabqaB050FHeUchC7Hg6';

// 3. Nome da aba dentro da planilha
var SHEET_NAME = 'Relatórios';
// ─────────────────────────────────────────────────────────────

var HEADERS = [
  'ID', 'Data', 'Hora', 'Gestora', 'Operação', 'Período',
  'Realizado (R$)', 'Meta (%)', 'Destaques', 'Riscos',
  'Foco Próx. Semana', 'Status Operacional', 'Link Google Doc'
];

function doPost(e) {
  try {
    // Suporta tanto application/x-www-form-urlencoded (URLSearchParams) quanto JSON puro
    var payload;
    if (e.parameter && e.parameter.data) {
      payload = JSON.parse(e.parameter.data);
    } else {
      payload = JSON.parse(e.postData.contents);
    }
    var sheet   = getOrCreateSheet_();
    var id      = generateId_(sheet);
    var now     = new Date();

    // Cria o Google Doc na pasta
    var docUrl = createDoc_(id, payload);

    // Adiciona linha na planilha
    sheet.appendRow([
      id,
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm'),
      payload.gestora   || '',
      payload.operacao  || '',
      payload.periodo   || '',
      payload.realizado || '',
      payload.pct       || '',
      payload.destaques || '',
      payload.riscos    || '',
      payload.foco      || '',
      payload.status    || '',
      docUrl
    ]);

    // Formata a linha recém adicionada
    formatLastRow_(sheet);

    return json_({ success: true, id: id, docUrl: docUrl });

  } catch (err) {
    return json_({ success: false, error: err.message });
  }
}

// Chamada GET para testar se o Web App está no ar
function doGet() {
  return json_({ status: 'ok', message: 'Painel da Líder — backend ativo.' });
}

// ── Funções internas ─────────────────────────────────────────

function getOrCreateSheet_() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    // Cabeçalho em negrito e cor de fundo
    var header = sheet.getRange(1, 1, 1, HEADERS.length);
    header.setFontWeight('bold')
          .setBackground('#1B2430')
          .setFontColor('#FFFFFF')
          .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1,  90);   // ID
    sheet.setColumnWidth(2,  90);   // Data
    sheet.setColumnWidth(3,  60);   // Hora
    sheet.setColumnWidth(4, 160);   // Gestora
    sheet.setColumnWidth(5,  80);   // Operação
    sheet.setColumnWidth(6, 160);   // Período
    sheet.setColumnWidth(7, 110);   // Realizado
    sheet.setColumnWidth(8,  70);   // Meta %
    sheet.setColumnWidth(9, 220);   // Destaques
    sheet.setColumnWidth(10, 220);  // Riscos
    sheet.setColumnWidth(11, 220);  // Foco
    sheet.setColumnWidth(12, 180);  // Status
    sheet.setColumnWidth(13, 100);  // Link
  }
  return sheet;
}

function generateId_(sheet) {
  var year = new Date().getFullYear();
  // Conta linhas de dados (exclui cabeçalho)
  var rows = Math.max(0, sheet.getLastRow() - 1);
  var seq  = String(rows + 1).padStart(4, '0');
  return 'REL-' + year + '-' + seq;
}

function createDoc_(id, payload) {
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var title  = id + ' — ' + (payload.operacao || 'Operação') + ' — ' + (payload.periodo || 'Sem período');
  var doc    = DocumentApp.create(title);

  // Move para a pasta correta
  DriveApp.getFileById(doc.getId()).moveTo(folder);

  var body = doc.getBody();

  // Cabeçalho
  var h = body.appendParagraph('RELATÓRIO EXECUTIVO — OPERAÇÃO ' + (payload.operacao || ''));
  h.setHeading(DocumentApp.ParagraphHeading.HEADING1);

  body.appendParagraph('ID: ' + id + '   |   ' + new Date().toLocaleDateString('pt-BR')).setItalic(true);
  body.appendParagraph('Para: ' + (payload.gestora || '') + '   |   Período: ' + (payload.periodo || ''));
  body.appendHorizontalRule();

  // Conteúdo completo do relatório
  body.appendParagraph(payload.conteudo || '');

  doc.saveAndClose();
  return doc.getUrl();
}

function formatLastRow_(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return;
  var row = sheet.getRange(last, 1, 1, HEADERS.length);
  // Linhas alternadas em cinza claro
  var bg = (last % 2 === 0) ? '#F5F5F5' : '#FFFFFF';
  row.setBackground(bg).setVerticalAlignment('middle');
  // Coluna de link como hiperlink clicável
  var linkCell = sheet.getRange(last, HEADERS.length);
  linkCell.setFontColor('#1155CC').setFontLine('underline');
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
