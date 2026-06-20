// ============================================================
// Painel de Controle da Líder — Backend Google Apps Script
// Salva relatórios na planilha e cria Google Docs timbrados
// ============================================================

// ── CONFIGURE AQUI ──────────────────────────────────────────
var SHEET_ID   = '1UFb2ZGLAWBwN0XUKyXakoKZPYiouTWlVTSyWwhWVuNg';
var FOLDER_ID  = '1t7Dw49g74DmKBabqaB050FHeUchC7Hg6';
var SHEET_NAME = 'Relatórios';
// ─────────────────────────────────────────────────────────────

// ── Identidade visual por escritório ────────────────────────
var BRAND = {
  'NG': {
    heading:         'ESCRITÓRIO NG',
    subheading:      'Relatório Executivo de Gestão',
    headerBg:        '#0055A5',
    headerText:      '#FFFFFF',
    accentColor:     '#0077CC',
    sectionColor:    '#0055A5',
    lightBg:         '#EBF3FC',
    separatorColor:  '#0055A5',
    bodyFont:        'Arial',
    titleFont:       'Arial',
    footerText:      'NG · Relatório Executivo'
  },
  'LAADV AM': {
    heading:         'LUIS ALBERT ADVOGADO',
    subheading:      'Relatório Executivo · Operação Amazonas',
    headerBg:        '#1B3A2D',
    headerText:      '#C9A84C',
    accentColor:     '#C9A84C',
    sectionColor:    '#1B3A2D',
    lightBg:         '#F8F4E8',
    separatorColor:  '#C9A84C',
    bodyFont:        'Georgia',
    titleFont:       'Georgia',
    footerText:      'Luis Albert Advogado · Operação AM'
  },
  'LAADV RJ': {
    heading:         'LUIS ALBERT ADVOGADO',
    subheading:      'Relatório Executivo · Operação Rio de Janeiro',
    headerBg:        '#1B3A2D',
    headerText:      '#C9A84C',
    accentColor:     '#C9A84C',
    sectionColor:    '#1B3A2D',
    lightBg:         '#F8F4E8',
    separatorColor:  '#C9A84C',
    bodyFont:        'Georgia',
    titleFont:       'Georgia',
    footerText:      'Luis Albert Advogado · Operação RJ'
  }
};

var HEADERS = [
  'ID', 'Data', 'Hora', 'Gestora', 'Operação', 'Período',
  'Realizado (R$)', 'Meta (%)', 'Destaques', 'Riscos',
  'Foco Próx. Semana', 'Status Operacional', 'Link Google Doc'
];

// ── Handler principal ────────────────────────────────────────
function doPost(e) {
  try {
    var payload;
    if (e.parameter && e.parameter.data) {
      payload = JSON.parse(e.parameter.data);
    } else {
      payload = JSON.parse(e.postData.contents);
    }

    var sheet  = getOrCreateSheet_();
    var id     = generateId_(sheet);
    var now    = new Date();
    var docUrl = createBrandedDoc_(id, payload, now);

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

    formatLastRow_(sheet);
    return json_({ success: true, id: id, docUrl: docUrl });

  } catch (err) {
    return json_({ success: false, error: err.message });
  }
}

function doGet() {
  return json_({ status: 'ok', message: 'Painel da Líder — backend ativo.' });
}

// ── Criação do documento timbrado ───────────────────────────
function createBrandedDoc_(id, payload, now) {
  var op     = payload.operacao || 'NG';
  var brand  = BRAND[op] || BRAND['NG'];
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var title  = id + ' — ' + op + ' — ' + (payload.periodo || 'Sem período');
  var doc    = DocumentApp.create(title);
  DriveApp.getFileById(doc.getId()).moveTo(folder);

  var body = doc.getBody();
  body.setMarginTop(36);
  body.setMarginBottom(36);
  body.setMarginLeft(54);
  body.setMarginRight(54);

  // ── 1. FAIXA DE CABEÇALHO ──────────────────────────────────
  var hTable = body.appendTable([['']]);
  hTable.setBorderColor(brand.headerBg);
  var hCell = hTable.getCell(0, 0);
  hCell.setBackgroundColor(brand.headerBg);
  hCell.setPaddingTop(22);
  hCell.setPaddingBottom(18);
  hCell.setPaddingLeft(28);
  hCell.setPaddingRight(28);

  var hTitle = hCell.getChild(0).asParagraph();
  hTitle.setText(brand.heading);
  hTitle.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  hTitle.setSpacingAfter(4);
  var hTitleText = hTitle.editAsText();
  hTitleText.setFontFamily(brand.titleFont);
  hTitleText.setFontSize(22);
  hTitleText.setBold(true);
  hTitleText.setForegroundColor(brand.headerText);

  var hSub = hCell.appendParagraph(brand.subheading);
  hSub.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  hSub.setSpacingBefore(0);
  var hSubText = hSub.editAsText();
  hSubText.setFontFamily(brand.titleFont);
  hSubText.setFontSize(10);
  hSubText.setBold(false);
  hSubText.setForegroundColor(
    brand.headerText === '#C9A84C' ? '#A8C5A0' : brand.accentColor
  );

  // ── 2. TABELA DE METADADOS ────────────────────────────────
  body.appendParagraph(' ').setSpacingBefore(0).setSpacingAfter(0);

  var mTable = body.appendTable([
    ['ID DO RELATÓRIO', 'DATA DE EMISSÃO'],
    [id, Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy  HH:mm')],
    ['PARA', 'PERÍODO'],
    [payload.gestora || '', payload.periodo || '']
  ]);
  mTable.setBorderColor('#DDDDDD');

  styleMetaTable_(mTable, brand);

  // ── 3. LINHA DECORATIVA ───────────────────────────────────
  var divider = body.appendParagraph(' ');
  divider.setSpacingBefore(10);
  divider.setSpacingAfter(10);
  var divTable = body.appendTable([['']]);
  divTable.setBorderColor(brand.separatorColor);
  var divCell = divTable.getCell(0, 0);
  divCell.setBackgroundColor(brand.separatorColor);
  divCell.setPaddingTop(2);
  divCell.setPaddingBottom(2);
  divCell.getChild(0).asParagraph().setText(' ');

  body.appendParagraph(' ').setSpacingBefore(0).setSpacingAfter(4);

  // ── 4. CONTEÚDO DO RELATÓRIO ─────────────────────────────
  var lines = (payload.conteudo || '').split('\n');
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    var para = body.appendParagraph(line || ' ');
    var txt  = para.editAsText();
    txt.setFontFamily(brand.bodyFont);
    txt.setFontSize(10.5);
    txt.setForegroundColor('#2C2C2C');

    if (/^\d+\.\s/.test(line) && line === line.toUpperCase()) {
      txt.setBold(true);
      txt.setFontSize(11.5);
      txt.setForegroundColor(brand.sectionColor);
      para.setSpacingBefore(12);
      para.setSpacingAfter(4);
    } else if (/^[─━-]{4,}/.test(line)) {
      txt.setForegroundColor(brand.separatorColor);
      para.setSpacingBefore(8);
      para.setSpacingAfter(8);
    } else if (/^\s{3}•/.test(line)) {
      para.setIndentStart(20);
      txt.setFontSize(10);
    } else if (/^Para:/.test(line)) {
      txt.setBold(true);
      txt.setForegroundColor(brand.sectionColor);
    }
  }

  // ── 5. RODAPÉ ────────────────────────────────────────────
  var footer = doc.addFooter();
  var fDivTable = footer.appendTable([['']]);
  fDivTable.setBorderColor(brand.separatorColor);
  fDivTable.getCell(0,0)
    .setBackgroundColor(brand.separatorColor)
    .setPaddingTop(1)
    .setPaddingBottom(1)
    .getChild(0).asParagraph().setText(' ');

  var fPara = footer.appendParagraph(
    brand.footerText + '   |   ' + id + '   |   ' +
    Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy')
  );
  fPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  var fText = fPara.editAsText();
  fText.setFontFamily('Arial');
  fText.setFontSize(8);
  fText.setForegroundColor('#888888');

  doc.saveAndClose();
  return doc.getUrl();
}

// ── Estiliza a tabela de metadados ───────────────────────────
function styleMetaTable_(table, brand) {
  table.setBorderColor('#DDDDDD');
  for (var r = 0; r < table.getNumRows(); r++) {
    var row = table.getRow(r);
    var isLabel = (r % 2 === 0);
    for (var c = 0; c < row.getNumCells(); c++) {
      var cell = row.getCell(c);
      cell.setPaddingTop(isLabel ? 8 : 4);
      cell.setPaddingBottom(isLabel ? 2 : 10);
      cell.setPaddingLeft(12);
      cell.setPaddingRight(12);
      if (isLabel) {
        cell.setBackgroundColor(brand.lightBg);
      }
      var para = cell.getChild(0).asParagraph();
      var txt  = para.editAsText();
      txt.setFontFamily('Arial');
      if (isLabel) {
        txt.setFontSize(7.5);
        txt.setBold(true);
        txt.setForegroundColor(brand.sectionColor);
      } else {
        txt.setFontSize(11);
        txt.setBold(true);
        txt.setForegroundColor('#1B2430');
      }
    }
  }
}

// ── Planilha ─────────────────────────────────────────────────
function getOrCreateSheet_() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    var header = sheet.getRange(1, 1, 1, HEADERS.length);
    header.setFontWeight('bold').setBackground('#1B2430').setFontColor('#FFFFFF').setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1,90);  sheet.setColumnWidth(2,90);  sheet.setColumnWidth(3,60);
    sheet.setColumnWidth(4,160); sheet.setColumnWidth(5,80);  sheet.setColumnWidth(6,160);
    sheet.setColumnWidth(7,110); sheet.setColumnWidth(8,70);  sheet.setColumnWidth(9,220);
    sheet.setColumnWidth(10,220); sheet.setColumnWidth(11,220); sheet.setColumnWidth(12,180);
    sheet.setColumnWidth(13,100);
  }
  return sheet;
}

function generateId_(sheet) {
  var year = new Date().getFullYear();
  var rows = Math.max(0, sheet.getLastRow() - 1);
  return 'REL-' + year + '-' + String(rows + 1).padStart(4, '0');
}

function formatLastRow_(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return;
  sheet.getRange(last, 1, 1, HEADERS.length)
       .setBackground(last % 2 === 0 ? '#F5F5F5' : '#FFFFFF')
       .setVerticalAlignment('middle');
  sheet.getRange(last, HEADERS.length)
       .setFontColor('#1155CC')
       .setFontLine('underline');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}
