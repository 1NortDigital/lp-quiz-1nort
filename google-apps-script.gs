// 1Nort - LP-Quiz para Google Sheets
// Recebe o POST da LP e grava cada lead numa linha (upsert por event_id).
//
// PUBLICAR: script.google.com (conta dona da planilha) > Novo projeto >
//   apaga tudo, cola este arquivo, Ctrl+S > Implantar > App da Web
//   (Executar como: Eu / Acesso: Qualquer pessoa) > autoriza > copia a URL /exec.

var SHEET_ID = '1NABMLnzCLeAOM6TOrmNTpmOHl96cJrogWcSLxkrqU24';
var ABA = 'Leads';

var COLUNAS = [
  'data_hora','tier','qualificado','top_tier','score',
  'nome','telefone','email','cidade',
  'papel','projetos','desafio','foco_vendedor','nao_atuo','trafego','instagram',
  'pagina','utm_source','utm_medium','utm_campaign','utm_content',
  'fbc','fbp','origem','event_id','parcial','respostas_json'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    var r = d.respostas || {};

    var sheet = pegarAba_();
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(COLUNAS);
      sheet.getRange(1, 1, 1, COLUNAS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    var linha = {
      data_hora:    new Date(),
      tier:         d.tier || '',
      qualificado:  d.qualificado === true ? 'SIM' : 'NAO',
      top_tier:     d.top_tier === true ? 'SIM' : '',
      score:        d.score != null ? d.score : '',
      nome:         d.nome || '',
      telefone:     d.telefone || '',
      email:        d.email || '',
      cidade:       d.cidade || '',
      papel:        d.papel || '',
      projetos:     r.projetos || '',
      desafio:      r.desafio || '',
      foco_vendedor:r.foco_vendedor || '',
      nao_atuo:     r.nao_atuo || '',
      trafego:      r.trafego || '',
      instagram:    r.instagram || '',
      pagina:       d.pagina || '',
      utm_source:   d.utm_source || '',
      utm_medium:   d.utm_medium || '',
      utm_campaign: d.utm_campaign || '',
      utm_content:  d.utm_content || '',
      fbc:          d.fbc || '',
      fbp:          d.fbp || '',
      origem:       d.origem || '',
      event_id:     d.event_id || '',
      parcial:      d.parcial === true ? 'SIM' : '',
      respostas_json: JSON.stringify(r)
    };

    var valores = COLUNAS.map(function(c){ return linha[c]; });

    // UPSERT por event_id: atualiza a linha do envio parcial em vez de duplicar.
    var idCol = COLUNAS.indexOf('event_id') + 1;
    var lastRow = sheet.getLastRow();
    var alvo = 0;
    if (d.event_id && lastRow > 1) {
      var ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (ids[i][0] === d.event_id) { alvo = i + 2; break; }
      }
    }
    if (alvo) sheet.getRange(alvo, 1, 1, valores.length).setValues([valores]);
    else      sheet.appendRow(valores);

    return resposta_({ ok: true, atualizado: alvo > 0 });
  } catch (err) {
    return resposta_({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return resposta_({ ok: true, msg: '1Nort LP-Quiz endpoint no ar' });
}

function pegarAba_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(ABA);
  if (!sheet) sheet = ss.insertSheet(ABA);
  return sheet;
}

function resposta_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
