// 1Nort - LP-Quiz para Google Sheets (ROTEADOR multi-segmento)
// Recebe o POST das LPs e grava na aba certa conforme o campo "pagina".
// Mesma planilha, mesmo script, mesma URL /exec pras duas LPs (Solar e Advoga).
//
// PUBLICAR/ATUALIZAR: script.google.com (conta dona da planilha) > cola este arquivo,
//   confere o SHEET_ID, Ctrl+S > Implantar > Gerenciar implantacoes > editar (lapis)
//   > Nova versao > Implantar. A URL /exec continua a mesma.

var SHEET_ID = '1NABMLnzCLeAOM6TOrmNTpmOHl96cJrogWcSLxkrqU24';

// um schema por segmento: aba destino + colunas (chaves) + titulos (cabecalho)
var SCHEMAS = {
  solar: {
    aba: 'Leads',
    colunas: [
      'data_hora','tier','qualificado','top_tier','score',
      'nome','telefone','email','cidade',
      'papel','projetos','desafio','foco_vendedor','nao_atuo','trafego','instagram',
      'pagina','utm_source','utm_medium','utm_campaign','utm_content',
      'fbc','fbp','origem','event_id','parcial','respostas_json'
    ],
    titulos: [
      'Data/Hora','Tier','Qualificado','Top Tier','Score',
      'Nome','Telefone','Email','Cidade',
      'Papel','Projetos','Desafio','Foco Vendedor','Nao Atua','Trafego','Instagram',
      'Pagina','UTM Source','UTM Medium','UTM Campaign','UTM Content',
      'FBC','FBP','Origem','Event ID','Parcial','Respostas (JSON)'
    ]
  },
  advoga: {
    aba: 'Advoga',
    colunas: [
      'data_hora','tier','qualificado','top_tier','score',
      'nome','telefone','email','cidade',
      'advogados','faturamento','investimento','instagram',
      'pagina','utm_source','utm_medium','utm_campaign','utm_content',
      'fbc','fbp','origem','event_id','parcial','respostas_json'
    ],
    titulos: [
      'Data/Hora','Tier','Qualificado','Top Tier','Score',
      'Nome','Telefone','Email','Cidade',
      'Advogados','Faturamento','Investimento','Instagram',
      'Pagina','UTM Source','UTM Medium','UTM Campaign','UTM Content',
      'FBC','FBP','Origem','Event ID','Parcial','Respostas (JSON)'
    ]
  }
};

function schemaFor_(pagina) {
  if ((pagina || '').indexOf('advoga') > -1) return SCHEMAS.advoga;
  return SCHEMAS.solar;
}

// resolve o valor de cada coluna: campos do topo do payload OU de respostas{}
function valueFor_(key, d, r) {
  if (key === 'data_hora')      return new Date();
  if (key === 'qualificado')    return d.qualificado === true ? 'SIM' : 'NAO';
  if (key === 'top_tier')       return d.top_tier === true ? 'SIM' : '';
  if (key === 'parcial')        return d.parcial === true ? 'SIM' : '';
  if (key === 'respostas_json') return JSON.stringify(r);
  if (d[key] != null && typeof d[key] !== 'object') return d[key];
  if (r[key] != null) return r[key];
  return '';
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    var r = d.respostas || {};
    var schema = schemaFor_(d.pagina);

    var sheet = pegarAba_(schema.aba);
    garantirCabecalho_(sheet, schema.titulos);

    var valores = schema.colunas.map(function(c){ return valueFor_(c, d, r); });

    // UPSERT por event_id: atualiza a linha do envio parcial em vez de duplicar.
    var idCol = schema.colunas.indexOf('event_id') + 1;
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

    return resposta_({ ok: true, aba: schema.aba, atualizado: alvo > 0 });
  } catch (err) {
    return resposta_({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return resposta_({ ok: true, msg: '1Nort LP-Quiz endpoint no ar (Solar + Advoga)' });
}

function garantirCabecalho_(sheet, titulos) {
  var faixa = sheet.getRange(1, 1, 1, titulos.length);
  var atual = faixa.getValues()[0];
  var precisa = false;
  for (var i = 0; i < titulos.length; i++) {
    if (atual[i] !== titulos[i]) { precisa = true; break; }
  }
  if (precisa) {
    faixa.setValues([titulos]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

function pegarAba_(nome) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(nome);
  if (!sheet) sheet = ss.insertSheet(nome);
  return sheet;
}

function resposta_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
