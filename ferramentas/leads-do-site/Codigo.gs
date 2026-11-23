/**
 * PROIMOB · LEADS DO SITE
 * Planilha que recebe os cadastros do site ProImob (simulações, visitas, propostas, consórcio)
 * e entrega a lista para o app PRO IMOBILIÁRIA - ANDERSON ALVES.
 *
 * Instale este código NO E-MAIL PRÓPRIO DO SITE (não no e-mail pessoal):
 *   1. Entre no Google com o e-mail do site e crie uma planilha: "ProImob · Leads do site".
 *   2. Na planilha: Extensões > Apps Script. Apague o que estiver lá e cole este código. Salve.
 *   3. No topo, escolha a função "configurar" e clique em Executar. Autorize.
 *      Abra "Registro de execução": a CHAVE aparece lá. Guarde-a (ela também fica na aba "Config").
 *   4. Implantar > Nova implantação > tipo "App da Web".
 *        Executar como: Eu   ·   Quem pode acessar: Qualquer pessoa
 *      Copie a URL que termina em /exec.
 *   5. A URL vai no site (config.js) e a URL + CHAVE vão no app, em Leads > Leads do site > Ligar a planilha.
 *
 * O site só consegue GRAVAR. Para LER a lista é preciso a CHAVE, que só o app conhece.
 */

const ABA = 'Leads do site';
const COLUNAS_BASE = ['id', 'data', 'nome', 'fone', 'email', 'tipo', 'intencao', 'imovel', 'imovelTitulo', 'link',
  'cidade', 'bairro', 'renda', 'mensagem', 'canal', 'horario', 'prazoCompra', 'faixaValor', 'formaPagamento',
  'estadoImovel', 'regiao', 'score', 'scoreMotivos', 'consentResposta', 'consentOfertas', 'consentData',
  'origem', 'termo', 'pagina'];

function configurar() {
  const props = PropertiesService.getScriptProperties();
  let chave = props.getProperty('CHAVE');
  if (!chave) {
    chave = 'pi-' + Utilities.getUuid().replace(/-/g, '').slice(0, 20);
    props.setProperty('CHAVE', chave);
  }
  aba_();
  const ss = SpreadsheetApp.getActive();
  const cfg = ss.getSheetByName('Config') || ss.insertSheet('Config');
  cfg.clear();
  cfg.getRange(1, 1, 4, 2).setValues([
    ['Chave do app (não compartilhe)', chave],
    ['Avisar por e-mail a cada lead novo', 'sim'],
    ['E-mail que recebe o aviso', Session.getEffectiveUser().getEmail()],
    ['Instruções', 'Para trocar a chave, apague a propriedade CHAVE em Configurações do projeto e rode "configurar" de novo.']
  ]);
  cfg.autoResizeColumns(1, 2);
  Logger.log('CHAVE DO APP: ' + chave);
  return chave;
}

function aba_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(ABA);
  if (!sh) {
    sh = ss.insertSheet(ABA, 0);
    sh.getRange(1, 1, 1, COLUNAS_BASE.length).setValues([COLUNAS_BASE]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** O site grava aqui (POST). */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!d.nome || !(d.fone || d.email)) return json_({ ok: false, erro: 'dados incompletos' });
    d.id = 'S' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
    d.data = d.data || new Date().toISOString();
    if (!d.origem && d.utm_source) d.origem = d.utm_source;

    const sh = aba_();
    let cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    const novas = Object.keys(d).filter(k => cab.indexOf(k) < 0);
    if (novas.length) {
      sh.getRange(1, cab.length + 1, 1, novas.length).setValues([novas]).setFontWeight('bold');
      cab = cab.concat(novas);
    }
    const linha = sh.getRange(sh.getLastRow() + 1, 1, 1, cab.length);
    linha.setNumberFormat('@'); /* grava como texto: telefone e data não viram número nem fórmula */
    linha.setValues([cab.map(k => d[k] === undefined || d[k] === null ? '' : String(d[k]))]);
    avisar_(d);
    return json_({ ok: true, id: d.id });
  } catch (err) {
    return json_({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** O app lê aqui (GET ?chave=...&desde=...). */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const chave = PropertiesService.getScriptProperties().getProperty('CHAVE');
  if (!chave || p.chave !== chave) return json_({ erro: 'chave inválida' });
  const sh = aba_();
  const vals = sh.getDataRange().getValues();
  const cab = vals.shift().map(String);
  const desde = String(p.desde || '');
  const leads = vals.map(r => {
    const o = {};
    cab.forEach((k, i) => { o[k] = r[i] instanceof Date ? r[i].toISOString() : r[i]; });
    o.data = String(o.data || '');
    return o;
  }).filter(o => o.id && (!desde || o.data > desde));
  return json_({ ok: true, total: leads.length, leads: leads });
}

function avisar_(d) {
  try {
    const cfg = SpreadsheetApp.getActive().getSheetByName('Config');
    if (cfg && String(cfg.getRange(2, 2).getValue()).toLowerCase() !== 'sim') return;
    const para = (cfg && cfg.getRange(3, 2).getValue()) || Session.getEffectiveUser().getEmail();
    const fone = String(d.fone || '').replace(/\D/g, '');
    const wa = fone ? 'https://wa.me/55' + fone.replace(/^55/, '') : '';
    MailApp.sendEmail({
      to: para,
      subject: 'Novo lead no site ProImob: ' + d.nome + ' · ' + (d.intencao || d.tipo || 'contato'),
      htmlBody:
        '<p><b>' + esc_(d.nome) + '</b> · ' + esc_(d.intencao || d.tipo || '') + '</p>' +
        '<p>WhatsApp: ' + esc_(d.fone || '—') + (wa ? ' · <a href="' + wa + '">abrir conversa</a>' : '') +
        '<br>E-mail: ' + esc_(d.email || '—') + '</p>' +
        (d.imovelTitulo ? '<p>Imóvel: ' + esc_(d.imovelTitulo) + '</p>' : '') +
        (d.mensagem ? '<p>' + esc_(d.mensagem) + '</p>' : '') +
        '<p style="color:#666">Abra o PRO IMOBILIÁRIA e toque em “Buscar novos leads” para trazer este cadastro para a lista.</p>'
    });
  } catch (err) { /* o lead já foi gravado; o aviso é opcional */ }
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

/** Teste rápido: rode esta função e veja se aparece uma linha na aba. Depois apague a linha. */
function testeGravar() {
  doPost({ postData: { contents: JSON.stringify({ nome: 'Teste ProImob', fone: '51999990000', email: 'teste@exemplo.com', tipo: 'simular', intencao: 'Teste de instalação' }) } });
}
