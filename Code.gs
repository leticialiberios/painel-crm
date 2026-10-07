/**
 * Backend do Painel de Propostas — Google Apps Script
 * Cole este arquivo inteiro em Extensões > Apps Script de uma Planilha Google.
 * Propriedades do script (Configurações do projeto > Propriedades do script):
 *   TOKEN  senha de acesso ao painel, igual para todos os sócios
 *   PINS   PIN pessoal de aprovação de cada sócio, em JSON. Ex.:
 *          {"Letícia Libério":"4821","Bruno Montenegro":"7390","Rafael Lucas Russo Daniel":"1155"}
 * Veja o LEIAME.md para o passo a passo.
 */

const VERSAO_BACKEND = 4;   // o painel confere este número para saber se a implantação está atualizada
const SHEET_ITEMS = 'Propostas';
const SHEET_CONFIG = 'Config';
const SHEET_REC = 'Recebiveis';
const REC_COLS = ['id', 'codigo', 'cliente', 'servico', 'empreendimento', 'parcela', 'valorBruto', 'imposto', 'desembolso',
                  'status', 'previsao', 'nf', 'emissao', 'dataPagamento', 'captacao', 'competencia', 'atualizadoEm', 'versao', 'json'];
const SHEET_PAG = 'Pagamentos';
const PAG_COLS = ['id', 'tipo', 'descricao', 'fornecedor', 'categoria', 'centroCusto', 'valor', 'vencimento', 'status',
                  'dataPagamento', 'empreendimento', 'recorrente', 'reembolso', 'atualizadoEm', 'versao', 'json'];
const COLS = ['id', 'codigo', 'cliente', 'objeto', 'valor', 'status', 'aprovacao', 'responsavel',
              'dataEntrada', 'prazo', 'atualizadoEm', 'atualizadoPor', 'versao', 'json'];

function doGet(e) { return handle_((e && e.parameter) || {}); }

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) {}
  return handle_(p);
}

function handle_(p) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('TOKEN');
  if (token && p.token !== token) return out_({ ok: false, error: 'Senha (TOKEN) inválida' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (p.action) {
      case 'list':       return out_({ ok: true, items: listItems_(), recebiveis: listRows_(SHEET_REC, REC_COLS), pagamentos: listRows_(SHEET_PAG, PAG_COLS), versaoBackend: VERSAO_BACKEND, config: getConfig_(), pinsConfigurados: Object.keys(pins_()) });
      case 'saveRec':    return out_(saveRec_(p.rec));
      case 'deleteRec':  return out_(deleteRow_(SHEET_REC, REC_COLS, p.id));
      case 'savePag':    return out_(saveIn_(SHEET_PAG, PAG_COLS, p.pag, 'pag'));
      case 'deletePag':  return out_(deleteRow_(SHEET_PAG, PAG_COLS, p.id));
      case 'save':       return out_(save_(p.item, p.user));
      case 'vote':       return out_(vote_(p));
      case 'delete':     return out_(deleteItem_(p.id));
      case 'saveConfig': setConfig_(p.config); return out_({ ok: true });
      default:           return out_({ ok: false, error: 'Ação desconhecida' });
    }
  } catch (err) {
    return out_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Rode esta função uma vez pelo editor (selecione "configurarPlanilha" e clique em Executar)
 * para criar todas as abas e autorizar o script. Não apaga nada.
 */
function configurarPlanilha() {
  sheet_(SHEET_ITEMS, COLS);
  sheet_(SHEET_REC, REC_COLS);
  sheet_(SHEET_PAG, PAG_COLS);
  sheet_(SHEET_CONFIG, ['chave', 'valor']);
  const abas = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
  console.log('Abas na planilha: ' + abas.join(', ') + ' · backend versão ' + VERSAO_BACKEND);
  return abas;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function pins_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('PINS') || '{}'); }
  catch (e) { throw new Error('A propriedade PINS não é um JSON válido.'); }
}

/* ---------- Planilha ---------- */
function sheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  }
  return sh;
}

function listRows_(name, cols) {
  const sh = sheet_(name, cols);
  const ji = cols.indexOf('json');
  return sh.getDataRange().getValues().slice(1)
    .filter(r => r[0])
    .map(r => { try { return JSON.parse(r[ji]); } catch (e) { return null; } })
    .filter(Boolean);
}

function writeRow_(sh, cols, row, obj) {
  obj.versao = (obj.versao || 0) + 1;
  const json = JSON.stringify(obj);
  if (json.length > 49000) throw new Error('Registro muito grande para uma célula da planilha.');
  const vals = cols.map(c => c === 'json' ? json : (obj[c] === undefined || obj[c] === null ? '' : String(obj[c])));
  const range = row > 0 ? sh.getRange(row, 1, 1, cols.length) : sh.getRange(sh.getLastRow() + 1, 1, 1, cols.length);
  range.setNumberFormat('@');
  range.setValues([vals]);
  return obj;
}

function saveRec_(rec) { return saveIn_(SHEET_REC, REC_COLS, rec, 'rec'); }

/* Grava um registro com controle de versão. key = nome do campo na resposta ('rec' ou 'pag'). */
function saveIn_(name, cols, obj, key) {
  if (!obj || !obj.id) return { ok: false, error: 'Registro sem id' };
  const sh = sheet_(name, cols);
  const row = findRow_(sh, obj.id);
  if (row > 0) {
    let cur = null;
    try { cur = JSON.parse(sh.getRange(row, cols.indexOf('json') + 1).getValue()); } catch (e) {}
    if (cur && (cur.versao || 0) !== (obj.versao || 0)) { const r = { ok: false, conflict: true }; r[key] = cur; return r; }
  }
  obj.atualizadoEm = new Date().toISOString();
  writeRow_(sh, cols, row, obj);
  const res = { ok: true }; res[key] = obj; return res;
}

function deleteRow_(name, cols, id) {
  const sh = sheet_(name, cols);
  const row = findRow_(sh, id);
  if (row > 0) sh.deleteRow(row);
  return { ok: true };
}

function listItems_() {
  const sh = sheet_(SHEET_ITEMS, COLS);
  const ji = COLS.indexOf('json');
  return sh.getDataRange().getValues().slice(1)
    .filter(r => r[0])
    .map(r => { try { return JSON.parse(r[ji]); } catch (e) { return null; } })
    .filter(Boolean);
}

function findRow_(sh, id) {
  const ids = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  for (let i = 1; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 1;
  return -1;
}

function readItem_(sh, row) {
  if (row < 0) return null;
  try { return JSON.parse(sh.getRange(row, COLS.indexOf('json') + 1).getValue()); } catch (e) { return null; }
}

function writeItem_(sh, row, item) {
  item.versao = (item.versao || 0) + 1;
  const json = JSON.stringify(item);
  if (json.length > 49000) throw new Error('Proposta muito grande para uma célula da planilha (reduza comentários ou cláusulas).');
  const ap = item.aprovacao || {};
  const rowValues = COLS.map(c => {
    if (c === 'json') return json;
    if (c === 'aprovacao') return ap.estado ? ap.estado + (ap.rodada ? ' (rodada ' + ap.rodada + ')' : '') : '';
    const v = item[c];
    return v === undefined || v === null ? '' : String(v);
  });
  const range = row > 0 ? sh.getRange(row, 1, 1, COLS.length) : sh.getRange(sh.getLastRow() + 1, 1, 1, COLS.length);
  range.setNumberFormat('@');
  range.setValues([rowValues]);
  return item;
}

/* ---------- Ações ---------- */
function save_(item, user) {
  if (!item || !item.id) return { ok: false, error: 'Proposta sem id' };
  const sh = sheet_(SHEET_ITEMS, COLS);
  const row = findRow_(sh, item.id);
  const cur = readItem_(sh, row);
  if (cur && (cur.versao || 0) !== (item.versao || 0)) return { ok: false, conflict: true, item: cur };

  const cfg = getConfig_();
  const now = new Date().toISOString();
  const r = applySave(cur, item, cfg, now, user || item.atualizadoPor || '');
  if (r.error) return { ok: false, error: r.error };
  writeItem_(sh, row, r.item);
  if (r.novaRodada) notifyRound_(r.item, cfg);

  // Proposta aceita pelo cliente: gera os recebíveis (uma vez)
  let recebiveis = [];
  if (r.item.status === 'aprovada' && (!cur || cur.status !== 'aprovada')) {
    const existentes = listRows_(SHEET_REC, REC_COLS).filter(x => x.propostaId === r.item.id);
    if (!existentes.length) {
      const shr = sheet_(SHEET_REC, REC_COLS);
      recebiveis = buildRecebiveis(r.item, now, user || '', () => Utilities.getUuid());
      recebiveis.forEach(rec => { rec.atualizadoEm = now; writeRow_(shr, REC_COLS, -1, rec); });
    }
  }
  return { ok: true, item: r.item, recebiveis: recebiveis };
}

function vote_(p) {
  const pins = pins_();
  if (!Object.keys(pins).length) return { ok: false, error: 'PINs de aprovação não configurados no Apps Script (propriedade PINS).' };
  const esperado = pins[p.socio];
  if (esperado === undefined || String(esperado).trim() === '') return { ok: false, error: 'Não há PIN cadastrado para ' + p.socio + '.' };
  if (String(esperado).trim() !== String(p.pin || '').trim()) return { ok: false, error: 'PIN incorreto.' };

  const sh = sheet_(SHEET_ITEMS, COLS);
  const row = findRow_(sh, p.id);
  const cur = readItem_(sh, row);
  if (!cur) return { ok: false, error: 'Proposta não encontrada.' };

  const cfg = getConfig_();
  const r = applyVote(cur, cfg, p.socio, p.decisao, p.comentario, p.hash, new Date().toISOString());
  if (r.error) return { ok: false, error: r.error };
  writeItem_(sh, row, r.item);
  if (r.item.aprovacao.estado !== 'pendente') notifyResult_(r.item, cfg, p.socio);
  return { ok: true, item: r.item };
}

function deleteItem_(id) {
  const sh = sheet_(SHEET_ITEMS, COLS);
  const row = findRow_(sh, id);
  if (row > 0) sh.deleteRow(row);
  return { ok: true };
}

function getConfig_() {
  const sh = sheet_(SHEET_CONFIG, ['chave', 'valor']);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (v[i][0] === 'config') { try { return JSON.parse(v[i][1]); } catch (e) {} }
  }
  return {};
}

function setConfig_(config) {
  const sh = sheet_(SHEET_CONFIG, ['chave', 'valor']);
  const v = sh.getDataRange().getValues();
  const json = JSON.stringify(config || {});
  for (let i = 1; i < v.length; i++) {
    if (v[i][0] === 'config') { sh.getRange(i + 1, 2).setValue(json); return; }
  }
  sh.appendRow(['config', json]);
}

/* ---------- Notificações por e-mail (opcional) ---------- */
function emailOf_(cfg, nome) {
  const s = (cfg.socios || []).find(x => x.nome === nome);
  return s && s.email ? s.email : '';
}

function send_(cfg, to, subject, body) {
  if (cfg.notificarEmail === false || !to) return;
  try {
    MailApp.sendEmail({ to: to, subject: subject, body: body + (cfg.urlPainel ? '\n\nAbrir o painel: ' + cfg.urlPainel : '') });
  } catch (e) { console.warn('Falha ao enviar e-mail: ' + e); }
}

function notifyRound_(item, cfg) {
  const ap = item.aprovacao;
  const to = requiredApprovers(cfg, ap.autor).map(n => emailOf_(cfg, n)).filter(Boolean).join(',');
  send_(cfg, to,
    'Aprovação pendente: ' + item.codigo + ' – ' + item.cliente,
    ap.autor + ' enviou a proposta ' + item.codigo + ' (' + item.cliente + ') para aprovação.\n\n' +
    'Abra o Painel de Propostas, revise a proposta e registre sua decisão na aba "Aprovação".');
}

function notifyResult_(item, cfg, socio) {
  const ap = item.aprovacao;
  const ok = ap.estado === 'aprovada';
  const voto = (ap.votos || {})[socio] || {};
  send_(cfg, emailOf_(cfg, ap.autor),
    (ok ? 'Proposta aprovada: ' : 'Ajustes solicitados: ') + item.codigo + ' – ' + item.cliente,
    ok ? 'Todos os sócios aprovaram a proposta ' + item.codigo + '. Ela já pode ser gerada em PDF e enviada ao cliente.'
       : socio + ' solicitou ajustes na proposta ' + item.codigo + ':\n\n' + (voto.comentario || ''));
}

/* ===== Regras de aprovação — bloco idêntico no painel (index.html) e no servidor (Code.gs) ===== */
var LOCKED_STATUS = ['enviada', 'negociacao', 'aprovada'];

function cyrb53(str, seed) {
  var h1 = 0xdeadbeef ^ (seed || 0), h2 = 0x41c6ce57 ^ (seed || 0);
  for (var i = 0, ch; i < str.length; i++) {
    ch = str.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507);
  h1 ^= Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507);
  h2 ^= Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(36);
}

/* Impressão digital do conteúdo que vai ao cliente. Qualquer mudança aqui anula aprovações. */
function contentHash(it) {
  return cyrb53(JSON.stringify([
    it.cliente || '', it.clienteCnpj || '', it.cidade || '', it.dataProposta || '',
    it.titulo || '', it.doc || null, it.signatarios || [],
    it.fin ? [it.fin.captacao || '', it.fin.divisao || null] : null
  ]));
}

/* ----- Financeiro ----- */
function numBR(v) { var n = parseFloat(String(v === undefined || v === null ? '' : v).replace(',', '.')); return isNaN(n) ? 0 : n; }
function round2(v) { return Math.round(v * 100) / 100; }
function docSubtotal(doc) { return ((doc && doc.itens) || []).reduce(function (a, r) { return a + numBR(r.valor); }, 0); }
/* desconto: {tipo:'pct'|'valor', valor, motivo} */
function docDescontoValor(doc) {
  var d = doc && doc.desconto;
  if (!d || !numBR(d.valor)) return 0;
  var s = docSubtotal(doc);
  var v = d.tipo === 'valor' ? numBR(d.valor) : s * numBR(d.valor) / 100;
  return round2(Math.min(v, s));
}
function docTotalLiquido(doc) { return round2(docSubtotal(doc) - docDescontoValor(doc)); }
function propostaTotal(it) {
  var t = docTotalLiquido(it.doc || {});
  return t || numBR(it.valor);
}
/* Gera os recebíveis (um por parcela) de uma proposta aceita. */
function buildRecebiveis(it, now, user, newId) {
  var fin = it.fin || {};
  var ps = (it.doc && it.doc.parcelas && it.doc.parcelas.length) ? it.doc.parcelas : [{ pct: 100, condicao: '' }];
  var fp = fin.parcelas || [];
  var total = propostaTotal(it), soma = 0, n = ps.length;
  return ps.map(function (p, i) {
    var v = i < n - 1 ? round2(total * numBR(p.pct) / 100) : round2(total - soma);
    soma += v;
    var f = fp[i] || {};
    return {
      id: newId(), propostaId: it.id, codigo: it.codigo || '', cliente: it.cliente || '',
      servico: fin.servico || '', empreendimento: fin.empreendimento || '',
      parcela: i + 1, deTotal: n, condicao: p.condicao || '',
      valorBruto: v, imposto: fin.imposto === undefined ? 12.5 : numBR(fin.imposto), desembolso: numBR(f.desembolso),
      captacao: fin.captacao || '', competencia: fin.competencia || '',
      divisao: JSON.parse(JSON.stringify(fin.divisao || {})), excecao: fin.excecao ? (fin.motivo || 'Exceção à regra') : '',
      previsao: f.previsao || '', status: 'a faturar', nf: '', emissao: '', dataPagamento: '', obs: '',
      log: [{ data: now, autor: user, texto: 'Gerado a partir da proposta ' + (it.codigo || '') + ' (parcela ' + (i + 1) + ' de ' + n + ')' }],
      versao: 0, criadoEm: now
    };
  });
}

function emptyApproval() {
  return { estado: 'rascunho', rodada: 0, autor: '', enviadoEm: '', hash: '', votos: {} };
}

/* Aprovadores = todos os sócios cadastrados, exceto quem elaborou */
function requiredApprovers(cfg, autor) {
  return ((cfg && cfg.socios) || [])
    .map(function (s) { return String(s.nome || '').trim(); })
    .filter(function (n) { return n && n !== autor; });
}

function computeApprovalState(ap, req) {
  var v = ap.votos || {};
  if (req.some(function (n) { return v[n] && v[n].decisao === 'ajustes'; })) return 'ajustes';
  if (req.length && req.every(function (n) { return v[n] && v[n].decisao === 'aprovado'; })) return 'aprovada';
  return 'pendente';
}

/* Aplica as regras ao salvar. prev = versão gravada (ou null), item = versão enviada. */
function applySave(prev, item, cfg, now, user) {
  var pa = (prev && prev.aprovacao) ? prev.aprovacao : emptyApproval();
  var ia = item.aprovacao || pa;
  var ap = JSON.parse(JSON.stringify(pa));
  var novaRodada = false;
  item.log = item.log || [];
  var h = contentHash(item);
  var prevStatus = prev ? prev.status : '';

  if (ia.estado === 'pendente' && (ia.rodada || 0) > (pa.rodada || 0)) {
    var autor = ia.autor || user;
    if (!requiredApprovers(cfg, autor).length) return { error: 'Cadastre os sócios em Configurações antes de enviar para aprovação.' };
    ap = { estado: 'pendente', rodada: ia.rodada, autor: autor, enviadoEm: now, hash: h, votos: {} };
    novaRodada = true;
  } else if (ia.estado === 'rascunho' && pa.estado === 'pendente' && ia.rodada === pa.rodada) {
    ap.estado = 'rascunho'; ap.votos = {};
  } else if ((pa.estado === 'pendente' || pa.estado === 'aprovada') && pa.hash && h !== pa.hash) {
    ap.estado = 'rascunho'; ap.votos = {};
    item.log.push({ data: now, autor: user, texto: 'Conteúdo da proposta alterado após o envio para aprovação — aprovações anuladas.' });
  }

  if (ap.estado === 'pendente' && ['nova', 'elaboracao', 'liberada'].indexOf(item.status) >= 0) item.status = 'revisao';
  if (ap.estado === 'aprovada' && item.status === 'revisao') item.status = 'liberada';
  if ((ap.estado === 'rascunho' || ap.estado === 'ajustes') && (item.status === 'revisao' || item.status === 'liberada')) item.status = 'elaboracao';
  if (LOCKED_STATUS.indexOf(item.status) >= 0 && ap.estado !== 'aprovada' && LOCKED_STATUS.indexOf(prevStatus) < 0) {
    return { error: 'A proposta precisa ser aprovada pelos sócios antes de ser marcada como enviada ou aceita.' };
  }
  item.aprovacao = ap;
  return { item: item, novaRodada: novaRodada };
}

/* Registra o voto de um sócio. hash = versão que ele revisou. */
function applyVote(item, cfg, socio, decisao, comentario, hash, now) {
  var ap = item.aprovacao;
  if (!ap || ap.estado !== 'pendente') return { error: 'Esta proposta não está aguardando aprovação.' };
  if (hash && hash !== ap.hash) return { error: 'A proposta foi alterada desde que você a abriu. Atualize e revise antes de votar.' };
  var req = requiredApprovers(cfg, ap.autor);
  if (req.indexOf(socio) < 0) return { error: socio + ' não está entre os aprovadores desta proposta.' };
  if (decisao !== 'aprovado' && decisao !== 'ajustes') return { error: 'Decisão inválida.' };
  comentario = String(comentario || '').trim();
  if (decisao === 'ajustes' && !comentario) return { error: 'Descreva os ajustes solicitados.' };
  ap.votos = ap.votos || {};
  ap.votos[socio] = { decisao: decisao, data: now, comentario: comentario };
  ap.estado = computeApprovalState(ap, req);
  item.log = item.log || [];
  item.log.push({ data: now, autor: socio, texto: (decisao === 'aprovado' ? 'Aprovou a proposta' : 'Solicitou ajustes') + ' (rodada ' + ap.rodada + ')' + (comentario ? ': ' + comentario : '') });
  if (comentario) { item.comentarios = item.comentarios || []; item.comentarios.push({ data: now, autor: socio, texto: (decisao === 'aprovado' ? '[Aprovação] ' : '[Ajustes] ') + comentario }); }
  if (ap.estado === 'aprovada' && item.status === 'revisao') item.status = 'liberada';
  if (ap.estado === 'ajustes' && (item.status === 'revisao' || item.status === 'liberada')) item.status = 'elaboracao';
  item.atualizadoEm = now; item.atualizadoPor = socio;
  return { item: item };
}
/* ===== fim do bloco compartilhado ===== */

