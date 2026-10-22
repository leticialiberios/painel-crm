/**
 * Painel de Licenças — Mannare / BMM / Vertis
 * Backend em Google Apps Script, vinculado a uma planilha própria
 * (separada do Painel de Propostas).
 *
 * Primeira instalação: rode a função `configurar` uma vez pelo editor.
 * Ela cria as abas, gera a chave de acesso (TOKEN) nas Propriedades do
 * script e ativa o aviso diário. Nenhum segredo fica neste código.
 */

const VERSAO_BACKEND = 1;
const TZ = 'America/Fortaleza';
const ABA_LIC = 'Licencas';
const ABA_CFG = 'Config';
const HORA_AVISO = 7; // aviso diário por volta das 7h (horário de Fortaleza)

const COLS = [
  'id', 'cliente', 'empreendimento', 'cnpj', 'tipo', 'numero', 'orgao', 'processo',
  'emissao', 'validade', 'aviso_dias', 'responsavel', 'situacao',
  'protocolo_data', 'protocolo_numero', 'documento_url', 'obs',
  'renova', 'renovada_por', 'ultimo_marco', 'criado_em', 'atualizado_em'
];
const COLS_TEXTO = ['id', 'renova', 'renovada_por', 'emissao', 'validade', 'protocolo_data', 'cnpj', 'numero', 'processo', 'protocolo_numero', 'ultimo_marco'];
const EDITAVEIS = [
  'cliente', 'empreendimento', 'cnpj', 'tipo', 'numero', 'orgao', 'processo',
  'emissao', 'validade', 'aviso_dias', 'responsavel', 'situacao',
  'protocolo_data', 'protocolo_numero', 'documento_url', 'obs'
];
const SITUACOES = ['Vigente', 'Renovação protocolada', 'Renovada', 'Encerrada'];
const MARCOS_PADRAO = [90, 60, 30, 15, 7, 0];

/* ------------------------------------------------------------------ */
/* Instalação                                                          */
/* ------------------------------------------------------------------ */

function configurar() {
  abaLicencas_();
  abaConfig_();
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TOKEN')) {
    props.setProperty('TOKEN', gerarToken_());
  }
  ativarAvisoDiario_();
  Logger.log('Pronto. Chave de acesso do painel (copie e guarde): ' + props.getProperty('TOKEN'));
}

function gerarToken_() {
  const bytes = Utilities.getUuid() + Utilities.getUuid();
  return bytes.replace(/-/g, '').slice(0, 32);
}

function ativarAvisoDiario_() {
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === 'verificarVencimentos'; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('verificarVencimentos')
    .timeBased().everyDays(1).atHour(HORA_AVISO).inTimezone(TZ).create();
}

function avisoDiarioAtivo_() {
  try {
    return ScriptApp.getProjectTriggers()
      .some(function (t) { return t.getHandlerFunction() === 'verificarVencimentos'; });
  } catch (e) {
    return null;
  }
}

/* ------------------------------------------------------------------ */
/* Planilha                                                            */
/* ------------------------------------------------------------------ */

function planilha_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function abaLicencas_() {
  const ss = planilha_();
  let sh = ss.getSheetByName(ABA_LIC);
  if (!sh) {
    sh = ss.insertSheet(ABA_LIC);
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]);
    sh.setFrozenRows(1);
  } else {
    // Compatibilidade: acrescenta colunas novas sem mexer nas existentes.
    const ult = Math.max(sh.getLastColumn(), 1);
    const cab = sh.getRange(1, 1, 1, ult).getValues()[0].map(String);
    const faltam = COLS.filter(function (c) { return cab.indexOf(c) < 0; });
    if (faltam.length) {
      const inicio = cab.filter(String).length + 1;
      sh.getRange(1, inicio, 1, faltam.length).setValues([faltam]);
    }
  }
  return sh;
}

function abaConfig_() {
  const ss = planilha_();
  let sh = ss.getSheetByName(ABA_CFG);
  if (!sh) {
    sh = ss.insertSheet(ABA_CFG);
    sh.getRange(1, 1, 1, 2).setValues([['chave', 'valor']]);
    sh.getRange(2, 1, 1, 2).setValues([['emails_aviso', '']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function cabecalho_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
}

function paraTextoData_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  }
  const s = String(v).trim();
  const br = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (br) return br[3] + '-' + ('0' + br[2]).slice(-2) + '-' + ('0' + br[1]).slice(-2);
  return s.slice(0, 10);
}

function lerLicencas_() {
  const sh = abaLicencas_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const cab = cabecalho_(sh);
  const dados = sh.getRange(2, 1, n, cab.length).getValues();
  const out = [];
  dados.forEach(function (linha, i) {
    const o = { _linha: i + 2 };
    cab.forEach(function (c, j) { if (c) o[c] = linha[j]; });
    if (!o.id) return;
    ['emissao', 'validade', 'protocolo_data'].forEach(function (c) { o[c] = paraTextoData_(o[c]); });
    ['criado_em', 'atualizado_em'].forEach(function (c) {
      if (Object.prototype.toString.call(o[c]) === '[object Date]') {
        o[c] = Utilities.formatDate(o[c], TZ, "yyyy-MM-dd'T'HH:mm:ss");
      }
    });
    o.id = String(o.id);
    o.aviso_dias = Number(o.aviso_dias) || 0;
    o.situacao = o.situacao || 'Vigente';
    ['cliente', 'empreendimento', 'cnpj', 'tipo', 'numero', 'orgao', 'processo', 'responsavel',
      'protocolo_numero', 'documento_url', 'obs', 'renova', 'renovada_por', 'ultimo_marco']
      .forEach(function (c) { o[c] = o[c] === undefined || o[c] === null ? '' : String(o[c]); });
    out.push(o);
  });
  return out;
}

function gravarLicenca_(lic) {
  const sh = abaLicencas_();
  const cab = cabecalho_(sh);
  const linha = cab.map(function (c) {
    const v = lic[c];
    if (v === undefined || v === null) return '';
    if (COLS_TEXTO.indexOf(c) >= 0 && v !== '') return "'" + String(v).replace(/^'/, '');
    return v;
  });
  if (lic._linha) {
    sh.getRange(lic._linha, 1, 1, cab.length).setValues([linha]);
  } else {
    sh.appendRow(linha);
    lic._linha = sh.getLastRow();
  }
}

function lerConfig_() {
  const sh = abaConfig_();
  const n = sh.getLastRow() - 1;
  const cfg = { emails_aviso: '' };
  if (n < 1) return cfg;
  sh.getRange(2, 1, n, 2).getValues().forEach(function (r) {
    if (r[0]) cfg[String(r[0])] = r[1] === null ? '' : String(r[1]);
  });
  return cfg;
}

function gravarConfig_(chave, valor) {
  const sh = abaConfig_();
  const n = sh.getLastRow() - 1;
  if (n >= 1) {
    const chaves = sh.getRange(2, 1, n, 1).getValues();
    for (let i = 0; i < chaves.length; i++) {
      if (String(chaves[i][0]) === chave) {
        sh.getRange(i + 2, 2).setValue(valor);
        return;
      }
    }
  }
  sh.appendRow([chave, valor]);
}

/* ------------------------------------------------------------------ */
/* Regras                                                              */
/* ------------------------------------------------------------------ */

function hojeISO_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

function diasEntre_(deISO, ateISO) {
  const a = deISO.split('-').map(Number), b = ateISO.split('-').map(Number);
  const ua = Date.UTC(a[0], a[1] - 1, a[2]), ub = Date.UTC(b[0], b[1] - 1, b[2]);
  return Math.round((ub - ua) / 86400000);
}

function dataValida_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s || '')) return false;
  const p = s.split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
  return d.getUTCFullYear() === p[0] && d.getUTCMonth() === p[1] - 1 && d.getUTCDate() === p[2];
}

/** Marco de aviso atingido hoje, ou '' se ainda fora do prazo de aviso. */
function marcoAtual_(lic, hoje) {
  if (lic.situacao !== 'Vigente' || !dataValida_(lic.validade)) return '';
  const dias = diasEntre_(hoje, lic.validade);
  if (dias < 0) return 'V' + Math.floor((-dias - 1) / 7); // vencida: lembrete semanal
  const aviso = Number(lic.aviso_dias) || 0;
  if (dias > aviso) return '';
  const marcos = [aviso].concat(MARCOS_PADRAO.filter(function (m) { return m < aviso; }));
  let atingido = aviso;
  marcos.forEach(function (m) { if (dias <= m && m < atingido) atingido = m; });
  return String(atingido);
}

function validar_(lic) {
  if (!String(lic.cliente || '').trim()) throw new Error('Informe o cliente.');
  if (!String(lic.tipo || '').trim()) throw new Error('Informe o tipo de licença.');
  if (!dataValida_(lic.validade)) throw new Error('Informe uma data de validade válida.');
  if (lic.emissao && !dataValida_(lic.emissao)) throw new Error('Data de emissão inválida.');
  if (lic.protocolo_data && !dataValida_(lic.protocolo_data)) throw new Error('Data do protocolo inválida.');
  const aviso = Number(lic.aviso_dias);
  if (!(aviso >= 0 && aviso <= 730 && Math.floor(aviso) === aviso)) {
    throw new Error('O prazo de aviso deve ser um número de dias entre 0 e 730.');
  }
  if (SITUACOES.indexOf(lic.situacao) < 0) throw new Error('Situação inválida.');
}

function limpar_(dados) {
  const o = {};
  EDITAVEIS.forEach(function (c) {
    if (dados[c] === undefined || dados[c] === null) return;
    o[c] = c === 'aviso_dias' ? Number(dados[c]) : String(dados[c]).trim();
  });
  return o;
}

function publico_(lic) {
  const o = {};
  Object.keys(lic).forEach(function (k) { if (k !== '_linha') o[k] = lic[k]; });
  return o;
}

function agora_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss"); }

/* ------------------------------------------------------------------ */
/* API                                                                 */
/* ------------------------------------------------------------------ */

function doGet() {
  return json_({ ok: true, app: 'licencas', versao: VERSAO_BACKEND });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, erro: 'Requisição inválida.' });
  }
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!token || String(req.token || '') !== token) {
    return json_({ ok: false, erro: 'Chave de acesso inválida.', codigo: 'auth' });
  }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const r = executar_(req);
    r.ok = true;
    r.versao = VERSAO_BACKEND;
    return json_(r);
  } catch (err) {
    return json_({ ok: false, erro: err && err.message ? err.message : String(err), versao: VERSAO_BACKEND });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function executar_(req) {
  switch (req.acao) {
    case 'listar': return listar_();
    case 'salvar': return salvar_(req.licenca || {});
    case 'renovar': return renovar_(req.id, req.nova || {});
    case 'excluir': return excluir_(req.id);
    case 'config': return salvarConfig_(req.config || {});
    case 'ativar_aviso': ativarAvisoDiario_(); return { aviso_ativo: avisoDiarioAtivo_() };
    case 'testar_aviso': return { resultado: verificarVencimentos({ teste: true }) };
    default: throw new Error('Ação desconhecida: ' + req.acao);
  }
}

function listar_() {
  const cfg = lerConfig_();
  return {
    licencas: lerLicencas_().map(publico_),
    config: { emails_aviso: cfg.emails_aviso || '', ultima_verificacao: cfg.ultima_verificacao || '' },
    aviso_ativo: avisoDiarioAtivo_(),
    hoje: hojeISO_()
  };
}

function salvar_(dados) {
  const lista = lerLicencas_();
  const limpo = limpar_(dados);
  let lic;
  if (dados.id) {
    lic = lista.filter(function (l) { return l.id === String(dados.id); })[0];
    if (!lic) throw new Error('Licença não encontrada. Atualize o painel.');
    const validadeAntes = lic.validade, avisoAntes = lic.aviso_dias, sitAntes = lic.situacao;
    Object.keys(limpo).forEach(function (k) { lic[k] = limpo[k]; });
    // Se a validade, o prazo ou a situação mudaram, recomeça a contagem dos avisos.
    if (lic.validade !== validadeAntes || Number(lic.aviso_dias) !== Number(avisoAntes) || lic.situacao !== sitAntes) {
      lic.ultimo_marco = '';
    }
  } else {
    lic = limpo;
    lic.id = Utilities.getUuid().slice(0, 8);
    lic.situacao = lic.situacao || 'Vigente';
    lic.criado_em = agora_();
    lic.renova = '';
    lic.renovada_por = '';
    lic.ultimo_marco = '';
  }
  validar_(lic);
  lic.atualizado_em = agora_();
  gravarLicenca_(lic);
  return { licenca: publico_(lic) };
}

function renovar_(id, nova) {
  const lista = lerLicencas_();
  const antiga = lista.filter(function (l) { return l.id === String(id); })[0];
  if (!antiga) throw new Error('Licença não encontrada. Atualize o painel.');
  if (antiga.situacao === 'Renovada') throw new Error('Esta licença já foi renovada.');
  const base = {};
  ['cliente', 'empreendimento', 'cnpj', 'tipo', 'orgao', 'processo', 'aviso_dias', 'responsavel']
    .forEach(function (c) { base[c] = antiga[c]; });
  const dados = Object.assign(base, limpar_(nova));
  dados.situacao = 'Vigente';
  dados.protocolo_data = '';
  dados.protocolo_numero = '';
  const nl = Object.assign({}, dados, {
    id: Utilities.getUuid().slice(0, 8),
    renova: antiga.id, renovada_por: '', ultimo_marco: '',
    criado_em: agora_(), atualizado_em: agora_()
  });
  validar_(nl);
  antiga.situacao = 'Renovada';
  antiga.renovada_por = nl.id;
  antiga.atualizado_em = agora_();
  gravarLicenca_(nl);
  gravarLicenca_(antiga);
  return { licenca: publico_(nl), anterior: publico_(antiga) };
}

function excluir_(id) {
  const lic = lerLicencas_().filter(function (l) { return l.id === String(id); })[0];
  if (!lic) throw new Error('Licença não encontrada.');
  abaLicencas_().deleteRow(lic._linha);
  return { excluida: lic.id };
}

function salvarConfig_(cfg) {
  if (cfg.emails_aviso !== undefined) {
    const emails = normalizarEmails_(cfg.emails_aviso);
    const ruins = emails.filter(function (e) { return !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e); });
    if (ruins.length) throw new Error('E-mail inválido: ' + ruins.join(', '));
    gravarConfig_('emails_aviso', emails.join(', '));
  }
  return listar_();
}

function normalizarEmails_(s) {
  return String(s || '').split(/[\s,;]+/).map(function (e) { return e.trim().toLowerCase(); }).filter(String);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ */
/* Aviso diário por e-mail                                             */
/* ------------------------------------------------------------------ */

/**
 * Roda todo dia (gatilho). Manda um único e-mail-resumo quando alguma
 * licença atinge um novo marco de aviso: o prazo definido na licença,
 * depois 90, 60, 30, 15, 7 e 0 dias, e semanalmente depois de vencida.
 * Licenças com renovação protocolada, renovadas ou encerradas não geram aviso.
 * Com { teste: true }, envia o resumo atual sem registrar marcos.
 */
function verificarVencimentos(opcoes) {
  const teste = !!(opcoes && opcoes.teste === true);
  const hoje = hojeISO_();
  const lista = lerLicencas_();
  const novos = [], todas = [];
  lista.forEach(function (lic) {
    const marco = marcoAtual_(lic, hoje);
    if (!marco) return;
    const item = { lic: lic, dias: diasEntre_(hoje, lic.validade) };
    todas.push(item);
    if (marco !== lic.ultimo_marco) {
      novos.push(item);
      if (!teste) {
        lic.ultimo_marco = marco;
        gravarLicenca_(lic);
      }
    }
  });
  if (!teste) gravarConfig_('ultima_verificacao', agora_());

  const emails = normalizarEmails_(lerConfig_().emails_aviso);
  if (!emails.length) {
    if (teste) throw new Error('Cadastre ao menos um e-mail para receber os avisos.');
    return { enviado: false, motivo: 'sem e-mails', novos: novos.length };
  }
  if (!teste && !novos.length) return { enviado: false, motivo: 'nada novo', novos: 0 };

  todas.sort(function (a, b) { return a.dias - b.dias; });
  const vencidas = todas.filter(function (i) { return i.dias < 0; }).length;
  const assunto = (teste ? '[Teste] ' : '') + 'Licenças: ' +
    (vencidas ? vencidas + ' vencida(s), ' : '') +
    (todas.length - vencidas) + ' a renovar';
  MailApp.sendEmail({
    to: emails.join(','),
    subject: assunto,
    htmlBody: corpoEmail_(todas, novos, teste)
  });
  return { enviado: true, para: emails, novos: novos.length, total: todas.length };
}

function corpoEmail_(todas, novos, teste) {
  const ids = novos.map(function (i) { return i.lic.id; });
  const esc = function (s) {
    return String(s || '').replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  };
  const br = function (iso) { const p = iso.split('-'); return p[2] + '/' + p[1] + '/' + p[0]; };
  const linhas = todas.map(function (i) {
    const l = i.lic;
    const prazo = i.dias < 0 ? 'Vencida há ' + (-i.dias) + ' dia(s)'
      : i.dias === 0 ? 'Vence hoje' : 'Faltam ' + i.dias + ' dia(s)';
    const cor = i.dias < 0 ? '#b42318' : '#b54708';
    const novo = ids.indexOf(l.id) >= 0 ? ' <span style="background:#fef0c7;color:#93370d;padding:1px 6px;border-radius:8px;font-size:11px">novo</span>' : '';
    return '<tr>' +
      '<td style="padding:8px;border-bottom:1px solid #eee">' + esc(l.cliente) +
      (l.empreendimento ? '<br><span style="color:#667085">' + esc(l.empreendimento) + '</span>' : '') + '</td>' +
      '<td style="padding:8px;border-bottom:1px solid #eee">' + esc(l.tipo) +
      (l.numero ? ' nº ' + esc(l.numero) : '') +
      (l.orgao ? '<br><span style="color:#667085">' + esc(l.orgao) + '</span>' : '') + '</td>' +
      '<td style="padding:8px;border-bottom:1px solid #eee">' + br(l.validade) + '</td>' +
      '<td style="padding:8px;border-bottom:1px solid #eee;color:' + cor + ';font-weight:600">' + prazo + novo + '</td>' +
      '<td style="padding:8px;border-bottom:1px solid #eee">' + esc(l.responsavel) + '</td></tr>';
  }).join('');
  return '<div style="font-family:Arial,sans-serif;font-size:14px;color:#101828">' +
    (teste ? '<p style="color:#667085">Este é um e-mail de teste enviado pelo painel.</p>' : '') +
    (todas.length ? '<p>Licenças que precisam de renovação:</p>'
      : '<p>Nenhuma licença está no prazo de aviso hoje.</p>') +
    '<table style="border-collapse:collapse;width:100%;max-width:820px">' +
    '<tr style="text-align:left;background:#f2f4f7">' +
    '<th style="padding:8px">Cliente</th><th style="padding:8px">Licença</th>' +
    '<th style="padding:8px">Validade</th><th style="padding:8px">Prazo</th><th style="padding:8px">Responsável</th></tr>' +
    linhas + '</table>' +
    '<p style="color:#667085;font-size:12px">Ao protocolar a renovação, marque a licença como “Renovação protocolada” no painel para parar os avisos.</p>' +
    '</div>';
}
