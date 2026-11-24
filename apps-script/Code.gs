/**
 * Leozin Barber Studio — back-end de agendamento (Google Apps Script)
 *
 * Instalação: veja README.md na raiz do repositório.
 * Nenhum dado real fica neste arquivo. E-mail de resposta, "pepper" das senhas e o
 * ID da planilha ficam nas Propriedades do Script (Configurações do projeto).
 */

const SHOP_NAME = 'Leozin Barber Studio';
const TOKEN_DIAS = 30;          // validade da sessão do cliente
const HASH_ROUNDS = 1000;       // rodadas de SHA-256 na senha
const MAX_FALHAS_LOGIN = 5;     // depois disso, bloqueia o e-mail por 15 min

const HEAD = {
  Agenda:    ['ID', 'Criado em', 'Unidade', 'Serviço', 'Barbeiro', 'Data', 'Hora', 'Cliente ID', 'Cliente', 'E-mail',
              'Status', 'Lembrete enviado', 'Atualizado em', 'Histórico'],
  Clientes:  ['ID', 'Nome', 'E-mail', 'Salt', 'Hash da senha', 'Aceite LGPD em', 'Criado em', 'Hash do token', 'Token expira em'],
  Config:    ['Tipo', 'ID', 'Nome', 'Valor', 'Telefone', 'Unidade', 'Ativo', 'Obs'],
  Bloqueios: ['Unidade', 'Barbeiro', 'Data', 'Início', 'Fim', 'Motivo']
};

// Valores iniciais da aba Config. Tudo marcado como [EXEMPLO]: troque direto na planilha.
const CONFIG_INICIAL = [
  ['parametro', 'abertura', 'Abertura (HH:mm)', '09:00', '', '', 'sim', ''],
  ['parametro', 'fechamento', 'Fechamento (HH:mm)', '19:00', '', '', 'sim', 'Último horário começa 30 min antes'],
  ['parametro', 'almoco_inicio', 'Início do almoço (HH:mm)', '12:00', '', '', 'sim', 'Deixe vazio se não houver'],
  ['parametro', 'almoco_fim', 'Fim do almoço (HH:mm)', '14:00', '', '', 'sim', ''],
  ['parametro', 'dias_funcionamento', 'Dias abertos (0=dom … 6=sáb)', '2,3,4,5,6', '', '', 'sim', 'Terça a sábado'],
  ['parametro', 'intervalo_min', 'Duração do bloco (min)', '30', '', '', 'sim', ''],
  ['parametro', 'prazo_alteracao_horas', 'Prazo para remarcar/cancelar (h)', '2', '', '', 'sim', ''],
  ['parametro', 'lembrete_min', 'Lembrete por e-mail (min antes)', '30', '', '', 'sim', ''],
  ['parametro', 'dias_agenda', 'Dias exibidos no calendário', '7', '', '', 'sim', ''],
  ['parametro', 'antecedencia_min', 'Antecedência mínima para agendar (min)', '15', '', '', 'sim', ''],
  ['parametro', 'max_ativos', 'Agendamentos futuros por cliente', '3', '', '', 'sim', ''],
  ['unidade', 'marilea', 'Jardim Marilea', 'Rua das Acácias, 120', '[a definir]', '', 'sim', '[EXEMPLO]'],
  ['unidade', 'praiana', 'Cidade Praiana', 'Av. Beira-Mar, 845', '[a definir]', '', 'sim', '[EXEMPLO]'],
  ['barbeiro', 'leozin', 'Leozin', '', '', 'marilea', 'sim', '[EXEMPLO]'],
  ['barbeiro', 'biel', 'Biel', '', '', 'marilea', 'sim', '[EXEMPLO]'],
  ['barbeiro', 'rafa', 'Rafa', '', '', 'praiana', 'sim', '[EXEMPLO]'],
  ['barbeiro', 'kaua', 'Kauã', '', '', 'praiana', 'sim', '[EXEMPLO]'],
  ['servico', 'freestyle', 'Freestyle risco + mid fade', '30', '', '', 'sim', '[EXEMPLO]'],
  ['servico', 'lowfade', 'Máquina low fade', '30', '', '', 'sim', '[EXEMPLO]'],
  ['servico', 'acabamento', 'Acabamento pezinho e lateral', '30', '', '', 'sim', '[EXEMPLO]'],
  ['servico', 'sobrancelha', 'Sobrancelha', '30', '', '', 'sim', '[EXEMPLO]'],
  ['servico', 'navalhado', 'Navalhado', '30', '', '', 'sim', '[EXEMPLO]'],
  ['servico', 'barba', 'Barba na toalha quente', '30', '', '', 'sim', '[EXEMPLO]'],
  ['servico', 'tesoura', 'Degradê na tesoura', '30', '', '', 'sim', '[EXEMPLO]'],
  ['servico', 'pigmentacao', 'Pigmentação', '30', '', '', 'sim', '[EXEMPLO]']
];

/* ===================================================================== */
/* INSTALAÇÃO                                                             */
/* ===================================================================== */

/** Rode uma vez pelo editor: cria as abas, os dados de exemplo e o gatilho de lembretes. */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
  if (!ss && props.getProperty('SPREADSHEET_ID')) ss = SpreadsheetApp.openById(props.getProperty('SPREADSHEET_ID'));
  if (!ss) ss = SpreadsheetApp.create(SHOP_NAME + ' — Agenda');
  props.setProperty('SPREADSHEET_ID', ss.getId());
  if (!props.getProperty('PEPPER')) props.setProperty('PEPPER', novoSegredo_());

  Object.keys(HEAD).forEach(function (nome) {
    let sh = ss.getSheetByName(nome);
    if (!sh) sh = ss.insertSheet(nome);
    sh.getRange(1, 1, sh.getMaxRows(), HEAD[nome].length).setNumberFormat('@');
    sh.getRange(1, 1, 1, HEAD[nome].length).setValues([HEAD[nome]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  const cfg = ss.getSheetByName('Config');
  if (cfg.getLastRow() < 2) cfg.getRange(2, 1, CONFIG_INICIAL.length, HEAD.Config.length).setValues(CONFIG_INICIAL);

  // Remove a aba padrão vazia ("Página1" / "Sheet1").
  ss.getSheets().forEach(function (sh) {
    if (!HEAD[sh.getName()] && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'enviarLembretes') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enviarLembretes').timeBased().everyMinutes(5).create();

  Logger.log('Planilha pronta: ' + ss.getUrl());
  Logger.log('E-mails respondem para: ' + replyTo_());
  return ss.getUrl();
}

/* ===================================================================== */
/* WEB APP                                                                */
/* ===================================================================== */

/*
 * CORS: o site envia POST com Content-Type "text/plain", que não dispara preflight.
 * O Apps Script responde por um redirecionamento para googleusercontent.com, que já
 * devolve Access-Control-Allow-Origin: *. Por isso não há cabeçalho manual aqui.
 */
function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return json_({ ok: false, erro: 'Requisição inválida.' }); }
  return json_(rotear_(body));
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.action) return json_({ ok: true, api: SHOP_NAME });
  if (['config', 'horarios'].indexOf(p.action) < 0) return json_({ ok: false, erro: 'Use POST para esta ação.' });
  return json_(rotear_(p));
}

const ROTAS = {
  config: acaoConfig_,
  horarios: acaoHorarios_,
  cadastrar: acaoCadastrar_,
  entrar: acaoEntrar_,
  sair: acaoSair_,
  meus: acaoMeus_,
  agendar: acaoAgendar_,
  remarcar: acaoRemarcar_,
  cancelar: acaoCancelar_
};

function rotear_(req) {
  try {
    const fn = ROTAS[req.action];
    if (!fn) throw new ErroUsuario('Ação desconhecida.');
    const r = fn(req) || {};
    r.ok = true;
    return r;
  } catch (err) {
    if (err instanceof ErroUsuario) return { ok: false, erro: err.message, codigo: err.codigo || null };
    console.error(err && err.stack || err);
    return { ok: false, erro: 'Não foi possível concluir agora. Tente de novo em instantes.' };
  }
}

function ErroUsuario(msg, codigo) { this.message = msg; this.codigo = codigo; }
ErroUsuario.prototype = Object.create(Error.prototype);

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ===================================================================== */
/* AÇÕES                                                                  */
/* ===================================================================== */

function acaoConfig_() {
  const c = lerConfig_();
  return {
    unidades: c.unidades, barbeiros: c.barbeiros, servicos: c.servicos,
    params: {
      abertura: c.p.abertura, fechamento: c.p.fechamento, almocoInicio: c.p.almocoInicio, almocoFim: c.p.almocoFim,
      dias: c.p.dias, intervalo: c.p.intervalo, prazoHoras: c.p.prazoHoras, diasAgenda: c.p.diasAgenda
    }
  };
}

/** Horários de um barbeiro nos próximos dias. Não expõe nenhum dado de clientes. */
function acaoHorarios_(req) {
  const c = lerConfig_();
  const unidade = exigir_(c.unidades, req.unidade, 'Unidade');
  const barbeiro = exigir_(c.barbeiros, req.barbeiro, 'Barbeiro');
  if (barbeiro.unidade !== unidade.id) throw new ErroUsuario('Esse barbeiro não atende nessa unidade.');
  const ag = lerAgenda_(), bl = lerBloqueios_();
  return { dias: proximosDias_(c).map(function (data) {
    const aberto = c.p.dias.indexOf(diaSemana_(data)) >= 0;
    return { data: data, aberto: aberto, slots: aberto ? slotsDoDia_(c, unidade.id, barbeiro.id, data, ag, bl, null) : [] };
  }) };
}

function acaoCadastrar_(req) {
  const nome = String(req.nome || '').trim().replace(/\s+/g, ' ');
  const email = normEmail_(req.email);
  const senha = String(req.senha || '');
  if (nome.length < 2 || nome.length > 80) throw new ErroUsuario('Digite seu nome.');
  if (!emailValido_(email)) throw new ErroUsuario('Digite um e-mail válido, como nome@email.com.');
  if (senha.length < 6 || senha.length > 100) throw new ErroUsuario('A senha precisa ter pelo menos 6 caracteres.');
  if (req.aceite !== true) throw new ErroUsuario('Para continuar, aceite o uso dos seus dados para o agendamento.');

  return comLock_(function () {
    const sh = aba_('Clientes');
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) if (normEmail_(rows[i][2]) === email) throw new ErroUsuario('Esse e-mail já tem cadastro. Use "Entrar".', 'EMAIL_EXISTE');
    const salt = novoSegredo_();
    const token = novoSegredo_() + novoSegredo_();
    const agora = agora_();
    sh.appendRow(['C' + novoSegredo_().slice(0, 10), nome, email, salt, hashSenha_(senha, salt), carimbo_(agora), carimbo_(agora),
      sha256_(token), String(agora.getTime() + TOKEN_DIAS * 864e5)]);
    return { token: token, cliente: { nome: nome, email: email } };
  });
}

function acaoEntrar_(req) {
  const email = normEmail_(req.email), senha = String(req.senha || '');
  if (!emailValido_(email) || !senha) throw new ErroUsuario('Digite seu e-mail e sua senha.');
  const cache = CacheService.getScriptCache(), chave = 'falha:' + sha256_(email).slice(0, 32);
  const falhas = Number(cache.get(chave) || 0);
  if (falhas >= MAX_FALHAS_LOGIN) throw new ErroUsuario('Muitas tentativas. Espere 15 minutos e tente de novo.');

  return comLock_(function () {
    const sh = aba_('Clientes');
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (normEmail_(rows[i][2]) !== email) continue;
      if (!igual_(hashSenha_(senha, String(rows[i][3])), String(rows[i][4]))) break;
      cache.remove(chave);
      const token = novoSegredo_() + novoSegredo_();
      sh.getRange(i + 1, 8, 1, 2).setValues([[sha256_(token), String(agora_().getTime() + TOKEN_DIAS * 864e5)]]);
      return { token: token, cliente: { nome: String(rows[i][1]), email: email } };
    }
    cache.put(chave, String(falhas + 1), 900);
    throw new ErroUsuario('E-mail ou senha incorretos.');
  });
}

function acaoSair_(req) {
  const cli = autenticar_(req.token, true);
  if (cli) aba_('Clientes').getRange(cli.linha, 8, 1, 2).setValues([['', '']]);
  return {};
}

function acaoMeus_(req) {
  const cli = autenticar_(req.token);
  const c = lerConfig_();
  const hoje = hoje_();
  const lista = lerAgenda_().filter(function (a) {
    return a.clienteId === cli.id && a.status === 'confirmado' && a.data >= hoje && inicio_(a) > agora_();
  }).sort(function (a, b) { return (a.data + a.hora).localeCompare(b.data + b.hora); });
  return { cliente: { nome: cli.nome, email: cli.email }, prazoHoras: c.p.prazoHoras, agendamentos: lista.map(function (a) { return publico_(a, c); }) };
}

function acaoAgendar_(req) {
  const cli = autenticar_(req.token);
  const c = lerConfig_();
  const unidade = exigir_(c.unidades, req.unidade, 'Unidade');
  const servico = exigir_(c.servicos, req.servico, 'Serviço');
  const barbeiro = exigir_(c.barbeiros, req.barbeiro, 'Barbeiro');
  if (barbeiro.unidade !== unidade.id) throw new ErroUsuario('Esse barbeiro não atende nessa unidade.');
  const data = String(req.data || ''), hora = String(req.hora || '');

  const novo = comLock_(function () {
    const ag = lerAgenda_();
    const ativos = ag.filter(function (a) { return a.clienteId === cli.id && a.status === 'confirmado' && inicio_(a) > agora_(); });
    if (ativos.length >= c.p.maxAtivos) throw new ErroUsuario('Você já tem ' + ativos.length + ' horários marcados. Cancele um para agendar outro.');
    exigirLivre_(c, unidade.id, barbeiro.id, data, hora, ag, null);
    const agora = carimbo_(agora_());
    const a = { id: 'A' + novoSegredo_().slice(0, 10), criadoEm: agora, unidade: unidade.id, servico: servico.id, barbeiro: barbeiro.id,
      data: data, hora: hora, clienteId: cli.id, cliente: cli.nome, email: cli.email, status: 'confirmado', lembrete: '',
      atualizadoEm: agora, historico: agora + ' agendado' };
    aba_('Agenda').appendRow(linhaAgenda_(a));
    return a;
  });
  enviarEmail_('confirmado', novo, c);
  return { agendamento: publico_(novo, c) };
}

function acaoRemarcar_(req) {
  const cli = autenticar_(req.token);
  const c = lerConfig_();
  const data = String(req.data || ''), hora = String(req.hora || '');
  const res = comLock_(function () {
    const ag = lerAgenda_();
    const a = meuAgendamento_(ag, req.id, cli);
    exigirPrazo_(a, c);
    const barbeiro = exigir_(c.barbeiros, req.barbeiro || a.barbeiro, 'Barbeiro');
    if (barbeiro.unidade !== a.unidade) throw new ErroUsuario('Escolha um barbeiro da mesma unidade.');
    if (barbeiro.id === a.barbeiro && data === a.data && hora === a.hora) throw new ErroUsuario('Esse já é o seu horário atual.');
    exigirLivre_(c, a.unidade, barbeiro.id, data, hora, ag, a.id);
    const antes = { data: a.data, hora: a.hora, barbeiro: a.barbeiro };
    const agora = carimbo_(agora_());
    a.barbeiro = barbeiro.id; a.data = data; a.hora = hora; a.lembrete = ''; a.atualizadoEm = agora;
    a.historico = (a.historico ? a.historico + ' | ' : '') + agora + ' remarcado de ' + antes.data + ' ' + antes.hora + ' (' + antes.barbeiro + ')';
    aba_('Agenda').getRange(a.linha, 1, 1, HEAD.Agenda.length).setValues([linhaAgenda_(a)]);
    return { a: a, antes: antes };
  });
  enviarEmail_('remarcado', res.a, c, res.antes);
  return { agendamento: publico_(res.a, c) };
}

function acaoCancelar_(req) {
  const cli = autenticar_(req.token);
  const c = lerConfig_();
  const a = comLock_(function () {
    const ag = lerAgenda_();
    const a = meuAgendamento_(ag, req.id, cli);
    exigirPrazo_(a, c);
    const agora = carimbo_(agora_());
    a.status = 'cancelado'; a.atualizadoEm = agora;
    a.historico = (a.historico ? a.historico + ' | ' : '') + agora + ' cancelado pelo cliente';
    aba_('Agenda').getRange(a.linha, 1, 1, HEAD.Agenda.length).setValues([linhaAgenda_(a)]);
    return a;
  });
  enviarEmail_('cancelado', a, c);
  return { cancelado: a.id };
}

/* ===================================================================== */
/* LEMBRETES (gatilho a cada 5 minutos)                                   */
/* ===================================================================== */

function enviarLembretes() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    const c = lerConfig_();
    const janela = c.p.lembreteMin * 60000;
    const agora = agora_().getTime();
    const sh = aba_('Agenda');
    lerAgenda_().forEach(function (a) {
      if (a.status !== 'confirmado' || a.lembrete) return;
      const falta = inicio_(a).getTime() - agora;
      if (falta > janela) return;
      let marca;
      if (falta <= 0) marca = 'não enviado (horário passou)';
      else if (inicio_(a).getTime() - lerCarimbo_(a.atualizadoEm) <= janela) marca = 'dispensado (marcado em cima da hora)';
      else marca = enviarEmail_('lembrete', a, c) ? carimbo_(agora_()) : '';
      if (marca) sh.getRange(a.linha, 12).setValue(marca);   // coluna "Lembrete enviado": impede duplicado
    });
  } finally {
    lock.releaseLock();
  }
}

/* ===================================================================== */
/* AGENDA E HORÁRIOS                                                      */
/* ===================================================================== */

function slotsDoDia_(c, unidadeId, barbeiroId, data, ag, bl, ignorarId) {
  const p = c.p, out = [];
  const limite = agora_().getTime() + p.antecedenciaMin * 60000;
  const ocupados = {};
  ag.forEach(function (a) {
    if (a.status === 'confirmado' && a.barbeiro === barbeiroId && a.data === data && a.id !== ignorarId) ocupados[a.hora] = true;
  });
  for (let m = min_(p.abertura); m + p.intervalo <= min_(p.fechamento); m += p.intervalo) {
    const fim = m + p.intervalo;
    if (p.almocoInicio && p.almocoFim && m < min_(p.almocoFim) && fim > min_(p.almocoInicio)) continue;
    if (bloqueado_(bl, unidadeId, barbeiroId, data, m, fim)) continue;
    const hora = hhmm_(m);
    const livre = !ocupados[hora] && dataHora_(data, hora).getTime() >= limite;
    out.push({ hora: hora, livre: livre });
  }
  return out;
}

function exigirLivre_(c, unidadeId, barbeiroId, data, hora, ag, ignorarId) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(data) || !/^\d{2}:\d{2}$/.test(hora)) throw new ErroUsuario('Escolha um dia e um horário.');
  if (proximosDias_(c).indexOf(data) < 0 || c.p.dias.indexOf(diaSemana_(data)) < 0) throw new ErroUsuario('Esse dia não está disponível para agendamento.');
  const slot = slotsDoDia_(c, unidadeId, barbeiroId, data, ag, lerBloqueios_(), ignorarId).filter(function (s) { return s.hora === hora; })[0];
  if (!slot || !slot.livre) throw new ErroUsuario('Esse horário acabou de ser ocupado. Escolha outro.', 'HORARIO_OCUPADO');
}

function exigirPrazo_(a, c) {
  if (inicio_(a).getTime() - agora_().getTime() < c.p.prazoHoras * 3600000) {
    throw new ErroUsuario('Remarcações e cancelamentos só até ' + c.p.prazoHoras + ' h antes do horário. Fale com a unidade.', 'FORA_DO_PRAZO');
  }
}

function meuAgendamento_(ag, id, cli) {
  const a = ag.filter(function (x) { return x.id === String(id || ''); })[0];
  if (!a || a.clienteId !== cli.id) throw new ErroUsuario('Agendamento não encontrado.');
  if (a.status !== 'confirmado') throw new ErroUsuario('Esse agendamento já foi cancelado.');
  return a;
}

function bloqueado_(bl, unidadeId, barbeiroId, data, ini, fim) {
  return bl.some(function (b) {
    if (b.unidade && b.unidade !== '*' && b.unidade !== unidadeId) return false;
    if (b.barbeiro && b.barbeiro !== '*' && b.barbeiro !== barbeiroId) return false;
    if (b.data && b.data !== data) return false;
    const bi = b.inicio ? min_(b.inicio) : 0, bf = b.fim ? min_(b.fim) : 24 * 60;
    return ini < bf && fim > bi;
  });
}

function proximosDias_(c) {
  const out = [], hoje = hoje_();
  for (let i = 0; i < c.p.diasAgenda; i++) out.push(somaDias_(hoje, i));
  return out;
}

function publico_(a, c) {
  const u = achar_(c.unidades, a.unidade) || { nome: a.unidade, endereco: '' };
  const b = achar_(c.barbeiros, a.barbeiro) || { nome: a.barbeiro };
  const s = achar_(c.servicos, a.servico) || { nome: a.servico, duracao: c.p.intervalo };
  return {
    id: a.id, data: a.data, hora: a.hora, unidade: a.unidade, unidadeNome: u.nome, endereco: u.endereco,
    barbeiro: a.barbeiro, barbeiroNome: b.nome, servico: a.servico, servicoNome: s.nome, duracao: s.duracao,
    podeAlterar: inicio_(a).getTime() - agora_().getTime() >= c.p.prazoHoras * 3600000
  };
}

/* ===================================================================== */
/* LEITURA DA PLANILHA                                                    */
/* ===================================================================== */

function planilha_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Rode a função setup() antes de usar o Web App.');
  return ss;
}

function aba_(nome) {
  const sh = planilha_().getSheetByName(nome);
  if (!sh) throw new Error('Aba "' + nome + '" não encontrada. Rode setup().');
  return sh;
}

function lerConfig_() {
  const rows = aba_('Config').getDataRange().getValues().slice(1);
  const ativo = function (r) { return !/^(n|não|nao|false|0)$/i.test(String(r[6]).trim()); };
  const par = {};
  const c = { unidades: [], barbeiros: [], servicos: [] };
  rows.forEach(function (r) {
    const tipo = String(r[0]).trim().toLowerCase(), id = String(r[1]).trim();
    if (!tipo || !id) return;
    if (tipo === 'parametro') { par[id] = texto_(r[3]); return; }
    if (!ativo(r)) return;
    if (tipo === 'unidade') c.unidades.push({ id: id, nome: texto_(r[2]), endereco: texto_(r[3]), telefone: texto_(r[4]) });
    if (tipo === 'barbeiro') c.barbeiros.push({ id: id, nome: texto_(r[2]), unidade: String(r[5]).trim() });
    if (tipo === 'servico') c.servicos.push({ id: id, nome: texto_(r[2]), duracao: Number(r[3]) || 30 });
  });
  const num = function (k, d) { const n = Number(par[k]); return isFinite(n) && par[k] !== '' && par[k] != null ? n : d; };
  c.p = {
    abertura: hora_(par.abertura) || '09:00',
    fechamento: hora_(par.fechamento) || '19:00',
    almocoInicio: hora_(par.almoco_inicio),
    almocoFim: hora_(par.almoco_fim),
    dias: String(par.dias_funcionamento || '2,3,4,5,6').split(/[,;\s]+/).filter(String).map(Number),
    intervalo: num('intervalo_min', 30),
    prazoHoras: num('prazo_alteracao_horas', 2),
    lembreteMin: num('lembrete_min', 30),
    diasAgenda: num('dias_agenda', 7),
    antecedenciaMin: num('antecedencia_min', 15),
    maxAtivos: num('max_ativos', 3)
  };
  return c;
}

function lerAgenda_() {
  const rows = aba_('Agenda').getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    out.push({ linha: i + 1, id: String(r[0]), criadoEm: texto_(r[1]), unidade: String(r[2]), servico: String(r[3]), barbeiro: String(r[4]),
      data: data_(r[5]), hora: hora_(r[6]), clienteId: String(r[7]), cliente: texto_(r[8]), email: String(r[9]),
      status: String(r[10]).trim().toLowerCase(), lembrete: texto_(r[11]), atualizadoEm: texto_(r[12]), historico: texto_(r[13]) });
  }
  return out;
}

function linhaAgenda_(a) {
  return [a.id, a.criadoEm, a.unidade, a.servico, a.barbeiro, a.data, a.hora, a.clienteId, a.cliente, a.email,
    a.status, a.lembrete, a.atualizadoEm, a.historico];
}

function lerBloqueios_() {
  return aba_('Bloqueios').getDataRange().getValues().slice(1).filter(function (r) { return r.join('') !== ''; }).map(function (r) {
    return { unidade: String(r[0]).trim(), barbeiro: String(r[1]).trim(), data: data_(r[2]), inicio: hora_(r[3]), fim: hora_(r[4]) };
  });
}

/* ===================================================================== */
/* CLIENTES, SENHAS E SESSÃO                                              */
/* ===================================================================== */

function autenticar_(token, opcional) {
  const t = String(token || '');
  if (t.length >= 32) {
    const h = sha256_(t), agora = agora_().getTime();
    const rows = aba_('Clientes').getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][7] && igual_(String(rows[i][7]), h) && Number(rows[i][8]) > agora) {
        return { linha: i + 1, id: String(rows[i][0]), nome: String(rows[i][1]), email: normEmail_(rows[i][2]) };
      }
    }
  }
  if (opcional) return null;
  throw new ErroUsuario('Sua sessão expirou. Entre de novo com seu e-mail e senha.', 'SESSAO');
}

/** SHA-256 iterado com salt por cliente + "pepper" guardado nas Propriedades do Script. */
function hashSenha_(senha, salt) {
  const pepper = PropertiesService.getScriptProperties().getProperty('PEPPER') || '';
  let h = sha256_(salt + ':' + pepper + ':' + senha);
  for (let i = 0; i < HASH_ROUNDS; i++) h = sha256_(h + salt);
  return h;
}

function sha256_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

function igual_(a, b) {
  if (a.length !== b.length) return false;
  let d = 0;
  for (let i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}

function novoSegredo_() { return Utilities.getUuid().replace(/-/g, ''); }

/* ===================================================================== */
/* E-MAILS                                                                */
/* ===================================================================== */

function replyTo_() {
  return PropertiesService.getScriptProperties().getProperty('REPLY_TO') || Session.getEffectiveUser().getEmail();
}

function enviarEmail_(tipo, a, c, antes) {
  const pub = publico_(a, c);
  const primeiro = String(a.cliente).split(' ')[0];
  const quando = dataLonga_(a.data) + ' às ' + a.hora;
  const titulos = {
    confirmado: ['Agendamento confirmado', 'Seu horário está marcado, ' + primeiro + '.'],
    remarcado:  ['Horário remarcado', 'Seu horário mudou, ' + primeiro + '. Confira os novos dados.'],
    cancelado:  ['Agendamento cancelado', 'Cancelamos o horário abaixo, ' + primeiro + '. Quando quiser, é só marcar de novo.'],
    lembrete:   ['Seu horário é daqui a ' + c.p.lembreteMin + ' minutos', primeiro + ', te esperamos daqui a pouco.']
  };
  const t = titulos[tipo];
  const linhas = [
    ['Cliente', a.cliente], ['Unidade', pub.unidadeNome + (pub.endereco ? ' — ' + pub.endereco : '')],
    ['Serviço', pub.servicoNome + ' (' + pub.duracao + ' min)'], ['Barbeiro', pub.barbeiroNome], ['Data e hora', quando]
  ];
  if (antes) linhas.push(['Antes', dataLonga_(antes.data) + ' às ' + antes.hora]);
  const site = PropertiesService.getScriptProperties().getProperty('SITE_URL');
  const rodape = tipo === 'cancelado' ? '' : 'Remarcações e cancelamentos pelo site, em "Meus agendamentos", até ' + c.p.prazoHoras + ' h antes do horário.';

  const html = '<div style="font-family:Arial,sans-serif;max-width:520px;margin:auto;color:#1a1207">' +
    '<div style="background:#0c0a08;color:#d9a441;padding:18px 22px;font-size:20px;font-weight:bold;letter-spacing:1px">' + SHOP_NAME + '</div>' +
    '<div style="padding:22px;border:1px solid #e6dccb;border-top:0">' +
    '<h2 style="margin:0 0 6px;font-size:20px">' + esc_(t[0]) + '</h2><p style="margin:0 0 16px;color:#555">' + esc_(t[1]) + '</p>' +
    '<table style="border-collapse:collapse;width:100%;font-size:14px">' + linhas.map(function (l) {
      return '<tr><td style="padding:7px 10px 7px 0;color:#777;width:110px;vertical-align:top">' + esc_(l[0]) + '</td><td style="padding:7px 0;font-weight:bold">' + esc_(l[1]) + '</td></tr>';
    }).join('') + '</table>' +
    (rodape ? '<p style="font-size:13px;color:#555;margin:18px 0 0">' + esc_(rodape) + '</p>' : '') +
    (site ? '<p style="margin:18px 0 0"><a href="' + esc_(site) + '#meus-agendamentos" style="background:#d9a441;color:#1a1207;padding:10px 16px;border-radius:20px;text-decoration:none;font-weight:bold;font-size:13px">Meus agendamentos</a></p>' : '') +
    '<p style="font-size:12px;color:#999;margin:22px 0 0">Responda este e-mail se precisar falar com a barbearia.</p></div></div>';
  const texto = t[0] + '\n\n' + t[1] + '\n\n' + linhas.map(function (l) { return l[0] + ': ' + l[1]; }).join('\n') + (rodape ? '\n\n' + rodape : '') + (site ? '\n\n' + site + '#meus-agendamentos' : '');

  try {
    MailApp.sendEmail({ to: a.email, subject: t[0] + ' · ' + SHOP_NAME, body: texto, htmlBody: html, name: SHOP_NAME, replyTo: replyTo_() });
    return true;
  } catch (err) {
    console.error('Falha ao enviar e-mail (' + tipo + ', ' + a.id + '): ' + err);
    return false;
  }
}

/* ===================================================================== */
/* UTILITÁRIOS                                                            */
/* ===================================================================== */

function comLock_(fn) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { throw new ErroUsuario('Muitos acessos ao mesmo tempo. Tente de novo em alguns segundos.'); }
  try { return fn(); } finally { lock.releaseLock(); }
}

function agora_() { return new Date(); }
function tz_() { return Session.getScriptTimeZone() || 'America/Sao_Paulo'; }
function hoje_() { return Utilities.formatDate(agora_(), tz_(), 'yyyy-MM-dd'); }
function carimbo_(d) { return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd HH:mm:ss'); }
function lerCarimbo_(s) { return s ? Utilities.parseDate(String(s), tz_(), 'yyyy-MM-dd HH:mm:ss').getTime() : 0; }
function dataHora_(data, hora) { return Utilities.parseDate(data + ' ' + hora, tz_(), 'yyyy-MM-dd HH:mm'); }
function inicio_(a) { return dataHora_(a.data, a.hora); }
function diaSemana_(data) { return new Date(data + 'T12:00:00Z').getUTCDay(); }
function somaDias_(data, n) { const d = new Date(data + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
function min_(hhmm) { const p = String(hhmm).split(':'); return Number(p[0]) * 60 + Number(p[1] || 0); }
function hhmm_(m) { return ('0' + Math.floor(m / 60)).slice(-2) + ':' + ('0' + (m % 60)).slice(-2); }
function texto_(v) { return v instanceof Date ? carimbo_(v) : String(v == null ? '' : v).trim(); }
function data_(v) { return v instanceof Date ? Utilities.formatDate(v, tz_(), 'yyyy-MM-dd') : String(v || '').trim(); }
function hora_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'HH:mm');
  const m = String(v || '').trim().match(/^(\d{1,2})(?::|h)?(\d{2})?$/);
  return m ? hhmm_(Number(m[1]) * 60 + Number(m[2] || 0)) : '';
}
function normEmail_(v) { return String(v || '').trim().toLowerCase(); }
function emailValido_(e) { return e.length <= 120 && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e); }
function achar_(lista, id) { return lista.filter(function (x) { return x.id === id; })[0] || null; }
function exigir_(lista, id, rotulo) { const x = achar_(lista, String(id || '')); if (!x) throw new ErroUsuario(rotulo + ' inválido(a).'); return x; }
function esc_(s) { return String(s).replace(/[&<>"']/g, function (ch) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]; }); }
function dataLonga_(data) {
  const dias = ['domingo', 'segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira', 'sábado'];
  const p = data.split('-');
  return dias[diaSemana_(data)] + ', ' + p[2] + '/' + p[1] + '/' + p[0];
}
