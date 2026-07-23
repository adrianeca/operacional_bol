// =============================================================================
// CONFIGURAÇÕES — Webapp Controle Operacional BOL
// =============================================================================

const OPERACIONAL_SHEET_ID = '1I8nss8jH0sv8qCmagWI2bITPEW2JWhi5OUsjVeBA_-M';
const HUB_SS_ID  = '1eZPbzhzjhjHoPwMhAW5YvOZgYiAvlTYc07dRan6Lyoc'; // mesmo Hub do app de Horas
const MEU_ACESSO = 'operacionalbol'; // precisa estar liberado na coluna ACESSOS da aba SESSOES do Hub
const HUB_URL     = 'https://script.google.com/a/macros/brasas.com/s/AKfycbyF7BArYMYFtcQY7_4RTGGPw89yNohAjR7eGptItP-EsnWhNfiZR2ISRaHdAkwlLSlr/exec';

const ABA_EVENTOS         = 'EVENTOS';
const ABA_ROTINAS         = 'ROTINAS';
const ABA_ROTINAS_LOG     = 'ROTINAS_LOG';
const ABA_ENTREGAS        = 'ENTREGAS';
const ABA_HORARIOS        = 'HORARIOS';
const ABA_ESCALA_SABADO   = 'ESCALA_SABADO';
const ABA_AJUSTES_HORARIO = 'AJUSTES_HORARIO';

const TZ = Session.getScriptTimeZone();

// Normaliza texto para comparação: minúsculo, sem acento, sem espaços nas bordas
function norm_(s) {
  return String(s || '').trim().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '');
}

function fmtDataHora_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'dd/MM/yyyy HH:mm');
  return String(v);
}

function fmtData_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v);
}

function getOpSS_() {
  return SpreadsheetApp.openById(OPERACIONAL_SHEET_ID);
}

// =============================================================================
// ENTRY POINT
// =============================================================================

function doGet(e) {
  const token = (e && e.parameter && e.parameter.s) ? e.parameter.s : '';
  const tmpl = HtmlService.createTemplateFromFile('Index');
  tmpl.token = token;
  return tmpl.evaluate()
    .setTitle('Controle Operacional BOL — BRASAS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =============================================================================
// AUTENTICAÇÃO (mesmo padrão do app de Horas de Professores)
// =============================================================================

// Colunas SESSOES: TOKEN(0)|EMAIL(1)|NOME(2)|ROLE(3)|UNIDADE(4)|CRIADO_EM(5)|EXPIRA_EM(6)|ACESSOS(7)

function getUserFromHub(token) {
  return requireUser_(token);
}

function getSessionUser_(token) {
  if (!token) return null;
  try {
    const ss       = SpreadsheetApp.openById(HUB_SS_ID);
    const sesSheet = ss.getSheetByName('SESSOES');
    if (!sesSheet) return null;

    const tok   = String(token).trim();
    const found = sesSheet.getRange(1, 1, sesSheet.getLastRow(), 1)
      .createTextFinder(tok).matchEntireCell(true).findNext();
    if (!found) return null;

    const row = sesSheet.getRange(found.getRow(), 1, 1, 8).getValues()[0];

    if (row[6] && new Date(row[6]) < new Date()) return null; // expirado

    const email = String(row[1] || '').trim().toLowerCase();
    if (!email) return null;

    const acessos = String(row[7] || '').toLowerCase()
      .split(',').map(function(a) { return a.trim(); });
    if (acessos.indexOf(MEU_ACESSO) === -1) {
      throw new Error('Você não tem permissão para acessar o Controle Operacional BOL. Contacte o administrador.');
    }

    return {
      email: email,
      nome:  String(row[2] || '').trim(),
      role:  String(row[3] || '').trim().toLowerCase()
    };
  } catch (e) {
    if (e.message && e.message.indexOf('permissão') !== -1) throw e;
    Logger.log('getSessionUser_: ' + e);
    return null;
  }
}

function requireUser_(token) {
  const user = getSessionUser_(token);
  if (!user) throw new Error('Sessão inválida ou expirada. Acesse novamente pelo Hub.');
  return user;
}

// =============================================================================
// DASHBOARD
// =============================================================================

function getDashboard(token) {
  requireUser_(token);

  const hoje    = new Date();
  const hojeStr = fmtData_(hoje);
  const diaSemana = hoje.getDay(); // 0=domingo ... 6=sábado

  const horarios    = getHorarios(token).filter(function(h) { return h.ativo; });
  const ajustesHoje  = getAjustesHorario(token).filter(function(a) { return a.data === hojeStr; });

  let quemTrabalha = [];
  if (diaSemana === 6) {
    const escalaHoje = getEscalaSabado(token).filter(function(e) { return e.data === hojeStr; })[0];
    if (escalaHoje) {
      quemTrabalha = [escalaHoje.pessoa1, escalaHoje.pessoa2, escalaHoje.pessoa3]
        .filter(Boolean)
        .map(function(nome) { return { nome: nome, horario: 'Escala de sábado', ajuste: escalaHoje.observacao || '' }; });
    }
  } else if (diaSemana !== 0) {
    const sexta = diaSemana === 5;
    quemTrabalha = horarios.map(function(h) {
      const ajuste = ajustesHoje.filter(function(a) {
        return a.atividade.toLowerCase().indexOf(h.nome.toLowerCase()) !== -1;
      })[0];
      return {
        nome: h.nome,
        horario: sexta ? h.horarioSexta : h.horarioPadrao,
        ajuste: ajuste ? (ajuste.status + ': ' + ajuste.atividade) : ''
      };
    });
  }

  const eventos  = getEventos(token);
  const hojeMs   = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate()).getTime();
  const limiteMs = hojeMs + 7 * 86400000;
  const eventosSemana = eventos.filter(function(ev) {
    if (!ev.dataInicio) return false;
    const ini = new Date(ev.dataInicio + 'T00:00:00').getTime();
    const fim = ev.dataFim ? new Date(ev.dataFim + 'T00:00:00').getTime() : ini;
    return fim >= hojeMs && ini <= limiteMs;
  });

  const mesAtual = hoje.getMonth() + 1, anoAtual = hoje.getFullYear();
  const entregasPendentes = getEntregas(token).filter(function(e) {
    return e.mes === mesAtual && e.ano === anoAtual && e.status !== 'NO PRAZO';
  });

  return { hoje: hojeStr, quemTrabalha: quemTrabalha, eventosSemana: eventosSemana, entregasPendentes: entregasPendentes };
}

// =============================================================================
// EVENTOS (Calendário)
// =============================================================================

function getEventosSheet_() {
  const sheet = getOpSS_().getSheetByName(ABA_EVENTOS);
  if (!sheet) throw new Error('Aba "' + ABA_EVENTOS + '" não encontrada.');
  const header = sheet.getRange(1, 1, 1, 8).getValues()[0];
  if (!header[7]) {
    sheet.getRange(1, 8).setValue('ID');
    const last = sheet.getLastRow();
    if (last > 1) {
      const ids = sheet.getRange(2, 8, last - 1, 1).getValues();
      sheet.getRange(2, 8, ids.length, 1).setValues(ids.map(function(r) { return [r[0] || Utilities.getUuid()]; }));
    }
  }
  return sheet;
}

function eventoFromRow_(r) {
  return {
    id: String(r[7]),
    dataInicio: fmtData_(r[0]),
    dataFim: fmtData_(r[1]),
    titulo: String(r[2] || '').trim(),
    categoria: String(r[3] || '').trim(),
    observacao: String(r[4] || '').trim(),
    criadoPor: String(r[5] || '').trim(),
    criadoEm: fmtDataHora_(r[6])
  };
}

function getEventos(token) {
  requireUser_(token);
  const rows = getEventosSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][2]) continue; // sem título
    out.push(eventoFromRow_(rows[i]));
  }
  out.sort(function(a, b) { return (a.dataInicio || '9999').localeCompare(b.dataInicio || '9999'); });
  return out;
}

function saveEvento(token, evento) {
  const user  = requireUser_(token);
  const sheet = getEventosSheet_();

  const titulo = String(evento.titulo || '').trim();
  if (!titulo) throw new Error('Informe o título do evento.');
  const dataInicio = evento.dataInicio ? new Date(evento.dataInicio + 'T00:00:00') : '';
  const dataFim     = evento.dataFim ? new Date(evento.dataFim + 'T00:00:00') : '';
  const categoria   = String(evento.categoria || 'EVENTO').trim();
  const observacao  = String(evento.observacao || '').trim();

  if (evento.id) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][7]) === String(evento.id)) {
        sheet.getRange(i + 1, 1, 1, 5).setValues([[dataInicio, dataFim, titulo, categoria, observacao]]);
        return getEventos(token);
      }
    }
    throw new Error('Evento não encontrado.');
  }

  sheet.appendRow([dataInicio, dataFim, titulo, categoria, observacao, user.email, new Date(), Utilities.getUuid()]);
  return getEventos(token);
}

function deleteEvento(token, id) {
  requireUser_(token);
  const sheet = getEventosSheet_();
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][7]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return getEventos(token);
}

// =============================================================================
// ROTINAS (checklist por período, com histórico em ROTINAS_LOG)
// =============================================================================

function getRotinasSheet_() {
  const sheet = getOpSS_().getSheetByName(ABA_ROTINAS);
  if (!sheet) throw new Error('Aba "' + ABA_ROTINAS + '" não encontrada.');
  const header = sheet.getRange(1, 1, 1, 5).getValues()[0];
  if (!header[4]) {
    sheet.getRange(1, 5).setValue('ID');
    const last = sheet.getLastRow();
    if (last > 1) {
      const ids = sheet.getRange(2, 5, last - 1, 1).getValues();
      sheet.getRange(2, 5, ids.length, 1).setValues(ids.map(function(r) { return [r[0] || Utilities.getUuid()]; }));
    }
  }
  return sheet;
}

function getRotinasLogSheet_() {
  const ss = getOpSS_();
  let sheet = ss.getSheetByName(ABA_ROTINAS_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_ROTINAS_LOG);
    sheet.appendRow(['Tarefa_ID', 'Periodo', 'Concluido_Em', 'Concluido_Por']);
  }
  return sheet;
}

// Número aproximado da semana do ano (não é ISO 8601 exato, mas é suficiente
// para agrupar conclusões semanais de forma estável dentro do mesmo ano)
function numeroSemana_(data) {
  const inicioAno = new Date(data.getFullYear(), 0, 1);
  const dias = Math.floor((data - inicioAno) / 86400000);
  return Math.floor(dias / 7) + 1;
}

// Calcula a chave do período corrente conforme a frequência da tarefa —
// cada frequência tem sua própria "janela" de conclusão.
function periodoAtual_(frequencia, data) {
  data = data || new Date();
  const freq = norm_(frequencia).toUpperCase();
  const ano = data.getFullYear();

  if (freq.indexOf('DIARIO') !== -1) return Utilities.formatDate(data, TZ, 'yyyy-MM-dd');
  if (freq.indexOf('SEMANAL') !== -1 && freq.indexOf('QUINZENAL') === -1) {
    return ano + '-W' + ('0' + numeroSemana_(data)).slice(-2);
  }
  if (freq.indexOf('QUINZENAL') !== -1) return ano + '-Q' + ('0' + Math.ceil(numeroSemana_(data) / 2)).slice(-2);
  if (freq.indexOf('MENSAL') !== -1)    return Utilities.formatDate(data, TZ, 'yyyy-MM');
  if (freq.indexOf('TRIMESTRAL') !== -1) return ano + '-T' + (Math.floor(data.getMonth() / 3) + 1);
  if (freq.indexOf('SEMESTRAL') !== -1)  return ano + '-S' + (Math.floor(data.getMonth() / 6) + 1);
  if (freq.indexOf('ANUAL') !== -1)      return String(ano);
  return Utilities.formatDate(data, TZ, 'yyyy-MM-dd');
}

function getRotinas(token) {
  requireUser_(token);
  const rows    = getRotinasSheet_().getDataRange().getValues();
  const logRows = getRotinasLogSheet_().getDataRange().getValues();
  const agora   = new Date();

  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const tarefa = String(r[1] || '').trim();
    if (!tarefa) continue;
    if (String(r[3]).trim().toUpperCase() !== 'SIM') continue; // só ativas

    const id = String(r[4]);
    const frequencia = String(r[0]).trim();
    const periodo = periodoAtual_(frequencia, agora);

    let concluido = false, concluidoEm = '', concluidoPor = '';
    for (let j = 1; j < logRows.length; j++) {
      if (String(logRows[j][0]) === id && String(logRows[j][1]) === periodo) {
        concluido    = true;
        concluidoEm  = fmtDataHora_(logRows[j][2]);
        concluidoPor = String(logRows[j][3] || '').trim();
        break;
      }
    }

    out.push({
      id: id, frequencia: frequencia, tarefa: tarefa,
      responsavel: String(r[2] || '').trim(), periodo: periodo,
      concluido: concluido, concluidoEm: concluidoEm, concluidoPor: concluidoPor
    });
  }
  return out;
}

function toggleRotina(token, id, periodo, concluido) {
  const user  = requireUser_(token);
  const sheet = getRotinasLogSheet_();
  const rows  = sheet.getDataRange().getValues();

  let rowIdx = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id) && String(rows[i][1]) === String(periodo)) { rowIdx = i + 1; break; }
  }

  if (concluido) {
    if (rowIdx === -1) sheet.appendRow([id, periodo, new Date(), user.email]);
  } else if (rowIdx !== -1) {
    sheet.deleteRow(rowIdx);
  }

  return getRotinas(token);
}

function saveRotina(token, rotina) {
  requireUser_(token);
  const sheet = getRotinasSheet_();

  const frequencia = String(rotina.frequencia || '').trim().toUpperCase();
  const tarefa      = String(rotina.tarefa || '').trim();
  const responsavel = String(rotina.responsavel || 'Todas').trim();
  if (!frequencia || !tarefa) throw new Error('Informe a frequência e a tarefa.');

  if (rotina.id) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][4]) === String(rotina.id)) {
        sheet.getRange(i + 1, 1, 1, 3).setValues([[frequencia, tarefa, responsavel]]);
        return getRotinas(token);
      }
    }
    throw new Error('Rotina não encontrada.');
  }

  sheet.appendRow([frequencia, tarefa, responsavel, 'SIM', Utilities.getUuid()]);
  return getRotinas(token);
}

// Desativa a tarefa (não apaga a linha, para preservar o histórico em ROTINAS_LOG)
function deleteRotina(token, id) {
  requireUser_(token);
  const sheet = getRotinasSheet_();
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][4]) === String(id)) {
      sheet.getRange(i + 1, 4).setValue('NÃO');
      break;
    }
  }
  return getRotinas(token);
}

// Seed único: importa as tarefas individuais da aba "Rotina/Feedback" (por
// pessoa: Elaine, Bruna, Nayara, Rafaela, Rosana, Caren) para a aba ROTINAS,
// como frequência "INDIVIDUAL". Rode manualmente uma vez pelo editor do Apps
// Script (menu Executar → popularRotinasIndividuais). Pode rodar de novo sem
// medo: tarefas já existentes (mesma tarefa + mesma pessoa) não são duplicadas.
function popularRotinasIndividuais() {
  const sheet = getRotinasSheet_();
  const rows  = sheet.getDataRange().getValues();

  const existentes = new Set();
  for (let i = 1; i < rows.length; i++) {
    existentes.add(norm_(rows[i][1]) + '|' + norm_(rows[i][2]));
  }

  const tarefasPorPessoa = {
    'Elaine': [
      'Horário conferido Alice até o Bram',
      'Appsheet Controle de alunos - verificar cancelamentos',
      'Envio da posição de estoque (SophiA x New Choice)',
      'Planilha de certificados (todos os níveis)',
      'Lista de cancelados do mês para Pri e Ge (todo dia 01)',
      'Recibo de pagamento Ana Beatriz dos Santos Reis dia 1',
      'Declaração de pagamento/demonstrativo e frequência Bruno Ferraz Pais Bittencourt dia 02',
      'Declaração Tais Soares de Freitas Dias dia 11',
      'Declaração de pagamento Gregory Pontes Oliveira dia 14',
      'Luise Soares Vieira - Declaração Mensal de Frequência'
    ],
    'Bruna': [
      'Horário Conferido Nathalie até Olivia',
      'Contato com os alunos em TI (às terças)',
      'Neide Aparecida Sousa Freitas - NF-e Bol e Polo todo dia 6',
      'Sharlene Barga de Campos NF-e BOL todo o dia 10',
      'Declaração BOL - Laercio Ribeiro Souza dia 08',
      'Declaração de pagamento dia 20 Jessica Fernanda Veiga Lopes',
      'Lilian Odeli - Enviar NF-e Bol e Polo todo dia 5 de cada mês',
      'Declaração de pagamento dia 15 Maria Luiza dos Santos Itaborahy',
      'Natalia Salvador - Enviar NF-e BOL e polo, após o dia 23'
    ],
    'Nayara': [
      'Horário Conferido Hudson até Mary',
      'Relatório de Cias',
      'Relatório de Bolsistas',
      'My BRASAS - verificar alunos sem convite',
      'Private - cobrança mensal',
      'Caixa de entrada e-mail Admin',
      'Conferência mensal appsheet de matrículas'
    ],
    'Rafaela': [
      'Horário Conferido Peter até Taylor',
      'Planilha de segunda chamada - toda quarta',
      'Contato com TI - terça e quinta',
      'Juliana Azevedo Cunha - Declaração dia 01',
      'Andre Gomes Valle Nery - NF-es todo o dia 1',
      'Luciane F. de Castro de Abreu - Declaração de Pagamento dia 2',
      'Pedro Gomes dos Santos NFs todo dia 02',
      'Rozania Pereira - dia 05 (Polo Macaé)'
    ],
    'Rosana': [
      'Horário Conferido Bunny até Houston',
      'Daniel Vitor Silva Declaração e NF-e dia 05'
    ],
    'Caren': [
      'Igor Aniceto dos Santos - Enviar NFs por e-mail todo dia 12',
      'Juliana Matias da Silva - NFS por email (Aluna Split) dia 16',
      'Leandro de Lucas Mendes - Enviar NFS por email todo dia 18'
    ]
  };

  let adicionadas = 0;
  Object.keys(tarefasPorPessoa).forEach(function(pessoa) {
    tarefasPorPessoa[pessoa].forEach(function(tarefa) {
      const chave = norm_(tarefa) + '|' + norm_(pessoa);
      if (existentes.has(chave)) return; // já existe, evita duplicar
      sheet.appendRow(['INDIVIDUAL', tarefa, pessoa, 'SIM', Utilities.getUuid()]);
      existentes.add(chave);
      adicionadas++;
    });
  });

  Logger.log(adicionadas + ' rotinas individuais adicionadas.');
  return adicionadas;
}

// =============================================================================
// ENTREGAS
// =============================================================================

function getEntregasSheet_() {
  const sheet = getOpSS_().getSheetByName(ABA_ENTREGAS);
  if (!sheet) throw new Error('Aba "' + ABA_ENTREGAS + '" não encontrada.');
  return sheet;
}

function entregaFromRow_(r) {
  return {
    atividade: String(r[0] || '').trim(),
    setor: String(r[1] || '').trim(),
    prazo: String(r[2] || '').trim(),
    mes: Number(r[3]),
    ano: Number(r[4]),
    status: String(r[5] || '').trim(),
    atualizadoEm: fmtDataHora_(r[6]),
    atualizadoPor: String(r[7] || '').trim()
  };
}

function getEntregas(token) {
  requireUser_(token);
  const rows = getEntregasSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    out.push(entregaFromRow_(rows[i]));
  }
  return out;
}

function saveEntregaStatus(token, payload) {
  const user  = requireUser_(token);
  const sheet = getEntregasSheet_();

  const atividade = String(payload.atividade || '').trim();
  const mes = Number(payload.mes), ano = Number(payload.ano);
  const status = String(payload.status || '').trim();

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[0]).trim() === atividade && Number(r[3]) === mes && Number(r[4]) === ano) {
      sheet.getRange(i + 1, 6, 1, 3).setValues([[status, new Date(), user.email]]);
      return getEntregas(token);
    }
  }
  throw new Error('Entrega não encontrada para "' + atividade + '" (' + mes + '/' + ano + ').');
}

// =============================================================================
// SECRETARIA — Horários fixos + Escala de sábado + Ajustes de horário
// =============================================================================

function getHorariosSheet_() {
  const sheet = getOpSS_().getSheetByName(ABA_HORARIOS);
  if (!sheet) throw new Error('Aba "' + ABA_HORARIOS + '" não encontrada.');
  return sheet;
}

function getHorarios(token) {
  requireUser_(token);
  const rows = getHorariosSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    out.push({
      nome: String(rows[i][0]).trim(),
      horarioPadrao: String(rows[i][1] || '').trim(),
      horarioSexta: String(rows[i][2] || '').trim(),
      ativo: String(rows[i][3] || '').trim().toUpperCase() === 'ATIVO'
    });
  }
  return out;
}

function saveHorario(token, h) {
  requireUser_(token);
  const sheet = getHorariosSheet_();

  const nome = String(h.nome || '').trim();
  if (!nome) throw new Error('Informe o nome.');
  const chaveOriginal = norm_(h.nomeOriginal || nome);
  const linha = [nome, h.horarioPadrao || '', h.horarioSexta || '', h.ativo ? 'ATIVO' : 'INATIVO'];

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (norm_(rows[i][0]) === chaveOriginal) {
      sheet.getRange(i + 1, 1, 1, 4).setValues([linha]);
      return getHorarios(token);
    }
  }
  sheet.appendRow(linha);
  return getHorarios(token);
}

function getEscalaSabadoSheet_() {
  const sheet = getOpSS_().getSheetByName(ABA_ESCALA_SABADO);
  if (!sheet) throw new Error('Aba "' + ABA_ESCALA_SABADO + '" não encontrada.');
  return sheet;
}

function escalaFromRow_(r) {
  return {
    mes: Number(r[0]), ano: Number(r[1]),
    data: fmtData_(r[2]),
    pessoa1: String(r[3] || '').trim(),
    pessoa2: String(r[4] || '').trim(),
    pessoa3: String(r[5] || '').trim(),
    observacao: String(r[6] || '').trim()
  };
}

function getEscalaSabado(token) {
  requireUser_(token);
  const rows = getEscalaSabadoSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][2]) continue;
    out.push(escalaFromRow_(rows[i]));
  }
  out.sort(function(a, b) { return a.data.localeCompare(b.data); });
  return out;
}

function saveEscalaSabado(token, e) {
  requireUser_(token);
  const sheet = getEscalaSabadoSheet_();

  const dataStr = String(e.data || '').trim();
  if (!dataStr) throw new Error('Informe a data do sábado.');
  const chaveOriginal = String(e.dataOriginal || dataStr).trim();
  const data = new Date(dataStr + 'T00:00:00');
  const linha = [data.getMonth() + 1, data.getFullYear(), data, e.pessoa1 || '', e.pessoa2 || '', e.pessoa3 || '', e.observacao || ''];

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (fmtData_(rows[i][2]) === chaveOriginal) {
      sheet.getRange(i + 1, 1, 1, 7).setValues([linha]);
      return getEscalaSabado(token);
    }
  }
  sheet.appendRow(linha);
  return getEscalaSabado(token);
}

function getAjustesSheet_() {
  const sheet = getOpSS_().getSheetByName(ABA_AJUSTES_HORARIO);
  if (!sheet) throw new Error('Aba "' + ABA_AJUSTES_HORARIO + '" não encontrada.');
  const header = sheet.getRange(1, 1, 1, 9).getValues()[0];
  if (!header[7]) {
    sheet.getRange(1, 8).setValue('ID');
    const last = sheet.getLastRow();
    if (last > 1) {
      const ids = sheet.getRange(2, 8, last - 1, 1).getValues();
      sheet.getRange(2, 8, ids.length, 1).setValues(ids.map(function(r) { return [r[0] || Utilities.getUuid()]; }));
    }
  }
  if (!header[8]) sheet.getRange(1, 9).setValue('Tags');
  return sheet;
}

function ajusteFromRow_(r) {
  return {
    id: String(r[7]),
    atividade: String(r[0] || '').trim(),
    data: fmtData_(r[1]),
    status: String(r[2] || '').trim(),
    mes: Number(r[3]), ano: Number(r[4]),
    registradoPor: String(r[5] || '').trim(),
    registradoEm: fmtDataHora_(r[6]),
    tags: String(r[8] || '').trim()
  };
}

function getAjustesHorario(token) {
  requireUser_(token);
  const rows = getAjustesSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    out.push(ajusteFromRow_(rows[i]));
  }
  out.sort(function(a, b) { return (b.data || '').localeCompare(a.data || ''); });
  return out;
}

function saveAjusteHorario(token, a) {
  const user  = requireUser_(token);
  const sheet = getAjustesSheet_();

  const atividade = String(a.atividade || '').trim();
  const dataStr = String(a.data || '').trim();
  if (!atividade || !dataStr) throw new Error('Informe a atividade e a data.');
  const data = new Date(dataStr + 'T00:00:00');
  const status = String(a.status || '').trim();
  const tags = String(a.tags || '').trim();

  if (a.id) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][7]) === String(a.id)) {
        sheet.getRange(i + 1, 1, 1, 5).setValues([[atividade, data, status, data.getMonth() + 1, data.getFullYear()]]);
        sheet.getRange(i + 1, 9).setValue(tags);
        return getAjustesHorario(token);
      }
    }
    throw new Error('Ajuste não encontrado.');
  }

  sheet.appendRow([atividade, data, status, data.getMonth() + 1, data.getFullYear(), user.email, new Date(), Utilities.getUuid(), tags]);
  return getAjustesHorario(token);
}

function deleteAjusteHorario(token, id) {
  requireUser_(token);
  const sheet = getAjustesSheet_();
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][7]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return getAjustesHorario(token);
}

function getSecretaria(token) {
  requireUser_(token);
  return {
    horarios: getHorarios(token),
    escalaSabado: getEscalaSabado(token),
    ajustes: getAjustesHorario(token)
  };
}
