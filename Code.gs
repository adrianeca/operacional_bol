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
const ABA_HORAS_EXTRAS    = 'HORAS_EXTRAS';
const ABA_FERIAS          = 'FERIAS';
const ABA_FUNCIONARIOS    = 'FUNCIONARIOS';
const ABA_CHECKLIST_ADMISSAO = 'ADMISSAO_CHECKLIST';

const RH_SS_ID    = '1BDiPjv0FqRJp5EwcvLdYXVvEAWesvwdEgbhYdnTlqPY';
const RH_SHEET_GID = 566990656;

// Itens do checklist de admissão, na ordem em que devem aparecer na tela.
const CHECKLIST_ADMISSAO_ITENS = [
  'Carta ou Exame',
  'Sophia',
  'E-mail',
  'Grupos',
  'Liberação Google Slides e provas (coord)',
  'Acesso ao E-MAIL ADM',
  'Canal My BRASAS',
  'Aprovar ação no My BRASAS',
  'Agenda',
  'Acesso as provas',
  'Transferência de turmas',
  'Classroom',
  'Classroom antigo',
  'Drive',
  'Form horas extras e cancelados',
  'Coaching Session'
];

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

// O Sheets converte texto tipo "07:00" (uma hora só, sem " - ") pra um valor de
// hora/data internamente quando salvo via setValues; sem isso, String(v) vira
// "Sat Dec 30 1899 07:00:00 GMT...".
function fmtHora_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'HH:mm');
  return String(v).trim();
}

// Cacheado por execução: cada chamada do cliente pode acionar dezenas de leituras
// de abas diferentes, e reabrir a planilha por ID a cada uma delas sobrecarrega o
// serviço de Planilhas (erro RESOURCE_EXHAUSTED). Abrindo uma vez só por execução.
let _opSS_ = null;
function getOpSS_() {
  if (!_opSS_) _opSS_ = SpreadsheetApp.openById(OPERACIONAL_SHEET_ID);
  return _opSS_;
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

// Cacheados por execução: requireUser_ roda no início de toda função exposta,
// inclusive nas chamadas aninhadas (ex.: getDashboard chama getHorarios, que
// chama requireUser_ de novo). Sem isso, um único carregamento do painel abre
// a planilha do Hub e refaz a busca na aba SESSOES dezenas de vezes, o que
// também contribui para o erro RESOURCE_EXHAUSTED.
let _hubSS_ = null;
const _sessionUserCache_ = {};

function getSessionUser_(token) {
  if (!token) return null;
  const tokenKey = String(token).trim();
  if (_sessionUserCache_.hasOwnProperty(tokenKey)) return _sessionUserCache_[tokenKey];
  const user = getSessionUserUncached_(token); // deixa erro de permissão propagar sem cachear
  _sessionUserCache_[tokenKey] = user;
  return user;
}

function getSessionUserUncached_(token) {
  try {
    if (!_hubSS_) _hubSS_ = SpreadsheetApp.openById(HUB_SS_ID);
    const ss       = _hubSS_;
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
  const ajustesHoje  = getAjustesHorario(token).filter(function(a) {
    return a.data === hojeStr && (!a.status || a.status === 'Pendente');
  });

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

  const feriasPeriodos = getFerias(token);
  quemTrabalha = quemTrabalha.map(function(p) {
    const deFerias = feriasPeriodos.some(function(f) {
      return norm_(f.pessoa) === norm_(p.nome) && f.dataInicio <= hojeStr && f.dataFim >= hojeStr;
    });
    return deFerias ? Object.assign({}, p, { ferias: true }) : p;
  });

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

// Eventos "virtuais" derivados da aba FERIAS, para aparecerem no Calendário e
// no Dashboard sem duplicar dado (a aba FERIAS é a fonte da verdade). Têm id
// prefixado com "ferias-" e não podem ser editados/excluídos por aqui.
function eventosVirtuaisFerias_(token) {
  return getFerias(token).map(function(f) {
    return {
      id: 'ferias-' + f.id,
      dataInicio: f.dataInicio,
      dataFim: f.dataFim,
      titulo: 'Férias · ' + f.pessoa,
      categoria: 'FERIAS',
      observacao: f.observacao,
      criadoPor: f.criadoPor,
      criadoEm: f.criadoEm
    };
  });
}

function getEventos(token) {
  requireUser_(token);
  const rows = getEventosSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][2]) continue; // sem título
    // Férias não vêm mais de EVENTOS lançados manualmente — só da aba FERIAS
    // (eventosVirtuaisFerias_ abaixo), pra não duplicar no Calendário.
    if (String(rows[i][3]).trim().toUpperCase() === 'FERIAS') continue;
    out.push(eventoFromRow_(rows[i]));
  }
  out.push.apply(out, eventosVirtuaisFerias_(token));
  out.sort(function(a, b) { return (a.dataInicio || '9999').localeCompare(b.dataInicio || '9999'); });
  return out;
}

function saveEvento(token, evento) {
  const user  = requireUser_(token);
  const sheet = getEventosSheet_();

  if (evento.id && String(evento.id).indexOf('ferias-') === 0) {
    throw new Error('Este evento vem da página Férias — edite ou exclua por lá.');
  }

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
  if (String(id).indexOf('ferias-') === 0) {
    throw new Error('Este evento vem da página Férias — edite ou exclua por lá.');
  }
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
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return getRotinas(token);
}

function deleteFuncionario(token, nome) {
  requireUser_(token);
  const sheet = getFuncionariosSheet_();
  const rows  = sheet.getDataRange().getValues();
  const nomeNorm = norm_(nome);
  for (let i = 1; i < rows.length; i++) {
    if (norm_(rows[i][0]) === nomeNorm) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
  throw new Error('Funcionário não encontrado: ' + nome);
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
  const header = sheet.getRange(1, 1, 1, 14).getValues()[0];
  if (!header[7]) {
    sheet.getRange(1, 8).setValue('ID');
    const last = sheet.getLastRow();
    if (last > 1) {
      const ids = sheet.getRange(2, 8, last - 1, 1).getValues();
      sheet.getRange(2, 8, ids.length, 1).setValues(ids.map(function(r) { return [r[0] || Utilities.getUuid()]; }));
    }
  }
  if (!header[8])  sheet.getRange(1, 9).setValue('Tags');
  if (!header[9])  sheet.getRange(1, 10).setValue('Tipo');
  if (!header[10]) sheet.getRange(1, 11).setValue('Quem sai');
  if (!header[11]) sheet.getRange(1, 12).setValue('Quem cobre');
  if (!header[12]) sheet.getRange(1, 13).setValue('Horário de');
  if (!header[13]) sheet.getRange(1, 14).setValue('Horário para');
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
    tags: String(r[8] || '').trim(),
    tipo: String(r[9] || '').trim(),
    quemSai: String(r[10] || '').trim(),
    quemCobre: String(r[11] || '').trim(),
    horarioDe: fmtHora_(r[12]),
    horarioPara: fmtHora_(r[13])
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

  const dataStr = String(a.data || '').trim();
  const quemSai = String(a.quemSai || '').trim();
  const quemCobre = String(a.quemCobre || '').trim();
  if (!dataStr) throw new Error('Informe a data.');
  if (!quemSai && !quemCobre) throw new Error('Informe quem sai e/ou quem cobre.');
  const data = new Date(dataStr + 'T00:00:00');
  const atividade = String(a.atividade || '').trim();
  const status = String(a.status || '').trim();
  const tipo = String(a.tipo || '').trim();
  const horarioDe = String(a.horarioDe || '').trim();
  const horarioPara = String(a.horarioPara || '').trim();
  const tags = [quemSai, quemCobre].filter(function(x) { return x; }).join(' / ');

  if (a.id) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][7]) === String(a.id)) {
        sheet.getRange(i + 1, 1, 1, 5).setValues([[atividade, data, status, data.getMonth() + 1, data.getFullYear()]]);
        sheet.getRange(i + 1, 9, 1, 4).setValues([[tags, tipo, quemSai, quemCobre]]);
        // Colunas de horário precisam ser texto puro, senão o Sheets converte
        // um valor tipo "07:00" (sem " - ") num serial de hora automaticamente.
        sheet.getRange(i + 1, 13, 1, 2).setNumberFormat('@').setValues([[horarioDe, horarioPara]]);
        return getAjustesHorario(token);
      }
    }
    throw new Error('Ajuste não encontrado.');
  }

  sheet.appendRow([atividade, data, status, data.getMonth() + 1, data.getFullYear(), user.email, new Date(), Utilities.getUuid(), tags, tipo, quemSai, quemCobre, '', '']);
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 13, 1, 2).setNumberFormat('@').setValues([[horarioDe, horarioPara]]);
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

function getHorasExtrasSheet_() {
  const ss = getOpSS_();
  let sheet = ss.getSheetByName(ABA_HORAS_EXTRAS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_HORAS_EXTRAS);
    sheet.appendRow(['Pessoa', 'Horas', 'Data_Pagamento', 'Status', 'Observacao', 'Criado_Por', 'Criado_Em', 'ID']);
  }
  return sheet;
}

function horaExtraFromRow_(r) {
  return {
    id: String(r[7]),
    pessoa: String(r[0] || '').trim(),
    horas: Number(r[1]) || 0,
    dataPagamento: fmtData_(r[2]),
    status: String(r[3] || '').trim(),
    observacao: String(r[4] || '').trim(),
    criadoPor: String(r[5] || '').trim(),
    criadoEm: fmtDataHora_(r[6])
  };
}

function getHorasExtras(token) {
  requireUser_(token);
  const rows = getHorasExtrasSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    out.push(horaExtraFromRow_(rows[i]));
  }
  out.sort(function(a, b) { return (b.dataPagamento || '').localeCompare(a.dataPagamento || ''); });
  return out;
}

function saveHoraExtra(token, h) {
  const user  = requireUser_(token);
  const sheet = getHorasExtrasSheet_();

  const pessoa = String(h.pessoa || '').trim();
  const horas  = Number(h.horas) || 0;
  if (!pessoa) throw new Error('Informe a pessoa.');
  if (!horas) throw new Error('Informe a quantidade de horas.');
  const dataPagamentoStr = String(h.dataPagamento || '').trim();
  const dataPagamento = dataPagamentoStr ? new Date(dataPagamentoStr + 'T00:00:00') : '';
  const status = String(h.status || '').trim();
  const observacao = String(h.observacao || '').trim();

  if (h.id) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][7]) === String(h.id)) {
        sheet.getRange(i + 1, 1, 1, 5).setValues([[pessoa, horas, dataPagamento, status, observacao]]);
        return getHorasExtras(token);
      }
    }
    throw new Error('Registro de hora extra não encontrado.');
  }

  sheet.appendRow([pessoa, horas, dataPagamento, status, observacao, user.email, new Date(), Utilities.getUuid()]);
  return getHorasExtras(token);
}

function deleteHoraExtra(token, id) {
  requireUser_(token);
  const sheet = getHorasExtrasSheet_();
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][7]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return getHorasExtras(token);
}

function getSecretaria(token) {
  requireUser_(token);
  return {
    horarios: getHorarios(token),
    escalaSabado: getEscalaSabado(token),
    ajustes: getAjustesHorario(token),
    horasExtras: getHorasExtras(token)
  };
}

// =============================================================================
// FÉRIAS + FUNCIONÁRIOS (data de admissão)
// =============================================================================

function getFuncionariosSheet_() {
  const ss = getOpSS_();
  let sheet = ss.getSheetByName(ABA_FUNCIONARIOS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_FUNCIONARIOS);
    sheet.appendRow(['Nome', 'Data_Admissao', 'Data_Demissao', 'Funcao', 'Apelido', 'ID', 'Ciclo_Inicio_Override', 'Ciclo_Fim_Override']);
  }
  return sheet;
}

function funcionarioFromRow_(r) {
  const dataDemissao = fmtData_(r[2]);
  return {
    nome: String(r[0] || '').trim(),
    dataAdmissao: fmtData_(r[1]),
    dataDemissao: dataDemissao,
    funcao: String(r[3] || '').trim(),
    apelido: String(r[4] || '').trim(),
    ativo: !dataDemissao,
    id: String(r[5] || ''),
    cicloInicioOverride: fmtData_(r[6]),
    cicloFimOverride: fmtData_(r[7])
  };
}

function getFuncionarios(token) {
  requireUser_(token);
  const rows = getFuncionariosSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    out.push(funcionarioFromRow_(rows[i]));
  }
  out.sort(function(a, b) { return a.nome.localeCompare(b.nome, 'pt-BR'); });
  return out;
}

function saveFuncionario(token, f) {
  requireUser_(token);
  const sheet = getFuncionariosSheet_();

  const nome = String(f.nome || '').trim();
  if (!nome) throw new Error('Informe o nome.');
  const idParam        = String(f.id || '').trim();
  const chaveOriginal  = norm_(f.nomeOriginal || nome);
  const dataAdmissao = f.dataAdmissao ? new Date(f.dataAdmissao + 'T00:00:00') : '';
  const dataDemissao = f.dataDemissao ? new Date(f.dataDemissao + 'T00:00:00') : '';
  const funcao = String(f.funcao || '').trim();
  const apelido = String(f.apelido || '').trim();

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const rowId   = String(rows[i][5] || '').trim();
    const rowNome = norm_(rows[i][0]);
    const match   = (idParam && rowId === idParam) || (!idParam && rowNome === chaveOriginal);
    if (match) {
      sheet.getRange(i + 1, 1, 1, 5).setValues([[nome, dataAdmissao, dataDemissao, funcao, apelido]]);
      return getFuncionarios(token);
    }
  }
  sheet.appendRow([nome, dataAdmissao, dataDemissao, funcao, apelido, Utilities.getUuid()]);
  return getFuncionarios(token);
}

// Lê a planilha de RH da BRASAS e retorna os funcionários com unidade ONLINE
// e status Ativo. Colunas: B=Unidade, C=Nome, J=DataAdmissão, K=Status.
function getFuncionariosRH(token) {
  requireUser_(token);
  const ss = SpreadsheetApp.openById(RH_SS_ID);
  const sheet = ss.getSheets().filter(function(s) { return s.getSheetId() === RH_SHEET_GID; })[0];
  if (!sheet) throw new Error('Aba RH não encontrada (gid=' + RH_SHEET_GID + ').');
  const rows = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const unidade     = String(rows[i][1] || '');
    const nome        = String(rows[i][2] || '').trim();
    const dataAdmissao = rows[i][9];
    const status      = String(rows[i][10] || '').trim();
    if (!nome) continue;
    if (norm_(status) !== 'ativo') continue;
    if (unidade.toUpperCase().indexOf('ONLINE') === -1) continue;
    out.push({ nome: nome, dataAdmissao: fmtData_(dataAdmissao) });
  }
  out.sort(function(a, b) { return a.nome.localeCompare(b.nome, 'pt-BR'); });
  return out;
}

function saveCicloOverride(token, nome, cicloInicio, cicloFim) {
  requireUser_(token);
  const sheet = getFuncionariosSheet_();
  const rows = sheet.getDataRange().getValues();
  const nomeNorm = norm_(nome);
  for (let i = 1; i < rows.length; i++) {
    if (norm_(rows[i][0]) === nomeNorm) {
      const ini = cicloInicio ? new Date(cicloInicio + 'T00:00:00') : '';
      const fim = cicloFim   ? new Date(cicloFim   + 'T00:00:00') : '';
      sheet.getRange(i + 1, 7, 1, 2).setValues([[ini, fim]]);
      return getFeriasPagina(token);
    }
  }
  throw new Error('Funcionário não encontrado: ' + nome);
}

function getFeriasSheet_() {
  const ss = getOpSS_();
  let sheet = ss.getSheetByName(ABA_FERIAS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_FERIAS);
    sheet.appendRow(['Pessoa', 'Data_Inicio', 'Data_Fim', 'Observacao', 'Criado_Por', 'Criado_Em', 'ID']);
  }
  return sheet;
}

// Diferença de dias entre duas datas ISO (yyyy-MM-dd), inclusive nas duas pontas
function diasEntre_(iso1, iso2) {
  if (!iso1 || !iso2) return 0;
  const a = new Date(iso1 + 'T00:00:00'), b = new Date(iso2 + 'T00:00:00');
  return Math.round((b - a) / 86400000) + 1;
}

function feriaFromRow_(r) {
  const dataInicio = fmtData_(r[1]);
  const dataFim = fmtData_(r[2]);
  return {
    id: String(r[6]),
    pessoa: String(r[0] || '').trim(),
    dataInicio: dataInicio,
    dataFim: dataFim,
    dias: diasEntre_(dataInicio, dataFim),
    observacao: String(r[3] || '').trim(),
    criadoPor: String(r[4] || '').trim(),
    criadoEm: fmtDataHora_(r[5])
  };
}

function getFerias(token) {
  requireUser_(token);
  const rows = getFeriasSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0] || !rows[i][1]) continue;
    out.push(feriaFromRow_(rows[i]));
  }
  out.sort(function(a, b) { return (b.dataInicio || '').localeCompare(a.dataInicio || ''); });
  return out;
}

function saveFeria(token, f) {
  const user  = requireUser_(token);
  const sheet = getFeriasSheet_();

  const pessoa     = String(f.pessoa || '').trim();
  const dataInicio = String(f.dataInicio || '').trim();
  const dataFim    = String(f.dataFim || '').trim();
  if (!pessoa || !dataInicio || !dataFim) throw new Error('Informe a pessoa e o período.');
  if (dataFim < dataInicio) throw new Error('A data final não pode ser anterior à data inicial.');
  const observacao = String(f.observacao || '').trim();
  const ini = new Date(dataInicio + 'T00:00:00');
  const fim = new Date(dataFim + 'T00:00:00');

  if (f.id) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][6]) === String(f.id)) {
        sheet.getRange(i + 1, 1, 1, 4).setValues([[pessoa, ini, fim, observacao]]);
        return getFerias(token);
      }
    }
    throw new Error('Período de férias não encontrado.');
  }

  sheet.appendRow([pessoa, ini, fim, observacao, user.email, new Date(), Utilities.getUuid()]);
  return getFerias(token);
}

function deleteFeria(token, id) {
  requireUser_(token);
  const sheet = getFeriasSheet_();
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][6]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return getFerias(token);
}

function addMeses_(data, meses) {
  const d = new Date(data);
  d.setMonth(d.getMonth() + meses);
  return d;
}

// Calcula o ciclo de férias em aberto (regra CLT: o período aquisitivo de 12
// meses mais recente já concluído libera 30 dias, a serem usados nos 12 meses
// seguintes — o período concessivo). Retorna o período concessivo atual, ou
// null se ainda não completou 1 ano de casa (nenhum período aquisitivo
// concluído ainda), ou se não há data de admissão cadastrada.
function cicloFeriasAtual_(dataAdmissaoIso, hojeIso) {
  if (!dataAdmissaoIso) return null;
  const admissao = new Date(dataAdmissaoIso + 'T00:00:00');
  const hoje = new Date(hojeIso + 'T00:00:00');

  // n = quantos períodos aquisitivos de 12 meses já foram concluídos até hoje
  let n = 0;
  while (addMeses_(admissao, (n + 1) * 12) <= hoje) n++;
  if (n === 0) return null;

  // o período concessivo (onde os dias podem/devem ser usados) começa quando
  // o n-ésimo período aquisitivo termina, e vai até o término do seguinte
  return {
    inicio: fmtData_(addMeses_(admissao, n * 12)),
    fim: fmtData_(addMeses_(admissao, (n + 1) * 12))
  };
}

// Para cada pessoa (cadastrada em FUNCIONARIOS e/ou com férias lançadas),
// calcula o saldo de dias do ciclo de férias atualmente em aberto.
function getFeriasResumo(token) {
  requireUser_(token);
  const funcionarios = getFuncionarios(token);
  const ferias = getFerias(token);
  const hojeIso = fmtData_(new Date());

  const pessoas = {};
  funcionarios.forEach(function(f) {
    pessoas[norm_(f.nome)] = { nome: f.nome, dataAdmissao: f.dataAdmissao, dataDemissao: f.dataDemissao, funcao: f.funcao, apelido: f.apelido, cicloInicioOverride: f.cicloInicioOverride || '', cicloFimOverride: f.cicloFimOverride || '' };
  });
  ferias.forEach(function(f) {
    const k = norm_(f.pessoa);
    if (!pessoas[k]) pessoas[k] = { nome: f.pessoa, dataAdmissao: '', dataDemissao: '', funcao: '', apelido: '', cicloInicioOverride: '', cicloFimOverride: '' };
  });

  return Object.keys(pessoas).map(function(k) {
    const p = pessoas[k];
    const temOverride = p.cicloInicioOverride && p.cicloFimOverride;
    const ciclo = temOverride
      ? { inicio: p.cicloInicioOverride, fim: p.cicloFimOverride }
      : cicloFeriasAtual_(p.dataAdmissao, hojeIso);
    let diasTirados = 0;
    if (ciclo) {
      ferias.filter(function(f) {
        return norm_(f.pessoa) === k && f.dataInicio >= ciclo.inicio && f.dataInicio < ciclo.fim;
      }).forEach(function(f) { diasTirados += f.dias; });
    }
    return {
      nome: p.nome,
      dataAdmissao: p.dataAdmissao,
      dataDemissao: p.dataDemissao,
      funcao: p.funcao,
      apelido: p.apelido,
      cicloInicio: ciclo ? ciclo.inicio : '',
      cicloFim: ciclo ? ciclo.fim : '',
      diasTirados: ciclo ? diasTirados : null,
      diasRestantes: ciclo ? Math.max(0, 30 - diasTirados) : null,
      cicloCustomizado: !!temOverride,
      cicloInicioOverride: p.cicloInicioOverride || '',
      cicloFimOverride: p.cicloFimOverride || ''
    };
  }).sort(function(a, b) { return a.nome.localeCompare(b.nome, 'pt-BR'); });
}

function getFeriasPagina(token) {
  requireUser_(token);
  return {
    funcionarios: getFuncionarios(token),
    ferias: getFerias(token),
    resumo: getFeriasResumo(token),
    checklist: getChecklist(token)
  };
}

// =============================================================================
// CHECKLIST DE ADMISSÃO / DEMISSÃO
// =============================================================================
// Mesma lista de 16 itens usada nos dois processos, mas com marcação
// independente — a mesma pessoa pode ter status diferente pro mesmo item
// dependendo se é o checklist de entrada (ADMISSAO) ou de saída (DEMISSAO).

function getChecklistSheet_() {
  const ss = getOpSS_();
  let sheet = ss.getSheetByName(ABA_CHECKLIST_ADMISSAO);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_CHECKLIST_ADMISSAO);
    sheet.appendRow(['Nome', 'Processo', 'Item', 'Status', 'Atualizado_Por', 'Atualizado_Em']);
  }
  return sheet;
}

// Retorna a lista canônica de itens (pra a tela não precisar hardcodar de
// novo) e as marcações já feitas (só as que têm status — item sem marcação
// não gera linha na planilha).
function getChecklist(token) {
  requireUser_(token);
  const rows = getChecklistSheet_().getDataRange().getValues();
  const marcacoes = [];
  for (let i = 1; i < rows.length; i++) {
    const status = String(rows[i][3] || '').trim();
    if (!rows[i][0] || !status) continue;
    marcacoes.push({
      nome: String(rows[i][0]).trim(),
      processo: String(rows[i][1] || '').trim().toUpperCase(),
      item: String(rows[i][2]).trim(),
      status: status
    });
  }
  return { itens: CHECKLIST_ADMISSAO_ITENS, marcacoes: marcacoes };
}

function saveChecklistItem(token, nome, processo, item, status) {
  const user  = requireUser_(token);
  const sheet = getChecklistSheet_();

  const nomeNorm = norm_(nome);
  const processoNorm = String(processo || '').trim().toUpperCase();
  const itemNorm = norm_(item);

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (norm_(rows[i][0]) === nomeNorm &&
        String(rows[i][1] || '').trim().toUpperCase() === processoNorm &&
        norm_(rows[i][2]) === itemNorm) {
      sheet.getRange(i + 1, 4, 1, 3).setValues([[status, user.email, new Date()]]);
      return getChecklist(token);
    }
  }
  sheet.appendRow([nome, processoNorm, item, status, user.email, new Date()]);
  return getChecklist(token);
}

// Seed único: registra as férias já cumpridas informadas em 2026. Rode
// manualmente uma vez pelo editor do Apps Script (menu Executar →
// popularFeriasIniciais). Pode rodar de novo sem medo: períodos já existentes
// (mesma pessoa + mesmo início) não são duplicados.
function popularFeriasIniciais() {
  const sheet = getFeriasSheet_();
  const rows  = sheet.getDataRange().getValues();

  const existentes = new Set();
  for (let i = 1; i < rows.length; i++) {
    existentes.add(norm_(rows[i][0]) + '|' + fmtData_(rows[i][1]));
  }

  const periodos = [
    ['Natasha', '2026-01-05', '2026-01-12'],
    ['Elaine',  '2026-03-16', '2026-03-28'],
    ['Aline',   '2026-03-09', '2026-03-18'],
    ['Germana', '2026-03-16', '2026-03-29'],
    ['Rafaela', '2026-04-06', '2026-04-11'],
    ['Mel',     '2026-04-06', '2026-04-19'],
    ['Natasha', '2026-05-11', '2026-05-15'],
    ['Rosana',  '2026-05-18', '2026-05-31'],
    ['Aline',   '2026-06-02', '2026-06-11'],
    ['Bruna',   '2026-08-10', '2026-08-14'],
    ['Elaine',  '2026-09-21', '2026-10-04'],
    ['Nayara',  '2026-10-14', '2026-10-28']
  ];

  let adicionadas = 0;
  periodos.forEach(function(p) {
    const chave = norm_(p[0]) + '|' + p[1];
    if (existentes.has(chave)) return;
    sheet.appendRow([p[0], new Date(p[1] + 'T00:00:00'), new Date(p[2] + 'T00:00:00'), '', '', new Date(), Utilities.getUuid()]);
    existentes.add(chave);
    adicionadas++;
  });

  Logger.log(adicionadas + ' períodos de férias adicionados.');
  return adicionadas;
}

// Seed único: registra funcionários (equipe administrativa ativa da unidade
// ONLINE + algumas professoras já desligadas), conforme a aba "RJ - UNIDADES"
// da planilha de RH da BRASAS — nome, apelido e função copiados literalmente
// das colunas C/G/F de lá, sem reformatar capitalização. Rode manualmente uma
// vez pelo editor do Apps Script (menu Executar → popularFuncionariosIniciais).
// Pode rodar de novo sem medo: nomes já cadastrados não são sobrescritos.
function popularFuncionariosIniciais() {
  const sheet = getFuncionariosSheet_();
  const rows  = sheet.getDataRange().getValues();

  const existentes = new Set();
  for (let i = 1; i < rows.length; i++) {
    existentes.add(norm_(rows[i][0]));
  }

  const funcionarios = [
    { nome: 'Natasha', dataAdmissao: '2016-05-11', funcao: 'SUPERVISOR ADMINISTRATIVO' },
    { nome: 'Elaine',  dataAdmissao: '2012-02-01', funcao: 'SECRETARIA' },
    { nome: 'Aline',   dataAdmissao: '2014-08-14', funcao: 'ASSISTENTE OPERACIONAL' },
    { nome: 'Germana', dataAdmissao: '2016-02-01', funcao: 'COORDENADOR' },
    { nome: 'Rafaela', dataAdmissao: '2023-01-02', funcao: 'SECRETARIA' },
    { nome: 'Mel',     dataAdmissao: '2011-02-10', funcao: 'COORDENADOR' },
    { nome: 'Rosana',  dataAdmissao: '2022-01-26', funcao: 'SECRETARIA' },
    { nome: 'Bruna',   dataAdmissao: '2022-02-04', funcao: 'SECRETARIA' },
    { nome: 'Nayara',  dataAdmissao: '2023-08-01', funcao: 'SECRETARIA' },
    // Desligadas
    { nome: 'Alana Tomazetti Carvalho',          apelido: 'Avril',   dataAdmissao: '2024-02-05', dataDemissao: '2025-07-03', funcao: 'PROFESSOR' },
    { nome: 'Caio Mascheroni Costa Gonçalves',   apelido: 'Fred',    dataAdmissao: '2023-02-01', dataDemissao: '2025-12-16', funcao: 'PROFESSOR' },
    { nome: 'Carolina Mourão Mello',             apelido: 'Phoenix', dataAdmissao: '2025-02-03', dataDemissao: '2025-12-17', funcao: 'PROFESSOR' },
    { nome: 'Larissa da Silva Cury',             apelido: 'Cury',    dataAdmissao: '2025-08-01', dataDemissao: '2026-06-01', funcao: 'PROFESSOR' },
    // Novas, vindas junto com o checklist de admissão
    { nome: 'Fairuzz Jabbour',          dataAdmissao: '2025-08-01' },
    { nome: 'Lucas Ventura dos Santos', dataAdmissao: '2026-02-03' },
    { nome: 'Taiza Aguiar Ferreira',    dataAdmissao: '2026-06-01' }
  ];

  let adicionadas = 0;
  funcionarios.forEach(function(f) {
    const chave = norm_(f.nome);
    if (existentes.has(chave)) return;
    sheet.appendRow([
      f.nome,
      new Date(f.dataAdmissao + 'T00:00:00'),
      f.dataDemissao ? new Date(f.dataDemissao + 'T00:00:00') : '',
      f.funcao || '',
      f.apelido || '',
      Utilities.getUuid()
    ]);
    existentes.add(chave);
    adicionadas++;
  });

  Logger.log(adicionadas + ' funcionários adicionados.');
  return adicionadas;
}

// Marcações trazidas dos controles de admissão/desligamento: valores na mesma
// ordem de CHECKLIST_ADMISSAO_ITENS (os 10 primeiros itens vinham como
// TRUE/FALSE na planilha de origem, os 6 últimos já como Feito/Não se
// aplica). TRUE e FEITO viram "Feito"; FALSE e célula vazia viram "ainda não
// feito" (sem marcação, não gera linha). A mesma pessoa pode ter valores
// diferentes pro mesmo item entre admissão e demissão — são checklists
// independentes.
const CHECKLIST_ADMISSAO_BRUTO_ = {
  'Rosana Lisboa da Silva':     ['TRUE','TRUE','TRUE','TRUE','TRUE','FALSE','TRUE','FALSE','TRUE','FALSE','NAO SE APLICA','NAO SE APLICA','NAO SE APLICA','NAO SE APLICA','NAO SE APLICA','NAO SE APLICA'],
  'Larissa da Silva Cury':      ['TRUE','TRUE','TRUE','TRUE','TRUE','FALSE','TRUE','FALSE','TRUE','FALSE','FEITO','FEITO','NAO SE APLICA','FEITO','FEITO','FEITO'],
  'Fairuzz Jabbour':            ['TRUE','TRUE','TRUE','TRUE','TRUE','FALSE','TRUE','FALSE','TRUE','FALSE','FEITO','FEITO','NAO SE APLICA','FEITO','FEITO','FEITO'],
  'Lucas Ventura dos Santos':   ['TRUE','TRUE','TRUE','TRUE','TRUE','FALSE','TRUE','FALSE','FALSE','FALSE','NAO SE APLICA','NAO SE APLICA','NAO SE APLICA','FEITO','FEITO','FEITO'],
  'Taiza Aguiar Ferreira':      ['TRUE','TRUE','TRUE','TRUE','TRUE','FALSE','TRUE','FALSE','FALSE','FALSE','FEITO','FEITO','NAO SE APLICA','FEITO','','FEITO'],
  'Caren Xavier Lucena da Silva': ['TRUE','TRUE','TRUE','FALSE','FALSE','TRUE','TRUE','TRUE','TRUE','TRUE','NAO SE APLICA','FEITO','NAO SE APLICA','FEITO','','NAO SE APLICA']
};

const CHECKLIST_DEMISSAO_BRUTO_ = {
  'Alana Tomazetti Carvalho':        ['TRUE','TRUE','TRUE','TRUE','TRUE','FALSE','TRUE','FALSE','TRUE','FALSE','FEITO','FEITO','FEITO','FEITO','FEITO','NAO SE APLICA'],
  'Caio Mascheroni Costa Gonçalves': ['TRUE','TRUE','TRUE','TRUE','TRUE','FALSE','TRUE','TRUE','TRUE','FALSE','','FEITO','FEITO','FEITO','FEITO','NAO SE APLICA'],
  'Carolina Mourão Mello':           ['TRUE','TRUE','TRUE','TRUE','TRUE','FALSE','TRUE','TRUE','TRUE','FALSE','FEITO','FEITO','NAO SE APLICA','FEITO','FEITO','NAO SE APLICA'],
  'Larissa da Silva Cury':           ['TRUE','TRUE','TRUE','TRUE','FALSE','FALSE','TRUE','FALSE','FALSE','FALSE','FEITO','FEITO','NAO SE APLICA','FEITO','','NAO SE APLICA']
};

function checklistStatusDeBruto_(raw) {
  const v = String(raw || '').trim().toUpperCase();
  if (v === 'TRUE' || v === 'FEITO') return 'Feito';
  if (v === 'NAO SE APLICA' || v === 'NÃO SE APLICA') return 'Não se aplica';
  return ''; // FALSE ou célula vazia = ainda não feito
}

function popularChecklistBruto_(processo, dados) {
  const sheet = getChecklistSheet_();
  const rows  = sheet.getDataRange().getValues();

  const existentes = new Set();
  for (let i = 1; i < rows.length; i++) {
    existentes.add(norm_(rows[i][0]) + '|' + String(rows[i][1] || '').trim().toUpperCase() + '|' + norm_(rows[i][2]));
  }

  let adicionadas = 0;
  Object.keys(dados).forEach(function(nome) {
    const valores = dados[nome];
    CHECKLIST_ADMISSAO_ITENS.forEach(function(item, i) {
      const status = checklistStatusDeBruto_(valores[i]);
      if (!status) return;
      const chave = norm_(nome) + '|' + processo + '|' + norm_(item);
      if (existentes.has(chave)) return;
      sheet.appendRow([nome, processo, item, status, '', new Date()]);
      existentes.add(chave);
      adicionadas++;
    });
  });

  return adicionadas;
}

// Seed único: registra as marcações de checklist de admissão já feitas pra
// essas pessoas. Rode manualmente uma vez pelo editor do Apps Script (menu
// Executar → popularChecklistAdmissaoInicial). Pode rodar de novo sem medo:
// marcações já existentes (mesma pessoa + processo + item) não são
// sobrescritas.
function popularChecklistAdmissaoInicial() {
  const adicionadas = popularChecklistBruto_('ADMISSAO', CHECKLIST_ADMISSAO_BRUTO_);
  Logger.log(adicionadas + ' marcações de checklist de admissão adicionadas.');
  return adicionadas;
}

// Seed único: registra as marcações de checklist de desligamento já feitas
// para as professoras desligadas. Rode manualmente uma vez pelo editor do
// Apps Script (menu Executar → popularChecklistDemissaoInicial). Pode rodar
// de novo sem medo: marcações já existentes (mesma pessoa + processo + item)
// não são sobrescritas.
function popularChecklistDemissaoInicial() {
  const adicionadas = popularChecklistBruto_('DEMISSAO', CHECKLIST_DEMISSAO_BRUTO_);
  Logger.log(adicionadas + ' marcações de checklist de demissão adicionadas.');
  return adicionadas;
}
