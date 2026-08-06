// =============================================================================
// CONFIGURAÇÕES — Webapp Controle Operacional BOL
// =============================================================================

const OPERACIONAL_SHEET_ID = '1I8nss8jH0sv8qCmagWI2bITPEW2JWhi5OUsjVeBA_-M';
const HUB_SS_ID  = '1eZPbzhzjhjHoPwMhAW5YvOZgYiAvlTYc07dRan6Lyoc'; // mesmo Hub do app de Horas
const MEU_ACESSO = 'operacionalbol'; // precisa estar liberado na coluna ACESSOS da aba SESSOES do Hub

// Acesso por página. A chave base MEU_ACESSO só deixa a pessoa ENTRAR no app;
// cada aba abaixo exige, ALÉM dela, a sua própria chave na mesma coluna
// ACESSOS da aba SESSOES do Hub — separadas por vírgula. Ex.: alguém que só
// cuida da escala recebe "operacionalbol, operacionalbol_secretaria".
// A chave do objeto é o mesmo nome usado no data-tab do Index.html.
//
// Minhas Férias é a única aba fora desta lista: é o autoatendimento da própria
// pessoa (as férias dela, os pedidos dela), então fica aberta a quem entra.
const ACESSOS_PAGINA_ = {
  dashboard:    'operacionalbol_dash',
  calendario:   'operacionalbol_calendario',
  rotinas:      'operacionalbol_rotinas',
  rotinaOp:     'operacionalbol_rotinaope',
  entregas:     'operacionalbol_entregas',
  funcionarios: 'operacionalbol_func',
  secretaria:   'operacionalbol_secretaria',
  projetos:     'operacionalbol_projetos'
};

// E-mails com acesso total ao painel: enxergam todas as abas sem precisar de
// nenhuma das chaves acima. Serve de atalho e de trava de segurança — evita
// alguém ficar trancado pra fora do próprio painel por uma chave esquecida.
const ADMIN_EMAILS_ = ['natasha@brasas.com', 'aline.castro@brasas.com', 'adriane@brasas.com', 'bruno@brasas.com', 'peter@brasas.com'];

function ehAdmin_(email) {
  return ADMIN_EMAILS_.indexOf(String(email || '').trim().toLowerCase()) !== -1;
}
const HUB_URL     = 'https://script.google.com/a/macros/brasas.com/s/AKfycbyF7BArYMYFtcQY7_4RTGGPw89yNohAjR7eGptItP-EsnWhNfiZR2ISRaHdAkwlLSlr/exec';

const ABA_EVENTOS         = 'EVENTOS';
const ABA_ROTINAS         = 'ROTINAS';
const ABA_ROTINAS_OPERACIONAIS = 'ROTINAS_OPERACIONAIS';
const ABA_ROTINAS_LOG     = 'ROTINAS_LOG';
const ABA_ENTREGAS        = 'ENTREGAS';
const ABA_HORARIOS        = 'HORARIOS';
const ABA_ESCALA_SABADO   = 'ESCALA_SABADO';
const ABA_AJUSTES_HORARIO = 'AJUSTES_HORARIO';
const ABA_FERIAS          = 'FERIAS';
const ABA_PERIODOS_AQUISITIVOS = 'PERIODOS_AQUISITIVOS';
const ABA_SOLICITACOES_FERIAS  = 'SOLICITACOES_FERIAS';
const ABA_FUNCIONARIOS    = 'FUNCIONARIOS';
const ABA_CHECKLIST_ADMISSAO = 'ADMISSAO_CHECKLIST';
const ABA_PROJETOS_COLUNAS = 'PROJETOS_COLUNAS';
const ABA_PROJETOS_CARTOES = 'PROJETOS_CARTOES';

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

// Tudo que a tela precisa pra abrir, numa chamada só. Antes o front fazia 11
// chamadas em fila (cada uma esperando a anterior voltar) e o painel só abria
// depois da última — além de lento, qualquer engasgo no meio deixava o app
// preso no "Carregando...". Como requireUser_ e getOpSS_ são cacheados por
// execução, juntar tudo aqui também abre a planilha bem menos vezes (era isso
// que causava o RESOURCE_EXHAUSTED).
function getBootstrap(token) {
  const user = requireUser_(token);
  const out = {
    user: user,
    minhasFerias: getMinhaFeriasPagina(token) // aberta a todos
  };
  // Só busca os dados das abas que esta pessoa realmente abre: quem não tem a
  // chave nem paga o custo de ler aquelas abas da planilha.
  const paginas = user.paginas || {};
  if (paginas.dashboard)  out.dashboard  = getDashboard(token);
  if (paginas.rotinas)    out.rotinas    = getRotinas(token);
  if (paginas.rotinaOp)   out.rotinaOp   = getRotinasOperacionais(token);
  if (paginas.calendario) out.eventos    = getEventos(token);
  if (paginas.entregas)   out.entregas   = getEntregas(token);
  if (paginas.secretaria) out.secretaria = getSecretaria(token);
  if (paginas.projetos)   out.projetos   = getProjetosPagina(token);
  if (paginas.funcionarios) {
    out.ferias       = getFeriasPagina(token);
    out.solicitacoes = getSolicitacoes(token);
    // A lista do RH vem de outra planilha — se ela falhar, o painel abre
    // mesmo assim (o seletor de novo funcionário fica com "Carregando lista
    // RH...", igual ao comportamento antigo).
    try { out.funcionariosRH = getFuncionariosRH(token); }
    catch (e) { out.funcionariosRH = []; }
  }
  return out;
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

    // Resolve de uma vez quais abas restritas esta pessoa alcança. Admin passa
    // em todas; os demais, só nas cujas chaves estiverem na coluna ACESSOS.
    const isAdmin = ehAdmin_(email);
    const paginas = {};
    Object.keys(ACESSOS_PAGINA_).forEach(function(pagina) {
      paginas[pagina] = isAdmin || acessos.indexOf(ACESSOS_PAGINA_[pagina]) !== -1;
    });

    return {
      email: email,
      nome:  String(row[2] || '').trim(),
      role:  String(row[3] || '').trim().toLowerCase(),
      isAdmin: isAdmin,
      paginas: paginas
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

// Igual requireUser_, mas exige também a chave da página em questão (ver
// ACESSOS_PAGINA_). Toda função de leitura ou gravação de uma aba restrita
// passa por aqui: esconder o botão no Index.html é só conforto visual, quem
// souber o nome da função ainda consegue chamá-la pelo console — é esta
// checagem no servidor que de fato barra. Dashboard, Rotinas e Minhas Férias
// não usam isso; são de todo mundo que entra no app.
function requirePagina_(token, pagina) {
  const user = requireUser_(token);
  if (!user.paginas || !user.paginas[pagina]) {
    throw new Error('Você não tem acesso a esta página. Contacte o administrador.');
  }
  return user;
}

// =============================================================================
// DASHBOARD
// =============================================================================

function getDashboard(token) {
  requirePagina_(token, 'dashboard');

  const hoje    = new Date();
  const hojeStr = fmtData_(hoje);
  const diaSemana = hoje.getDay(); // 0=domingo ... 6=sábado

  const horarios    = getHorariosImpl_().filter(function(h) { return h.ativo; });
  const ajustesHoje  = getAjustesHorarioImpl_().filter(function(a) {
    return a.data === hojeStr && (!a.status || a.status === 'Pendente');
  });

  let quemTrabalha = [];
  if (diaSemana === 6) {
    const escalaHoje = getEscalaSabadoImpl_().filter(function(e) { return e.data === hojeStr; })[0];
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

  const feriasPeriodos = getFeriasImpl_();
  quemTrabalha = quemTrabalha.map(function(p) {
    const deFerias = feriasPeriodos.some(function(f) {
      return norm_(f.pessoa) === norm_(p.nome) && f.dataInicio <= hojeStr && f.dataFim >= hojeStr;
    });
    return deFerias ? Object.assign({}, p, { ferias: true }) : p;
  });

  const eventos  = getEventosImpl_();
  const hojeMs   = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate()).getTime();
  // Próximos eventos: o que ainda não terminou, até o fim do mês vigente. Não
  // precisa de corte por quantidade — o próprio mês já limita. Um evento que
  // começou antes e ainda está rolando conta, porque acontece durante o mês.
  // O Calendário continua mostrando o ano inteiro.
  const fimDoMesMs = new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0).getTime();
  const proximosEventos = eventos
    .filter(function(ev) {
      if (!ev.dataInicio) return false;
      const ini = new Date(ev.dataInicio + 'T00:00:00').getTime();
      const fim = ev.dataFim ? new Date(ev.dataFim + 'T00:00:00').getTime() : ini;
      return fim >= hojeMs && ini <= fimDoMesMs;
    })
    .sort(function(a, b) {
      return a.dataInicio < b.dataInicio ? -1 : (a.dataInicio > b.dataInicio ? 1 : 0);
    });

  const mesAtual = hoje.getMonth() + 1, anoAtual = hoje.getFullYear();
  const entregasPendentes = getEntregasImpl_().filter(function(e) {
    return e.mes === mesAtual && e.ano === anoAtual && e.status !== 'NO PRAZO';
  });

  return { hoje: hojeStr, quemTrabalha: quemTrabalha, proximosEventos: proximosEventos, entregasPendentes: entregasPendentes };
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
function eventosVirtuaisFerias_() {
  return getFeriasImpl_().map(function(f) {
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

// Implementação sem checagem de admin — usada pelo Calendário (admin) e
// também internamente pelo Dashboard (aberto a qualquer funcionário).
function getEventosImpl_() {
  const rows = getEventosSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][2]) continue; // sem título
    // Férias não vêm mais de EVENTOS lançados manualmente — só da aba FERIAS
    // (eventosVirtuaisFerias_ abaixo), pra não duplicar no Calendário.
    if (String(rows[i][3]).trim().toUpperCase() === 'FERIAS') continue;
    out.push(eventoFromRow_(rows[i]));
  }
  out.push.apply(out, eventosVirtuaisFerias_());
  out.sort(function(a, b) { return (a.dataInicio || '9999').localeCompare(b.dataInicio || '9999'); });
  return out;
}

function getEventos(token) {
  requirePagina_(token, 'calendario');
  return getEventosImpl_();
}

function saveEvento(token, evento) {
  const user  = requirePagina_(token, 'calendario');
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
  requirePagina_(token, 'calendario');
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

// Há duas listas de rotina, idênticas em comportamento: mudam só a aba onde
// moram e a chave de acesso que exigem. O histórico de conclusão é o mesmo
// ROTINAS_LOG para as duas — cada linha é identificada pelo ID (UUID) da
// tarefa, então não há como uma lista enxergar a conclusão da outra.
const LISTAS_ROTINA_ = {
  rotinas:  { aba: ABA_ROTINAS,              pagina: 'rotinas'  },
  rotinaOp: { aba: ABA_ROTINAS_OPERACIONAIS, pagina: 'rotinaOp' }
};

function listaRotina_(lista) {
  const cfg = LISTAS_ROTINA_[lista];
  if (!cfg) throw new Error('Lista de rotina desconhecida: ' + lista);
  return cfg;
}

// ROTINAS veio da migração; ROTINAS_OPERACIONAIS nasce vazia na primeira vez
// que alguém abre a aba, com o mesmo cabeçalho.
function getRotinasSheetDe_(nomeAba) {
  const ss = getOpSS_();
  let sheet = ss.getSheetByName(nomeAba);
  if (!sheet) {
    sheet = ss.insertSheet(nomeAba);
    sheet.appendRow(['Frequência', 'Tarefa', 'Responsável', 'Ativo', 'ID']);
    return sheet;
  }
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

function getRotinasSheet_() { return getRotinasSheetDe_(ABA_ROTINAS); }

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

function rotinasDaLista_(lista) {
  const rows    = getRotinasSheetDe_(listaRotina_(lista).aba).getDataRange().getValues();
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

function toggleRotinaDaLista_(lista, user, id, periodo, concluido) {
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

  return rotinasDaLista_(lista);
}

function saveRotinaDaLista_(lista, rotina) {
  const sheet = getRotinasSheetDe_(listaRotina_(lista).aba);

  const frequencia  = String(rotina.frequencia || '').trim().toUpperCase();
  const tarefa      = String(rotina.tarefa || '').trim();
  const responsavel = String(rotina.responsavel || 'Todas').trim();
  if (!frequencia || !tarefa) throw new Error('Informe a frequência e a tarefa.');

  if (rotina.id) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][4]) === String(rotina.id)) {
        sheet.getRange(i + 1, 1, 1, 3).setValues([[frequencia, tarefa, responsavel]]);
        return rotinasDaLista_(lista);
      }
    }
    throw new Error('Rotina não encontrada.');
  }

  sheet.appendRow([frequencia, tarefa, responsavel, 'SIM', Utilities.getUuid()]);
  return rotinasDaLista_(lista);
}

function deleteRotinaDaLista_(lista, id) {
  const sheet = getRotinasSheetDe_(listaRotina_(lista).aba);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][4]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return rotinasDaLista_(lista);
}

// ── Rotinas (aba ROTINAS) ─────────────────────────────

function getRotinas(token) {
  requirePagina_(token, 'rotinas');
  return rotinasDaLista_('rotinas');
}

function toggleRotina(token, id, periodo, concluido) {
  const user = requirePagina_(token, 'rotinas');
  return toggleRotinaDaLista_('rotinas', user, id, periodo, concluido);
}

function saveRotina(token, rotina) {
  requirePagina_(token, 'rotinas');
  return saveRotinaDaLista_('rotinas', rotina);
}

function deleteRotina(token, id) {
  requirePagina_(token, 'rotinas');
  return deleteRotinaDaLista_('rotinas', id);
}

// ── Rotina Operacional (aba ROTINAS_OPERACIONAIS) ─────
// Lista à parte, com chave de acesso própria. Mesmas operações da de cima —
// o que muda é só de qual aba os dados saem e quem pode chegar até elas.

function getRotinasOperacionais(token) {
  requirePagina_(token, 'rotinaOp');
  return rotinasDaLista_('rotinaOp');
}

function toggleRotinaOperacional(token, id, periodo, concluido) {
  const user = requirePagina_(token, 'rotinaOp');
  return toggleRotinaDaLista_('rotinaOp', user, id, periodo, concluido);
}

function saveRotinaOperacional(token, rotina) {
  requirePagina_(token, 'rotinaOp');
  return saveRotinaDaLista_('rotinaOp', rotina);
}

function deleteRotinaOperacional(token, id) {
  requirePagina_(token, 'rotinaOp');
  return deleteRotinaDaLista_('rotinaOp', id);
}

function deleteFuncionario(token, nome) {
  requirePagina_(token, 'funcionarios');
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

// Sem checagem de admin — usada pela tela de Entregas (admin) e internamente
// pelo Dashboard (aberto a qualquer funcionário).
function getEntregasImpl_() {
  const rows = getEntregasSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    out.push(entregaFromRow_(rows[i]));
  }
  return out;
}

function getEntregas(token) {
  requirePagina_(token, 'entregas');
  return getEntregasImpl_();
}

function saveEntregaStatus(token, payload) {
  const user  = requirePagina_(token, 'entregas');
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

// Sem checagem de admin — usada pela Secretaria (admin) e internamente pelo
// Dashboard (aberto a qualquer funcionário).
function getHorariosImpl_() {
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

function getHorarios(token) {
  requirePagina_(token, 'secretaria');
  return getHorariosImpl_();
}

function saveHorario(token, h) {
  requirePagina_(token, 'secretaria');
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

// Sem checagem de admin — usada pela Secretaria (admin) e internamente pelo
// Dashboard (aberto a qualquer funcionário, pra saber a escala de sábado).
function getEscalaSabadoImpl_() {
  const rows = getEscalaSabadoSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][2]) continue;
    out.push(escalaFromRow_(rows[i]));
  }
  out.sort(function(a, b) { return a.data.localeCompare(b.data); });
  return out;
}

function getEscalaSabado(token) {
  requirePagina_(token, 'secretaria');
  return getEscalaSabadoImpl_();
}

function saveEscalaSabado(token, e) {
  requirePagina_(token, 'secretaria');
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
  const header = sheet.getRange(1, 1, 1, 16).getValues()[0];
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
  if (!header[14]) sheet.getRange(1, 15).setValue('Anexo_URL');
  if (!header[15]) sheet.getRange(1, 16).setValue('Anexo_Nome');
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
    horarioPara: fmtHora_(r[13]),
    anexoUrl: String(r[14] || '').trim(),
    anexoNome: String(r[15] || '').trim()
  };
}

// Pasta do Drive onde ficam os anexos (atestados, documentos etc.) dos
// ajustes de horário: https://drive.google.com/drive/folders/1IRDW35AMspLcsIf3ke-DvWbxRbJtI4nP
const ANEXOS_AJUSTES_FOLDER_ID = '1IRDW35AMspLcsIf3ke-DvWbxRbJtI4nP';

// Recebe o arquivo em base64 (vindo do FileReader do navegador), salva na
// pasta de anexos e libera visualização por link pra quem tem conta no
// domínio da BRASAS — sem isso, só o dono do arquivo (a conta por trás do
// Apps Script) conseguiria abrir o link.
function uploadAnexoAjuste_(base64, nome, mimeType) {
  const folder = DriveApp.getFolderById(ANEXOS_AJUSTES_FOLDER_ID);
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', nome || 'anexo');
  const file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    Logger.log('uploadAnexoAjuste_: não foi possível liberar compartilhamento por link: ' + e);
  }
  return { url: file.getUrl(), nome: file.getName() };
}

// Sem checagem de admin — usada pela Secretaria (admin) e internamente pelo
// Dashboard (aberto a qualquer funcionário, pra saber trocas/coberturas de hoje).
function getAjustesHorarioImpl_() {
  const rows = getAjustesSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    out.push(ajusteFromRow_(rows[i]));
  }
  out.sort(function(a, b) { return (b.data || '').localeCompare(a.data || ''); });
  return out;
}

function getAjustesHorario(token) {
  requirePagina_(token, 'secretaria');
  return getAjustesHorarioImpl_();
}

function saveAjusteHorario(token, a) {
  const user  = requirePagina_(token, 'secretaria');
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

  // anexo === null: não mexe no que já tinha. {url,nome}: grava (novo upload,
  // ou {url:'',nome:''} quando o usuário pediu pra remover o anexo atual).
  let anexo = null;
  if (a.anexoBase64) {
    anexo = uploadAnexoAjuste_(a.anexoBase64, a.anexoNome, a.anexoMimeType);
  } else if (a.removerAnexo) {
    anexo = { url: '', nome: '' };
  }

  if (a.id) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][7]) === String(a.id)) {
        sheet.getRange(i + 1, 1, 1, 5).setValues([[atividade, data, status, data.getMonth() + 1, data.getFullYear()]]);
        sheet.getRange(i + 1, 9, 1, 4).setValues([[tags, tipo, quemSai, quemCobre]]);
        // Colunas de horário precisam ser texto puro, senão o Sheets converte
        // um valor tipo "07:00" (sem " - ") num serial de hora automaticamente.
        sheet.getRange(i + 1, 13, 1, 2).setNumberFormat('@').setValues([[horarioDe, horarioPara]]);
        if (anexo) sheet.getRange(i + 1, 15, 1, 2).setValues([[anexo.url, anexo.nome]]);
        return getAjustesHorario(token);
      }
    }
    throw new Error('Ajuste não encontrado.');
  }

  sheet.appendRow([
    atividade, data, status, data.getMonth() + 1, data.getFullYear(), user.email, new Date(), Utilities.getUuid(),
    tags, tipo, quemSai, quemCobre, '', '',
    anexo ? anexo.url : '', anexo ? anexo.nome : ''
  ]);
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 13, 1, 2).setNumberFormat('@').setValues([[horarioDe, horarioPara]]);
  return getAjustesHorario(token);
}

function deleteAjusteHorario(token, id) {
  requirePagina_(token, 'secretaria');
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
  requirePagina_(token, 'secretaria');
  return {
    horarios: getHorarios(token),
    escalaSabado: getEscalaSabado(token),
    ajustes: getAjustesHorario(token)
  };
}

// =============================================================================
// FÉRIAS + FUNCIONÁRIOS (data de admissão)
// =============================================================================

const FUNCIONARIOS_HEADERS_ = ['Nome', 'Data_Admissao', 'Data_Demissao', 'Funcao', 'Apelido', 'ID', 'Ciclo_Inicio_Override', 'Ciclo_Fim_Override', 'Rastreado_Ferias', 'Rastreado_Admissao', 'Rastreado_Demissao', 'Email'];
const FUNCIONARIOS_COL_RASTREADO_FERIAS_ = 9; // a partir daqui o backfill de migrarRastreamentoInicial_ só deve rodar uma vez

function getFuncionariosSheet_() {
  const ss = getOpSS_();
  let sheet = ss.getSheetByName(ABA_FUNCIONARIOS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_FUNCIONARIOS);
    sheet.appendRow(FUNCIONARIOS_HEADERS_);
    return sheet;
  }
  const lastCol = sheet.getLastColumn();
  if (lastCol < FUNCIONARIOS_HEADERS_.length) {
    // Só faz o backfill de Rastreado_* na primeira vez que essas colunas são
    // criadas — se elas já existiam (e só uma coluna posterior, como Email,
    // está sendo adicionada agora), rodar de novo apagaria marcações que os
    // usuários já ajustaram manualmente no app.
    const precisaBackfillRastreamento = lastCol < FUNCIONARIOS_COL_RASTREADO_FERIAS_ + 2;
    sheet.getRange(1, lastCol + 1, 1, FUNCIONARIOS_HEADERS_.length - lastCol).setValues([FUNCIONARIOS_HEADERS_.slice(lastCol)]);
    if (precisaBackfillRastreamento) migrarRastreamentoInicial_(sheet);
  }
  return sheet;
}

// Roda uma única vez, no momento em que as colunas Rastreado_* são criadas:
// replica o comportamento antigo (mesma lista compartilhada entre Férias e
// Admissão, com pessoas escondidas via arrays de exclusão hardcoded no
// front-end) como estado inicial, pra ninguém sumir das telas na migração.
// Daqui pra frente, cada aba passa a ter sua própria lista independente.
function migrarRastreamentoInicial_(sheet) {
  const FERIAS_EXCLUIDOS = ['fairuzz', 'lucas ventura', 'taiza aguiar'];
  const ADMISSAO_EXCLUIDOS = ['aline', 'bruna', 'elaine', 'germana', 'mel', 'natasha', 'nayara', 'rafaela', 'rosana'];
  const rows = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const nome = String(rows[i][0] || '').trim();
    if (!nome) { out.push(['', '', '']); continue; }
    const nomeNorm = norm_(nome);
    const primeiroNome = norm_(nome.split(' ')[0]);
    const temDemissao = !!rows[i][2];
    const excluidoFerias = FERIAS_EXCLUIDOS.some(function(ex) { return nomeNorm === ex || nomeNorm.indexOf(ex + ' ') === 0; });
    out.push([
      !temDemissao && !excluidoFerias,
      !temDemissao && ADMISSAO_EXCLUIDOS.indexOf(primeiroNome) === -1,
      temDemissao
    ]);
  }
  if (out.length) sheet.getRange(2, 9, out.length, 3).setValues(out);
}

function bool_(v) {
  return v === true || String(v).trim().toUpperCase() === 'TRUE';
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
    rastreadoFerias: bool_(r[8]),
    rastreadoAdmissao: bool_(r[9]),
    rastreadoDemissao: bool_(r[10]),
    email: String(r[11] || '').trim().toLowerCase()
  };
}

// Sem checagem de admin — usada pelo Cadastro (admin) e internamente por
// Minhas Férias / solicitarFerias (abertos a qualquer funcionário, que só
// precisam achar o próprio registro pelo e-mail).
function getFuncionariosImpl_() {
  const rows = getFuncionariosSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    out.push(funcionarioFromRow_(rows[i]));
  }
  out.sort(function(a, b) { return a.nome.localeCompare(b.nome, 'pt-BR'); });
  return out;
}

function getFuncionarios(token) {
  requirePagina_(token, 'funcionarios');
  return getFuncionariosImpl_();
}

function saveFuncionario(token, f) {
  requirePagina_(token, 'funcionarios');
  const sheet = getFuncionariosSheet_();

  const nome = String(f.nome || '').trim();
  if (!nome) throw new Error('Informe o nome.');
  const idParam        = String(f.id || '').trim();
  const chaveOriginal  = norm_(f.nomeOriginal || nome);
  const dataAdmissao = f.dataAdmissao ? new Date(f.dataAdmissao + 'T00:00:00') : '';
  const dataDemissao = f.dataDemissao ? new Date(f.dataDemissao + 'T00:00:00') : '';
  const funcao = String(f.funcao || '').trim();
  const apelido = String(f.apelido || '').trim();
  const email = String(f.email || '').trim().toLowerCase();

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const rowId   = String(rows[i][5] || '').trim();
    const rowNome = norm_(rows[i][0]);
    const match   = (idParam && rowId === idParam) || (!idParam && rowNome === chaveOriginal);
    if (match) {
      const nomeAntigo = String(rows[i][0] || '').trim();
      sheet.getRange(i + 1, 1, 1, 5).setValues([[nome, dataAdmissao, dataDemissao, funcao, apelido]]);
      sheet.getRange(i + 1, 12).setValue(email);
      if (norm_(nomeAntigo) !== norm_(nome)) propagarRenomeacaoPessoa_(nomeAntigo, nome);
      return getFuncionariosImpl_();
    }
  }
  sheet.appendRow([nome, dataAdmissao, dataDemissao, funcao, apelido, Utilities.getUuid(), '', '', '', '', '', email]);
  return getFuncionariosImpl_();
}

// FERIAS, PERIODOS_AQUISITIVOS e SOLICITACOES_FERIAS referenciam a pessoa
// pelo nome (não pelo ID do funcionário). Renomear alguém só no Cadastro
// deixaria essas três abas apontando para um nome que não existe mais —
// férias e períodos aquisitivos sumiriam da tela sem aviso. Por isso o
// rename é propagado para todas elas.
function propagarRenomeacaoPessoa_(nomeAntigo, nomeNovo) {
  const chave = norm_(nomeAntigo);
  if (!chave) return;
  [getFeriasSheet_(), getPeriodosAquisitivosSheet_(), getSolicitacoesSheet_()].forEach(function(sheet) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (norm_(rows[i][0]) === chave) sheet.getRange(i + 1, 1).setValue(nomeNovo);
    }
  });
}

// Lê a planilha de RH da BRASAS e retorna os funcionários com unidade ONLINE
// e status Ativo. Colunas: B=Unidade, C=Nome, F=Função, G=Apelido, H=E-mail,
// J=DataAdmissão, K=Status.
function getFuncionariosRH(token) {
  requirePagina_(token, 'funcionarios');
  return lerFuncionariosRHOnlineAtivos_();
}

function lerFuncionariosRHOnlineAtivos_() {
  const ss = SpreadsheetApp.openById(RH_SS_ID);
  const sheet = ss.getSheets().filter(function(s) { return s.getSheetId() === RH_SHEET_GID; })[0];
  if (!sheet) throw new Error('Aba RH não encontrada (gid=' + RH_SHEET_GID + ').');
  const rows = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const unidade     = String(rows[i][1] || '');
    const nome        = String(rows[i][2] || '').trim();
    const funcao      = String(rows[i][5] || '').trim();
    const apelido     = String(rows[i][6] || '').trim();
    // Coluna Z ("E-mail BRASAS") — é o e-mail institucional, o mesmo usado no
    // login do Hub. A coluna H ("E-MAIL") às vezes guarda e-mail pessoal.
    const email       = String(rows[i][25] || '').trim().toLowerCase();
    const dataAdmissao = rows[i][9];
    const status      = String(rows[i][10] || '').trim();
    if (!nome) continue;
    if (norm_(status) !== 'ativo') continue;
    if (unidade.toUpperCase().indexOf('ONLINE') === -1) continue;
    out.push({ nome: nome, funcao: funcao, apelido: apelido, email: email, dataAdmissao: fmtData_(dataAdmissao) });
  }
  out.sort(function(a, b) { return a.nome.localeCompare(b.nome, 'pt-BR'); });
  return out;
}

// Sincroniza o Cadastro de Funcionários (aba FUNCIONARIOS) com a planilha de
// RH: atualiza admissão/função/e-mail de quem já está cadastrado e adiciona
// quem ainda não está. Não marca ninguém em Férias/Admissão/Demissão — isso
// continua sendo sempre uma ação manual e separada em cada aba.
// Roda dentro de uma trava: sem ela, cliques repetidos no botão viravam
// execuções simultâneas que liam a aba no mesmo estado e cada uma adicionava
// as mesmas pessoas de novo — foi assim que o cadastro chegou a ter a mesma
// pessoa 5 vezes. A trava faz o segundo clique esperar o primeiro terminar
// (e aí ele já encontra todo mundo cadastrado e não adiciona nada).
function sincronizarFuncionariosRH(token) {
  requirePagina_(token, 'funcionarios');
  const lock = LockService.getScriptLock();
  lock.waitLock(120000);
  try {
    return sincronizarFuncionariosRHImpl_();
  } finally {
    lock.releaseLock();
  }
}

function sincronizarFuncionariosRHImpl_() {
  const sheet = getFuncionariosSheet_();
  const duplicatasRemovidas = removerFuncionariosDuplicados_(sheet);

  const rh = lerFuncionariosRHOnlineAtivos_();
  const rows = sheet.getDataRange().getValues();

  const linhaPorNome = {};
  for (let i = 1; i < rows.length; i++) {
    const nome = String(rows[i][0] || '').trim();
    if (nome) linhaPorNome[norm_(nome)] = i + 1;
  }

  let adicionados = 0, atualizados = 0;
  const novas = [];
  rh.forEach(function(p) {
    const chave = norm_(p.nome);
    const linha = linhaPorNome[chave];
    if (linha === undefined) {
      novas.push([
        p.nome,
        p.dataAdmissao ? new Date(p.dataAdmissao + 'T00:00:00') : '',
        '',
        p.funcao || '',
        p.apelido || '',
        Utilities.getUuid(),
        '', '', '', '', '',
        p.email || ''
      ]);
      linhaPorNome[chave] = -1; // se o RH repetir o mesmo nome, não duplica
      adicionados++;
      return;
    }
    if (linha === -1) return; // nome repetido dentro do próprio RH
    const atual = sheet.getRange(linha, 1, 1, 5).getValues()[0];
    const emailAtual = sheet.getRange(linha, 12).getValue();
    const novaAdmissao = p.dataAdmissao ? new Date(p.dataAdmissao + 'T00:00:00') : atual[1];
    const novaFuncao   = p.funcao || atual[3];
    const novoApelido  = atual[4] || p.apelido;
    const novoEmail    = emailAtual || p.email || '';
    sheet.getRange(linha, 2, 1, 4).setValues([[novaAdmissao, atual[2], novaFuncao, novoApelido]]);
    sheet.getRange(linha, 12).setValue(novoEmail);
    atualizados++;
  });

  // Uma escrita só pra todas as linhas novas — bem mais rápido que appendRow
  // por pessoa, o que também encurta a janela pra alguém clicar de novo.
  if (novas.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, novas.length, novas[0].length).setValues(novas);
  }

  return {
    funcionarios: getFuncionariosImpl_(),
    adicionados: adicionados,
    atualizados: atualizados,
    duplicatasRemovidas: duplicatasRemovidas
  };
}

// Apaga cadastros repetidos (mesmo nome normalizado), mantendo a primeira
// linha de cada nome. As duplicatas nasceram de cliques simultâneos no
// Sincronizar, então são sempre as linhas mais novas — appendadas depois da
// original, sem marcação de rastreamento feita à mão.
function removerFuncionariosDuplicados_(sheet) {
  const rows = sheet.getDataRange().getValues();
  const vistos = {};
  const excluir = [];
  for (let i = 1; i < rows.length; i++) {
    const chave = norm_(rows[i][0]);
    if (!chave) continue;
    if (vistos[chave]) excluir.push(i + 1);
    else vistos[chave] = true;
  }
  // De baixo pra cima, senão cada exclusão deslocaria os índices seguintes.
  for (let k = excluir.length - 1; k >= 0; k--) {
    sheet.deleteRow(excluir[k]);
  }
  return excluir.length;
}

const FERIAS_HEADERS_ = ['Pessoa', 'Data_Inicio', 'Data_Fim', 'Observacao', 'Criado_Por', 'Criado_Em', 'ID', 'Periodo_Aquisitivo_ID'];

function getFeriasSheet_() {
  const ss = getOpSS_();
  let sheet = ss.getSheetByName(ABA_FERIAS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_FERIAS);
    sheet.appendRow(FERIAS_HEADERS_);
    return sheet;
  }
  // Periodo_Aquisitivo_ID diz de qual período aquisitivo esses dias são
  // abatidos. Fica em branco nos lançamentos antigos (e em qualquer um que o
  // usuário não queira abater de nenhum período) — nesse caso os dias
  // aparecem no histórico mas não entram no saldo de nenhum período.
  const lastCol = sheet.getLastColumn();
  if (lastCol < FERIAS_HEADERS_.length) {
    sheet.getRange(1, lastCol + 1, 1, FERIAS_HEADERS_.length - lastCol)
      .setValues([FERIAS_HEADERS_.slice(lastCol)]);
  }
  return sheet;
}

// =============================================================================
// PERÍODOS AQUISITIVOS
// =============================================================================
// Cada período aquisitivo são os 12 meses trabalhados que dão direito a 30
// dias de férias; a data limite de gozo (quando esses dias têm que ser
// usados, sob risco de pagamento em dobro) cai 11 meses após o fim do
// período aquisitivo. Todos os períodos são cadastrados manualmente — o
// sistema não gera nenhum sozinho a partir da data de admissão.
// A coluna Data_Limite_Gozo guarda um limite escrito à mão; quando está em
// branco, vale o cálculo de MESES_ATE_LIMITE_GOZO_.

const MESES_ATE_LIMITE_GOZO_ = 11;

const PERIODOS_AQUISITIVOS_HEADERS_ = ['Pessoa', 'Data_Inicio', 'Data_Fim', 'ID', 'Data_Limite_Gozo'];

function getPeriodosAquisitivosSheet_() {
  const ss = getOpSS_();
  let sheet = ss.getSheetByName(ABA_PERIODOS_AQUISITIVOS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_PERIODOS_AQUISITIVOS);
    sheet.appendRow(PERIODOS_AQUISITIVOS_HEADERS_);
    migrarPeriodosAquisitivosLegado_(sheet);
    return sheet;
  }
  const lastCol = sheet.getLastColumn();
  if (lastCol < PERIODOS_AQUISITIVOS_HEADERS_.length) {
    sheet.getRange(1, lastCol + 1, 1, PERIODOS_AQUISITIVOS_HEADERS_.length - lastCol)
      .setValues([PERIODOS_AQUISITIVOS_HEADERS_.slice(lastCol)]);
  }
  return sheet;
}

// Data limite de gozo de um período: o que foi escrito à mão, ou o cálculo
// padrão (fim do aquisitivo + MESES_ATE_LIMITE_GOZO_) quando não há override.
function limiteGozoDe_(dataFimIso, overrideIso) {
  if (overrideIso) return overrideIso;
  if (!dataFimIso) return '';
  return fmtData_(addMeses_(new Date(dataFimIso + 'T00:00:00'), MESES_ATE_LIMITE_GOZO_));
}

// Migração única: converte o antigo override de ciclo (colunas
// Ciclo_Inicio_Override/Ciclo_Fim_Override da aba FUNCIONARIOS, que guardava
// o período concessivo — a janela de uso, não o aquisitivo) num período
// aquisitivo equivalente na aba nova, pra quem já tinha um customizado.
// O início do período aquisitivo é deduzido subtraindo 12 meses do início do
// concessivo salvo (que corresponde ao fim do aquisitivo) — 12 aqui é a
// duração do próprio período aquisitivo, não tem relação com o prazo de gozo.
function migrarPeriodosAquisitivosLegado_(sheet) {
  const funcSheet = getOpSS_().getSheetByName(ABA_FUNCIONARIOS);
  if (!funcSheet) return;
  const rows = funcSheet.getDataRange().getValues();
  const novas = [];
  for (let i = 1; i < rows.length; i++) {
    const nome = String(rows[i][0] || '').trim();
    const overrideInicio = rows[i][6];
    if (!nome || !overrideInicio) continue;
    const aquisitivoFim = new Date(overrideInicio);
    const aquisitivoInicio = addMeses_(aquisitivoFim, -12);
    novas.push([nome, aquisitivoInicio, aquisitivoFim, Utilities.getUuid()]);
  }
  if (novas.length) sheet.getRange(2, 1, novas.length, 4).setValues(novas);
}

function periodoAquisitivoFromRow_(r) {
  const dataFim = fmtData_(r[2]);
  const limiteOverride = fmtData_(r[4]);
  return {
    id: String(r[3] || ''),
    pessoa: String(r[0] || '').trim(),
    dataInicio: fmtData_(r[1]),
    dataFim: dataFim,
    dataLimiteGozo: limiteGozoDe_(dataFim, limiteOverride),
    limiteCustomizado: !!limiteOverride
  };
}

// Sem checagem de admin — usada pelo admin (sidebar de Férias) e
// internamente por Minhas Férias (aberto a qualquer funcionário, que só
// enxerga os próprios períodos depois de filtrar por nome).
function getPeriodosAquisitivosImpl_() {
  const rows = getPeriodosAquisitivosSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0] || !rows[i][1] || !rows[i][2]) continue;
    out.push(periodoAquisitivoFromRow_(rows[i]));
  }
  return out;
}

function getPeriodosAquisitivos(token) {
  requirePagina_(token, 'funcionarios');
  return getPeriodosAquisitivosImpl_();
}

function savePeriodoAquisitivo(token, p) {
  requirePagina_(token, 'funcionarios');
  const sheet = getPeriodosAquisitivosSheet_();

  const pessoa = String(p.pessoa || '').trim();
  const dataInicio = String(p.dataInicio || '').trim();
  const dataFim = String(p.dataFim || '').trim();
  if (!pessoa || !dataInicio || !dataFim) throw new Error('Informe a pessoa e o período.');
  if (dataFim < dataInicio) throw new Error('O fim não pode ser antes do início.');
  const ini = new Date(dataInicio + 'T00:00:00');
  const fim = new Date(dataFim + 'T00:00:00');

  // Só guarda a data limite quando ela difere do cálculo padrão: assim, quem
  // não mexeu no campo continua acompanhando a regra automaticamente se o
  // fim do período aquisitivo mudar depois.
  const limiteInformado = String(p.dataLimiteGozo || '').trim();
  const limiteOverride = (limiteInformado && limiteInformado !== limiteGozoDe_(dataFim, ''))
    ? new Date(limiteInformado + 'T00:00:00')
    : '';
  if (limiteInformado && limiteInformado < dataFim) throw new Error('A data limite de gozo não pode ser anterior ao fim do período aquisitivo.');

  if (p.id) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][3]) === String(p.id)) {
        sheet.getRange(i + 1, 1, 1, 3).setValues([[pessoa, ini, fim]]);
        sheet.getRange(i + 1, 5).setValue(limiteOverride);
        return getFeriasPagina(token);
      }
    }
    throw new Error('Período aquisitivo não encontrado.');
  }
  sheet.appendRow([pessoa, ini, fim, Utilities.getUuid(), limiteOverride]);
  return getFeriasPagina(token);
}

function deletePeriodoAquisitivo(token, id) {
  requirePagina_(token, 'funcionarios');
  const sheet = getPeriodosAquisitivosSheet_();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][3]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return getFeriasPagina(token);
}

// =============================================================================
// SOLICITAÇÕES DE FÉRIAS (autoatendimento do funcionário + aprovação)
// =============================================================================
// Qualquer funcionário autenticado cujo e-mail esteja vinculado a um
// registro no Cadastro pode solicitar um período de férias. Fica "Pendente"
// até um admin aprovar (o que cria automaticamente o período em FERIAS,
// contando no saldo) ou recusar (com motivo).

function getSolicitacoesSheet_() {
  const ss = getOpSS_();
  let sheet = ss.getSheetByName(ABA_SOLICITACOES_FERIAS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_SOLICITACOES_FERIAS);
    sheet.appendRow(['Pessoa', 'Email', 'Data_Inicio', 'Data_Fim', 'Observacao', 'Status', 'Data_Solicitacao', 'Aprovado_Por', 'Data_Decisao', 'Motivo_Recusa', 'ID']);
  }
  return sheet;
}

function solicitacaoFromRow_(r) {
  const dataInicio = fmtData_(r[2]);
  const dataFim = fmtData_(r[3]);
  return {
    id: String(r[10] || ''),
    pessoa: String(r[0] || '').trim(),
    email: String(r[1] || '').trim().toLowerCase(),
    dataInicio: dataInicio,
    dataFim: dataFim,
    dias: diasEntre_(dataInicio, dataFim),
    observacao: String(r[4] || '').trim(),
    status: String(r[5] || '').trim() || 'Pendente',
    dataSolicitacao: fmtDataHora_(r[6]),
    aprovadoPor: String(r[7] || '').trim(),
    dataDecisao: fmtDataHora_(r[8]),
    motivoRecusa: String(r[9] || '').trim()
  };
}

function getSolicitacoesImpl_() {
  const rows = getSolicitacoesSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0] || !rows[i][2] || !rows[i][3]) continue;
    out.push(solicitacaoFromRow_(rows[i]));
  }
  out.sort(function(a, b) { return (b.dataSolicitacao || '').localeCompare(a.dataSolicitacao || ''); });
  return out;
}

// Todas as solicitações — visão do admin, pra aprovar/recusar.
function getSolicitacoes(token) {
  requirePagina_(token, 'funcionarios');
  return getSolicitacoesImpl_();
}

function resolverFuncionarioPorEmail_(email) {
  const emailNorm = norm_(email);
  if (!emailNorm) return null;
  return getFuncionariosImpl_().filter(function(f) { return f.email && norm_(f.email) === emailNorm; })[0] || null;
}

// Funcionário solicita um período de férias — vira uma linha "Pendente" até
// um admin decidir. Resolve automaticamente quem é a pessoa a partir do
// e-mail da sessão (não do nome digitado), pra ninguém solicitar em nome de
// outra pessoa.
function solicitarFerias(token, payload) {
  const user = requireUser_(token);
  const meu = resolverFuncionarioPorEmail_(user.email);
  if (!meu) throw new Error('Seu e-mail (' + user.email + ') não está vinculado a nenhum funcionário no Cadastro. Peça pro administrador vincular.');

  const dataInicio = String(payload.dataInicio || '').trim();
  const dataFim = String(payload.dataFim || '').trim();
  if (!dataInicio || !dataFim) throw new Error('Informe o período desejado.');
  if (dataFim < dataInicio) throw new Error('A data final não pode ser anterior à inicial.');
  const observacao = String(payload.observacao || '').trim();

  const sheet = getSolicitacoesSheet_();
  sheet.appendRow([
    meu.nome, user.email, new Date(dataInicio + 'T00:00:00'), new Date(dataFim + 'T00:00:00'),
    observacao, 'Pendente', new Date(), '', '', '', Utilities.getUuid()
  ]);
  return getMinhaFeriasPagina(token);
}

// periodoAquisitivoId: de qual período aquisitivo esses dias serão abatidos.
// Quem aprova escolhe na hora; em branco significa "registra as férias mas
// não abate de nenhum período" (aparece marcado no histórico da pessoa).
function aprovarSolicitacao(token, id, periodoAquisitivoId) {
  const admin = requirePagina_(token, 'funcionarios');
  const sheet = getSolicitacoesSheet_();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][10]) === String(id)) {
      const pessoa = String(rows[i][0] || '').trim();
      const dataInicio = fmtData_(rows[i][2]);
      const dataFim = fmtData_(rows[i][3]);
      const observacao = String(rows[i][4] || '').trim();

      sheet.getRange(i + 1, 6).setValue('Aprovado');
      sheet.getRange(i + 1, 8).setValue(admin.email);
      sheet.getRange(i + 1, 9).setValue(new Date());

      // Aprovar já registra o período como férias efetivamente tiradas, pra
      // contar automaticamente no saldo — sem precisar lançar de novo à mão.
      getFeriasSheet_().appendRow([
        pessoa, new Date(dataInicio + 'T00:00:00'), new Date(dataFim + 'T00:00:00'), observacao,
        admin.email, new Date(), Utilities.getUuid(), String(periodoAquisitivoId || '').trim()
      ]);

      return getSolicitacoesImpl_();
    }
  }
  throw new Error('Solicitação não encontrada.');
}

function recusarSolicitacao(token, id, motivo) {
  const admin = requirePagina_(token, 'funcionarios');
  const sheet = getSolicitacoesSheet_();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][10]) === String(id)) {
      sheet.getRange(i + 1, 6).setValue('Recusado');
      sheet.getRange(i + 1, 8).setValue(admin.email);
      sheet.getRange(i + 1, 9).setValue(new Date());
      sheet.getRange(i + 1, 10).setValue(String(motivo || '').trim());
      return getSolicitacoesImpl_();
    }
  }
  throw new Error('Solicitação não encontrada.');
}

// Página de autoatendimento do funcionário: período aquisitivo, dias
// restantes, histórico de férias tiradas e as próprias solicitações — nunca
// depende do flag rastreadoFerias (isso é só sobre o que o admin optou por
// acompanhar na tela dele; o funcionário sempre vê o próprio saldo real).
function getMinhaFeriasPagina(token) {
  const user = requireUser_(token);
  const meu = resolverFuncionarioPorEmail_(user.email);
  if (!meu) {
    return { vinculado: false, email: user.email, nome: user.nome || '' };
  }

  const ferias = getFeriasImpl_();
  const periodosManuais = getPeriodosAquisitivosImpl_();

  const saldos = periodosDaPessoa_(meu.nome, ferias, periodosManuais)
    .map(function(per) { return Object.assign({ dataAdmissao: meu.dataAdmissao }, per); });
  const historico = ferias
    .filter(function(f) { return norm_(f.pessoa) === norm_(meu.nome); })
    .sort(function(a, b) { return (b.dataInicio || '').localeCompare(a.dataInicio || ''); });
  const solicitacoes = getSolicitacoesImpl_()
    .filter(function(s) { return norm_(s.email) === norm_(user.email); });

  return {
    vinculado: true,
    nome: meu.nome, funcao: meu.funcao, apelido: meu.apelido, dataAdmissao: meu.dataAdmissao,
    saldos: saldos, historico: historico, solicitacoes: solicitacoes
  };
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
    criadoEm: fmtDataHora_(r[5]),
    periodoAquisitivoId: String(r[7] || '').trim()
  };
}

// Sem checagem de admin — usada pela aba Férias (admin), pelo Dashboard e
// por Minhas Férias (ambos abertos a qualquer funcionário autenticado, que
// só enxergam o histórico da própria pessoa depois de filtrar por nome).
function getFeriasImpl_() {
  const rows = getFeriasSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0] || !rows[i][1]) continue;
    out.push(feriaFromRow_(rows[i]));
  }
  out.sort(function(a, b) { return (b.dataInicio || '').localeCompare(a.dataInicio || ''); });
  return out;
}

function getFerias(token) {
  requirePagina_(token, 'funcionarios');
  return getFeriasImpl_();
}

// Cria/edita um período de férias já tirado — lançamento manual, feito pelo
// admin (o fluxo do funcionário é sempre via solicitarFerias + aprovação).
function saveFeria(token, f) {
  const user  = requirePagina_(token, 'funcionarios');
  const sheet = getFeriasSheet_();

  const pessoa     = String(f.pessoa || '').trim();
  const dataInicio = String(f.dataInicio || '').trim();
  const dataFim    = String(f.dataFim || '').trim();
  if (!pessoa || !dataInicio || !dataFim) throw new Error('Informe a pessoa e o período.');
  if (dataFim < dataInicio) throw new Error('A data final não pode ser anterior à data inicial.');
  const observacao = String(f.observacao || '').trim();
  const periodoAquisitivoId = String(f.periodoAquisitivoId || '').trim();
  const ini = new Date(dataInicio + 'T00:00:00');
  const fim = new Date(dataFim + 'T00:00:00');

  if (f.id) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][6]) === String(f.id)) {
        sheet.getRange(i + 1, 1, 1, 4).setValues([[pessoa, ini, fim, observacao]]);
        sheet.getRange(i + 1, 8).setValue(periodoAquisitivoId);
        return getFerias(token);
      }
    }
    throw new Error('Período de férias não encontrado.');
  }

  sheet.appendRow([pessoa, ini, fim, observacao, user.email, new Date(), Utilities.getUuid(), periodoAquisitivoId]);
  return getFerias(token);
}

function deleteFeria(token, id) {
  requirePagina_(token, 'funcionarios');
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

// Para cada pessoa (cadastrada em FUNCIONARIOS e/ou com férias lançadas),
// só os dados cadastrais + flags de rastreamento — usado por Cadastro,
// Admissão e Demissão. O saldo de férias em si (períodos aquisitivos) está
// em getFeriasSaldos, que é o que a sub-aba Férias usa.
function getFeriasResumo(token) {
  requirePagina_(token, 'funcionarios');
  const funcionarios = getFuncionariosImpl_();
  const ferias = getFeriasImpl_();

  const pessoas = {};
  funcionarios.forEach(function(f) {
    pessoas[norm_(f.nome)] = {
      nome: f.nome, dataAdmissao: f.dataAdmissao, dataDemissao: f.dataDemissao, funcao: f.funcao, apelido: f.apelido,
      rastreadoFerias: f.rastreadoFerias, rastreadoAdmissao: f.rastreadoAdmissao, rastreadoDemissao: f.rastreadoDemissao
    };
  });
  // Quem tem férias lançadas mas não está no Cadastro entra aqui só pra
  // alimentar listas de nomes conhecidos — nunca como rastreado, senão
  // apareceria sozinho nas sub-abas sem ninguém ter adicionado.
  ferias.forEach(function(f) {
    const k = norm_(f.pessoa);
    if (!pessoas[k]) pessoas[k] = { nome: f.pessoa, dataAdmissao: '', dataDemissao: '', funcao: '', apelido: '', rastreadoFerias: false, rastreadoAdmissao: false, rastreadoDemissao: false };
  });

  return Object.keys(pessoas).map(function(k) {
    const p = pessoas[k];
    return {
      nome: p.nome,
      dataAdmissao: p.dataAdmissao,
      dataDemissao: p.dataDemissao,
      funcao: p.funcao,
      apelido: p.apelido,
      rastreadoFerias: !!p.rastreadoFerias,
      rastreadoAdmissao: !!p.rastreadoAdmissao,
      rastreadoDemissao: !!p.rastreadoDemissao
    };
  }).sort(function(a, b) { return a.nome.localeCompare(b.nome, 'pt-BR'); });
}

// Soma os dias de férias abatidos de um período aquisitivo, separando o que
// já foi gozado (tirados) do que está só marcado pro futuro (agendados) —
// lançamento com início depois de hoje é agendamento; em andamento conta como
// tirado. O vínculo é explícito (coluna Periodo_Aquisitivo_ID da aba FERIAS,
// escolhida por quem lança as férias) — o sistema não tenta adivinhar por
// data a qual período cada férias pertence, porque isso errava sempre que as
// férias caíam fora da janela de gozo (antecipadas, por exemplo).
function diasDoPeriodo_(feriasPessoa, periodoId) {
  const saldo = { tirados: 0, agendados: 0 };
  if (!periodoId) return saldo;
  const hoje = fmtData_(new Date());
  feriasPessoa.forEach(function(f) {
    if (f.periodoAquisitivoId !== periodoId) return;
    if (f.dataInicio && f.dataInicio > hoje) saldo.agendados += f.dias;
    else saldo.tirados += f.dias;
  });
  return saldo;
}

function periodoAquisitivoComSaldo_(periodo, feriasPessoa) {
  const dias = diasDoPeriodo_(feriasPessoa, periodo.id);
  return {
    id: periodo.id || '',
    manual: !!periodo.id,
    periodoAquisitivoInicio: periodo.dataInicio,
    periodoAquisitivoFim: periodo.dataFim,
    dataLimiteGozo: periodo.dataLimiteGozo,
    limiteCustomizado: !!periodo.limiteCustomizado,
    diasTirados: dias.tirados,
    diasAgendados: dias.agendados,
    diasRestantes: Math.max(0, 30 - dias.tirados - dias.agendados)
  };
}

// Calcula os períodos aquisitivos (com saldo) de UMA pessoa — reaproveitado
// tanto pelo saldo geral do admin (getFeriasSaldos) quanto pela própria
// visão do funcionário (getMinhaFeriasPagina). Não há cálculo automático a
// partir da data de admissão — só entram aqui os períodos que alguém
// cadastrou manualmente em PERIODOS_AQUISITIVOS (o admin decide quando abrir
// um período novo, não o sistema sozinho).
function periodosDaPessoa_(nome, ferias, periodosManuais) {
  const k = norm_(nome);
  const feriasPessoa = ferias.filter(function(f) { return norm_(f.pessoa) === k; });
  const manuais = periodosManuais.filter(function(m) { return norm_(m.pessoa) === k; });

  let periodos = manuais
    .map(function(m) { return periodoAquisitivoComSaldo_(m, feriasPessoa); })
    .sort(function(a, b) { return a.periodoAquisitivoInicio.localeCompare(b.periodoAquisitivoInicio); });

  if (!periodos.length) {
    periodos = [{ id: '', manual: false, periodoAquisitivoInicio: '', periodoAquisitivoFim: '', dataLimiteGozo: '', diasTirados: null, diasAgendados: null, diasRestantes: null }];
  }
  return periodos;
}

// Saldo de férias por pessoa rastreada em Férias (admin) — um item por
// período aquisitivo em aberto (pode ser mais de um por pessoa).
function getFeriasSaldos(token) {
  requirePagina_(token, 'funcionarios');
  const funcionarios = getFuncionariosImpl_();
  const ferias = getFeriasImpl_();
  const periodosManuais = getPeriodosAquisitivosImpl_();

  // Só quem está no Cadastro entra aqui: ter férias lançadas não coloca
  // ninguém automaticamente nesta aba — a inclusão é sempre manual.
  const pessoas = {};
  funcionarios.forEach(function(f) {
    pessoas[norm_(f.nome)] = { nome: f.nome, funcao: f.funcao, apelido: f.apelido, dataAdmissao: f.dataAdmissao, dataDemissao: f.dataDemissao, rastreadoFerias: f.rastreadoFerias };
  });

  const out = [];
  Object.keys(pessoas).forEach(function(k) {
    const p = pessoas[k];
    if (!p.rastreadoFerias) return;
    periodosDaPessoa_(p.nome, ferias, periodosManuais).forEach(function(per) {
      out.push({
        nome: p.nome, funcao: p.funcao, apelido: p.apelido, dataAdmissao: p.dataAdmissao, dataDemissao: p.dataDemissao,
        id: per.id, manual: per.manual,
        periodoAquisitivoInicio: per.periodoAquisitivoInicio, periodoAquisitivoFim: per.periodoAquisitivoFim,
        dataLimiteGozo: per.dataLimiteGozo, limiteCustomizado: per.limiteCustomizado,
        diasTirados: per.diasTirados, diasAgendados: per.diasAgendados, diasRestantes: per.diasRestantes
      });
    });
  });

  out.sort(function(a, b) {
    return a.nome.localeCompare(b.nome, 'pt-BR') || (a.periodoAquisitivoInicio || '').localeCompare(b.periodoAquisitivoInicio || '');
  });
  return out;
}

// Marca/desmarca o rastreamento de uma pessoa numa das 3 sub-abas
// (Férias/Admissão/Demissão), independentemente das outras. Cadastrar
// alguém no Cadastro de Funcionários não mexe nesses flags — a inclusão em
// cada aba é sempre uma ação explícita e separada.
function setRastreamento(token, nome, aba, valor) {
  requirePagina_(token, 'funcionarios');
  const col = aba === 'ferias' ? 9 : aba === 'admissao' ? 10 : aba === 'demissao' ? 11 : null;
  if (!col) throw new Error('Aba inválida: ' + aba);
  const sheet = getFuncionariosSheet_();
  const rows = sheet.getDataRange().getValues();
  const nomeNorm = norm_(nome);
  for (let i = 1; i < rows.length; i++) {
    if (norm_(rows[i][0]) === nomeNorm) {
      sheet.getRange(i + 1, col).setValue(!!valor);
      return getFeriasPagina(token);
    }
  }
  throw new Error('Funcionário não encontrado: ' + nome);
}

function getFeriasPagina(token) {
  requirePagina_(token, 'funcionarios');
  return {
    funcionarios: getFuncionariosImpl_(),
    ferias: getFeriasImpl_(),
    resumo: getFeriasResumo(token),
    feriasSaldos: getFeriasSaldos(token),
    periodosAquisitivos: getPeriodosAquisitivosImpl_(),
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
  requirePagina_(token, 'funcionarios');
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
  const user  = requirePagina_(token, 'funcionarios');
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

// =============================================================================
// PROJETOS (quadro estilo Trello)
// =============================================================================
// Duas abas: PROJETOS_COLUNAS guarda as raias do quadro (nome + ordem) e
// PROJETOS_CARTOES os cartões, cada um apontando para uma coluna. A posição
// dentro da coluna vem da coluna Ordem — quem arrasta um cartão na tela manda
// a lista inteira de IDs já na ordem nova, e o backend só regrava os índices.
// Nenhuma coluna é criada automaticamente: o quadro nasce vazio.

function getProjetosColunasSheet_() {
  const ss = getOpSS_();
  let sheet = ss.getSheetByName(ABA_PROJETOS_COLUNAS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_PROJETOS_COLUNAS);
    sheet.appendRow(['Nome', 'Ordem', 'ID']);
  }
  return sheet;
}

function getProjetosCartoesSheet_() {
  const ss = getOpSS_();
  let sheet = ss.getSheetByName(ABA_PROJETOS_CARTOES);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_PROJETOS_CARTOES);
    sheet.appendRow(['Coluna_ID', 'Titulo', 'Descricao', 'Responsavel', 'Prazo', 'Etiqueta', 'Etiqueta_Cor', 'Ordem', 'Criado_Por', 'Criado_Em', 'ID']);
  }
  return sheet;
}

function colunaProjetoFromRow_(r) {
  return {
    id: String(r[2] || ''),
    nome: String(r[0] || '').trim(),
    ordem: Number(r[1]) || 0
  };
}

function cartaoProjetoFromRow_(r) {
  return {
    id: String(r[10] || ''),
    colunaId: String(r[0] || '').trim(),
    titulo: String(r[1] || '').trim(),
    descricao: String(r[2] || '').trim(),
    responsavel: String(r[3] || '').trim(),
    prazo: fmtData_(r[4]),
    etiqueta: String(r[5] || '').trim(),
    etiquetaCor: String(r[6] || '').trim(),
    ordem: Number(r[7]) || 0,
    criadoPor: String(r[8] || '').trim(),
    criadoEm: fmtDataHora_(r[9])
  };
}

function getProjetosPagina(token) {
  requirePagina_(token, 'projetos');

  const colRows = getProjetosColunasSheet_().getDataRange().getValues();
  const colunas = [];
  for (let i = 1; i < colRows.length; i++) {
    if (!colRows[i][0]) continue;
    colunas.push(colunaProjetoFromRow_(colRows[i]));
  }
  colunas.sort(function(a, b) { return a.ordem - b.ordem; });

  const cardRows = getProjetosCartoesSheet_().getDataRange().getValues();
  const cartoes = [];
  for (let i = 1; i < cardRows.length; i++) {
    if (!cardRows[i][1]) continue; // sem título
    cartoes.push(cartaoProjetoFromRow_(cardRows[i]));
  }
  cartoes.sort(function(a, b) { return a.ordem - b.ordem; });

  return { colunas: colunas, cartoes: cartoes };
}

function saveColunaProjeto(token, c) {
  requirePagina_(token, 'projetos');
  const sheet = getProjetosColunasSheet_();
  const nome = String(c.nome || '').trim();
  if (!nome) throw new Error('Informe o nome da coluna.');

  const rows = sheet.getDataRange().getValues();
  if (c.id) {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][2]) === String(c.id)) {
        sheet.getRange(i + 1, 1).setValue(nome);
        return getProjetosPagina(token);
      }
    }
    throw new Error('Coluna não encontrada.');
  }

  // Nova coluna entra no fim do quadro.
  let maiorOrdem = 0;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) maiorOrdem = Math.max(maiorOrdem, Number(rows[i][1]) || 0);
  }
  sheet.appendRow([nome, maiorOrdem + 1, Utilities.getUuid()]);
  return getProjetosPagina(token);
}

// Recusa apagar coluna que ainda tem cartão: excluir junto apagaria trabalho
// registrado sem a pessoa perceber.
function deleteColunaProjeto(token, id) {
  requirePagina_(token, 'projetos');
  const cartoes = getProjetosCartoesSheet_().getDataRange().getValues();
  let usados = 0;
  for (let i = 1; i < cartoes.length; i++) {
    if (String(cartoes[i][0]) === String(id) && cartoes[i][1]) usados++;
  }
  if (usados) throw new Error('Esta coluna ainda tem ' + usados + ' cartão(ões). Mova ou exclua antes de apagar a coluna.');

  const sheet = getProjetosColunasSheet_();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return getProjetosPagina(token);
}

// idsOrdenados = todas as colunas na ordem nova, como a tela mostra.
function reordenarColunasProjeto(token, idsOrdenados) {
  requirePagina_(token, 'projetos');
  const sheet = getProjetosColunasSheet_();
  const rows = sheet.getDataRange().getValues();
  const posicao = {};
  (idsOrdenados || []).forEach(function(id, i) { posicao[String(id)] = i + 1; });
  for (let i = 1; i < rows.length; i++) {
    const nova = posicao[String(rows[i][2])];
    if (nova) sheet.getRange(i + 1, 2).setValue(nova);
  }
  return getProjetosPagina(token);
}

function saveCartaoProjeto(token, c) {
  const user = requirePagina_(token, 'projetos');
  const sheet = getProjetosCartoesSheet_();

  const titulo = String(c.titulo || '').trim();
  const colunaId = String(c.colunaId || '').trim();
  if (!titulo) throw new Error('Informe o título do cartão.');
  if (!colunaId) throw new Error('Escolha a coluna do cartão.');

  const descricao = String(c.descricao || '').trim();
  const responsavel = String(c.responsavel || '').trim();
  const prazo = c.prazo ? new Date(c.prazo + 'T00:00:00') : '';
  const etiqueta = String(c.etiqueta || '').trim();
  const etiquetaCor = String(c.etiquetaCor || '').trim();

  const rows = sheet.getDataRange().getValues();
  if (c.id) {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][10]) === String(c.id)) {
        sheet.getRange(i + 1, 1, 1, 7).setValues([[colunaId, titulo, descricao, responsavel, prazo, etiqueta, etiquetaCor]]);
        return getProjetosPagina(token);
      }
    }
    throw new Error('Cartão não encontrado.');
  }

  // Cartão novo entra no fim da coluna escolhida.
  let maiorOrdem = 0;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === colunaId) maiorOrdem = Math.max(maiorOrdem, Number(rows[i][7]) || 0);
  }
  sheet.appendRow([colunaId, titulo, descricao, responsavel, prazo, etiqueta, etiquetaCor, maiorOrdem + 1, user.email, new Date(), Utilities.getUuid()]);
  return getProjetosPagina(token);
}

function deleteCartaoProjeto(token, id) {
  requirePagina_(token, 'projetos');
  const sheet = getProjetosCartoesSheet_();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][10]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return getProjetosPagina(token);
}

// Chamada ao soltar um cartão: idsOrdenadosDestino são todos os cartões da
// coluna de destino já na ordem final que a tela mostra (incluindo o que
// acabou de chegar), então uma passada só resolve mover e reordenar.
function moverCartaoProjeto(token, cartaoId, colunaDestinoId, idsOrdenadosDestino) {
  requirePagina_(token, 'projetos');
  const sheet = getProjetosCartoesSheet_();
  const rows = sheet.getDataRange().getValues();

  const posicao = {};
  (idsOrdenadosDestino || []).forEach(function(id, i) { posicao[String(id)] = i + 1; });

  for (let i = 1; i < rows.length; i++) {
    const id = String(rows[i][10]);
    if (id === String(cartaoId)) sheet.getRange(i + 1, 1).setValue(String(colunaDestinoId));
    const nova = posicao[id];
    if (nova) sheet.getRange(i + 1, 8).setValue(nova);
  }
  return getProjetosPagina(token);
}

// Os seeds popularFeriasIniciais e popularFuncionariosIniciais foram removidos
// depois de executados (julho/2026): eles usavam nomes curtos ("Natasha",
// "Elaine"), e o cadastro migrou pra nomes completos — rodar de novo passaria
// pela checagem de duplicata sem casar nada e recriaria tudo em dobro. Além
// disso, como toda função global sem checagem de token, ficavam chamáveis por
// qualquer pessoa via google.script.run. O histórico está no git.

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
