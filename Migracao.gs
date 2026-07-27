// =============================================================================
// MIGRAÇÃO DE DADOS — Controle Operacional BOL
// 1. Abra o Apps Script da nova planilha (Extensões → Apps Script)
// 2. Cole este arquivo inteiro num arquivo .gs novo
// 3. Rode a função migrarDadosBOL()
// =============================================================================

function migrarDadosBOL() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];

  _criarAba_(ss, 'EVENTOS',         ['Data_Inicio','Data_Fim','Titulo','Categoria','Observacao','Criado_Por','Criado_Em']);
  _criarAba_(ss, 'ROTINAS',         ['Frequencia','Tarefa','Responsavel','Ativo']);
  _criarAba_(ss, 'ENTREGAS',        ['Atividade','Setor','Prazo_Descricao','Mes','Ano','Status','Atualizado_Em','Atualizado_Por']);
  _criarAba_(ss, 'HORARIOS',        ['Nome','Horario_Padrao','Horario_Sexta','Ativo']);
  _criarAba_(ss, 'ESCALA_SABADO',   ['Mes','Ano','Data_Sabado','Pessoa1','Pessoa2','Pessoa3','Observacao']);
  _criarAba_(ss, 'AJUSTES_HORARIO', ['Atividade','Data','Status','Mes','Ano','Registrado_Por','Registrado_Em']);
  log.push('✅ Abas criadas');

  _popularEventos_(ss);     log.push('✅ EVENTOS');
  _popularRotinas_(ss);     log.push('✅ ROTINAS');
  _popularEntregas_(ss);    log.push('✅ ENTREGAS');
  _popularHorarios_(ss);    log.push('✅ HORARIOS');
  _popularEscala_(ss);      log.push('✅ ESCALA_SABADO');
  _popularAjustes_(ss);     log.push('✅ AJUSTES_HORARIO');

  // Remove a aba vazia padrão "Plan1" ou "Sheet1" se existir
  ['Plan1','Sheet1','Planilha1'].forEach(function(nome) {
    const a = ss.getSheetByName(nome);
    if (a && ss.getSheets().length > 1) try { ss.deleteSheet(a); } catch(e) {}
  });

  SpreadsheetApp.getUi().alert('Migração concluída!\n\n' + log.join('\n'));
}

// =============================================================================
// HELPERS
// =============================================================================

function _criarAba_(ss, nome, cab) {
  let sheet = ss.getSheetByName(nome);
  if (sheet) {
    sheet.clearContents();
  } else {
    sheet = ss.insertSheet(nome);
  }
  const range = sheet.getRange(1, 1, 1, cab.length);
  range.setValues([cab])
       .setFontWeight('bold')
       .setBackground('#0a1628')
       .setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  return sheet;
}

// Converte "DD/MM/YYYY" em Date; retorna '' se vazio ou inválido
function _d_(str) {
  if (!str || str.trim() === '') return '';
  const p = str.trim().split('/');
  if (p.length !== 3) return '';
  const d = new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10));
  return isNaN(d.getTime()) ? '' : d;
}

function _escrever_(ss, abaName, dados, colsData) {
  if (!dados.length) return;
  const sheet = ss.getSheetByName(abaName);
  sheet.getRange(2, 1, dados.length, dados[0].length).setValues(dados);
  // Formata colunas de data (base 1)
  if (colsData) {
    colsData.forEach(function(c) {
      sheet.getRange(2, c, dados.length, 1).setNumberFormat('dd/MM/yyyy');
    });
  }
}

// =============================================================================
// EVENTOS
// =============================================================================

function _popularEventos_(ss) {
  // [Data_Inicio_str, Data_Fim_str, Titulo, Categoria, Observacao]
  const SRC = [
    // ─── 2025 ──────────────────────────────────────────────────────────────
    ['02/01/2025', '',           'Turmas de imersão',                         'OPERACIONAL',     ''],
    ['',           '',           'Tabela nova BRASAS',                         'OPERACIONAL',     'Janeiro 2025'],
    ['',           '',           'Volta às aulas',                             'OPERACIONAL',     'Janeiro 2025'],
    ['27/02/2025', '',           'BRASAS Start',                               'BRASAS_START',    ''],
    ['',           '',           'Início das turmas novas',                    'OPERACIONAL',     'Fevereiro 2025'],
    ['12/03/2025', '',           'BRASAS Start',                               'BRASAS_START',    ''],
    ['10/04/2025', '30/04/2025', 'Special Classes',                            'SPECIAL_CLASSES', ''],
    ['25/04/2025', '',           'BRASAS DAY - Administrativo e Apoio',        'BRASAS_DAY',      ''],
    ['28/05/2025', '29/05/2025', 'BFDF',                                       'EVENTO',          ''],
    ['30/05/2025', '',           'BRASAS DAY - Online',                        'BRASAS_DAY',      ''],
    ['01/06/2025', '31/08/2025', 'Novo período de meta e promoção',            'OPERACIONAL',     ''],
    ['02/06/2025', '06/06/2025', 'Férias Aline',                               'FERIAS',          ''],
    ['',           '',           'Processo Seletivo',                          'OPERACIONAL',     'Junho 2025'],
    ['',           '',           'Aplicação TOEIC BOL - TJ',                   'OPERACIONAL',     'Junho 2025'],
    ['04/06/2025', '26/06/2025', 'Special Classes',                            'SPECIAL_CLASSES', ''],
    ['27/06/2025', '',           'Festa Junina EC',                            'EVENTO',          ''],
    ['08/07/2025', '',           'BRASAS Start',                               'BRASAS_START',    ''],
    ['11/07/2025', '',           'BRASAS DAY - Comercial',                     'BRASAS_DAY',      ''],
    ['16/07/2025', '',           'Nova Contratação - Rosana',                  'OPERACIONAL',     ''],
    ['16/07/2025', '31/07/2025', 'Ajuste de Horário - Aline',                  'OPERACIONAL',     ''],
    ['16/07/2025', '19/07/2025', 'Bring a Friend',                             'EVENTO',          ''],
    ['',           '',           'Recesso de Julho',                           'RECESSO',         'Julho 2025'],
    ['04/08/2025', '',           'BRASAS Start',                               'BRASAS_START',    ''],
    ['29/08/2025', '',           'Jam Session BOL',                            'EVENTO',          ''],
    ['01/09/2025', '31/10/2025', 'Novo período de meta e promoção',            'OPERACIONAL',     ''],
    ['01/09/2025', '',           'Novo Horário - Bruna 10h',                   'OPERACIONAL',     ''],
    ['01/09/2025', '15/09/2025', 'Férias Mel',                                 'FERIAS',          ''],
    ['02/09/2025', '16/09/2025', 'Férias Nayara',                              'FERIAS',          ''],
    ['',           '',           'Novo Horário - Nayara 08h',                  'OPERACIONAL',     'Setembro 2025'],
    ['11/09/2025', '30/09/2025', 'Férias Nathy',                               'FERIAS',          ''],
    ['11/09/2025', '25/09/2025', 'Special Classes',                            'SPECIAL_CLASSES', ''],
    ['01/10/2025', '',           'Special Classes',                            'SPECIAL_CLASSES', ''],
    ['07/10/2025', '',           'BRASAS Start - PFF',                         'BRASAS_START',    ''],
    ['',           '',           'Curso Preparatório de Certificações',         'OPERACIONAL',     'Outubro 2025'],
    ['24/10/2025', '',           'BRASAS DAY - Online',                        'BRASAS_DAY',      ''],
    ['31/10/2025', '',           'On boarding',                                'OPERACIONAL',     ''],
    ['01/11/2025', '30/11/2025', 'Novo período de meta e promoção',            'OPERACIONAL',     ''],
    ['01/11/2025', '30/11/2025', 'Black Friday Imersão BOL',                   'OPERACIONAL',     ''],
    ['06/11/2025', '12/11/2025', 'Férias Bruna',                               'FERIAS',          ''],
    ['12/11/2025', '',           'Special Classes',                            'SPECIAL_CLASSES', ''],
    ['18/11/2025', '',           'Special Classes',                            'SPECIAL_CLASSES', ''],
    ['28/11/2025', '',           'On boarding',                                'OPERACIONAL',     ''],
    ['01/12/2025', '31/03/2026', 'Novo período de meta e promoção',            'OPERACIONAL',     ''],
    ['05/12/2025', '',           'Formatura',                                  'EVENTO',          ''],
    ['08/12/2025', '14/12/2025', 'Férias ADM (Elaine)',                        'FERIAS',          ''],
    ['16/12/2025', '',           'Último dia de aula do ano',                  'OPERACIONAL',     ''],
    ['17/12/2025', '15/01/2026', 'Férias Professores',                         'FERIAS',          ''],
    ['17/12/2025', '05/01/2026', 'Férias ADM (Nay, Rafa e Bruna)',             'FERIAS',          ''],
    ['17/12/2025', '02/01/2026', 'Férias Coordenação (Ge e Mel)',              'FERIAS',          ''],
    ['18/12/2025', '',           'Festa EC',                                   'EVENTO',          ''],
    // ─── 2026 ──────────────────────────────────────────────────────────────
    ['02/01/2026', '08/01/2026', 'Turmas de imersão',                          'OPERACIONAL',     ''],
    ['23/01/2026', '',           'Treinamento com a secretária BOL',            'TREINAMENTO',     'Redução do Tempo de Atendimento'],
    ['26/01/2026', '',           'Tabela nova BRASAS',                          'OPERACIONAL',     ''],
    ['28/01/2026', '',           'Volta às aulas',                             'OPERACIONAL',     ''],
    ['03/02/2026', '',           'BRASAS Start',                               'BRASAS_START',    ''],
    ['04/02/2026', '07/02/2026', 'Início das turmas novas',                    'OPERACIONAL',     ''],
    ['03/03/2026', '',           'BRASAS Start',                               'BRASAS_START',    ''],
    ['04/03/2026', '09/03/2026', 'Início das turmas novas',                    'OPERACIONAL',     ''],
    ['12/03/2026', '',           'Dia Internacional da Mulher no EC',           'EVENTO',          ''],
    ['09/03/2026', '18/03/2026', 'Férias Aline',                               'FERIAS',          ''],
    ['16/03/2026', '28/03/2026', 'Férias Elaine',                              'FERIAS',          ''],
    ['16/03/2026', '29/03/2026', 'Férias Ge',                                  'FERIAS',          ''],
    ['',           '',           'Treinamento com a Coordenação',              'TREINAMENTO',     'Março 2026'],
    ['31/03/2026', '',           'Special Classes',                            'SPECIAL_CLASSES', ''],
    ['06/04/2026', '11/04/2026', 'Férias Rafa',                                'FERIAS',          ''],
    ['08/04/2026', '',           'Special Classes',                            'SPECIAL_CLASSES', ''],
    ['14/04/2026', '',           'Special Classes',                            'SPECIAL_CLASSES', ''],
    ['29/04/2026', '',           'Special Classes',                            'SPECIAL_CLASSES', ''],
    ['11/05/2026', '15/05/2026', 'Férias Nat',                                 'FERIAS',          ''],
    ['18/05/2026', '31/05/2026', 'Férias Rô',                                  'FERIAS',          ''],
    ['27/05/2026', '28/05/2026', 'BFDF',                                       'EVENTO',          ''],
    ['01/06/2026', '10/06/2026', 'Férias Aline',                               'FERIAS',          ''],
    ['02/06/2026', '',           'BRASAS Start',                               'BRASAS_START',    ''],
    ['03/06/2026', '',           'Special Classes',                            'SPECIAL_CLASSES', ''],
    ['11/06/2026', '',           'Special Classes',                            'SPECIAL_CLASSES', ''],
    ['12/06/2026', '',           'BRASAS DAY - Online',                        'BRASAS_DAY',      ''],
    ['17/06/2026', '',           'Special Classes',                            'SPECIAL_CLASSES', ''],
    ['25/06/2026', '',           'Special Classes',                            'SPECIAL_CLASSES', ''],
    ['',           '',           'BRASAS Start',                               'BRASAS_START',    'Julho 2026'],
    ['',           '',           'Nova Contratação - Caren',                   'OPERACIONAL',     'Julho 2026'],
    ['',           '',           'Ajuste de Horário - Nay',                    'OPERACIONAL',     'Julho 2026'],
    ['',           '',           'Bring a Friend',                             'EVENTO',          'Julho 2026'],
    ['',           '',           'Recesso de Julho',                           'RECESSO',         'Julho 2026'],
  ];

  const agora = new Date();
  const rows  = SRC.map(function(e) {
    return [_d_(e[0]), _d_(e[1]), e[2], e[3], e[4], 'Migração', agora];
  });

  _escrever_(ss, 'EVENTOS', rows, [1, 2, 7]);
}

// =============================================================================
// ROTINAS
// =============================================================================

function _popularRotinas_(ss) {
  // [Frequencia, Tarefa, Responsavel, Ativo]
  const dados = [
    ['DIÁRIO', 'Suporte para o Atendimento', 'Todas', 'SIM'],
    ['DIÁRIO', 'Acertos enviados por e-mail pelo financeiro e secretaria', 'Todas', 'SIM'],
    ['DIÁRIO', 'Parcelamento, acerto interno, renegociação, estorno, devolução', 'Todas', 'SIM'],
    ['DIÁRIO', 'Solicitação de emissão de nota fiscal', 'Todas', 'SIM'],
    ['DIÁRIO', 'Conferência diária de recebimentos Vindi x Efí x Sophia', 'Todas', 'SIM'],
    ['DIÁRIO', 'My BRASAS: verificar alunos sem convite', 'Nayara', 'SIM'],
    ['DIÁRIO', 'Conferência do form de cancelados', 'Todas', 'SIM'],
    ['DIÁRIO', 'Verificar emails sem assinaturas no Board (Gmelius)', 'Todas', 'SIM'],
    ['DIÁRIO', 'Verificar email ADM - solicitação de compartilhamento', 'Nayara', 'SIM'],
    ['DIÁRIO', 'Verificar no Drive a pasta Meet Recordings (organizar gravações)', 'Todas', 'SIM'],

    ['SEMANAL', 'Verificar quantidade de vagas na turma (salas do sistema)', 'Todas', 'SIM'],
    ['SEMANAL', 'Atualizar planilha de PFF', 'Todas', 'SIM'],
    ['SEMANAL', 'Encerramento de turmas (de acordo com coordenação)', 'Todas', 'SIM'],
    ['SEMANAL', 'Preencher planilha de certificados', 'Elaine', 'SIM'],
    ['SEMANAL', 'Detratores', 'Todas', 'SIM'],
    ['SEMANAL', 'Agendar TN e matrículas de Cias (todas)', 'Todas', 'SIM'],
    ['SEMANAL', 'Abertura de turmas novas', 'Todas', 'SIM'],

    ['QUINZENAL', 'Horário Conferido (Todas)', 'Todas', 'SIM'],
    ['QUINZENAL', 'Duplicidades no drive das aulas', 'Todas', 'SIM'],

    ['MENSAL', 'Conferência e preenchimento da planilha de faturamento BOL', 'Todas', 'SIM'],
    ['MENSAL', 'Reunião alinhamento da equipe', 'Nathy / Aline', 'SIM'],
    ['MENSAL', 'Horas dos Professores', 'Nathy', 'SIM'],
    ['MENSAL', 'Verificar se os alunos cancelados foram "cancelados" na planilha de controle', 'Elaine', 'SIM'],
    ['MENSAL', 'Abrir/criar turmas', 'Todas', 'SIM'],
    ['MENSAL', 'Planilha de Conceitos', 'Todas', 'SIM'],
    ['MENSAL', 'Planilha de Cias PERMUTA', 'Todas', 'SIM'],
    ['MENSAL', 'Envio de planilha de conceito', 'Todas', 'SIM'],
    ['MENSAL', 'Gerar as mensalidades do mês seguinte', 'Todas', 'SIM'],
    ['MENSAL', 'Baixar mensalidade de CIAS', 'Todas', 'SIM'],
    ['MENSAL', 'Private (cobrança mensal)', 'Nayara', 'SIM'],
    ['MENSAL', 'Conferência Appsheet / Controle de alunos BOL', 'Todas', 'SIM'],
    ['MENSAL', 'Primeiro dia útil: conferir faturamento', 'Todas', 'SIM'],
    ['MENSAL', 'Verificar se todos os títulos recebidos estão com notas', 'Todas', 'SIM'],
    ['MENSAL', 'Sophia X Nota Carioca (Todos os Métodos)', 'Todas', 'SIM'],
    ['MENSAL', 'Alimentar a planilha do Operacional', 'Todas', 'SIM'],
    ['MENSAL', 'Alimentar a planilha do Financeiro', 'Todas', 'SIM'],
    ['MENSAL', 'Disponibilizar a escala de folgas', 'Todas', 'SIM'],

    ['TRIMESTRAL', 'Conferência AppSheet Split (Polos)', 'Todas', 'SIM'],

    ['SEMESTRAL', 'Avaliação de desempenho', 'Todas', 'SIM'],
    ['SEMESTRAL', 'Treinamento Atendimento', 'Todas', 'SIM'],

    ['ANUAL', 'Elaboração Férias', 'Todas', 'SIM'],
    ['ANUAL', 'Revisão de Tarefas', 'Todas', 'SIM'],
    ['ANUAL', 'Alterar a Tabela BOL no final de janeiro (após agendamento dos pagamentos)', 'Todas', 'SIM'],
    ['ANUAL', 'Alterar a nomenclatura do plano de pagamento e valor em Janeiro', 'Todas', 'SIM'],
    ['ANUAL', 'Verificar os feriados para o próximo ano', 'Todas', 'SIM'],
    ['ANUAL', 'Diário de classe Online', 'Todas', 'SIM'],
    ['ANUAL', 'Form banco de horas', 'Todas', 'SIM'],
    ['ANUAL', 'Enviar Circular de Renovação para os alunos', 'Todas', 'SIM'],
    ['ANUAL', 'Limpeza do Drive', 'Todas', 'SIM'],
  ];

  _escrever_(ss, 'ROTINAS', dados, []);
}

// =============================================================================
// ENTREGAS
// =============================================================================

function _popularEntregas_(ss) {
  // Status histórico 2025 (Jan=1 … Dez=12), dado original: Jan-Jul preenchidos
  // índice 0 = Janeiro
  const hist2025 = {
    'Cias':                      ['NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','EM ATRASO','','','','',''],
    'Bolsistas':                  ['NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','','','','',''],
    'Horário Conferido':          ['NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','','','','',''],
    'Conferência de Faturamento': ['NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','','','','',''],
    'Horas':                      ['NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','','','','',''],
    'Caju':                       ['NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','NO PRAZO','','','','',''],
    'Dropout':                    ['','','','','','','','','','','',''],
  };

  const atividades = [
    { nome: 'Cias',                      setor: 'Secretaria',   prazo: '1º dia útil' },
    { nome: 'Bolsistas',                  setor: 'Secretaria',   prazo: 'Dia 05' },
    { nome: 'Horário Conferido',          setor: 'Secretaria',   prazo: 'Quinzenal' },
    { nome: 'Conferência de Faturamento', setor: 'Operacional',  prazo: '1º dia útil' },
    { nome: 'Horas',                      setor: 'Nathy',        prazo: 'Dia 10' },
    { nome: 'Caju',                       setor: 'Nathy',        prazo: 'Dia 10' },
    { nome: 'Dropout',                    setor: 'Coordenação',  prazo: '1º dia útil' },
  ];

  const rows = [];

  // 2025 — pré-popular com status histórico
  atividades.forEach(function(a) {
    const statusMeses = hist2025[a.nome] || new Array(12).fill('');
    for (let mes = 1; mes <= 12; mes++) {
      rows.push([a.nome, a.setor, a.prazo, mes, 2025, statusMeses[mes - 1] || '', '', '']);
    }
  });

  // 2026 — em branco para a equipe preencher
  atividades.forEach(function(a) {
    for (let mes = 1; mes <= 12; mes++) {
      rows.push([a.nome, a.setor, a.prazo, mes, 2026, '', '', '']);
    }
  });

  _escrever_(ss, 'ENTREGAS', rows, [7]);
}

// =============================================================================
// HORÁRIOS DE SECRETARIA
// =============================================================================

function _popularHorarios_(ss) {
  const dados = [
    ['Elaine',  '07:00 - 16:00', '07:00 - 16:00', 'ATIVO'],
    ['Bruna',   '09:00 - 18:00', '08:00 - 17:00', 'ATIVO'],
    ['Nayara',  '11:00 - 20:00', '08:00 - 17:00', 'ATIVO'],
    ['Caren',   '12:30 - 21:30', '11:00 - 20:00', 'ATIVO'],
    ['Rafa',    '12:30 - 21:30', '11:00 - 20:00', 'ATIVO'],
    ['Rosana',  '12:30 - 21:30', '11:00 - 20:00', 'ATIVO'],
  ];
  _escrever_(ss, 'HORARIOS', dados, []);
}

// =============================================================================
// ESCALA DE SÁBADO
// =============================================================================

function _popularEscala_(ss) {
  // [Mes, Ano, Data_DD/MM/YYYY, Pessoa1, Pessoa2, Pessoa3, Observacao]
  const dados = [
    // ─── JANEIRO 2026 ───────────────────────────────────
    [1, 2026, '03/01/2026', 'Rosana',  'Nayara', '',     ''],
    [1, 2026, '10/01/2026', 'Elaine',  'Bruna',  '',     ''],
    [1, 2026, '17/01/2026', 'Rafa',    'Aline',  '',     ''],
    [1, 2026, '24/01/2026', 'Rosana',  'Nayara', '',     ''],
    [1, 2026, '31/01/2026', 'Elaine',  'Bruna',  '',     ''],
    // ─── FEVEREIRO 2026 ─────────────────────────────────
    [2, 2026, '07/02/2026', 'Rafa',    'Aline',  '',     'OBS sexta 13/02: Elaine e Nay 7h-16h, Bruna 8h-17h, Rafa e Rô 9h-18h'],
    [2, 2026, '14/02/2026', 'Feriado', '',       '',     ''],
    [2, 2026, '21/02/2026', 'Rosana',  'Nayara', '',     ''],
    [2, 2026, '28/02/2026', 'Elaine',  'Bruna',  '',     ''],
    // ─── MARÇO 2026 ─────────────────────────────────────
    [3, 2026, '07/03/2026', 'Rafa',    'Aline',  '',     ''],
    [3, 2026, '14/03/2026', 'Rosana',  'Nayara', '',     ''],
    [3, 2026, '21/03/2026', 'Elaine',  'Bruna',  '',     ''],
    [3, 2026, '28/03/2026', 'Rafa',    'Aline',  '',     ''],
    // ─── ABRIL 2026 ─────────────────────────────────────
    [4, 2026, '04/04/2026', 'Feriado', '',       '',     ''],
    [4, 2026, '11/04/2026', 'Rosana',  'Nayara', '',     ''],
    [4, 2026, '18/04/2026', 'Elaine',  'Bruna',  '',     ''],
    [4, 2026, '25/04/2026', 'Rafa',    'Aline',  '',     ''],
    // ─── MAIO 2026 ──────────────────────────────────────
    [5, 2026, '02/05/2026', 'Rosana',  'Nayara', '',     ''],
    [5, 2026, '09/05/2026', 'Elaine',  'Bruna',  '',     ''],
    [5, 2026, '16/05/2026', 'Rafa',    'Aline',  '',     ''],
    [5, 2026, '23/05/2026', 'Rosana',  'Nayara', '',     ''],
    [5, 2026, '30/05/2026', 'Elaine',  'Bruna',  '',     ''],
    // ─── JUNHO 2026 ─────────────────────────────────────
    [6, 2026, '06/06/2026', 'Rafa',    'Aline',  '',     ''],
    [6, 2026, '13/06/2026', 'Rosana',  'Nayara', '',     ''],
    [6, 2026, '20/06/2026', 'Elaine',  'Bruna',  '',     ''],
    [6, 2026, '27/06/2026', 'Rafa',    'Aline',  'Caren',''],
    // ─── JULHO 2026 ─────────────────────────────────────
    [7, 2026, '04/07/2026', 'Rosana',  'Nayara', 'Caren',''],
    [7, 2026, '11/07/2026', 'Elaine',  'Bruna',  '',     ''],
    [7, 2026, '18/07/2026', 'Rafa',    'Aline',  'Caren',''],
    [7, 2026, '25/07/2026', 'Rosana',  'Nayara', '',     ''],
    // ─── AGOSTO 2026 ────────────────────────────────────
    [8, 2026, '01/08/2026', 'Elaine',  'Bruna',  '',     ''],
    [8, 2026, '08/08/2026', 'Rafa',    'Aline',  'Caren',''],
    [8, 2026, '15/08/2026', 'Rosana',  'Nayara', '',     ''],
    [8, 2026, '22/08/2026', 'Elaine',  'Bruna',  '',     ''],
    [8, 2026, '29/08/2026', 'Rafa',    'Aline',  'Caren',''],
    // ─── SETEMBRO 2026 ──────────────────────────────────
    [9, 2026, '05/09/2026', 'Rosana',  'Nayara', '',     ''],
    [9, 2026, '12/09/2026', 'Elaine',  'Bruna',  '',     ''],
    [9, 2026, '19/09/2026', 'Rafa',    'Aline',  '',     ''],
    [9, 2026, '26/09/2026', 'Rosana',  'Nayara', '',     ''],
    // ─── OUTUBRO 2026 ───────────────────────────────────
    [10, 2026, '10/10/2026', 'Elaine', 'Bruna',  '',     ''],
    [10, 2026, '17/10/2026', 'Rafa',   'Aline',  '',     ''],
    [10, 2026, '24/10/2026', 'Rosana', 'Nayara', '',     ''],
    [10, 2026, '31/10/2026', 'Elaine', 'Bruna',  '',     ''],
    // ─── NOVEMBRO 2026 ──────────────────────────────────
    [11, 2026, '07/11/2026', 'Rafa',   'Aline',  '',     ''],
    [11, 2026, '14/11/2026', 'Rosana', 'Nayara', '',     ''],
    [11, 2026, '21/11/2026', 'Elaine', 'Bruna',  '',     ''],
    [11, 2026, '28/11/2026', 'Rafa',   'Aline',  '',     ''],
    // ─── DEZEMBRO 2026 ──────────────────────────────────
    [12, 2026, '05/12/2026', 'Rosana', 'Nayara', '',     ''],
    [12, 2026, '12/12/2026', 'Elaine', 'Bruna',  '',     ''],
    [12, 2026, '19/12/2026', '',       '',       '',     ''],
    [12, 2026, '26/12/2026', '',       '',       '',     ''],
  ];

  const rows = dados.map(function(r) {
    return [r[0], r[1], _d_(r[2]), r[3], r[4], r[5], r[6]];
  });

  _escrever_(ss, 'ESCALA_SABADO', rows, [3]);
}

// =============================================================================
// AJUSTES DE HORÁRIO
// =============================================================================

function _popularAjustes_(ss) {
  const agora = new Date();
  // [Atividade, Data_DD/MM/YYYY, Status, Mes, Ano]
  const SRC = [
    // ─── OUTUBRO 2025 ───────────────────────────────────────────────────────
    ['Elaine OFF',                                        '16/10/2025', 'Feito', 10, 2025],
    ['Elaine OFF',                                        '17/10/2025', 'Feito', 10, 2025],
    ['Elaine OFF',                                        '23/10/2025', 'Feito', 10, 2025],
    ['Nay fez o horário das 12:30 às 21:30',              '16/10/2025', 'Feito', 10, 2025],
    ['Nay fez o horário das 7h às 16h',                   '17/10/2025', 'Feito', 10, 2025],
    ['Nay fez o horário das 7h às 16h',                   '23/10/2025', 'Feito', 10, 2025],
    ['Bruna fez o horário das 9h às 17h',                 '16/10/2025', 'Feito', 10, 2025],
    ['Bruna fez o horário das 9h às 17h',                 '17/10/2025', 'Feito', 10, 2025],
    // ─── NOVEMBRO 2025 ──────────────────────────────────────────────────────
    ['Elaine trabalhou (Bruna de férias)',                 '08/11/2025', 'Feito', 11, 2025],
    ['Rosana com atestado',                               '03/11/2025', 'Feito', 11, 2025],
    ['Rosana com atestado',                               '04/11/2025', 'Feito', 11, 2025],
    ['Elaine OFF',                                        '14/11/2025', 'Feito', 11, 2025],
    ['Nay fez o horário das 7h às 16h',                   '14/11/2025', 'Feito', 11, 2025],
    ['Nay pegou às 7h',                                   '28/11/2025', 'Feito', 11, 2025],
    // ─── DEZEMBRO 2025 ──────────────────────────────────────────────────────
    ['Dia 08/12 feriado na cidade da Nay',                '08/12/2025', 'Feito', 12, 2025],
    ['Rafa trocou de horário com Nay',                    '03/12/2025', 'Feito', 12, 2025],
    ['Bruna fez o horário das 8h às 17h',                 '08/12/2025', 'Feito', 12, 2025],
    ['Nay cobriu férias da Elaine (a partir 09/12)',       '09/12/2025', 'Feito', 12, 2025],
    ['Bruna fez o horário das 9h às 18h (a partir 09/12)','09/12/2025', 'Feito', 12, 2025],
    ['Bruna e Rosana trocaram sábado 06/12 por 13/12',    '06/12/2025', 'Feito', 12, 2025],
    // ─── JANEIRO 2026 ───────────────────────────────────────────────────────
    ['Elaine pegou às 7h',                                '05/01/2026', 'Feito',  1, 2026],
    ['Elaine vai pegar às 7h (tem médico)',                '07/01/2026', 'Feito',  1, 2026],
    ['Rô com internet oscilando (13:30)',                  '07/01/2026', '',       1, 2026],
    ['Bru se desconectou, passando mal',                  '07/01/2026', 'Feito',  1, 2026],
    ['Nay e Rô com exame periódico',                      '09/01/2026', 'Feito',  1, 2026],
    ['Nay pegou às 7h (tem médico)',                      '16/01/2026', 'Feito',  1, 2026],
    // ─── FEVEREIRO 2026 ─────────────────────────────────────────────────────
    ['Rafa trocou de horário com Bruna',                  '12/02/2026', 'Feito',  2, 2026],
    ['Nay pegou às 7h',                                   '20/02/2026', 'Feito',  2, 2026],
    // ─── ABRIL 2026 ─────────────────────────────────────────────────────────
    ['Bruna fez o horário das 11:30 às 20:30',            '06/04/2026', 'Feito',  4, 2026],
    ['Bruna fez o horário das 11:30 às 20:30',            '07/04/2026', 'Feito',  4, 2026],
    ['Nay fez o horário das 11:30 às 20:30',              '08/04/2026', 'Feito',  4, 2026],
    ['Nay trocou de horário com Rafa',                    '15/04/2026', 'Feito',  4, 2026],
    ['Bruna trocou de horário com Nay',                   '20/04/2026', 'Feito',  4, 2026],
    // ─── MAIO 2026 ──────────────────────────────────────────────────────────
    ['Rafa com atestado por 3 dias',                      '11/05/2026', 'Feito',  5, 2026],
    ['Bru fez o horário das 11h às 20h',                  '12/05/2026', 'Feito',  5, 2026],
    ['Nay fez o horário das 12:30 às 21:30',              '13/05/2026', 'Feito',  5, 2026],
    ['Nay cobriu férias da Rô às seg e qua',              '18/05/2026', 'Feito',  5, 2026],
    ['Bru cobriu férias da Rô às ter e qui',              '18/05/2026', 'Feito',  5, 2026],
    ['Bru OFF (problema pessoal)',                         '26/05/2026', 'Feito',  5, 2026],
    ['Nay pegou às 7h (tem médico às 17h)',               '28/05/2026', 'Feito',  5, 2026],
    // ─── JUNHO 2026 ─────────────────────────────────────────────────────────
    ['Bru e Rafa trocaram folgas: Rafa dia 06 e Bru dia 20', '06/06/2026', 'Feito', 6, 2026],
    ['Rafa chegou atrasada (14:46)',                       '01/06/2026', 'Feito',  6, 2026],
    ['Rafa chegou atrasada (13:30)',                       '16/06/2026', 'Feito',  6, 2026],
    ['Nay trocou de horário com Rafa',                    '22/06/2026', 'Feito',  6, 2026],
    // ─── JULHO 2026 ─────────────────────────────────────────────────────────
    ['Rafa fez o horário das 9h às 18h',                  '01/07/2026', 'Feito',  7, 2026],
    ['Rafa chegou às 13h',                                '08/07/2026', '',       7, 2026],
    ['Feriado SP: Elaine trabalhou; meio período 10/07; folga 11/07', '09/07/2026', 'Feito', 7, 2026],
    ['Rosana chegou atrasada (12:58)',                     '15/07/2026', '',       7, 2026],
    ['Nay vai fazer o horário das 7h às 16h',             '21/07/2026', '',       7, 2026],
    ['Rafa vai fazer o horário das 11h às 20h',           '22/07/2026', '',       7, 2026],
  ];

  const rows = SRC.map(function(r) {
    return [r[0], _d_(r[1]), r[2], r[3], r[4], 'Migração', agora];
  });

  _escrever_(ss, 'AJUSTES_HORARIO', rows, [2, 7]);
}

// =============================================================================
// MIGRAÇÃO 2 — Reclassifica os ajustes de horário já lançados (texto livre)
// pro novo modelo estruturado (Tipo / Quem sai / Quem cobre / Horário para).
// Não mexe na coluna Atividade (fica como observação) nem em Registrado_Por/Em.
//
// 1. Abra o Apps Script da planilha (Extensões → Apps Script)
// 2. Cole este arquivo inteiro (ou só esta função, se o resto já estiver lá)
// 3. Rode a função migrarAjustesParaNovoModelo()
// =============================================================================

function migrarAjustesParaNovoModelo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AJUSTES_HORARIO');
  if (!sheet) { SpreadsheetApp.getUi().alert('Aba AJUSTES_HORARIO não encontrada.'); return; }

  // Garante que as colunas novas existem (mesma lógica do getAjustesSheet_ do Code.gs)
  const header = sheet.getRange(1, 1, 1, 14).getValues()[0];
  if (!header[8])  sheet.getRange(1, 9).setValue('Tags');
  if (!header[9])  sheet.getRange(1, 10).setValue('Tipo');
  if (!header[10]) sheet.getRange(1, 11).setValue('Quem sai');
  if (!header[11]) sheet.getRange(1, 12).setValue('Quem cobre');
  if (!header[12]) sheet.getRange(1, 13).setValue('Horário de');
  if (!header[13]) sheet.getRange(1, 14).setValue('Horário para');

  // Chave: ID da linha (coluna H). Valor: [status novo, tipo, quemSai, quemCobre, horarioPara]
  const MAPA = {
    '57a3a0b6-4eee-4579-bba9-91504d5a3208': ['Confirmado', 'Falta', 'Elaine', '', ''],
    '531d055c-2a8c-4c1f-a437-f3ca47bb624b': ['Confirmado', 'Falta', 'Elaine', '', ''],
    '3225fd2c-08c8-4368-8f90-9fdc3fc211bf': ['Confirmado', 'Falta', 'Elaine', '', ''],
    '4b1bbf8f-8263-4836-81e6-81c97f016f55': ['Confirmado', 'Horário pontual', '', 'Nayara', '12:30 - 21:30'],
    '43e2acb6-e9f6-41db-b088-4275e465fd3f': ['Confirmado', 'Horário pontual', '', 'Nayara', '07:00 - 16:00'],
    'dc51735c-71f3-4c65-8457-b40ef439b2a5': ['Confirmado', 'Horário pontual', '', 'Nayara', '07:00 - 16:00'],
    '2095c024-c697-47eb-b8be-cdcbf276266e': ['Confirmado', 'Horário pontual', '', 'Bruna', '09:00 - 17:00'],
    '6082aeba-13da-4b9b-b7cf-fe676d53022c': ['Confirmado', 'Horário pontual', '', 'Bruna', '09:00 - 17:00'],
    '28acacfb-05aa-480f-a4f1-87476f549b7b': ['Confirmado', 'Cobertura', 'Bruna', 'Elaine', ''],
    'd004d16d-a243-40e0-a439-8009c57411bb': ['Confirmado', 'Falta', 'Rosana', '', ''],
    '986b9804-4c4d-429f-be47-245715fd5653': ['Confirmado', 'Falta', 'Rosana', '', ''],
    'aca56a1d-fe73-4641-b53c-0c70b92b61c5': ['Confirmado', 'Falta', 'Elaine', '', ''],
    'd8d0704c-0d0a-4833-8988-af59208af137': ['Confirmado', 'Horário pontual', '', 'Nayara', '07:00 - 16:00'],
    '43d42d8c-b884-4fef-9249-364f27e3c0a1': ['Confirmado', 'Horário pontual', '', 'Nayara', '07:00'],
    '3c0819f3-0853-4ff2-a18e-d30f9da7a78f': ['Confirmado', 'Falta', 'Nayara', '', ''],
    '6e6320c6-78d8-42df-a6b9-a158a4e32c43': ['Confirmado', 'Troca', 'Rafaela', 'Nayara', ''],
    '0c7ac77d-fddb-4426-b661-cc2d57857e90': ['Confirmado', 'Horário pontual', '', 'Bruna', '08:00 - 17:00'],
    '64bc779d-f385-49f6-a51e-d769686b84ee': ['Confirmado', 'Cobertura', 'Elaine', 'Nayara', ''],
    '695dfb17-cb48-419a-abe2-e597ab8fa112': ['Confirmado', 'Horário pontual', '', 'Bruna', '09:00 - 18:00'],
    '2426af8e-e038-4598-901f-411afabe626f': ['Confirmado', 'Troca', 'Bruna', 'Rosana', ''],
    '100b4194-bd45-490f-a883-25122f1706f9': ['Confirmado', 'Horário pontual', '', 'Elaine', '07:00'],
    'e0d27866-4455-403a-bded-f30dcb9d3833': ['Confirmado', 'Horário pontual', '', 'Elaine', '07:00'],
    '72a75d0b-968e-4b88-b81b-4ab8a786971f': ['A confirmar', '', 'Rosana', '', ''],
    'c174deb6-a94b-4004-b2ac-d46a0d8cef1b': ['Confirmado', 'Falta', 'Bruna', '', ''],
    '1fd0876e-7ee3-42eb-b194-8bcbaab9741c': ['Confirmado', 'Falta', 'Nayara / Rosana', '', ''],
    '4273e349-ce89-43a6-9973-7a0d16098e5d': ['Confirmado', 'Horário pontual', '', 'Nayara', '07:00'],
    'f87dad6a-dff1-44ed-bb3d-f1cd1670e6a0': ['Confirmado', 'Troca', 'Rafaela', 'Bruna', ''],
    'eeca90fb-e226-4b78-8195-dc6a9d55383f': ['Confirmado', 'Horário pontual', '', 'Nayara', '07:00'],
    'b223db8a-c39e-40e3-919b-6f35387acb5d': ['Confirmado', 'Horário pontual', '', 'Bruna', '11:30 - 20:30'],
    '9ca35fd6-9946-4498-b185-8ee65d0e57cc': ['Confirmado', 'Horário pontual', '', 'Bruna', '11:30 - 20:30'],
    'a3ecfc96-deea-4dd8-b159-2eaa9efa656e': ['Confirmado', 'Horário pontual', '', 'Nayara', '11:30 - 20:30'],
    '5d84b65b-a3e8-4fbd-ae1d-4cd801c6243d': ['Confirmado', 'Troca', 'Nayara', 'Rafaela', ''],
    'bda64574-6474-44f9-87b5-2b942d03734b': ['Confirmado', 'Troca', 'Bruna', 'Nayara', ''],
    '41d6ad31-8e1f-4d31-bf11-bc20b07930f3': ['Confirmado', 'Falta', 'Rafaela', '', ''],
    'c40c46a1-522d-477d-bcca-9fbbf2cc6c17': ['Confirmado', 'Horário pontual', '', 'Bruna', '11:00 - 20:00'],
    '1bb61d20-65a7-4b45-9a83-28e3cfe86af1': ['Confirmado', 'Horário pontual', '', 'Nayara', '12:30 - 21:30'],
    '2fa986d8-abc2-4aef-9d38-fd4b6eb58e18': ['Confirmado', 'Cobertura', 'Rosana', 'Nayara', ''],
    'cc7273df-21fa-4a13-ab4a-12c0df909d0d': ['Confirmado', 'Cobertura', 'Rosana', 'Bruna', ''],
    '3af420e0-0007-488d-93f4-27d5361dbd0f': ['Confirmado', 'Falta', 'Bruna', '', ''],
    '1fcb4d1c-532d-4ffe-ab4d-e05ee0d8b07e': ['Confirmado', 'Horário pontual', '', 'Nayara', '07:00'],
    'ae42476f-89af-489d-ae56-0bec5e16ab69': ['Confirmado', 'Troca', 'Bruna', 'Rafaela', ''],
    'af12c7ea-6ee9-4bc0-b60a-5ab7ff5e827a': ['Confirmado', 'Horário pontual', '', 'Rafaela', '14:46'],
    '4ab2a129-7188-419d-a4ae-579083c6dd1d': ['Confirmado', 'Horário pontual', '', 'Rafaela', '13:30'],
    '3a30969d-2f8a-4ae4-a897-7c0e8e812fd5': ['Confirmado', 'Troca', 'Nayara', 'Rafaela', ''],
    '5a8e5749-59b3-4cda-bf71-ee47931b6a87': ['Confirmado', 'Horário pontual', '', 'Rafaela', '09:00 - 18:00'],
    '5b69f7ba-625a-4083-9d76-af256146b64f': ['A confirmar', 'Horário pontual', '', 'Rafaela', '13:00'],
    '8d12c746-dc92-4198-9de1-a1ed2c5c406b': ['Confirmado', 'Horário pontual', '', 'Elaine', ''],
    '6cd1f4ad-5f0f-401d-9cce-8f32da87d8a3': ['A confirmar', 'Horário pontual', '', 'Rosana', '12:58'],
    '0f2d2a4d-8023-445f-abc6-097eeffa3ccc': ['A confirmar', 'Horário pontual', '', 'Nayara', '07:00 - 16:00'],
    '042ca589-b3c5-4730-a085-501c999b4eaf': ['A confirmar', 'Horário pontual', '', 'Rafaela', '11:00 - 20:00']
  };

  const rows = sheet.getDataRange().getValues();
  let atualizados = 0;
  const naoEncontrados = [];
  const idsUsados = {};

  for (let i = 1; i < rows.length; i++) {
    const id = String(rows[i][7] || '');
    if (!id) continue;
    const dado = MAPA[id];
    if (!dado) { naoEncontrados.push(id); continue; }
    idsUsados[id] = true;

    const status = dado[0];
    const tipo = dado[1];
    const quemSai = dado[2];
    const quemCobre = dado[3];
    const horarioPara = dado[4];

    sheet.getRange(i + 1, 3).setValue(status);
    sheet.getRange(i + 1, 10, 1, 5).setValues([[tipo, quemSai, quemCobre, '', horarioPara]]);
    atualizados++;
  }

  const idsNoMapaNaoAplicados = Object.keys(MAPA).filter(function(id) { return !idsUsados[id]; });

  Logger.log(
    'Migração de ajustes concluída!\n\n' +
    'Linhas atualizadas: ' + atualizados + '\n' +
    'Linhas na planilha sem correspondência no mapa: ' + naoEncontrados.length +
    (naoEncontrados.length ? '\n(' + naoEncontrados.join(', ') + ')' : '') + '\n' +
    'IDs do mapa que não foram encontrados na planilha: ' + idsNoMapaNaoAplicados.length +
    (idsNoMapaNaoAplicados.length ? '\n(' + idsNoMapaNaoAplicados.join(', ') + ')' : '')
  );
}

// =============================================================================
// MIGRAÇÃO 3 — Corrige células de "Horário de"/"Horário para" que o Sheets
// converteu sozinho num valor de data/hora (ex: "07:00" sem " - " vira
// internamente "Sat Dec 30 1899 07:00:00..."). Reescreve como texto "HH:mm"
// e marca as colunas como texto puro pra não acontecer de novo.
//
// Rode migrarAjustesParaNovoModelo() antes desta, se ainda não rodou.
// =============================================================================

function corrigirFormatoHorariosAjustes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AJUSTES_HORARIO');
  if (!sheet) { Logger.log('Aba AJUSTES_HORARIO não encontrada.'); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Nenhuma linha de dado.'); return; }

  const range = sheet.getRange(2, 13, lastRow - 1, 2); // colunas M:N = Horário de / Horário para
  const values = range.getValues();
  const tz = Session.getScriptTimeZone();
  let corrigidos = 0;

  const novos = values.map(function(row) {
    return row.map(function(v) {
      if (v instanceof Date) {
        corrigidos++;
        return Utilities.formatDate(v, tz, 'HH:mm');
      }
      return v;
    });
  });

  range.setNumberFormat('@').setValues(novos);
  Logger.log('Células de horário corrigidas para texto: ' + corrigidos);
}
