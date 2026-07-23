/**
 * Lembretes automáticos de preenchimento de Conceito para professores.
 * Script vinculado à planilha "Planilha de Conceitos de CIAS/BOLSISTAS BOL".
 * (Projeto Apps Script separado do webapp Controle Operacional BOL.)
 */

// ====== CONFIGURAÇÃO ======
const ADMIN_EMAIL = 'adriane@brasas.com';
const NOME_REMETENTE = 'BRASAS English Course';
const EMAILS_DUVIDAS = ['natasha@brasas.com', 'aline.castro@brasas.com'];
const DIAS_ENVIO = [25, 28, 4]; // dias do mês em que os lembretes são disparados
const DIA_ENVIO_GERAL = 25;     // neste dia, envia para TODOS os professores do mês (preenchido ou não)
const DIA_ENVIO_MES_ANTERIOR = 4; // neste dia, o lembrete se refere ao mês anterior (ainda pendente)
const ABA_PROFESSORES = 'Professores';
const PREFIXO_CIAS = 'CIAS';
const PREFIXO_BOLSISTAS = 'BOLSISTAS';
const COL_PROFESSOR = 4; // coluna D
const COL_CONCEITO = 5;  // coluna E

const MESES = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO',
  'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO'];

// ====== MENU ======
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Lembretes de Conceito')
    .addItem('Enviar lembretes agora', 'enviarLembretes')
    .addItem('Testar envio (ignora dia do mês)', 'enviarLembretesTeste')
    .addItem('Enviar exemplo de e-mail para mim', 'enviarEmailExemplo')
    .addToUi();
}

// ====== FUNÇÕES PRINCIPAIS ======
function enviarLembretes() {
  processarLembretes_(false);
}

function enviarLembretesTeste() {
  processarLembretes_(true);
}

function processarLembretes_(ignorarData) {
  const hoje = new Date();
  const diaHoje = hoje.getDate();

  if (!ignorarData && DIAS_ENVIO.indexOf(diaHoje) === -1) {
    Logger.log('Hoje é dia %s, não é um dos dias de envio (%s).', diaHoje, DIAS_ENVIO.join(', '));
    return;
  }

  // Envio geral (para todos os professores do mês) só acontece no DIA_ENVIO_GERAL.
  const enviarParaTodos = !ignorarData && diaHoje === DIA_ENVIO_GERAL;

  // No dia DIA_ENVIO_MES_ANTERIOR o lembrete ainda se refere ao mês anterior
  // (o mês cujo prazo venceu no fim do mês passado).
  const dataReferencia = (!ignorarData && diaHoje === DIA_ENVIO_MES_ANTERIOR)
    ? new Date(hoje.getFullYear(), hoje.getMonth() - 1, 1)
    : hoje;

  const nomeCias = nomeAbaMesAtual_(PREFIXO_CIAS, dataReferencia);
  const nomeBolsistas = nomeAbaMesAtual_(PREFIXO_BOLSISTAS, dataReferencia);
  const abaCias = encontrarAba_(nomeCias);
  const abaBolsistas = encontrarAba_(nomeBolsistas);

  if (!abaCias && !abaBolsistas) {
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: '[Lembretes de Conceito] Abas do mês não encontradas',
      body: `Não encontrei as abas "${nomeCias}" nem "${nomeBolsistas}" na planilha. ` +
        'Verifique se elas já foram criadas para este mês (nome precisa ser exatamente ' +
        '"CIAS <MÊS> <ANO>" / "BOLSISTAS <MÊS> <ANO>").'
    });
    return;
  }

  const destinatariosNomes = enviarParaTodos
    ? new Set([
        ...getProfessoresDoMes_(abaCias),
        ...getProfessoresDoMes_(abaBolsistas)
      ])
    : new Set([
        ...getProfessoresPendentes_(abaCias),
        ...getProfessoresPendentes_(abaBolsistas)
      ]);

  if (destinatariosNomes.size === 0) {
    Logger.log('Nenhum professor para notificar hoje.');
    return;
  }

  const mapaEmails = getMapaProfessores_();
  const semCorrespondencia = [];
  const mesReferencia = `${capitalizar_(MESES[dataReferencia.getMonth()])}/${dataReferencia.getFullYear()}`;
  const prazoTexto = ultimoDiaDoMesTexto_(dataReferencia);
  const linkPlanilha = SpreadsheetApp.getActive().getUrl();
  const corpo = montarCorpoEmail_(mesReferencia, prazoTexto, linkPlanilha, enviarParaTodos);
  const assunto = enviarParaTodos
    ? `Lembrete: preenchimento de Conceitos de ${mesReferencia}`
    : `Lembrete: Conceitos de ${mesReferencia} ainda não preenchidos`;

  destinatariosNomes.forEach(professor => {
    const email = mapaEmails[professor];
    if (!email) {
      semCorrespondencia.push(professor);
      return;
    }
    MailApp.sendEmail({
      to: email,
      name: NOME_REMETENTE,
      replyTo: EMAILS_DUVIDAS.join(', '),
      subject: assunto,
      htmlBody: corpo
    });
  });

  if (semCorrespondencia.length > 0) {
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: '[Lembretes de Conceito] Professores sem e-mail cadastrado',
      body: 'Os professores abaixo precisavam ser notificados, mas o nome deles não foi ' +
        `encontrado (correspondência exata) na aba "${ABA_PROFESSORES}":\n\n` +
        semCorrespondencia.join('\n')
    });
  }
}

function enviarEmailExemplo() {
  const hoje = new Date();
  const mesReferencia = `${capitalizar_(MESES[hoje.getMonth()])}/${hoje.getFullYear()}`;
  const prazoTexto = ultimoDiaDoMesTexto_(hoje);
  const linkPlanilha = SpreadsheetApp.getActive().getUrl();
  const corpo = montarCorpoEmail_(mesReferencia, prazoTexto, linkPlanilha, false);
  const destinatario = Session.getActiveUser().getEmail() || ADMIN_EMAIL;

  MailApp.sendEmail({
    to: destinatario,
    name: NOME_REMETENTE,
    replyTo: EMAILS_DUVIDAS.join(', '),
    subject: `[EXEMPLO] Lembrete: Conceitos de ${mesReferencia} ainda não preenchidos`,
    htmlBody: corpo
  });

  SpreadsheetApp.getUi().alert(`Exemplo enviado para ${destinatario}.`);
}

// ====== HELPERS DE DADOS ======
function nomeAbaMesAtual_(prefixo, data) {
  return `${prefixo} ${MESES[data.getMonth()]} ${data.getFullYear()}`;
}

function encontrarAba_(nomeEsperado) {
  const alvo = nomeEsperado.trim().toUpperCase();
  const abas = SpreadsheetApp.getActive().getSheets();
  for (const aba of abas) {
    if (aba.getName().trim().toUpperCase() === alvo) return aba;
  }
  return null;
}

function getMapaProfessores_() {
  const aba = SpreadsheetApp.getActive().getSheetByName(ABA_PROFESSORES);
  if (!aba) throw new Error(`Aba "${ABA_PROFESSORES}" não encontrada.`);
  const dados = aba.getDataRange().getValues();
  const mapa = {};
  for (let i = 1; i < dados.length; i++) {
    const nome = String(dados[i][0] || '').trim();
    const email = String(dados[i][1] || '').trim();
    if (nome && email) mapa[nome] = email;
  }
  return mapa;
}

function getProfessoresDoMes_(aba) {
  const nomes = new Set();
  if (!aba) return nomes;
  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    const professor = String(dados[i][COL_PROFESSOR - 1] || '').trim();
    if (professor) nomes.add(professor);
  }
  return nomes;
}

function getProfessoresPendentes_(aba) {
  const pendentes = new Set();
  if (!aba) return pendentes;
  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    const professor = String(dados[i][COL_PROFESSOR - 1] || '').trim();
    const conceito = String(dados[i][COL_CONCEITO - 1] || '').trim();
    if (professor && !conceito) pendentes.add(professor);
  }
  return pendentes;
}

function ultimoDiaDoMesTexto_(data) {
  const ultimoDia = new Date(data.getFullYear(), data.getMonth() + 1, 0);
  return Utilities.formatDate(ultimoDia, Session.getScriptTimeZone(), 'dd/MM');
}

function capitalizar_(palavra) {
  return palavra.charAt(0) + palavra.slice(1).toLowerCase();
}

// ====== TEMPLATE DO E-MAIL ======
function montarCorpoEmail_(mesReferencia, prazoTexto, linkPlanilha, geral) {
  const tituloTopo = geral
    ? 'Lembrete: preenchimento de Conceitos'
    : 'Lembrete: Conceitos ainda não preenchidos';
  const tituloFaixa = geral
    ? 'Lembrete — Preenchimento de Conceitos'
    : 'Lembrete — Preenchimento de Conceitos Pendente';
  const corFaixa = geral ? '#1a1a2e' : '#c0522d';
  const introTexto = geral
    ? `Este é o lembrete para o preenchimento dos conceitos na planilha de alunos de empresas ` +
      `e bolsistas, referente a <b>${mesReferencia}</b>.`
    : `Identificamos que ainda há conceitos pendentes de preenchimento na planilha de alunos ` +
      `de empresas e bolsistas, referente a <b>${mesReferencia}</b>.`;

  return `
  <div style="font-family:Arial,Helvetica,sans-serif;max-width:640px;margin:0 auto;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden;">
    <div style="background:#1a1a2e;color:#ffffff;padding:16px 24px;">
      <div style="font-size:15px;opacity:0.85;">${tituloTopo}</div>
    </div>
    <div style="background:${corFaixa};color:#ffffff;padding:14px 24px;font-size:18px;font-weight:bold;">
      ${tituloFaixa}
    </div>
    <div style="padding:24px;">
      <div style="background:#fff8e1;border:1px solid #f0d693;border-radius:6px;padding:12px 16px;margin-bottom:20px;font-size:13px;color:#6b5900;">
        ⚠️ Este é um comunicado automático. Em caso de dúvidas, entre em contato pelo e-mail
        <a href="mailto:${EMAILS_DUVIDAS[0]}">${EMAILS_DUVIDAS[0]}</a> ou
        <a href="mailto:${EMAILS_DUVIDAS[1]}">${EMAILS_DUVIDAS[1]}</a>.
      </div>

      <p>Olá, tudo bem?</p>
      <p>${introTexto}</p>

      <div style="background:#fdecea;border:1px solid #f3b3ac;border-radius:6px;padding:12px 16px;margin:16px 0;">
        <b>Prazo final para preencher: dia ${prazoTexto}, até as 10h.</b>
      </div>

      <p><b>Como preencher aproveitamento/conceito:</b></p>
      <table style="border-collapse:collapse;font-size:13px;" cellpadding="6">
        <tr><td style="border:1px solid #ccc;">E (90/100)</td><td style="border:1px solid #ccc;">Excelente</td></tr>
        <tr><td style="border:1px solid #ccc;">MB (80/89)</td><td style="border:1px solid #ccc;">Muito bom</td></tr>
        <tr><td style="border:1px solid #ccc;">B (70/79)</td><td style="border:1px solid #ccc;">Bom</td></tr>
        <tr><td style="border:1px solid #ccc;color:#c0392b;">R (60/69)</td><td style="border:1px solid #ccc;">Regular</td></tr>
        <tr><td style="border:1px solid #ccc;color:#c0392b;">I (&lt;60)</td><td style="border:1px solid #ccc;">Insuficiente</td></tr>
        <tr><td style="border:1px solid #ccc;color:#c0392b;">N (1 ou nenhuma aula assistida)</td><td style="border:1px solid #ccc;">Nulo</td></tr>
      </table>

      <p style="margin-top:20px;"><b>Atenção, pontos importantes:</b></p>
      <ul style="font-size:13px;">
        <li>Conceitos <b>I, N ou R</b> precisam de <b>JUSTIFICATIVA</b> no campo de Observações.
        Caso o aluno não tenha assistido aulas no período, informe a data da última tentativa
        de contato e se houve retorno.</li>
        <li>No primeiro mês de aula não podemos ter nenhum conceito I (Insuficiente) ou
        E (Excelente) para nenhum aluno novo.</li>
        <li>Caso haja avaliação no mês, o conceito deverá estar alinhado à nota do MT ou à
        média das notas finais.</li>
      </ul>

      <div style="text-align:center;margin:28px 0;">
        <a href="${linkPlanilha}" style="background:#1a1a2e;color:#ffffff;text-decoration:none;padding:12px 28px;border-radius:6px;font-size:14px;display:inline-block;">Preencher a planilha</a>
      </div>

      <p>Qualquer dúvida, estamos à disposição!</p>
    </div>
  </div>`;
}

// ====== SETUP (rodar manualmente uma única vez) ======
function criarGatilhoDiario() {
  removerGatilhosExistentes_();
  ScriptApp.newTrigger('enviarLembretes')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
}

function removerGatilhosExistentes_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'enviarLembretes') {
      ScriptApp.deleteTrigger(t);
    }
  });
}
