# Controle Operacional BOL — BRASAS Online

Webapp em Google Apps Script para gestão operacional da unidade BOL.

## Planilha de dados
`https://docs.google.com/spreadsheets/d/1I8nss8jH0sv8qCmagWI2bITPEW2JWhi5OUsjVeBA_-M`

## Estrutura da planilha (abas)

| Aba | Descrição |
|---|---|
| `EVENTOS` | Calendário anual de eventos (BRASAS Start, Special Classes, Férias, etc.) |
| `ROTINAS` | Lista de tarefas por frequência (Diário, Semanal, Mensal, etc.) |
| `ROTINAS_LOG` | Histórico de conclusão das rotinas por período (criada automaticamente pelo `Code.gs`) |
| `ENTREGAS` | Status mensal de entregas por atividade (NO PRAZO / EM ATRASO) |
| `HORARIOS` | Horários fixos da secretaria por pessoa |
| `ESCALA_SABADO` | Escala de sábados (rodízio mensal) |
| `AJUSTES_HORARIO` | Trocas, coberturas e ajustes pontuais de horário |
| `FERIAS` | Períodos de férias já tirados por pessoa (criada automaticamente pelo `Code.gs`) |
| `FUNCIONARIOS` | Nome + data de admissão/demissão de cada pessoa (criada automaticamente pelo `Code.gs`) |

## Arquivos

- `Migracao.gs` — Script de migração dos dados da planilha original para a nova planilha estruturada. **Já executado.**
- `Code.gs` — Backend do webapp.
- `Index.html` — Frontend do webapp.
- `LembretesConceito.gs` — Script à parte, de outro projeto Apps Script (lembretes de preenchimento de Conceito de professores, vinculado à planilha de CIAS/BOLSISTAS). Não faz parte do webapp Controle Operacional BOL.

## Módulos do Webapp

1. **Dashboard** — Visão do dia (quem trabalha, eventos da semana, entregas pendentes)
2. **Calendário** — Criação/edição de eventos, filtro por ano/mês
3. **Rotinas** — Checklist por frequência (Diário, Semanal, Mensal...), com histórico em `ROTINAS_LOG`
4. **Entregas** — Tabela anual editável com status por mês
5. **Funcionários** — três sub-abas:
   - *Férias*: períodos tirados por pessoa e saldo do ciclo atual (30 dias por ciclo de 12 meses a partir da admissão); períodos lançados aqui aparecem automaticamente no Calendário
   - *Admissão*: funcionários ativos e suas datas de admissão
   - *Demissão*: funcionários desligados e suas datas de admissão/demissão
6. **Secretaria** — Horários fixos + Escala de sábado + Ajustes de horário

## Auth
Autenticação via Hub BRASAS BI (token de sessão), mesmo padrão do app de Horas de Professores.
Chave de acesso deste app: `operacionalbol` — precisa estar liberada na coluna `ACESSOS`
da aba `SESSOES` da planilha do Hub para cada e-mail da equipe, senão o login barra com
"sem permissão".

## Deploy
No Apps Script vinculado à planilha `Controle Operacional BOL` (ID no topo deste README),
cole `Code.gs` e `Index.html`, publique como Web App ("Executar como: eu",
"Quem tem acesso: qualquer pessoa") e atualize a constante `HUB_URL` em ambos os arquivos
se a URL do Hub mudar.
