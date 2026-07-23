# Controle Operacional BOL — BRASAS Online

Webapp em Google Apps Script para gestão operacional da unidade BOL.

## Planilha de dados
`https://docs.google.com/spreadsheets/d/1I8nss8jH0sv8qCmagWI2bITPEW2JWhi5OUsjVeBA_-M`

## Estrutura da planilha (abas)

| Aba | Descrição |
|---|---|
| `EVENTOS` | Calendário anual de eventos (BRASAS Start, Special Classes, Férias, etc.) |
| `ROTINAS` | Lista de tarefas por frequência (Diário, Semanal, Mensal, etc.) |
| `ENTREGAS` | Status mensal de entregas por atividade (NO PRAZO / EM ATRASO) |
| `HORARIOS` | Horários fixos da secretaria por pessoa |
| `ESCALA_SABADO` | Escala de sábados (rodízio mensal) |
| `AJUSTES_HORARIO` | Trocas, coberturas e ajustes pontuais de horário |

## Arquivos

- `Migracao.gs` — Script de migração dos dados da planilha original para a nova planilha estruturada. **Já executado.**
- `Code.gs` — Backend do webapp (a criar)
- `Index.html` — Frontend do webapp (a criar)
- `LembretesConceito.gs` — Script à parte, de outro projeto Apps Script (lembretes de preenchimento de Conceito de professores, vinculado à planilha de CIAS/BOLSISTAS). Não faz parte do webapp Controle Operacional BOL.

## Módulos planejados para o Webapp

1. **Dashboard** — Visão do dia (quem trabalha, eventos da semana, entregas pendentes)
2. **Calendário** — Views: Ano / Mês / Semana
3. **Rotinas** — Checklist por frequência (Diário, Semanal, Mensal...)
4. **Entregas** — Tabela anual editável com status por mês
5. **Secretaria** — Horários fixos + Escala de sábado + Ajustes de horário

## Auth
Autenticação via Hub BRASAS BI (token de sessão), mesmo padrão do app de Horas de Professores.
