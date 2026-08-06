# Controle Operacional BOL — BRASAS Online

Webapp em Google Apps Script para gestão operacional da unidade BOL.

## Planilha de dados
`https://docs.google.com/spreadsheets/d/1I8nss8jH0sv8qCmagWI2bITPEW2JWhi5OUsjVeBA_-M`

## Estrutura da planilha (abas)

| Aba | Descrição |
|---|---|
| `EVENTOS` | Calendário anual de eventos (BRASAS Start, Special Classes, Férias, etc.) |
| `ROTINAS` | Lista de tarefas por frequência (Diário, Semanal, Mensal, etc.) |
| `ROTINAS_OPERACIONAIS` | Mesma estrutura da `ROTINAS`, para a aba Rotina Operacional — criada automaticamente pelo `Code.gs` |
| `ROTINAS_LOG` | Histórico de conclusão das duas listas de rotina por período (criada automaticamente pelo `Code.gs`) |
| `ENTREGAS` | Status mensal de entregas por atividade (NO PRAZO / EM ATRASO) |
| `HORARIOS` | Horários fixos da secretaria por pessoa |
| `ESCALA_SABADO` | Escala de sábados (rodízio mensal) |
| `AJUSTES_HORARIO` | Trocas, coberturas e ajustes pontuais de horário |
| `FERIAS` | Períodos de férias já tirados por pessoa, com a coluna `Periodo_Aquisitivo_ID` dizendo de qual período aquisitivo aqueles dias são abatidos (criada automaticamente pelo `Code.gs`) |
| `PERIODOS_AQUISITIVOS` | Períodos aquisitivos cadastrados manualmente, com `Data_Limite_Gozo` preenchida só quando alguém sobrescreve o cálculo padrão — criada automaticamente pelo `Code.gs` |
| `FUNCIONARIOS` | Cadastro de funcionários: nome, admissão/demissão, função, apelido, e-mail (liga a pessoa ao login dela), flags de rastreamento — criada automaticamente pelo `Code.gs` |
| `SOLICITACOES_FERIAS` | Pedidos de férias feitos pelos funcionários em Minhas Férias, com status Pendente/Aprovado/Recusado — criada automaticamente pelo `Code.gs` |
| `PROJETOS_COLUNAS` | Colunas do quadro de Projetos (nome + ordem) — criada automaticamente pelo `Code.gs` |
| `PROJETOS_CARTOES` | Cartões do quadro de Projetos (título, descrição, responsável, prazo, etiqueta e posição na coluna) — criada automaticamente pelo `Code.gs` |

## Arquivos

- `Migracao.gs` — Script de migração dos dados da planilha original para a nova planilha estruturada. **Já executado.**
- `Code.gs` — Backend do webapp.
- `Index.html` — Frontend do webapp.
- `LembretesConceito.gs` — Script à parte, de outro projeto Apps Script (lembretes de preenchimento de Conceito de professores, vinculado à planilha de CIAS/BOLSISTAS). Não faz parte do webapp Controle Operacional BOL.

## Módulos do Webapp

1. **Dashboard** — Visão do dia (quem trabalha, próximos eventos, entregas pendentes)
2. **Calendário** — Criação/edição de eventos, filtro por ano/mês
3. **Rotinas** — Checklist por frequência (Diário, Semanal, Mensal...), com histórico em `ROTINAS_LOG`.
   O Responsável é escolhido num seletor de tags: sugere os nomes já usados no app e permite
   criar um novo digitando. Aceita mais de uma pessoa (gravado como `Ana / Bia`); nenhuma
   escolhida equivale a `Todas`
4. **Rotina Operacional** — segunda lista de rotinas, idêntica à de cima em tudo (sub-abas
   Gerais/Individuais, checkbox por período, criar/editar/excluir), só que sobre a aba
   `ROTINAS_OPERACIONAIS` e com chave de acesso própria. O histórico de conclusão é o mesmo
   `ROTINAS_LOG` das duas: cada linha é identificada pelo ID da tarefa, então não há mistura
5. **Entregas** — Tabela anual editável com status por mês
6. **Funcionários** — quatro sub-abas:
   - *Cadastro*: lista mestre de funcionários (nome, função, apelido, admissão, demissão), sincronizável com a planilha de RH (unidade ONLINE). Cadastrar alguém aqui **não** inclui a pessoa em nenhuma das outras sub-abas — isso é sempre feito manualmente, uma a uma
   - *Férias*: um período aquisitivo é sempre 12 meses de casa que dão direito a 30 dias; a data limite de gozo é sugerida como 11 meses após o fim do período aquisitivo (`MESES_ATE_LIMITE_GOZO_` em `Code.gs`) e pode ser alterada à mão em cada período. **Não há cálculo automático pela data de admissão** — cada período aquisitivo é cadastrado manualmente por pessoa (podendo ter mais de um por pessoa), e só aparece saldo depois que alguém cadastra o período. Ao lançar férias tiradas, escolhe-se explicitamente de qual período aquisitivo os dias são abatidos (o sistema não deduz isso por data); férias sem período vinculado aparecem no histórico marcadas como "não abatido". Fica em vermelho quando faltam 30 dias ou menos pra data limite e ainda há dias a tirar. Só mostra quem foi explicitamente adicionado nesta sub-aba; períodos de férias lançados aqui aparecem automaticamente no Calendário
   - *Admissão*: checklist de admissão, só para quem foi explicitamente adicionado nesta sub-aba
   - *Demissão*: checklist de demissão, só para quem foi explicitamente adicionado nesta sub-aba

   As três últimas são independentes entre si — a inclusão em uma não afeta as outras.
7. **Secretaria** — Horários fixos + Escala de sábado + Ajustes de horário (tipo livre — Troca/Atraso/Falta/Hora Extra ou outro texto — e status Pendente/A Confirmar/Confirmado/Realizado, com filtro de tipo e de status). "Quem sai" e "quem cobre" são dropdowns com os nomes já usados no app, com opção `+ Nova tag…` que abre um campo pra digitar um nome novo. Cada ajuste aceita um anexo opcional (atestado, documento etc.), enviado direto pro Drive

   As tags de pessoa não são um cadastro à parte: são os nomes que já aparecem em rotinas,
   ajustes e escala de sábado (`getTodasTagsConhecidas_` no `Index.html`). Um nome novo passa
   a ser sugerido nos outros campos depois que o registro que o usa é salvo.
8. **Projetos** — quadro kanban estilo Trello: colunas criadas por você (nomear, reordenar,
   excluir) e cartões com título, etiqueta colorida, responsável, prazo e descrição. Cartões são
   arrastados entre colunas com o mouse; em telas de toque, dá pra mudar a coluna pelo próprio
   cartão. O quadro começa vazio — nenhuma coluna é criada automaticamente.
9. **Minhas Férias** — autoatendimento do funcionário: vê o próprio período aquisitivo, data
   limite de gozo e dias restantes, histórico de férias já tiradas, e pode solicitar um novo
   período (fica "Pendente" até um admin aprovar ou recusar em Funcionários → Solicitações).
   Aprovar já registra o período em `FERIAS` automaticamente.

## Auth e permissões
Autenticação via Hub BRASAS BI (token de sessão), mesmo padrão do app de Horas de Professores.
Tudo é controlado pela coluna `ACESSOS` da aba `SESSOES` da planilha do Hub, que aceita
várias chaves separadas por vírgula.

**Chave base — `operacionalbol`.** Sem ela o login barra com "sem permissão". Quem tem só
ela vê apenas **Minhas Férias** — o autoatendimento da própria pessoa, única aba que não
depende de chave, já que só mostra as férias e os pedidos de quem está logado.

**Chaves por página.** Todas as outras abas exigem, *além* da chave base, a sua própria
(mapeadas em `ACESSOS_PAGINA_`, no topo do `Code.gs`):

| Aba | Chave |
|---|---|
| Dashboard | `operacionalbol_dash` |
| Calendário | `operacionalbol_calendario` |
| Rotinas | `operacionalbol_rotinas` |
| Rotina Operacional | `operacionalbol_rotinaope` |
| Entregas | `operacionalbol_entregas` |
| Funcionários (todas as sub-abas) | `operacionalbol_func` |
| Secretaria | `operacionalbol_secretaria` |
| Projetos | `operacionalbol_projetos` |

Exemplo: quem só cuida da escala recebe `operacionalbol, operacionalbol_secretaria` e
enxerga Secretaria e Minhas Férias — mais nada.

O app abre sempre na primeira aba que a pessoa pode ver, seguindo a ordem da barra —
não adianta o Dashboard estar marcado como ativo no HTML se ela não tem a chave dele.

**Admin** — e-mails na lista `ADMIN_EMAILS_` (`Code.gs`) enxergam todas as abas sem precisar
de nenhuma chave de página. Serve de atalho e de trava: evita alguém ficar trancado pra fora
do painel por uma chave esquecida.

A checagem vale nos dois lados: o `Index.html` esconde o botão da aba, e toda função do
backend daquela aba passa por `requirePagina_(token, '<pagina>')`. Esconder o botão é só
conforto visual — é a checagem no servidor que de fato barra quem tentar chamar a função
pelo console. O `getBootstrap` também só lê da planilha as abas que a pessoa abre.

Para criar uma aba restrita nova: acrescente a entrada em `ACESSOS_PAGINA_` (a chave do
objeto tem que ser igual ao `data-tab` do botão no `Index.html`), chame
`requirePagina_(token, '<pagina>')` nas funções dela, e inclua o nome em `ABAS_RESTRITAS_`
no `Index.html`.

Pra "Minhas Férias" funcionar, a pessoa precisa ter o e-mail preenchido no Cadastro de
Funcionários (campo E-mail, sincronizável da planilha de RH) — sem isso o funcionário vê um
aviso pra pedir ao admin vincular o e-mail, em vez dos próprios dados.

## Anexos de Ajustes de horário
Arquivos enviados em Secretaria → Ajustes de horário (atestados, documentos etc.) são salvos
na pasta do Drive `1IRDW35AMspLcsIf3ke-DvWbxRbJtI4nP` (constante `ANEXOS_AJUSTES_FOLDER_ID`
em `Code.gs`) e compartilhados com "qualquer pessoa do domínio com o link" pra abrir direto
pelo ícone 📎 na tabela. Na primeira vez que isso rodar, o Apps Script vai pedir autorização
extra (escopo do Google Drive) — é só aceitar.

## Deploy
No Apps Script vinculado à planilha `Controle Operacional BOL` (ID no topo deste README),
cole `Code.gs` e `Index.html`, publique como Web App ("Executar como: eu",
"Quem tem acesso: qualquer pessoa") e atualize a constante `HUB_URL` em ambos os arquivos
se a URL do Hub mudar.
