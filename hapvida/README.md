# Conciliação Cadastral e Financeira — Convênio Médico Hapvida

Ferramenta de Gestão de Benefícios/RH que cruza a **base do convênio Hapvida** com a
**base do sistema interno da empresa**, identifica pessoa a pessoa quem está em cada
base e calcula o impacto financeiro de regularizar os cadastros.

## Arquivos

| Arquivo | O que é |
| --- | --- |
| `gerar_planilha.py` | Gera a pasta de trabalho a partir do arquivo com as duas bases |
| `compactar.py` | Converte colunas repetidas em `sharedFormula` (roda no fim da geração) |
| `verificar.py` | Recalcula o modelo em Python e confere campo a campo contra a planilha |
| `recalcular.py` | Recalcula no LibreOffice e aponta células com erro de fórmula |
| `dados/bases_origem.xlsx` | Abas HAPVIDA e SISTEMA JB (não versionado — ver *Privacidade*) |
| `Conciliacao_Hapvida_x_Sistema.xlsx` | Planilha entregue (não versionada) |

## Como gerar e conferir

```bash
pip install openpyxl
python3 gerar_planilha.py                       # usa dados/bases_origem.xlsx
python3 recalcular.py Conciliacao_Hapvida_x_Sistema.xlsx      # ~26 min
python3 verificar.py dados/bases_origem.xlsx Conciliacao_Hapvida_x_Sistema.xlsx
```

O `verificar.py` reconstrói o modelo inteiro em Python — inclusive a cascata de
correspondência de nomes — e compara 18 campos por titular, mais a simulação.

## Abas

1. **INSTRUÇÕES** — como atualizar as bases, como os nomes são comparados, legenda das situações
2. **RESUMO GERENCIAL** — painel: colaboradores, dependentes, agregados, faixa etária, financeiro, qualidade do cruzamento e impacto
3. **FICHA DO COLABORADOR** — digite um RE e a ficha se monta: titular, **um único quadro com toda a composição familiar**, custo atual, cenário e impacto
4. **CONCILIAÇÃO** — uma linha por pessoa, cruzando as duas bases, com situação e observação
5. **TITULARES** — uma linha por RE: quantidades, valores, situação e alertas
6. **SIMULAÇÃO DE IMPACTO** — quanto custaria incluir no Hapvida tudo o que já consta no sistema
7. **RESUMO POR CONTRATO** — titulares, dependentes, agregados e valores por contrato
8. **ALERTAS** — só os REs que exigem ação
9. **HAPVIDA** / **SISTEMA** — as duas bases (cole aqui para atualizar tudo)
10. **PARÂMETROS** — valores, idade limite e a tabela de RELAÇÃO → TIPO

## Como os nomes são comparados

O mesmo dependente aparece escrito de formas diferentes nas duas bases —
`MARIA SILVA` no sistema e `MARIA APARECIDA DA SILVA` no Hapvida. Comparar nome
completo apontaria essa pessoa como não cadastrada.

Dentro de cada RE, a busca roda em quatro níveis:

| Nível | Chave | Resultado |
| --- | --- | --- |
| 1 | data de nascimento | 🟢 FORTE |
| 2 | nome completo padronizado | 🟢 FORTE |
| 3 | primeiro nome + último sobrenome | 🟢 PROVÁVEL (nome abreviado) |
| 4 | só o primeiro nome | 🟡 **VERIFICAR CORRESPONDÊNCIA** |

O nome padronizado (maiúsculas, sem acento, sem pontuação) vive só em colunas
auxiliares — **os nomes originais das duas bases nunca são alterados**.

Quem cai no nível 4 **não é tratado como não cadastrado e fica fora da simulação**:
não se cria custo em cima de dúvida. O grau de confiança de cada pessoa fica visível
na coluna *Correspondência* das abas HAPVIDA e SISTEMA, e o RESUMO GERENCIAL traz o
bloco *Qualidade do cruzamento de nomes*.

**Limite conhecido:** a busca devolve a primeira pessoa que casa. Se duas pessoas do
mesmo RE tiverem o mesmo primeiro nome e nada mais as separar, ambas caem em
VERIFICAR — que é exatamente o sinal de conferir à mão.

## Regras financeiras

- **Titular:** R$ 151,73 — 100% custeado pela empresa.
- **Combo Familiar:** definido pela **faixa da quantidade total de dependentes**
  (1 → 239,93 · 2 → 446,62 · 3 → 579,83 · 4 → 663,57 · 5 ou mais → 846,86).
  Nunca valor unitário × quantidade. Responsabilidade do colaborador.
- **Agregados:** R$ 544,95 cada, somados individualmente e sempre separados do Combo.
  Responsabilidade do colaborador.

A regra foi confirmada contra o próprio arquivo: na quase totalidade dos titulares o
campo `cobrado` é exatamente `151,73 + Combo(nº de dependentes)`, e todo agregado é
cobrado a 544,95.

## Decisões de modelagem

Todas visíveis e configuráveis na aba PARÂMETROS:

- **Tabela RELAÇÃO → TIPO.** Cada relação das duas bases (`FILHO(A)`, `CONJUGE`,
  `EX CONJUGE`, `AGREGADO(A)`, `COMPANHEIRO(A)`, `FILHO UNIVERS`, `MENOR POBRE`…)
  vira TITULAR, DEPENDENTE, AGREGADO ou NÃO ELEGÍVEL. O que não estiver na tabela
  entra como DEPENDENTE sujeito à regra de idade.
- **EX CONJUGE = NÃO ELEGÍVEL.** São 265 pessoas na base do sistema. Ex-cônjuge não é
  dependente de plano de saúde; incluí-las inflaria a simulação em centenas de
  milhares de reais. Para mudar, basta trocar o TIPO na tabela.
- **A regra dos 25 anos não se aplica ao cônjuge.** Ao pé da letra, ela converteria
  mais de mil cônjuges em agregados. A coluna *Sujeito à regra de idade?* vem com NÃO
  para `CONJUGE` e `COMPANHEIRO(A)`.
- **Idades recalculadas** da data de nascimento com a data-base de PARÂMETROS — a
  coluna de idade dos arquivos é o retrato da data da extração.
- **RE ausente na base da empresa não é erro.** O vínculo sai do CONTRATO do Hapvida
  (PJ, Acordo, Sindicato) e, sem identificação, fica `NÃO LOCALIZADO – VERIFICAR`.

## O que as bases mostram hoje

| | |
| --- | --- |
| Titulares · dependentes · agregados no Hapvida | 5.860 · 3.235 · 63 |
| Custo mensal atual | R$ 1.636.842,45 |
| Empresa (titulares) · Colaboradores (combo + agregados) | R$ 889.137,80 · R$ 747.704,65 |
| Pessoas do sistema sem cadastro no Hapvida | 7.773 |
| A incluir: dependentes · agregados | 5.451 · 1.446 |
| **Impacto se tudo for regularizado** | **R$ 1.877.871,34/mês · R$ 22.534.456,08/ano** |
| Colaboradores impactados | 3.369 |

Dos 1.446 que entrariam como agregado, **1.431 são por passarem dos 25 anos** —
sozinhos representam cerca de R$ 780 mil por mês. É o principal vetor do impacto.

## Problemas encontrados nas bases

Levantados pela própria planilha (aba ALERTAS):

- **21 registros duplicados** no Hapvida (mesmo RE, nascimento e nome) — excluídos das
  contagens e sinalizados.
- **2 matrículas compartilhadas por titulares diferentes** (REs 18858 e 18959).
- **1 RE sem linha de titular** no Hapvida (RE 3988).
- **152 titulares com cobrança divergente** do contrato, somando R$ 8.009,85/mês; cerca
  de 100 com cobrança exatamente dobrada, quase todos em admissões recentes.
- **914 pessoas com divergência de dados** entre as bases — em geral nome abreviado ou
  com erro de digitação (casadas pela data de nascimento, não por nome), e **45 casos em
  que o Hapvida cobra como AGREGADO alguém que o sistema registra como FILHO(A)**.
- **531 colaboradores no sistema sem nenhum convênio** e **136 REs do Hapvida que não
  existem na base da empresa** e cujo contrato não permite identificar o vínculo.

## Notas técnicas

- **Tudo é fórmula viva.** Ao substituir as bases, todas as abas se recalculam.
  Capacidades: 9.700 linhas (Hapvida), 17.200 (sistema), 6.200 titulares,
  9.000 pessoas no bloco "somente sistema", 6.200 alertas.
- **Layout do export do sistema:** uma linha por pessoa. Na linha do próprio
  colaborador `RELACAO = TITULAR` e `NOME DEPENDENTE` fica vazio (o nome sai da
  coluna B); nas demais a pessoa é o `NOME DEPENDENTE`. A data de nascimento da
  pessoa é sempre a coluna G. A planilha resolve isso sozinha.
- **Compatibilidade:** só ÍNDICE/CORRESP, CONT.SES, SOMASES, DESLOC e SEERRO — abre em
  Excel 2016 em diante, LibreOffice e Google Sheets. Sem `PROCX` (ausente no Excel
  2016/2019) e sem `AGREGAR` (não interpretado por todas as implementações).
- **Datas em texto** DD/MM/AAAA são convertidas por fórmula explícita, e não por
  `DATA.VALOR`, que depende do idioma e leria "05/09/2002" como 9 de maio.
- **Fórmulas compartilhadas.** São ~985 mil fórmulas. Gravadas uma a uma, dariam
  260 MB de XML e minutos para abrir. O `compactar.py` grava cada coluna repetida
  uma única vez usando `sharedFormula` (mecanismo do próprio formato OOXML): o
  arquivo cai de 19 MB para 8 MB e o XML de 262 MB para 93 MB, sem alterar
  nenhum resultado.
- **Tempo de recálculo.** No LibreOffice o recálculo completo leva ~26 minutos,
  porque ele varre cada CORRESP linha a linha. O Excel mantém índice de busca e
  faz o mesmo trabalho em segundos — é para ele que a pasta foi construída.

## Privacidade

`dados/*.xlsx`, `dados/*.csv` e o `.xlsx` gerado **não são versionados**: contêm nome,
nome da mãe e data de nascimento de mais de 25.000 pessoas. Para regenerar, coloque o
arquivo com as duas abas em `hapvida/dados/bases_origem.xlsx`.

## Verificação

- **985.034 fórmulas · zero células com erro** (recalculadas no LibreOffice).
- `verificar.py` reconstrói o modelo inteiro em Python — inclusive a cascata de
  correspondência de nomes — e confere os **5.860 titulares** em 18 campos cada
  (nome, contrato, quantidades nas duas bases, encontrados, só no Hapvida, a incluir
  como dependente e como agregado, elegíveis, Combo, agregados, total, cobrado,
  divergências, a verificar), mais valor novo, aumento mensal e aumento anual da
  simulação: **0 divergências**.
- O exemplo do enunciado é reproduzido: 1 → 2 dependentes = **R$ 206,69/mês** e
  **R$ 2.480,28/ano**.
