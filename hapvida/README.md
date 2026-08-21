# Conciliação Cadastral e Financeira — Convênio Médico Hapvida

Ferramenta de Gestão de Benefícios/RH que cruza a **base do convênio Hapvida** com a
**base do sistema interno da empresa**, identifica quem está divergente (pessoa a pessoa,
não só quantidade) e calcula o impacto financeiro de regularizar os cadastros.

## Arquivos

| Arquivo | O que é |
| --- | --- |
| `gerar_planilha.py` | Gera a pasta de trabalho a partir do CSV do Hapvida |
| `verificar.py` | Recalcula o modelo em Python e confere célula a célula contra a planilha |
| `dados/base_hapvida.csv` | Exportação do convênio (não versionada — ver *Privacidade*) |
| `Conciliacao_Hapvida_x_Sistema.xlsx` | Planilha entregue (não versionada — ver *Privacidade*) |

## Como gerar

```bash
pip install openpyxl
python3 gerar_planilha.py            # usa dados/base_hapvida.csv
python3 gerar_planilha.py <csv> <saida.xlsx>
```

Conferência independente (exige a planilha já recalculada):

```bash
python3 verificar.py dados/base_hapvida.csv Conciliacao_Hapvida_x_Sistema.xlsx
```

## Abas

1. **INSTRUÇÕES** — como atualizar as bases, o que cada aba faz, legenda das situações
2. **RESUMO GERENCIAL** — painel: colaboradores, dependentes, agregados, faixa etária, financeiro, impacto
3. **FICHA DO COLABORADOR** — digite um RE e a ficha inteira se monta (titular, dependentes, agregados, cadastro do sistema, divergências, custo e simulação)
4. **CONCILIAÇÃO** — uma linha por pessoa, cruzando as duas bases, com situação e observação
5. **TITULARES** — uma linha por RE: quantidades, valores, situação e alertas
6. **SIMULAÇÃO DE IMPACTO** — quanto custaria incluir no Hapvida tudo o que já consta no sistema
7. **RESUMO POR CONTRATO** — titulares, dependentes, agregados e valores por contrato
8. **ALERTAS** — só os REs que exigem ação
9. **HAPVIDA** / **SISTEMA** — as duas bases (cole aqui para atualizar tudo)
10. **PARÂMETROS** — valores do contrato e regras (único lugar onde se mudam preços)

## Regras financeiras

- **Titular:** R$ 151,73 — 100% custeado pela empresa.
- **Combo Familiar:** definido pela **faixa da quantidade total de dependentes**
  (1 → 239,93 · 2 → 446,62 · 3 → 579,83 · 4 → 663,57 · 5 ou mais → 846,86).
  Nunca valor unitário × quantidade. Responsabilidade do colaborador.
- **Agregados:** R$ 544,95 cada, somados individualmente e sempre separados do Combo.
  Responsabilidade do colaborador.

A regra foi confirmada contra o próprio arquivo: em 97,7% dos titulares o campo `cobrado`
é exatamente `151,73 + Combo(nº de dependentes)`, e todo agregado é cobrado a 544,95.

## Decisões de modelagem

Três pontos em que a planilha se afasta de uma leitura literal da especificação, todos
visíveis e configuráveis:

- **Regra de idade e cônjuge.** A especificação diz que dependente acima de 25 anos vira
  agregado. Aplicada ao pé da letra, ela converteria **1.491 cônjuges** em agregados
  (R$ 812.000/mês a mais), porque cônjuge é dependente independentemente da idade.
  A planilha aplica a regra apenas a filhos/enteados/tutelados — restando **32 casos
  reais**. O comportamento é configurável em `PARÂMETROS` ("Aplicar a regra de idade ao
  CÔNJUGE?").
- **Idades recalculadas.** A coluna `idade` do arquivo é o retrato da data de extração e
  já diverge da data de nascimento em 176 linhas. Toda idade da pasta é recalculada a
  partir do nascimento, usando a data-base de `PARÂMETROS`.
- **RE ausente na base da empresa não é erro.** O vínculo é classificado pelo CONTRATO do
  Hapvida (PJ, Acordo, Sindicato) e, quando não é possível identificar, fica como
  `NÃO LOCALIZADO – VERIFICAR`.

## Problemas encontrados na base do Hapvida

Levantados pela própria planilha (aba ALERTAS) e confirmados fora dela:

- **21 registros duplicados** (mesmo RE, mesma data de nascimento, mesmo nome) — excluídos
  das contagens e sinalizados.
- **2 matrículas compartilhadas por titulares diferentes** (REs 18858 e 18959) — a planilha
  mantém uma linha por RE, conforme a regra de agrupamento, e emite alerta.
- **1 RE sem linha de titular** (RE 3988: um cônjuge sem titular vinculado).
- **152 titulares com cobrança divergente** do contrato, somando **R$ 8.009,85/mês**.
  Cerca de 100 deles têm cobrança exatamente dobrada (303,46 = 2 × 151,73;
  783,32 = 2 × 391,66; 1.196,70 = 2 × 598,35), quase todos em admissões recentes —
  compatível com cobrança retroativa, mas vale conferir com o convênio.

## Notas técnicas

- **Tudo é fórmula viva.** Ao substituir as bases nas abas HAPVIDA e SISTEMA, todas as
  demais abas se recalculam. Capacidades pré-preenchidas: 9.700 linhas (Hapvida),
  10.000 (sistema), 6.200 titulares, 3.000 alertas.
- **Chave de cruzamento:** 1º RE + data de nascimento; 2º RE + nome normalizado
  (maiúsculas, sem acentos e sem pontuação). Quem casa pela data mas não pelo nome vira
  `⚠️ DIVERGÊNCIA DE DADOS`.
- **Compatibilidade:** só ÍNDICE/CORRESP, CONT.SES, SOMASES, DESLOC e IFERROR — abre em
  Excel 2016 em diante, LibreOffice e Google Sheets. Sem `PROCX`/`XLOOKUP` (ausente no
  Excel 2016/2019) e sem `AGREGAR`/`AGGREGATE` (não interpretado por todas as
  implementações).
- **Datas em texto** no formato DD/MM/AAAA são convertidas por uma fórmula explícita, e não
  por `DATA.VALOR`/`DATEVALUE` — que depende do idioma do Excel e leria "05/09/2002" como
  9 de maio.

## Verificação

- 804.419 fórmulas, **zero erros** de fórmula (recalculadas no LibreOffice).
- `verificar.py` reconstrói o modelo em Python a partir do CSV e confere os **5.860
  titulares** em nome, contrato, nº de dependentes, nº de agregados, Combo, valor de
  agregados, total, valor cobrado e elegíveis a agregado: **0 divergências**.
- Teste de conciliação com base do sistema sintética cobrindo os 6 cenários (cadastro
  idêntico, nome grafado diferente, dependente novo até 25 anos, dependente novo acima de
  25, dependentes faltando no sistema, RE ausente): **300 REs, 0 divergências**, incluindo
  valor novo, aumento mensal e aumento anual.
- O exemplo do próprio enunciado é reproduzido: 1 → 2 dependentes = **R$ 206,69/mês** e
  **R$ 2.480,28/ano**.

## Privacidade

`dados/*.csv` e o `.xlsx` gerado **não são versionados** (`.gitignore`): contêm nome,
nome da mãe e data de nascimento de mais de 9.000 pessoas. Para regenerar, coloque a
exportação do convênio em `hapvida/dados/base_hapvida.csv`.
