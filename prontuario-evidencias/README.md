# Evidências — Portal Prontuário (Sambaíba)

Ferramenta que captura as telas do Portal de Evidências
(`https://nimer-prontuario.sambaibasp.cloud`) e monta um PowerPoint de
demonstração do projeto.

## Como rodar

```bash
pip install playwright requests
playwright install chromium
npm install pptxgenjs

export PORTAL_USER='seu_usuario'
export PORTAL_PASS='sua_senha'

python3 capturar.py        # login no portal -> shots/*.png (2x, página inteira)
node montar_deck.js        # shots/ -> Evidencias_Portal_Prontuario.pptx
```

## Estrutura do deck (12 slides, 16:9)

1. Capa
2. Problema → solução
3. Visão geral em números (KPIs da tela Início)
4. Diagrama do fluxo (Nova evidência → Validação ADH → Fila → Enviada → Retorno)
5–10. Uma tela por módulo (Nova evidência, Evidências, Validação, Fila Nimer,
   Retorno Nimer, Usuários)
11. Status atual e próximos passos

## Privacidade (LGPD)

As telas de Evidências, Validação e Retorno Nimer exibem **dados pessoais e
disciplinares de colaboradores** (nome, RE, descrição da falha, fotos). Por
isso `shots/` e o `.pptx` gerado **não são versionados** (ver `.gitignore`).
Só o código da ferramenta fica no repositório. Cuidado ao compartilhar o
PowerPoint gerado.
