# Comparativo e Conciliação de Excel

Aplicação web estática para análise de planilhas Excel diretamente no navegador.

## Funcionalidades

- Upload de arquivos `.xlsx`, `.xls`, `.xlsm` ou `.csv`.
- Processamento local no navegador, sem envio da planilha para servidor.
- Seleção manual das colunas necessárias após o upload. Se existir uma aba chamada `Base unificada`, ela é selecionada automaticamente como origem:
  - data;
  - histórico;
  - débito;
  - crédito;
  - valor pronto opcional;
  - filtros adicionais opcionais.
- Seleção do tipo de análise:
  - comparação;
  - conciliação.
- Leitura automática dos períodos pela coluna de data.
- Visões mensal, trimestral, semestral ou anual.
- Extração de informações do histórico:
  - nome;
  - nota fiscal;
  - CNPJ/CPF;
  - competência;
  - contrato;
  - pedido/OC;
  - parcela;
  - documento;
  - extrações personalizadas por regex.
- Consolidação de `AUTONOMO`, `AUTONOMOS` e `INSS S/ PF` como `AUTONOMO`.
- Exportação para Excel com abas de resultado.

## Ajuste visual da versão v11

Os seletores de coluna exibem **somente a letra da coluna e o cabeçalho detectado**.

Exemplo:

```text
A — DATA
B — CONTA CONTABIL
C — CENTRO CUSTO
E — HISTORICO
H — DEBITO
I — CREDITO
```

A aplicação não mostra mais amostras das primeiras linhas em menus, títulos, tooltips ou mapa de colunas.

Quando a planilha não possui cabeçalho detectável, os seletores exibem apenas:

```text
A — Coluna A
B — Coluna B
C — Coluna C
```

## Publicação no GitHub Pages

Suba estes arquivos na raiz do repositório:

```text
index.html
styles.css
app.js
README.md
.nojekyll
```

Depois vá em:

```text
Settings > Pages > Build and deployment
```

Configure:

```text
Source: Deploy from a branch
Branch: main
Folder: / (root)
```

Salve e aguarde a publicação.

## Observação

A aplicação usa SheetJS via CDN. O navegador precisa ter acesso ao script externo para ler e gerar arquivos Excel.
