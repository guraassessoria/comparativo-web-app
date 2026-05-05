# Motor de comparação e conciliação por histórico

Aplicação web estática para processar planilhas Excel no navegador, extrair informações da coluna de histórico e exportar análises tratadas em Excel.

## Publicação

A aplicação não precisa de backend. Publique os arquivos abaixo na raiz do GitHub Pages, Netlify, Vercel ou qualquer hospedagem estática:

- `index.html`
- `styles.css`
- `app.js`
- `.nojekyll`

A biblioteca SheetJS é carregada via CDN em `index.html`.

## Fluxo de uso

1. Faça upload de uma planilha `.xlsx`, `.xls`, `.xlsm` ou `.csv`.
2. Selecione a aba de origem.
3. Escolha o tipo de análise:
   - **Comparação**
   - **Conciliação**
4. Selecione:
   - coluna de data;
   - coluna de histórico;
   - coluna de débito;
   - coluna de crédito;
   - ou coluna de valor pronto, se houver.
5. Opcionalmente, selecione filtros adicionais.
6. Opcionalmente, selecione campos a extrair do histórico.
7. Clique em **Analisar planilha**.
8. Clique em **Exportar Excel tratado**.

## Modo Comparação

Compara automaticamente os dois anos mais recentes encontrados na coluna de data.

Exemplo: se a base tiver 2026 e 2025, a aplicação gera o comparativo 2026 x 2025.

A visão de período pode ser:

- mensal;
- trimestral;
- semestral;
- anual.

Abas exportadas no modo comparação:

- `Resumo`
- `Comparativo_Nome`
- `Resumo_Periodo`
- `Resumo_Categoria`
- `Resumo_Filtros`
- `Resumo_NF`
- `Resumo_Extracoes`
- `Extracao_Nome`
- `Base_Tratada`
- `Duplicidades`

## Modo Conciliação

Agrupa os lançamentos por uma ou mais chaves e verifica se os valores se compensam dentro de uma tolerância.

Chaves disponíveis:

- nome extraído;
- nota fiscal;
- categoria tratada;
- ano;
- período;
- filtros adicionais selecionados;
- campos extraídos do histórico.

Configurações disponíveis:

- tolerância de valor;
- coluna opcional de origem/lado, como Banco/Contábil ou Extrato/Razão;
- opção para exigir sinais opostos dentro do grupo.

Classificações possíveis:

- `Conciliado`
- `Conciliado - um lado`
- `Diferença`
- `Sem contraparte`
- `Sem sinais opostos`

Abas exportadas no modo conciliação:

- `Resumo`
- `Conciliacao_Grupos`
- `Conciliacao_Linhas`
- `Conciliacao_Pares`
- `Resumo_NF`
- `Resumo_Extracoes`
- `Base_Tratada`
- `Duplicidades`

## Campos extraídos do histórico

Extrações padrão disponíveis:

- Nota Fiscal
- CNPJ/CPF
- Competência
- Contrato
- Pedido/OC
- Parcela
- Documento

Também é possível criar extrações personalizadas, uma por linha, no formato:

```text
Projeto=PROJ(?:ETO)?\s*[:\-]?\s*([A-Z0-9 ._/-]+)
Centro=CC\s*[:\-]?\s*([0-9.]+)
```

A primeira captura entre parênteses será exportada.

## Regras fixas

- `AUTONOMO`, `AUTONOMOS` e `INSS S/ PF` são tratados como `AUTONOMO`.
- Valor líquido padrão: `Débito - Crédito`.
- Quando selecionada uma coluna de valor pronto, ela substitui o cálculo `Débito - Crédito`.
- A aplicação pode remover blocos duplicados automaticamente.
- O processamento ocorre localmente no navegador.
