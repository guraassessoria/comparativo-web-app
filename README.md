# Motor de análise e conciliação por histórico

Aplicação web estática para análise contábil a partir de arquivos Excel/CSV.

## O que faz

- Upload de um ou mais arquivos `.xlsx`, `.xls`, `.xlsm` ou `.csv`.
- Processamento local no navegador, sem backend.
- Mapeamento manual das colunas de:
  - data;
  - histórico;
  - débito;
  - crédito;
  - valor pronto opcional;
  - filtros adicionais opcionais.
- Extração automática de informações do histórico:
  - nome;
  - nota fiscal;
  - CNPJ/CPF;
  - competência;
  - contrato;
  - pedido/OC;
  - parcela;
  - documento;
  - campos personalizados por regex.
- Tratamento de `AUTONOMO`, `AUTONOMOS` e `INSS S/ PF` como `AUTONOMO`.
- Exportação de Excel tratado.

## Tipos de análise

### 1. Comparação

Compara automaticamente os dois anos mais recentes encontrados na coluna de data. O usuário pode escolher a visão:

- mensal;
- trimestral;
- semestral;
- anual.

### 2. Conciliação

A conciliação pode ser feita de duas formas.

#### Dentro de um único arquivo/aba

Agrupa lançamentos pelas chaves selecionadas e verifica se os valores se compensam dentro da tolerância.

Uso típico:

- provisão x baixa;
- débito x crédito;
- imposto retido x nota fiscal;
- lançamentos duplicados ou pendentes dentro da mesma base.

#### Entre arquivos/abas

Permite carregar dois ou mais arquivos, ou usar abas diferentes de um mesmo arquivo, e comparar totais por origem usando as chaves selecionadas.

Uso típico:

- razão contábil x extrato;
- contas a pagar x contabilidade;
- sistema operacional x ERP;
- base do cliente x base interna.

Neste modo, cada fonte selecionada vira uma origem/lado. A aplicação pode conciliar por:

- **saldo do grupo zerado**, útil quando os valores aparecem com sinais opostos; ou
- **totais absolutos por origem**, útil quando os arquivos têm valores com o mesmo sinal.

### Tratamento de divergência de valor

Quando a conciliação usa **Nota Fiscal + Nome extraído** como chave e encontra contraparte pela chave, mas o valor fica diferente acima da tolerância, o grupo passa a ser marcado como **Divergência de valor**.

Na aba `Conciliacao_Pares`, o par correspondente é marcado como **Verificar - valor divergente**, em vez de aparecer apenas como linha sem par. Isso permite auditar rapidamente casos em que a identificação bate, mas o valor não fecha.

## Publicação no GitHub Pages

1. Suba estes arquivos na raiz do repositório:
   - `index.html`
   - `styles.css`
   - `app.js`
   - `README.md`
   - `.nojekyll`
2. Vá em **Settings > Pages**.
3. Em **Build and deployment**, selecione **Deploy from a branch**.
4. Escolha:
   - branch: `main`
   - folder: `/ (root)`
5. Clique em **Save**.

## Observação importante

Na conciliação entre arquivos/abas, a aplicação usa o mesmo mapeamento de colunas para todas as fontes selecionadas. Portanto, o melhor resultado ocorre quando os arquivos têm layout equivalente ou colunas nas mesmas posições.

## Ajuste de rótulos das colunas

Quando a planilha não possui cabeçalho detectável, os seletores exibem apenas a letra da coluna, por exemplo `A — Coluna A`, `B — Coluna B` etc. As amostras das células não são mais concatenadas no rótulo do seletor, para evitar opções confusas em abas resumidas, pivotadas ou com várias linhas de título.

Se a aba tiver cabeçalho claro, o seletor exibirá a letra e o nome do cabeçalho, por exemplo `D — Histórico`.
