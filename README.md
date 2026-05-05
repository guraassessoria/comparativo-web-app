# Comparativo automático por histórico

Aplicação web estática para analisar planilhas Excel, extrair nomes da coluna de histórico, comparar automaticamente os dois anos mais recentes encontrados na coluna de data e exportar um novo Excel tratado.

## Como usar

1. Abra `index.html` no navegador ou publique a pasta em uma hospedagem estática.
2. Faça upload do arquivo Excel.
3. Selecione:
   - Aba de origem.
   - Visão do período: mensal, trimestral, semestral ou anual.
   - Coluna de data.
   - Coluna de histórico.
   - Coluna de débito e crédito, ou uma coluna de valor pronto opcional.
   - Colunas de filtros adicionais opcionais. Pode selecionar mais de uma usando Ctrl/Cmd.
4. Clique em **Analisar planilha**.
5. Clique em **Exportar Excel tratado**.

## Regras aplicadas

- Os anos são lidos automaticamente pela coluna de data.
- A aplicação compara automaticamente os dois anos mais recentes encontrados. Exemplo: se existirem 2026 e 2025, o comparativo será 2026 x 2025.
- A visão mensal, trimestral, semestral ou anual controla como os períodos são agrupados.
- `AUTONOMO`, `AUTONOMOS` e `INSS S/ PF` são tratados como uma categoria única: `AUTONOMO`.
- Para autônomos, a aplicação consolida como `AUTONOMOS`, exceto quando o histórico indicar nome explícito em provisões.
- Quando existe referência de NF, a aplicação tenta extrair o nome após a NF.
- Para retenções como IRRF, ISS, INSS e PCC, a aplicação tenta buscar o nome completo usando a NF correspondente.
- Valor líquido padrão: `Débito - Crédito`.
- Se uma coluna de valor pronto for selecionada, ela substitui o cálculo `Débito - Crédito`.
- Blocos duplicados de pelo menos 20 linhas consecutivas são removidos automaticamente, quando a opção estiver marcada.

## Abas exportadas

- `Resumo`
- `Comparativo_Nome`
- `Resumo_Periodo`
- `Resumo_Categoria`
- `Resumo_Filtros`
- `Base_Tratada`
- `Duplicidades`

## Publicação online

A aplicação não precisa de backend. Pode ser publicada em:

- Netlify
- Vercel
- GitHub Pages
- Cloudflare Pages
- Qualquer hospedagem estática

A aplicação usa SheetJS via CDN:

```html
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
```

Por isso, a página publicada precisa ter acesso à internet para carregar essa biblioteca.
