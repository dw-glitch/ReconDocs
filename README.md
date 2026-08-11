# ReconDocs — Conferência SGP × SIGEM

Aplicativo web para cruzar três planilhas:

- Lista do SGP;
- Consulta Geral do SIGEM;
- Documentos Previstos.

O **ReconDocs** compara SGP e SIGEM. A planilha **Documentos Previstos** é usada somente para informar `Alocado` quando o código for encontrado e `Não alocado` quando não for encontrado.

## Regras principais

- Comparação por índices em memória, adequada a bases com mais de 20.000 linhas.
- Detecção automática da aba, linha de cabeçalho e colunas; o usuário pode revisar todos os mapeamentos.
- Separação do código e da revisão quando ambos aparecem na mesma célula; cada revisão é exibida e comparada em sua própria coluna.
- Escopo selecionável: todos, ET, CV ou N-1710; dentro de ET, todos os ET, Doc RIR ou Doc de C&M.
- Documentos ET são procurados na forma informada e na forma alternativa com/sem `nt-` minúsculo no início do 7º grupo.
- O resultado informa explicitamente se encontrou com `nt-`, sem `nt-` ou nas duas formas, identificando as fontes.
- Documentos N-1710 não recebem a regra `nt-`.
- Nenhum código, revisão, título, status ou detalhe de Documentos Previstos é exibido no relatório.
- Diferenças entre SGP e SIGEM são mostradas ao lado da coluna correta: código, revisão, título ou status.
- Os arquivos são processados localmente no navegador.

## Relatório Excel

A exportação cria as abas `Resumo`, `Todos`, `Diferenças`, `Não alocados`, `Revisar` e `Metodologia`. As colunas são agrupadas por assunto e comparam somente SGP e SIGEM; Documentos Previstos aparece apenas no resultado `Alocado?`.

## Executar

```bash
npm install
npm run dev
```

Para validar:

```bash
npm test
```
