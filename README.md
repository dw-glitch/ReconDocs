# ReconDocs

Aplicativo web com dois módulos independentes:

| Rota | Módulo | Para que serve |
| --- | --- | --- |
| `/` | Conferência SGP × SIGEM | Conferência fechada das três planilhas do processo |
| `/cruzamento` | Cruzamento inteligente de planilhas | Cruzamento dinâmico de N planilhas quaisquer |

Os arquivos são processados localmente no navegador nos dois módulos.

## Módulo 1 — Conferência SGP × SIGEM (`/`)

Aplicativo web para cruzar três planilhas:

- Lista do SGP;
- Consulta Geral do SIGEM;
- Documentos Previstos.

O **ReconDocs** compara SGP e SIGEM. A planilha **Documentos Previstos** é usada somente para informar `Alocado` quando o código for encontrado e `Não alocado` quando não for encontrado.

### Regras principais

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

### Relatório Excel

A exportação cria as abas `Resumo`, `Todos`, `Diferenças`, `Não alocados`, `Revisar` e `Metodologia`. As colunas são agrupadas por assunto e comparam somente SGP e SIGEM; Documentos Previstos aparece apenas no resultado `Alocado?`.

## Módulo 2 — Cruzamento inteligente de planilhas (`/cruzamento`)

Cruzamento dinâmico de quantas planilhas forem necessárias, sem depender de um layout específico.

### Planilhas

- **Consulta Geral** — documentos já postados no SIGEM (número, status, data e demais colunas).
- **Documentos Previstos** — documentos planejados; presente na planilha significa `Alocado = Sim`, ausente significa `Alocado = Não`. Atualizada manualmente todo dia.
- **Planilhas adicionais** — quantas o usuário quiser, com o botão `Adicionar planilha`.

Cada planilha carregada permite escolher a aba, o papel no cruzamento, o nome usado no relatório e o mapeamento das colunas: documento (obrigatório), status, data e quantas outras colunas relevantes forem marcadas. A aba, a linha de cabeçalho e as colunas são detectadas automaticamente e podem ser corrigidas à mão, então qualquer planilha nova é aceita.

### Comparação dos documentos

- **Inteligente** (padrão) — ignora caminho, extensão e revisão colada ao fim do código, reaproveitando a normalização do módulo de conferência.
- **Exata** — compara o texto normalizado (acentos, espaços e caixa).

### Resultado na tela

Para cada documento o cruzamento mostra em que planilhas ele existe, o status em cada uma, se está alocado, o que está ausente e as observações da linha. Os filtros cobrem: todos, em todas as planilhas, em apenas algumas, só na Consulta Geral, só nos Documentos Previstos, não alocados, ausentes na Consulta Geral e divergências de status.

Indicadores exibidos: total analisado, total na Consulta Geral, total nos Documentos Previstos, total em comum, só na Consulta Geral, só nos Documentos Previstos, alocados, não alocados, divergências de status, presentes em todas as planilhas, presentes em apenas algumas e ausentes na Consulta Geral.

### Relatório Excel

| Aba | Conteúdo |
| --- | --- |
| `Resumo Executivo` | Todos os indicadores consolidados, o total por planilha e o critério de comparação usado |
| `Resultado Consolidado` | Documento, status na Consulta Geral, existe na Consulta Geral, existe nos Documentos Previstos, alocado, status das outras planilhas, observações, presença, data, complementos e um par `Existe`/`Status` por planilha carregada |
| `Somente Consulta Geral` | Documentos encontrados apenas na Consulta Geral |
| `Somente Documentos Previstos` | Documentos encontrados apenas nos Documentos Previstos |
| `Divergências` | Documentos com status diferente entre as planilhas, com uma coluna de status por planilha |
| `Planilhas Adicionais` | Criada quando há planilhas além das duas principais |

## Executar

```bash
npm install
npm run dev
```

Para validar:

```bash
npm test        # build + testes de renderização e dos motores
npm run test:engine   # somente os motores, sem build
```
