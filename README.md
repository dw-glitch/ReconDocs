# ReconDocs

Aplicativo web com dois módulos independentes:

| Rota | Módulo | Para que serve |
| --- | --- | --- |
| `/` | Conferência SGP × SIGEM | Conferência fechada das três planilhas do processo |
| `/cruzamento` | Cruzamento de planilhas | Cruzamento de N planilhas quaisquer, de qualquer layout |

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

## Módulo 2 — Cruzamento de planilhas (`/cruzamento`)

Cruzamento de quantas planilhas você quiser, de qualquer origem e layout. Não é uma segunda tela do processo SGP × SIGEM: nada no módulo é fixo em uma base específica.

### Planilhas

Carregue quantas planilhas precisar — não há limite de quantidade. Para cada uma você define:

- **Nome no relatório** — o nome que você der aparece em todos os títulos, colunas e indicadores.
- **Papel no cruzamento** — `Planilha comum` (padrão), `Base de referência` ou `Base de alocação`.
- **Aba analisada** e o **mapeamento das colunas**: documento (obrigatório), status, data e quantas outras colunas relevantes forem marcadas.

A aba, a linha de cabeçalho e as colunas são detectadas automaticamente em qualquer layout e podem ser corrigidas à mão.

### Papéis opcionais

Os dois papéis especiais são opcionais e existem só para quem precisa deles:

- **Base de referência** — ativa as colunas `Existe` e `Status` daquela base, o filtro de ausentes e o indicador correspondente.
- **Base de alocação** — estar nesta planilha significa `Alocado = Sim`; não estar, `Alocado = Não`.

Sem nenhum papel atribuído, o cruzamento é puramente genérico: presença por planilha, status por planilha, exclusivos e divergências. As colunas e abas que dependem dos papéis simplesmente não aparecem.

No processo SGP × SIGEM, por exemplo, a Consulta Geral entra como base de referência e os Documentos Previstos como base de alocação — mas isso é uma configuração, não uma regra do módulo.

### Comparação dos documentos

- **Inteligente** (padrão) — ignora caminho, extensão e revisão colada ao fim do código.
- **Exata** — compara o texto normalizado (acentos, espaços e caixa).

### Resultado na tela

Para cada documento: em quais planilhas existe, o status em cada uma, o que está ausente e as observações da linha. Filtros: todos, em todas, em algumas, exclusivos e divergências — mais os filtros de referência e alocação quando esses papéis existirem. Os indicadores são gerados a partir das planilhas carregadas, um por planilha, com os nomes que você deu.

### Relatório Excel

| Aba | Quando aparece | Conteúdo |
| --- | --- | --- |
| `Resumo Executivo` | sempre | Indicadores consolidados, total por planilha e critério de comparação |
| `Resultado Consolidado` | sempre | Documento, observações, presença, complementos e uma dupla `Existe`/`Status` por planilha; as colunas de referência e de alocação entram quando esses papéis existem |
| `Somente <base de referência>` | com base de referência | Documentos encontrados apenas nessa base |
| `Somente <base de alocação>` | com base de alocação | Documentos encontrados apenas nessa base |
| `Exclusivos por planilha` | sempre | Documentos presentes em uma única planilha, com o nome dela |
| `Divergências` | sempre | Documentos com status diferente entre as planilhas, uma coluna de status por planilha |
| `Planilhas Adicionais` | com planilhas além dos papéis | Detalhe por planilha |

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
