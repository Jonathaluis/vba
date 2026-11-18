# Micro-cursos de VBA — estrutura e plano de venda

Série de **6 micro-cursos independentes** de Excel VBA, cada um resolvendo uma necessidade específica ("compro só o que preciso"), mais um **Combo** com todos.

## Mapa dos módulos

| # | Micro-curso | O aluno sai sabendo… | Aulas | Duração alvo | Pasta |
|---|---|---|---|---|---|
| 01 | **Onde escrever macros** | Módulo × ThisWorkbook × Planilha × Function | 7 | ~40 min | `modulo-01-onde-escrever-macros` |
| 02 | **Barra de carregamento** | 4 técnicas de progress bar (status, célula, formas, UserForm) | 7 | ~40 min | `modulo-02-barra-de-carregamento` |
| 03 | **Formas: quadrado, círculo e animações** | Criar/colorir/mover/animar shapes, semáforo, botões | 9 | ~45 min | `modulo-03-formas-quadrado-circulo` |
| 04 | **Eventos: macros que rodam sozinhas** | Change, Open, DoubleClick, log de alterações | 7 | ~40 min | `modulo-04-eventos` |
| 05 | **Busca rápida** | Find, array, Dictionary, AutoFilter, busca ao digitar | 7 | ~40 min | `modulo-05-busca-rapida` |
| 06 | **Velocidade e boas práticas** | Modo turbo, arrays, erro seguro, checklist | 6 | ~35 min | `modulo-06-velocidade-e-boas-praticas` |

Cada pasta tem: `roteiro.md` (aula a aula, com o que mostrar e falar), `exercicios.md` (com gabarito) e `codigo/` (arquivos prontos para importar).

**Ordem sugerida / dependências:** 01 → (02, 03, 04, 05 em qualquer ordem) → 06. Os módulos 03+04 se combinam (semáforo atualizado por evento) e 05+04 (busca ao digitar). Cada módulo é compreensível sozinho.

## Como usar o código (aviso importante)

- Arquivos `.bas`: no editor VBA, **Arquivo > Importar arquivo…**. Já estão em **ANSI (Windows-1252) com quebra de linha CRLF**, que é o formato que o editor espera (acentos corretos).
- Arquivos `.txt`: são trechos para **colar** em `ThisWorkbook`, em uma planilha ou em um UserForm — o cabeçalho de cada um diz onde.
- Salve sempre como `.xlsm` (ou `.xlsb`).
- ⚠️ **Teste antes de gravar.** Este material foi escrito e revisado sem acesso ao Excel (ambiente Linux). Antes de gravar cada aula, importe o código no Excel (Windows) e rode **cada macro**, para ter certeza de que se comporta como no roteiro. Corrija qualquer detalhe de versão. Registre a versão do Excel usada (ex.: Microsoft 365) na descrição do curso.
- `Scripting.Dictionary` (módulo 05) e a janela "Confiar no acesso ao modelo de objeto" (aula 1.6) são específicos do Windows — avise na descrição.

## Plano comercial para a Udemy

### Regras que afetam o formato (confirme no painel, mudam com o tempo)
- A Udemy costuma exigir **no mínimo ~30 min de vídeo e 5 aulas** por curso para publicar. Os módulos acima foram dimensionados para isso (35–45 min). Se algum ficar abaixo de 30 min depois de gravado, acrescente uma aula prática ou o exercício resolvido em vídeo.
- Preços são escolhidos em **faixas fixas (tiers)** definidas pela Udemy; o preço mais baixo permitido costuma ser perto de **R$ 30** — confira a tabela vigente no painel. Os descontos de campanha da Udemy podem reduzir o preço de lista, e você recebe uma fatia menor nas vendas feitas via promoção/plataforma. Considere isso na conta.
- **Se quiser preços realmente baixos** (R$ 9–19 por módulo), compare com plataformas de infoprodutos (Hotmart, Eduzz, Kiwify etc.), onde você define o preço livremente e fica com a maior parte, mas precisa trazer o tráfego. Uma estratégia comum: **Udemy para descoberta + site próprio para o Combo.**

### Estratégia de preço sugerida
- Cada micro-curso: preço da faixa mínima/baixa da Udemy.
- **Combo (curso "Completo")**: 6 módulos num curso único, com preço de 2,5 a 3× um módulo (mais barato que comprar tudo separado). Na Udemy, o Combo é um curso separado que reaproveita os vídeos — não use o mesmo título; deixe claro "contém os módulos 1 a 6".
- Atualize o Combo conforme novos módulos sejam lançados (aumenta a percepção de valor).
- Uma **aula de prévia gratuita** por curso (marcada nos roteiros) — a melhor vendedora é a demonstração do resultado final.

### Títulos e subtítulos (modelo)

| Módulo | Título (≤ 60 car.) | Subtítulo |
|---|---|---|
| 01 | VBA: Onde Colocar Cada Macro (Módulo, Planilha, Pasta) | Entenda de uma vez onde escrever o código e pare de se perder no editor. |
| 02 | VBA: Barra de Carregamento Profissional no Excel | 4 técnicas prontas: barra de status, célula, formas e UserForm. |
| 03 | VBA: Formas, Círculos, Animações e Semáforos no Excel | Crie e controle formas por código para painéis e botões interativos. |
| 04 | VBA: Eventos — Macros que Rodam Sozinhas | Faça o Excel reagir ao abrir, salvar e editar células, sem botões. |
| 05 | VBA: Busca Rápida em Milhares de Linhas | Find, arrays, Dictionary e filtro: escolha a técnica certa. |
| 06 | VBA: Deixe Suas Macros Mais Rápidas e Seguras | Modo turbo, arrays e tratamento de erro em 35 minutos. |
| Combo | VBA Prático: 6 Micro-Cursos em 1 (Barras, Formas, Eventos, Busca) | Todos os módulos com projetos prontos para copiar e adaptar. |

### Descrição-modelo (módulo 02 — adapte para os demais)

> **Sua macro demora e o usuário acha que o Excel travou?**
> Neste curso rápido você aprende a criar barras de carregamento profissionais em VBA, com 4 técnicas diferentes: na Barra de Status, dentro de uma célula, com formas (Shapes) e em UserForm.
>
> **O que você vai aprender**
> - Calcular e exibir progresso em qualquer macro longa
> - Criar uma barra de progresso com formas e com UserForm
> - Tratar erros para nunca deixar o Excel "preso"
> - Aplicar o código pronto às suas próprias planilhas
>
> **Para quem é:** quem já sabe rodar uma macro e quer deixar seus arquivos mais profissionais.
> **Pré-requisitos:** Excel para Windows com guia Desenvolvedor (ensinamos a ativar).
> **Material incluso:** códigos prontos, exercícios com gabarito, apostila em PDF.

### Palavras-chave / tags
`vba excel`, `macro excel`, `automação excel`, `userform`, `barra de progresso excel`, `excel vba formas`, `eventos vba`, `busca vba`, `excel avançado`.

### Requisitos de qualidade (checklist de gravação)
- Áudio limpo (microfone USB/lapela, sem eco); vídeo em **1080p**, fonte do editor VBA **grande** (16–18 pt), zoom do Excel ≥ 125%.
- Aulas de 4–8 minutos, **uma ideia por aula**, sempre terminando com o resultado funcionando.
- Mostrar o resultado **antes** de explicar o código (gancho).
- Legenda em português (a Udemy permite enviar `.srt`).
- Material complementar em cada módulo: arquivos `.bas`/`.xlsm` e a apostila.
- Imagem de capa 750×422 px com o resultado visual (barra, formas, semáforo) — alto contraste.
- Vídeo promocional de 60–90 s mostrando o resultado final.

### Próximos módulos (ideias para ampliar o catálogo)
UserForms de cadastro (CRUD) · Enviar e-mail pelo Outlook com VBA · Gerar PDFs automaticamente · Ler/gravar arquivos CSV e TXT · Controles de data (calendário) · Gráficos por VBA · Proteger planilhas e código · Conectar com Access/SQL (ADO).

## Apostila

`apostila/apostila-vba-micro-cursos.md` (e versões `.docx`/`.pdf`): guia de consulta rápida com os conceitos e códigos-chave de cada módulo, para entregar como material complementar.
