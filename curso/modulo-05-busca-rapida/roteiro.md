# Módulo 05 — Busca rápida (Find, arrays, Dictionary e filtro)

**Duração alvo:** ~40 min · **Aulas:** 7 · **Nível:** já usa VBA básico · **Arquivos:** `codigo/mod05_Busca.bas`, `codigo/Planilha_BuscaAoDigitar.txt`

**Promessa:** buscar em dezenas de milhares de linhas em fração de segundo — e saber escolher a técnica certa para cada situação.

**Gancho:** "Percorrer 50 mil células uma por uma leva segundos. Com um array leva centésimos. Vamos medir isso ao vivo."

---

## Aula 5.1 — Preparando 50 mil linhas de teste (4 min) · *prévia gratuita*

**Na tela:** `GerarDadosTeste`.
**Roteiro:** gerar dados em memória (`dados(1 To QTD, 1 To 3)`) e escrever **de uma vez** em `Range.Resize(...).Value`. Isso já adianta a lição do módulo.

## Aula 5.2 — Find e FindNext: achar todas as ocorrências (7 min)

**Na tela:** `BuscarTodasAsOcorrencias`.
**Roteiro:** parâmetros `LookIn`, `LookAt` (`xlPart` × `xlWhole`), `MatchCase`; guardar o `Address` da primeira para saber quando a busca "deu a volta"; acumular com `Union`; selecionar tudo. Armadilha: o VBA **não** faz curto-circuito em `And` → por isso o `If achado Is Nothing Then Exit Do`.
**Atenção:** `Find` mantém as configurações da última busca (inclusive a do usuário) → sempre informe os parâmetros.

## Aula 5.3 — Array em memória: a busca veloz (7 min)

**Na tela:** `BuscarComArray`.
**Roteiro:** `dados = Range("B2:B50001").Value` cria um array 2D (linha, coluna) base 1; `UBound(dados, 1)`; `InStr` com `vbTextCompare`. Por que é rápido: ler/escrever na planilha é caro, a memória é barata.

## Aula 5.4 — O duelo: célula a célula × array (5 min)

**Na tela:** `DueloDeVelocidade`.
**Roteiro:** rodar e comentar os tempos (varia com a máquina; normalmente o array é dezenas de vezes mais rápido). Use `Timer` para medir.

## Aula 5.5 — Dictionary: o PROCV instantâneo (8 min)

**Na tela:** `BuscarComDicionario`.
**Roteiro:** `CreateObject("Scripting.Dictionary")`, `.Exists`, `.Add`, `.Item`. Monta o índice **uma vez**, depois cada busca é imediata — ideal para várias consultas ou cruzamento entre bases. Chave duplicada: use `Exists` antes de `Add`. `CompareMode = vbTextCompare` ignora maiúsculas. **Aviso:** Dictionary é recurso do Windows (no Mac use `Collection`).

## Aula 5.6 — Filtro automático (AutoFilter) por texto (5 min)

**Na tela:** `FiltrarPorTermo` e `DemoFiltro`.
**Roteiro:** `Range.CurrentRegion.AutoFilter Field:=2, Criteria1:="*" & termo & "*"`; `AutoFilterMode = False` remove; coringas `*` e `?`.

## Aula 5.7 — Busca ao digitar + escolhendo a técnica (7 min)

**Na tela:** `Planilha_BuscaAoDigitar.txt` — digitar em F1 e ver a tabela filtrar.
**Roteiro:** junta Módulo 04 (evento) + AutoFilter, com `EnableEvents` e `ScreenUpdating` protegidos. Feche com a **tabela de decisão**:

| Situação | Técnica |
|---|---|
| Achar todas as células com um texto | Find/FindNext |
| Contar/varrer muitos dados | Array |
| Várias consultas por chave | Dictionary |
| Mostrar ao usuário só as linhas que combinam | AutoFilter |
