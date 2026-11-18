# Módulo 04 — Eventos: macros que rodam sozinhas

**Duração alvo:** ~40 min · **Aulas:** 7 · **Nível:** já fez o Módulo 01 (ou sabe onde colar código) · **Arquivos:** `codigo/ThisWorkbook_eventos.txt`, `codigo/Planilha_eventos.txt`

**Promessa:** fazer o Excel reagir sozinho — ao abrir, salvar, alterar célula, dar duplo clique — sem o usuário apertar botão nenhum.

**Gancho:** "E se toda vez que alguém digitar um nome, o Excel já corrigisse para maiúsculo e tirasse os espaços sobrando?"

---

## Aula 4.1 — O que é um evento (4 min) · *prévia gratuita*

**Na tela:** demo final (MAIÚSCULO automático + check por duplo clique + log de alterações).
**Roteiro:** evento = "gatilho" que o Excel dispara. Dois grupos: **Workbook** (arquivo inteiro) e **Worksheet** (uma planilha). Mostre como descobrir eventos: caixas Objeto (esquerda) e Procedimento (direita) no topo do editor.

## Aula 4.2 — Eventos da pasta: Open, BeforeSave, BeforeClose (6 min)

**Na tela:** `ThisWorkbook_eventos.txt`.
**Roteiro:** teste um a um; o parâmetro `Cancel` (cancelar o salvamento/fechamento); registrar data/hora do último salvamento.

## Aula 4.3 — Worksheet_Change: o evento mais importante (8 min)

**Na tela:** `Worksheet_Change` do arquivo `Planilha_eventos.txt`.
**Roteiro:**
1. `Target` = células alteradas.
2. `Intersect(Target, Range(...))` — filtra só a área que interessa.
3. **O perigo do laço infinito:** alterar uma célula dentro de `Change` dispara `Change` de novo → `Application.EnableEvents = False`.
4. **Regra de ouro:** sempre religar no rótulo de saída, mesmo com erro. Demonstre o que acontece se o evento ficar desligado (tudo "morre" até reabrir o Excel) e como religar manualmente na janela Imediata: `Application.EnableEvents = True`.
5. Trate `Target` com várias células (colar dados) com `For Each`.

## Aula 4.4 — SelectionChange e BeforeDoubleClick (6 min)

**Na tela:** barra de status mostrando endereço; duplo clique marcando ✔.
**Roteiro:** `Cancel = True` evita entrar em modo de edição. `ChrW(10004)` para o símbolo ✔. Ideia de uso: lista de tarefas/checklist sem formulário.

## Aula 4.5 — BeforeRightClick e bloqueios (4 min)

**Na tela:** menu bloqueado na coluna D.
**Roteiro:** desabilitar menus/ações em áreas protegidas; quando **não** usar (acessibilidade e senso comum).

## Aula 4.6 — Eventos da pasta que valem para TODAS as planilhas (6 min)

**Na tela:** `Workbook_SheetChange` com planilha "Log".
**Roteiro:** parâmetro `Sh` (qual planilha); montar um **log de alterações** (data/hora, planilha, endereço); evitar log do próprio log; `SheetActivate` para informar onde o usuário está; `NewSheet`.

## Aula 4.7 — Boas práticas e como depurar eventos (6 min)

**Roteiro:**
- Checklist de segurança: `EnableEvents` religado? `ScreenUpdating` religado?
- Depuração: `F8` passo a passo, `Debug.Print`, ponto de interrupção (`F9`).
- Desempenho: não faça trabalho pesado dentro de `Change`; filtre com `Intersect`.
- Eventos **não** disparam quando a alteração vem de VBA com `EnableEvents = False`, nem em células calculadas por fórmula (use `Worksheet_Calculate`).
- Fecha: "juntando tudo" — o semáforo do Módulo 03 atualizado por `Worksheet_Change` (código no exercício 5).
