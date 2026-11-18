# Módulo 06 — Velocidade e boas práticas

**Duração alvo:** ~35 min · **Aulas:** 6 · **Nível:** quem já escreve macros · **Arquivo:** `codigo/mod06_Turbo.bas`

**Promessa:** deixar suas macros muito mais rápidas e robustas, sem quebrar o Excel do usuário quando algo der errado.

**Gancho:** "A mesma macro: 14 segundos de um jeito, meio segundo de outro. Eu mostro a diferença."

---

## Aula 6.1 — O que deixa o VBA lento (5 min) · *prévia gratuita*

**Roteiro:** (1) Excel redesenhando a tela a cada alteração; (2) recalculando fórmulas a cada célula; (3) disparando eventos; (4) falar com a planilha **célula por célula** (a pior). Mostrar `TesteEscrita` rodando e os 3 tempos.

## Aula 6.2 — Modo turbo: ScreenUpdating, Calculation, EnableEvents (7 min)

**Na tela:** `LigarTurbo` / `DesligarTurbo`.
**Roteiro:** guardar o **estado original** do cálculo para devolver como estava (não presumir automático). Por que usar variável de módulo `turboLigado`.

## Aula 6.3 — Escrita e leitura em bloco (arrays) (7 min)

**Na tela:** terceira parte de `TesteEscrita`.
**Roteiro:** `Range.Value = array` e `array = Range.Value`; arrays 2D base 1; `ReDim`; `Resize`. Ligação com o Módulo 05.

## Aula 6.4 — Tratamento de erro que restaura o Excel (6 min)

**Na tela:** `ModeloDeMacroRobusta`.
**Roteiro:** o padrão `On Error GoTo Tratar` + rótulo `Sair:` que **sempre** restaura + `Resume Sair`. Demonstrar um erro proposital (divisão por zero) e ver tudo voltando ao normal. Mostrar o desastre sem tratamento: tela congelada, status bar presa, eventos desligados.

## Aula 6.5 — Última linha, última coluna e referências seguras (5 min)

**Na tela:** `UltimaLinhaEColuna`.
**Roteiro:** `Cells(Rows.Count, 1).End(xlUp).Row`; evitar `Select/Activate`; usar variáveis `Worksheet` e `Range`; qualificar sempre (`ws.Range`, não só `Range`).

## Aula 6.6 — Checklist final e hábitos (5 min)

**Slide-checklist:**
- [ ] `Option Explicit` no topo
- [ ] Nomes claros de variáveis e macros
- [ ] Sem `Select`/`Activate`
- [ ] Turbo ligado e **sempre** desligado (padrão com `Sair:`)
- [ ] Leitura/escrita em bloco
- [ ] Barra de progresso para macros longas (Módulo 02)
- [ ] Comentários onde o "porquê" não é óbvio
- [ ] Testar com poucos dados antes dos 50 mil

Indique os outros módulos e o **Combo** com desconto.
