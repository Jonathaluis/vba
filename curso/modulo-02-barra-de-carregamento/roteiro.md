# Módulo 02 — Barra de carregamento (4 técnicas)

**Duração alvo:** ~40 min · **Aulas:** 7 · **Nível:** quem já roda uma macro · **Arquivos:** `codigo/mod02_BarraStatusBarECelula.bas`, `mod02_BarraComForma.bas`, `frmProgresso_codigo.txt`, `mod02_DemoFormProgresso.bas`

**Promessa:** adicionar uma barra de progresso profissional em qualquer macro demorada, escolhendo entre 4 técnicas (da mais simples à mais bonita).

**Gancho:** "Sua macro demora 30 segundos e o usuário acha que o Excel travou. Uma barra de carregamento resolve isso em 10 linhas."

---

## Aula 2.1 — Por que (e quando) usar barra de progresso (4 min) · *prévia gratuita*

**Na tela:** macro lenta sem feedback × macro com barra.
**Roteiro:** o usuário aceita esperar quando **vê** que algo acontece. Mostre as 4 técnicas rodando em sequência (teaser) e o quadro comparativo:

| Técnica | Dificuldade | Visual | Quando usar |
|---|---|---|---|
| Barra de Status | ★ | simples | macros rápidas/internas |
| Barra em célula | ★ | simples | relatórios/dashboards |
| Barra com formas | ★★ | bom | dashboard, sem janela |
| UserForm | ★★★ | profissional | sistemas entregues a terceiros |

## Aula 2.2 — A ideia por trás de toda barra: percentual (5 min)

**Na tela:** `i / TOTAL`.
**Roteiro:** toda barra é `pct = feito / total` (entre 0 e 1). Mostre `Format(pct, "0%")`. Explique `DoEvents` (deixa o Excel redesenhar) e a função auxiliar `Pausa` (para simular trabalho nas demonstrações — troque pelo seu código real).

## Aula 2.3 — Barra na Barra de Status (6 min)

**Na tela:** `BarraNaStatusBar`.
**Roteiro:** `Application.StatusBar`, `MontarBarra` com `ChrW(9608)` (█) e `ChrW(9617)` (░).
**Armadilha:** se a macro der erro e a barra de status não for liberada, o texto fica preso → mostre `Application.StatusBar = False` no rótulo `Limpa:` e o padrão `On Error GoTo`.

## Aula 2.4 — Barra dentro de uma célula (4 min)

**Na tela:** `BarraNaCelula`.
**Roteiro:** reaproveita `MontarBarra`; fonte monoespaçada (Consolas) deixa a barra alinhada. Extra: mostre a versão sem VBA: `=REPT("█";ROUND(A1*20;0))`.

## Aula 2.5 — Barra com formas (Shapes) (8 min)

**Na tela:** `BarraComForma`.
**Roteiro:**
1. Duas formas: `PB_Fundo` e `PB_Barra`.
2. A "barra" cresce alterando `.Width = fundo.Width * pct`.
3. Por que `If larg < 1 Then larg = 1`.
4. `RemoverBarra` apaga no final (mesmo se houver erro).
5. Mostre como posicionar a barra em qualquer célula (`Range("B2").Left/.Top`).

**Gancho para o Módulo 03:** "Você acabou de manipular formas — no módulo 03 vamos muito além."

## Aula 2.6 — Barra em UserForm: montando o formulário (7 min)

**Na tela:** Inserir > UserForm → seguir a tabela de `frmProgresso_codigo.txt`.
**Roteiro:** mostre **a ordem dos Labels** (quem vem depois fica por cima) e as propriedades (Name, Left, Top, Width, Height, BackColor). Explique `lblBarra.Width = lblFundo.Width * pct`.
**Dica de gravação:** deixe a tabela de propriedades na tela por tempo suficiente e coloque no material complementar em PDF.

## Aula 2.7 — Barra em UserForm: usando na sua macro (6 min)

**Na tela:** `DemoFormProgresso` e `ExemploReal_MaiusculasComProgresso`.
**Roteiro:**
- `Show vbModeless` — por que a macro continua rodando.
- `Atualizar pct, mensagem` em cada passo.
- **Performance:** atualizar a barra a cada N linhas (`i Mod 10 = 0`), não em todas.
- Bloqueio do "X" com `QueryClose`.
- `Unload` no final, inclusive em caso de erro.

**Fechamento:** checklist para qualquer barra (slide): 1) calcular `pct`; 2) atualizar a cada N itens; 3) tratar erro; 4) limpar no final. Convite para o Módulo 06 (Velocidade).
