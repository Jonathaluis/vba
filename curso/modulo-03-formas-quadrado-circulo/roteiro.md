# Módulo 03 — Manipulando formas: quadrado, círculo e além

**Duração alvo:** ~45 min · **Aulas:** 9 · **Nível:** quem já roda uma macro · **Arquivo:** `codigo/mod03_Formas.bas`

**Promessa:** controlar qualquer forma da planilha por código — criar, colorir, mover, girar, animar e usar como botão ou indicador (semáforo).

**Gancho:** "Quer um painel com bolinhas verdes e vermelhas que mudam sozinhas? Isso é só uma propriedade: `.Fill.ForeColor.RGB`."

---

## Aula 3.1 — O que é uma Shape e o que dá para fazer (5 min) · *prévia gratuita*

**Na tela:** inserir formas manualmente → mostrar o que vamos automatizar (demo final do módulo: semáforo + bola animada + botão ON/OFF).
**Roteiro:** coleção `ActiveSheet.Shapes`; cada forma tem **Name**, **Left/Top**, **Width/Height**, **Fill**, **Line**, **Rotation**. Medidas em **pontos** (1 pt = 1/72 pol). `Left` e `Top` contam a partir do canto da planilha.

## Aula 3.2 — Criando um quadrado e um círculo (6 min)

**Na tela:** `CriarQuadradoECirculo`.
**Roteiro:** `AddShape(Tipo, Esq, Topo, Larg, Alt)` → `msoShapeRectangle` e `msoShapeOval`. Círculo = elipse com largura = altura. **Sempre dar `.Name`** (e por que o prefixo `demo_`).

## Aula 3.3 — Cor, transparência, borda, rotação (6 min)

**Na tela:** `ModificarFormas`.
**Roteiro:** `RGB(r,g,b)`, `Fill.Transparency`, `Line.Weight`, `Rotation`, `Shadow`. Truque: gravar uma macro com o gravador, alterar a forma manualmente e ver o código gerado. Lista de `AutoShapeType` mais úteis: seta, triângulo, estrela, balão, coração.

## Aula 3.4 — Encontrando e listando formas (5 min)

**Na tela:** `ListarFormas` + painel de seleção (Alt+F10).
**Roteiro:** `For Each s In Shapes`, acessar por nome `Shapes("demo_quadrado")` ou por índice, `Debug.Print`. Cuidado: os nomes automáticos ("Retângulo 4") mudam.

## Aula 3.5 — Prendendo a forma a uma célula (5 min)

**Na tela:** `FormaNaCelula`.
**Roteiro:** usar `Range.Left/Top/Height` para alinhar. `Placement = xlMoveAndSize`. Caso de uso: bolinha de status ao lado de cada linha de uma tabela.

## Aula 3.6 — Animação: movendo e girando (6 min)

**Na tela:** `AnimarBola`, `GirarQuadrado`.
**Roteiro:** laço `For x = 20 To 420 Step 4` alterando `.Left`; `Pausa` controla a velocidade; `Rotation + 5`. Por que a tela precisa estar atualizando (`ScreenUpdating = True`) para ver a animação. Apagar a bola no `Fim:` mesmo em caso de erro.

## Aula 3.7 — Semáforo ligado a um valor (7 min)

**Na tela:** `CriarSemaforo`, `AtualizarSemaforo`, `DemoSemaforo`.
**Roteiro:** 3 círculos dentro de uma caixa escura; `AtualizarSemaforo(valor)` apaga todos e acende o certo (`If / ElseIf`). **Ligar a uma célula:** no código da planilha chame `AtualizarSemaforo Me.Range("B2").Value` dentro de `Worksheet_Change` (aproveite o Módulo 04).

## Aula 3.8 — Forma como botão (ON/OFF) (4 min)

**Na tela:** `TransformarQuadradoEmBotao` e `ClicouNaForma`.
**Roteiro:** `.OnAction = "NomeDaMacro"`; `Application.Caller` devolve o nome da forma clicada → **uma macro serve para várias formas**. Troca de cor/texto ON/OFF.

## Aula 3.9 — Limpando tudo (e fechamento) (3 min)

**Na tela:** `ApagarFormasDemo`.
**Roteiro:** apagar com laço **de trás para frente** (`Step -1`) e `Like "demo_*"`. Recapitule o módulo e proponha o desafio final: "Painel de KPIs com 5 semáforos".
