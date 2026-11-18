---
title: "VBA na Prática — Guia de Consulta Rápida"
subtitle: "Material complementar dos micro-cursos"
lang: pt-BR
---

# Como usar esta apostila

Este guia reúne, em formato de consulta rápida, o essencial dos 6 micro-cursos. Os códigos completos estão nos arquivos de cada módulo. Salve seus arquivos como `.xlsm`, ative a guia **Desenvolvedor** e use `Alt+F11` para abrir o editor.

**Atalhos úteis do editor:** `F5` executar · `F8` passo a passo · `F9` ponto de interrupção · `Ctrl+G` janela Imediata · `Ctrl+Espaço` autocompletar · `Alt+F8` lista de macros.

# Módulo 01 — Onde escrever macros

| Quero… | Coloque em… |
|---|---|
| Macro para botão, atalho ou reuso geral | **Módulo padrão** (Inserir > Módulo) |
| Fórmula personalizada para usar na célula | **Function** em módulo padrão |
| Reagir a algo do arquivo inteiro (abrir, salvar, fechar) | **ThisWorkbook** |
| Reagir a algo de uma planilha (alterar, selecionar) | **Código da planilha** (Plan1 etc.) |
| Uma janela para o usuário | **UserForm** |
| Objetos reutilizáveis | **Módulo de classe** |

**Public × Private:** `Public Sub` aparece em `Alt+F8` e pode ser chamada de outros módulos; `Private Sub` é uma peça interna do módulo.

**ThisWorkbook × ActiveWorkbook:** `ThisWorkbook` é a pasta onde o código mora; `ActiveWorkbook` é a que está em foco.

```vba
Option Explicit

Public Function DobroDe(ByVal numero As Double) As Double
    DobroDe = numero * 2
End Function
```

# Módulo 02 — Barra de carregamento

Toda barra é `pct = feito / total` (0 a 1).

| Técnica | Dificuldade | Quando usar |
|---|---|---|
| Barra de Status | ★ | macros rápidas, uso interno |
| Barra em célula | ★ | relatórios e dashboards |
| Barra com formas | ★★ | dashboard sem janela |
| UserForm | ★★★ | sistemas entregues a terceiros |

```vba
' Barra de texto (blocos cheios e vazios)
Private Function MontarBarra(ByVal pct As Double, ByVal largura As Long) As String
    Dim cheios As Long
    cheios = CLng(pct * largura)
    MontarBarra = String(cheios, ChrW(9608)) & String(largura - cheios, ChrW(9617))
End Function

' Uso:
Application.StatusBar = MontarBarra(i / total, 25) & " " & Format(i / total, "0%")
' ... e no final, SEMPRE:
Application.StatusBar = False
```

**Checklist:** 1) calcular o percentual · 2) atualizar a cada N itens, não a todos · 3) tratar erro · 4) limpar no final.

# Módulo 03 — Formas (Shapes)

```vba
Set s = ActiveSheet.Shapes.AddShape(msoShapeRectangle, esq, topo, larg, alt)  ' quadrado/retângulo
Set s = ActiveSheet.Shapes.AddShape(msoShapeOval, esq, topo, 100, 100)        ' círculo
s.Name = "demo_quadrado"
s.Fill.ForeColor.RGB = RGB(41, 128, 185)
s.Fill.Transparency = 0.3
s.Line.Visible = msoFalse
s.Rotation = 15
s.TextFrame2.TextRange.Text = "Texto"
s.OnAction = "MinhaMacro"        ' forma vira botão
```

| Propriedade | Para que serve |
|---|---|
| `.Left` `.Top` | posição (pontos) |
| `.Width` `.Height` | tamanho |
| `.Fill.ForeColor.RGB` | cor de preenchimento |
| `.Line.Weight` / `.Line.Visible` | borda |
| `.Rotation` | giro em graus |
| `.Placement` | como se comporta ao mover/redimensionar células |

**Apagar com segurança:** percorra de trás para frente (`For i = .Count To 1 Step -1`) e filtre pelo nome (`Like "demo_*"`).

**Qual forma foi clicada?** `Application.Caller` devolve o nome da forma que chamou a macro.

# Módulo 04 — Eventos

| Evento | Onde | Quando dispara |
|---|---|---|
| `Workbook_Open` | ThisWorkbook | ao abrir o arquivo |
| `Workbook_BeforeSave` | ThisWorkbook | antes de salvar |
| `Workbook_BeforeClose` | ThisWorkbook | antes de fechar |
| `Workbook_SheetChange` | ThisWorkbook | alteração em qualquer planilha |
| `Worksheet_Change` | Planilha | célula alterada |
| `Worksheet_SelectionChange` | Planilha | outra seleção |
| `Worksheet_BeforeDoubleClick` | Planilha | duplo clique |
| `Worksheet_BeforeRightClick` | Planilha | botão direito |

**Padrão seguro para `Change`:**

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Me.Range("B2:B1000")) Is Nothing Then Exit Sub
    On Error GoTo Fim
    Application.EnableEvents = False
    ' ... altere células aqui ...
Fim:
    Application.EnableEvents = True     ' nunca esqueça
End Sub
```

Se os eventos "morrerem", digite na janela Imediata: `Application.EnableEvents = True`.

# Módulo 05 — Busca rápida

| Situação | Técnica |
|---|---|
| Achar todas as células com um texto | `Find` / `FindNext` |
| Contar/varrer muitos dados | Array (`dados = Range.Value`) |
| Várias consultas por chave | `Scripting.Dictionary` |
| Mostrar só linhas que combinam | `AutoFilter` |

```vba
dados = ws.Range("B2:B" & ultima).Value           ' lê uma vez
For i = 1 To UBound(dados, 1)
    If InStr(1, dados(i, 1), termo, vbTextCompare) > 0 Then achou = achou + 1
Next i
```

```vba
Set dic = CreateObject("Scripting.Dictionary")
dic.CompareMode = vbTextCompare
If Not dic.Exists(chave) Then dic.Add chave, linha
If dic.Exists(busca) Then MsgBox dados(dic(busca), 2)
```

**Atenção:** o VBA **não** faz curto-circuito em `And`/`Or`: `If x Is Nothing Or x.Valor > 0` ainda avalia `x.Valor`. Aninhe os `If`s.

# Módulo 06 — Velocidade e boas práticas

**Modo turbo (com restauração):**

```vba
Public Sub MinhaMacro()
    On Error GoTo Tratar
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' ... seu código ...

Sair:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
Tratar:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical
    Resume Sair
End Sub
```

**Última linha da coluna A:** `ws.Cells(ws.Rows.Count, 1).End(xlUp).Row`

**Escrever em bloco:** monte um array e faça `ws.Range("A1").Resize(n, m).Value = dados`.

## Checklist final

- [ ] `Option Explicit` no topo
- [ ] Sem `Select` / `Activate`
- [ ] Variáveis `Worksheet` e `Range` qualificadas (`ws.Range`)
- [ ] Turbo ligado e **sempre** desligado
- [ ] Leitura/escrita em bloco
- [ ] Barra de progresso em macros longas
- [ ] Testar com poucos dados antes

# Glossário rápido

**Macro:** sequência de instruções VBA. **Sub:** macro que executa ações. **Function:** devolve um valor. **Evento:** gatilho disparado pelo Excel. **Shape:** forma/objeto gráfico. **Array:** lista de valores em memória. **Dictionary:** coleção chave→valor. **Target:** célula(s) afetada(s) por um evento. **Escopo:** onde um nome (variável/macro) é visível.
