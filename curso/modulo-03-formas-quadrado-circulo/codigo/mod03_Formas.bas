Attribute VB_Name = "mod03_Formas"
Option Explicit

' ============================================================
' MÓDULO 03 - MANIPULANDO FORMAS (SHAPES)
' Criar, nomear, colorir, mover, girar, animar, ligar a células
' e transformar formas em botões.
' Todas as formas criadas aqui começam com o prefixo "demo_"
' para que a limpeza apague só o que é nosso.
' ============================================================

Private Const PREFIXO As String = "demo_"

' ---------- Aula 3.2: criar quadrado e círculo ----------
Public Sub CriarQuadradoECirculo()
    Dim ws As Worksheet, quadrado As Shape, circulo As Shape
    Set ws = ActiveSheet

    ' AddShape(Tipo, Esquerda, Topo, Largura, Altura)  -> medidas em PONTOS
    Set quadrado = ws.Shapes.AddShape(msoShapeRectangle, 40, 60, 100, 100)
    quadrado.Name = PREFIXO & "quadrado"
    quadrado.Fill.ForeColor.RGB = RGB(41, 128, 185)
    quadrado.Line.Visible = msoFalse
    quadrado.TextFrame2.TextRange.Text = "Quadrado"

    ' Um círculo é uma elipse com largura = altura
    Set circulo = ws.Shapes.AddShape(msoShapeOval, 180, 60, 100, 100)
    circulo.Name = PREFIXO & "circulo"
    circulo.Fill.ForeColor.RGB = RGB(192, 57, 43)
    circulo.Line.Visible = msoFalse
    circulo.TextFrame2.TextRange.Text = "Círculo"
End Sub

' ---------- Aula 3.3: propriedades ----------
Public Sub ModificarFormas()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    With ws.Shapes(PREFIXO & "quadrado")
        .Fill.ForeColor.RGB = RGB(142, 68, 173)   ' cor
        .Fill.Transparency = 0.3                  ' 0 = opaco, 1 = invisível
        .Rotation = 15                            ' graus
        .Line.Visible = msoTrue
        .Line.Weight = 3
        .Line.ForeColor.RGB = RGB(44, 62, 80)
        .Shadow.Visible = msoTrue
    End With

    With ws.Shapes(PREFIXO & "circulo")
        .Width = 140: .Height = 140               ' manter proporção de círculo
        .Left = 200
        .Fill.ForeColor.RGB = RGB(243, 156, 18)
    End With
End Sub

' ---------- Aula 3.4: listar e encontrar formas ----------
Public Sub ListarFormas()
    Dim s As Shape
    For Each s In ActiveSheet.Shapes
        Debug.Print s.Name, "Tipo=" & s.AutoShapeType, _
                    "Esq=" & Round(s.Left), "Topo=" & Round(s.Top), _
                    "Larg=" & Round(s.Width), "Alt=" & Round(s.Height)
    Next s
    MsgBox "Veja a lista na janela Imediata (Ctrl+G).", vbInformation
End Sub

' ---------- Aula 3.5: prender a forma a uma célula ----------
' Faz um círculo do tamanho exato da célula B8 (útil p/ "bolinhas de status")
Public Sub FormaNaCelula()
    Dim ws As Worksheet, alvo As Range, s As Shape
    Set ws = ActiveSheet
    Set alvo = ws.Range("B8")

    Set s = ws.Shapes.AddShape(msoShapeOval, alvo.Left + 2, alvo.Top + 2, _
                               alvo.Height - 4, alvo.Height - 4)
    s.Name = PREFIXO & "bolinha_B8"
    s.Fill.ForeColor.RGB = RGB(39, 174, 96)
    s.Line.Visible = msoFalse
    s.Placement = xlMoveAndSize       ' acompanha a célula se linhas/colunas mudarem
End Sub

' ---------- Aula 3.6: animação ----------
Public Sub AnimarBola()
    Dim ws As Worksheet, bola As Shape, x As Double

    Set ws = ActiveSheet
    Set bola = ws.Shapes.AddShape(msoShapeOval, 20, 200, 40, 40)
    bola.Name = PREFIXO & "bola"
    bola.Fill.ForeColor.RGB = RGB(231, 76, 60)

    On Error GoTo Fim
    For x = 20 To 420 Step 4           ' vai
        bola.Left = x
        Pausa 0.005
    Next x
    For x = 420 To 20 Step -4          ' volta
        bola.Left = x
        Pausa 0.005
    Next x
Fim:
    bola.Delete
End Sub

Public Sub GirarQuadrado()
    Dim s As Shape, i As Long
    Set s = ActiveSheet.Shapes(PREFIXO & "quadrado")
    For i = 1 To 72
        s.Rotation = s.Rotation + 5     ' 72 x 5 = 360 graus
        Pausa 0.01
    Next i
End Sub

' ---------- Aula 3.7: semáforo ligado a um valor ----------
Public Sub CriarSemaforo()
    Dim ws As Worksheet, caixa As Shape, i As Long, s As Shape
    Dim nomes As Variant
    Set ws = ActiveSheet
    nomes = Array("vermelho", "amarelo", "verde")

    Set caixa = ws.Shapes.AddShape(msoShapeRoundedRectangle, 340, 40, 60, 170)
    caixa.Name = PREFIXO & "sem_caixa"
    caixa.Fill.ForeColor.RGB = RGB(44, 62, 80)
    caixa.Line.Visible = msoFalse

    For i = 0 To 2
        Set s = ws.Shapes.AddShape(msoShapeOval, 350, 50 + i * 53, 40, 40)
        s.Name = PREFIXO & "sem_" & nomes(i)
        s.Line.Visible = msoFalse
        s.Fill.ForeColor.RGB = RGB(90, 90, 90)   ' apagada
    Next i
End Sub

' valor de 0 a 100: >=80 verde | >=50 amarelo | <50 vermelho
Public Sub AtualizarSemaforo(ByVal valor As Double)
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ws.Shapes(PREFIXO & "sem_vermelho").Fill.ForeColor.RGB = RGB(90, 90, 90)
    ws.Shapes(PREFIXO & "sem_amarelo").Fill.ForeColor.RGB = RGB(90, 90, 90)
    ws.Shapes(PREFIXO & "sem_verde").Fill.ForeColor.RGB = RGB(90, 90, 90)

    If valor >= 80 Then
        ws.Shapes(PREFIXO & "sem_verde").Fill.ForeColor.RGB = RGB(46, 204, 113)
    ElseIf valor >= 50 Then
        ws.Shapes(PREFIXO & "sem_amarelo").Fill.ForeColor.RGB = RGB(241, 196, 15)
    Else
        ws.Shapes(PREFIXO & "sem_vermelho").Fill.ForeColor.RGB = RGB(231, 76, 60)
    End If
End Sub

Public Sub DemoSemaforo()
    Dim resp As String
    resp = InputBox("Digite um valor de 0 a 100:", "Semáforo", "75")
    If Len(resp) = 0 Then Exit Sub
    If Not IsNumeric(resp) Then MsgBox "Valor inválido.": Exit Sub
    AtualizarSemaforo CDbl(resp)
End Sub

' ---------- Aula 3.8: forma como BOTÃO ----------
' Liga a macro "ClicouNaForma" à forma "demo_quadrado"
Public Sub TransformarQuadradoEmBotao()
    ActiveSheet.Shapes(PREFIXO & "quadrado").OnAction = "ClicouNaForma"
End Sub

' Application.Caller devolve o NOME da forma clicada
Public Sub ClicouNaForma()
    Dim s As Shape
    Set s = ActiveSheet.Shapes(Application.Caller)

    If s.Fill.ForeColor.RGB = RGB(39, 174, 96) Then
        s.Fill.ForeColor.RGB = RGB(192, 57, 43)
        s.TextFrame2.TextRange.Text = "OFF"
    Else
        s.Fill.ForeColor.RGB = RGB(39, 174, 96)
        s.TextFrame2.TextRange.Text = "ON"
    End If
End Sub

' ---------- Aula 3.9: limpeza ----------
' Percorre de trás para frente: apagar mudando a contagem exige isso.
Public Sub ApagarFormasDemo()
    Dim ws As Worksheet, i As Long
    Set ws = ActiveSheet
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).Name Like PREFIXO & "*" Then ws.Shapes(i).Delete
    Next i
End Sub

Private Sub Pausa(ByVal segundos As Double)
    Dim t0 As Double
    t0 = Timer
    Do While Timer >= t0 And Timer < t0 + segundos
        DoEvents
    Loop
End Sub
