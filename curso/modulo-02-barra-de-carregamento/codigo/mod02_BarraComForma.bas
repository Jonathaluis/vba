Attribute VB_Name = "mod02_BarraComForma"
Option Explicit

' ============================================================
' Técnica 3: barra feita com FORMAS (Shapes) na própria planilha.
' Duas formas: um "fundo" cinza e uma "barra" que cresce.
' ============================================================

Private Const NOME_FUNDO As String = "PB_Fundo"
Private Const NOME_BARRA As String = "PB_Barra"

Public Sub BarraComForma()
    Const TOTAL As Long = 100
    Dim ws As Worksheet
    Dim i As Long

    Set ws = ActiveSheet
    On Error GoTo Limpa

    CriarBarra ws, ws.Range("B2").Left, ws.Range("B2").Top, 300, 22

    For i = 1 To TOTAL
        AtualizarBarra ws, i / TOTAL
        Pausa 0.02          ' <- troque pelo SEU trabalho real
    Next i

Limpa:
    RemoverBarra ws
    If Err.Number = 0 Then MsgBox "Concluído!", vbInformation
End Sub

' Cria as duas formas
Public Sub CriarBarra(ByVal ws As Worksheet, ByVal esq As Double, ByVal topo As Double, _
                      ByVal largura As Double, ByVal altura As Double)
    Dim fundo As Shape, barra As Shape

    RemoverBarra ws    ' garante que não sobrou barra antiga

    Set fundo = ws.Shapes.AddShape(msoShapeRoundedRectangle, esq, topo, largura, altura)
    fundo.Name = NOME_FUNDO
    fundo.Fill.ForeColor.RGB = RGB(230, 230, 230)
    fundo.Line.ForeColor.RGB = RGB(170, 170, 170)

    Set barra = ws.Shapes.AddShape(msoShapeRoundedRectangle, esq, topo, 1, altura)
    barra.Name = NOME_BARRA
    barra.Fill.ForeColor.RGB = RGB(39, 174, 96)
    barra.Line.Visible = msoFalse
End Sub

' pct entre 0 e 1
Public Sub AtualizarBarra(ByVal ws As Worksheet, ByVal pct As Double)
    Dim larg As Double
    If pct < 0 Then pct = 0
    If pct > 1 Then pct = 1

    larg = ws.Shapes(NOME_FUNDO).Width * pct
    If larg < 1 Then larg = 1                 ' largura 0 gera erro

    ws.Shapes(NOME_BARRA).Width = larg

    With ws.Shapes(NOME_FUNDO).TextFrame2.TextRange
        .Text = Format(pct, "0%")
        .Font.Size = 10
        .Font.Fill.ForeColor.RGB = RGB(60, 60, 60)
    End With
    DoEvents
End Sub

' Apaga as formas (sem erro se elas não existirem)
Public Sub RemoverBarra(ByVal ws As Worksheet)
    Dim s As Shape
    For Each s In ws.Shapes
        If s.Name = NOME_FUNDO Or s.Name = NOME_BARRA Then s.Delete
    Next s
End Sub

Private Sub Pausa(ByVal segundos As Double)
    Dim t0 As Double
    t0 = Timer
    Do While Timer >= t0 And Timer < t0 + segundos
        DoEvents
    Loop
End Sub
