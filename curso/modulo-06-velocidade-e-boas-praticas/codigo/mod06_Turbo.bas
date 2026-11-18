Attribute VB_Name = "mod06_Turbo"
Option Explicit

' ============================================================
' MÓDULO 06 - VELOCIDADE E BOAS PRÁTICAS
' 1) Modo turbo (liga/desliga seguro)
' 2) Teste: laço célula a célula x escrita em bloco
' 3) Padrão de tratamento de erro que SEMPRE restaura o Excel
' 4) Última linha / última coluna com segurança
' ============================================================

' Guardamos o estado original para devolver exatamente como estava
Private estadoCalculo As XlCalculation
Private turboLigado As Boolean

Public Sub LigarTurbo()
    If turboLigado Then Exit Sub
    estadoCalculo = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    turboLigado = True
End Sub

Public Sub DesligarTurbo()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If turboLigado Then Application.Calculation = estadoCalculo
    Application.StatusBar = False
    turboLigado = False
End Sub

' ---------- Teste: lento x rápido ----------
Public Sub TesteEscrita()
    Const LINHAS As Long = 20000
    Dim ws As Worksheet, i As Long, t0 As Double
    Dim tLento As Double, tTurbo As Double, tBloco As Double
    Dim dados() As Variant

    Set ws = ActiveSheet
    ws.Range("A1:A" & LINHAS).ClearContents

    ' 1) Lento: célula a célula, com tela atualizando
    t0 = Timer
    For i = 1 To LINHAS
        ws.Cells(i, 1).Value = i * 2
    Next i
    tLento = Timer - t0

    ' 2) Turbo: célula a célula, mas com Excel "quieto"
    ws.Range("A1:A" & LINHAS).ClearContents
    t0 = Timer
    On Error GoTo Restaura
    LigarTurbo
    For i = 1 To LINHAS
        ws.Cells(i, 1).Value = i * 2
    Next i
    DesligarTurbo
    tTurbo = Timer - t0

    ' 3) Bloco: monta em memória e escreve de uma vez
    ws.Range("A1:A" & LINHAS).ClearContents
    t0 = Timer
    ReDim dados(1 To LINHAS, 1 To 1)
    For i = 1 To LINHAS
        dados(i, 1) = i * 2
    Next i
    ws.Range("A1").Resize(LINHAS, 1).Value = dados
    tBloco = Timer - t0

    MsgBox "Célula a célula:           " & Format(tLento, "0.000") & " s" & vbCrLf & _
           "Célula a célula + turbo:   " & Format(tTurbo, "0.000") & " s" & vbCrLf & _
           "Escrita em bloco (array):  " & Format(tBloco, "0.000") & " s", _
           vbInformation, "Resultado"
    Exit Sub

Restaura:
    DesligarTurbo
    MsgBox "Erro: " & Err.Description, vbCritical
End Sub

' ---------- Padrão de macro robusta ----------
' Copie este esqueleto para suas macros longas.
Public Sub ModeloDeMacroRobusta()
    On Error GoTo Tratar
    LigarTurbo

    ' ... SEU CÓDIGO AQUI ...

Sair:
    DesligarTurbo              ' roda sempre: com ou sem erro
    Exit Sub

Tratar:
    MsgBox "Ocorreu um erro (" & Err.Number & "): " & Err.Description, _
           vbCritical, "Atenção"
    Resume Sair
End Sub

' ---------- Última linha / coluna ----------
Public Sub UltimaLinhaEColuna()
    Dim ws As Worksheet, ultLin As Long, ultCol As Long
    Set ws = ActiveSheet

    ' Última linha preenchida da coluna A (de baixo para cima)
    ultLin = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ' Última coluna preenchida da linha 1 (da direita para a esquerda)
    ultCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    MsgBox "Última linha (col. A): " & ultLin & vbCrLf & _
           "Última coluna (linha 1): " & ultCol, vbInformation
End Sub
