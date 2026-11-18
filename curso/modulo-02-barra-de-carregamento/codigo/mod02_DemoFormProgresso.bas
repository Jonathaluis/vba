Attribute VB_Name = "mod02_DemoFormProgresso"
Option Explicit

' ============================================================
' Usa o UserForm "frmProgresso" (veja frmProgresso_codigo.txt).
' Este módulo só funciona DEPOIS de montar o formulário.
' ============================================================

' Demonstração com pausa artificial
Public Sub DemoFormProgresso()
    Const TOTAL As Long = 100
    Dim i As Long

    frmProgresso.Show vbModeless      ' vbModeless = a macro continua rodando
    On Error GoTo Limpa

    For i = 1 To TOTAL
        frmProgresso.Atualizar i / TOTAL, "Processando item " & i & " de " & TOTAL
        Pausa 0.02
    Next i

Limpa:
    Unload frmProgresso
    If Err.Number = 0 Then MsgBox "Concluído!", vbInformation
End Sub

' Exemplo REAL: percorre as linhas da coluna A da planilha ativa
' e converte tudo para maiúsculas, mostrando o progresso.
Public Sub ExemploReal_MaiusculasComProgresso()
    Dim ws As Worksheet, ultima As Long, i As Long

    Set ws = ActiveSheet
    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultima < 2 Then MsgBox "Preencha dados na coluna A (a partir da linha 2).": Exit Sub

    frmProgresso.Show vbModeless
    On Error GoTo Limpa

    For i = 2 To ultima
        ws.Cells(i, 1).Value = UCase$(CStr(ws.Cells(i, 1).Value))

        ' Dica de performance: atualizar a barra a cada N linhas,
        ' e não em TODAS, para não deixar o processo lento.
        If i Mod 10 = 0 Or i = ultima Then
            frmProgresso.Atualizar (i - 1) / (ultima - 1), "Linha " & i & " de " & ultima
        End If
    Next i

Limpa:
    Unload frmProgresso
    If Err.Number = 0 Then MsgBox "Pronto!", vbInformation
End Sub

Private Sub Pausa(ByVal segundos As Double)
    Dim t0 As Double
    t0 = Timer
    Do While Timer >= t0 And Timer < t0 + segundos
        DoEvents
    Loop
End Sub
