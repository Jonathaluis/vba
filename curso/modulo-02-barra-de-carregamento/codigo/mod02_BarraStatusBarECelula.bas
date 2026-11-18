Attribute VB_Name = "mod02_BarraStatusBarECelula"
Option Explicit

' ============================================================
' MÓDULO 02 - BARRA DE CARREGAMENTO
' Técnicas 1 e 2: barra na Barra de Status e barra em célula.
' São as mais simples: sem UserForm, sem formas.
' ============================================================

' ---------- Técnica 1: Barra de Status ----------
Public Sub BarraNaStatusBar()
    Const TOTAL As Long = 50
    Dim i As Long

    On Error GoTo Limpa
    For i = 1 To TOTAL
        Application.StatusBar = "Processando  " & MontarBarra(i / TOTAL, 25) & _
                                "  " & Format(i / TOTAL, "0%")
        Pausa 0.05          ' <- troque pelo SEU trabalho real
    Next i

Limpa:
    Application.StatusBar = False   ' SEMPRE devolver o controle ao Excel
    If Err.Number = 0 Then MsgBox "Concluído!", vbInformation
End Sub

' ---------- Técnica 2: Barra dentro de uma célula ----------
' Escreve a barra em C2 da planilha ativa.
Public Sub BarraNaCelula()
    Const TOTAL As Long = 40
    Dim i As Long, celula As Range
    Set celula = ActiveSheet.Range("C2")
    celula.Font.Name = "Consolas"

    For i = 1 To TOTAL
        celula.Value = MontarBarra(i / TOTAL, 30) & " " & Format(i / TOTAL, "0%")
        Pausa 0.05
    Next i
    celula.ClearContents
End Sub

' ---------- Auxiliares ----------
' pct entre 0 e 1; largura = quantidade de "blocos"
Private Function MontarBarra(ByVal pct As Double, ByVal largura As Long) As String
    Dim cheios As Long
    If pct < 0 Then pct = 0
    If pct > 1 Then pct = 1
    cheios = CLng(pct * largura)
    MontarBarra = String(cheios, ChrW(9608)) & String(largura - cheios, ChrW(9617))
End Function

' Espera "segundos" sem travar a tela (Timer reinicia à meia-noite: tratado)
Private Sub Pausa(ByVal segundos As Double)
    Dim t0 As Double
    t0 = Timer
    Do While Timer >= t0 And Timer < t0 + segundos
        DoEvents
    Loop
End Sub
