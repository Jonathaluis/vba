Attribute VB_Name = "mod01_OndeEscrever"
Option Explicit

' ============================================================
' MÓDULO 01 - ONDE ESCREVER MACROS
' Este arquivo é um MÓDULO PADRÃO (Inserir > Módulo).
' Macros aqui são "gerais": servem para botões, atalhos,
' outras macros e funções usadas direto nas células.
' ============================================================

' Variável de módulo: vive enquanto a pasta estiver aberta
' e só é vista por macros DESTE módulo (Private).
Private contadorDeUso As Long

' Constante pública: qualquer módulo pode usar.
Public Const NOME_SISTEMA As String = "Meu Sistema VBA"

' ---------- Aula 1.2: a primeira macro ----------
Public Sub OlaDoModulo()
    MsgBox "Esta macro mora em um módulo padrão." & vbCrLf & _
           "Ela pode ser chamada por botão, atalho ou por outra macro.", _
           vbInformation, NOME_SISTEMA
End Sub

' ---------- Aula 1.3: Public x Private ----------
' Public Sub aparece em Alt+F8 e pode ser ligada a botões.
' Private Sub fica escondida: é "ferramenta interna".
Public Sub MostrarContador()
    IncrementarContador
    MsgBox "Você executou esta macro " & contadorDeUso & " vez(es) " & _
           "desde que abriu o arquivo.", vbInformation, NOME_SISTEMA
End Sub

Private Sub IncrementarContador()
    contadorDeUso = contadorDeUso + 1
End Sub

' ---------- Aula 1.4: ThisWorkbook x ActiveWorkbook ----------
' ThisWorkbook  = a pasta que CONTÉM este código.
' ActiveWorkbook = a pasta que está em foco na tela agora.
Public Sub DiferencaEntreThisEActive()
    MsgBox "ThisWorkbook   = " & ThisWorkbook.Name & vbCrLf & _
           "ActiveWorkbook = " & ActiveWorkbook.Name & vbCrLf & _
           "Planilha ativa = " & ActiveSheet.Name, _
           vbInformation, "Quem é quem?"
End Sub

' ---------- Aula 1.5: Function (fórmula personalizada) ----------
' Uma Function pode ser usada na célula: =DobroDe(A1)
Public Function DobroDe(ByVal numero As Double) As Double
    DobroDe = numero * 2
End Function

Public Function SaudacaoDoDia(ByVal nome As String) As String
    Dim h As Long
    h = Hour(Now)
    Select Case h
        Case Is < 12: SaudacaoDoDia = "Bom dia, " & nome
        Case Is < 18: SaudacaoDoDia = "Boa tarde, " & nome
        Case Else:    SaudacaoDoDia = "Boa noite, " & nome
    End Select
End Function

' ---------- Aula 1.6: onde cada coisa deve morar ----------
' Chame esta macro para imprimir um "mapa" na janela Imediata
' (Ctrl+G no editor).
Public Sub MapaDoProjeto()
    Dim comp As Object
    ' Requer: Arquivo > Opções > Central de Confiabilidade >
    ' Configurações de Macro > "Confiar no acesso ao modelo de
    ' objeto de projeto do VBA".
    On Error GoTo Falha
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Debug.Print comp.Name & "  (tipo " & comp.Type & ")"
    Next comp
    Exit Sub
Falha:
    MsgBox "Ative 'Confiar no acesso ao modelo de objeto de projeto do VBA' " & _
           "para listar os componentes.", vbExclamation
End Sub
