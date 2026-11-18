Attribute VB_Name = "mod05_Busca"
Option Explicit

' ============================================================
' MÓDULO 05 - BUSCA RÁPIDA
' 1) Find / FindNext     -> achar TODAS as ocorrências
' 2) Array em memória    -> varrer muitos dados rapidamente
' 3) Dictionary          -> "PROCV" instantâneo para consultas repetidas
' 4) AutoFilter          -> filtrar a tabela conforme o que digitam
' + gerador de dados de teste para comparar velocidades
' Observação: Scripting.Dictionary existe no Excel para Windows.
' ============================================================

' ---------- Dados de teste (50.000 linhas) ----------
' A: Código | B: Nome | C: Valor
Public Sub GerarDadosTeste()
    Const QTD As Long = 50000
    Dim ws As Worksheet, dados() As Variant, i As Long
    Dim nomes As Variant

    nomes = Array("Ana", "Bruno", "Carla", "Diego", "Elisa", "Fabio", "Gabi", "Hugo")
    Set ws = ActiveSheet
    ws.Cells.Clear

    ReDim dados(1 To QTD, 1 To 3)
    For i = 1 To QTD
        dados(i, 1) = "COD" & Format(i, "00000")
        dados(i, 2) = nomes(i Mod 8) & " " & nomes((i * 3) Mod 8)
        dados(i, 3) = Round(Rnd * 1000, 2)
    Next i

    ws.Range("A1:C1").Value = Array("Código", "Nome", "Valor")
    ws.Range("A1:C1").Font.Bold = True
    ws.Range("A2").Resize(QTD, 3).Value = dados      ' escrita em bloco = rápido
    ws.Columns("A:C").AutoFit
    MsgBox QTD & " linhas geradas.", vbInformation
End Sub

' ---------- 1) Find / FindNext ----------
Public Sub BuscarTodasAsOcorrencias()
    Dim ws As Worksheet, termo As String
    Dim achado As Range, primeiro As String, todos As Range

    Set ws = ActiveSheet
    termo = InputBox("O que procurar?", "Busca rápida")
    If Len(termo) = 0 Then Exit Sub

    Set achado = ws.UsedRange.Find(What:=termo, LookIn:=xlValues, _
                                   LookAt:=xlPart, MatchCase:=False)
    If achado Is Nothing Then
        MsgBox "Nada encontrado para '" & termo & "'.", vbInformation
        Exit Sub
    End If

    primeiro = achado.Address
    Do
        If todos Is Nothing Then
            Set todos = achado
        Else
            Set todos = Union(todos, achado)
        End If
        Set achado = ws.UsedRange.FindNext(achado)
        If achado Is Nothing Then Exit Do         ' VBA não faz "curto-circuito" no And
    Loop While achado.Address <> primeiro

    todos.Select
    MsgBox todos.Cells.Count & " ocorrência(s) selecionada(s).", vbInformation
End Sub

' ---------- 2) Array em memória ----------
' Procura na coluna B (Nome) e devolve quantas linhas contêm o termo.
Public Sub BuscarComArray()
    Dim ws As Worksheet, ultima As Long, dados As Variant
    Dim termo As String, i As Long, achou As Long, t0 As Double

    Set ws = ActiveSheet
    ultima = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If ultima < 3 Then MsgBox "Gere os dados de teste primeiro.": Exit Sub

    termo = InputBox("Nome contém:", "Busca com array", "Ana")
    If Len(termo) = 0 Then Exit Sub

    t0 = Timer
    dados = ws.Range("B2:B" & ultima).Value          ' 1 leitura só
    For i = 1 To UBound(dados, 1)
        If InStr(1, dados(i, 1), termo, vbTextCompare) > 0 Then achou = achou + 1
    Next i

    MsgBox achou & " linha(s) contêm '" & termo & "'." & vbCrLf & _
           "Tempo: " & Format(Timer - t0, "0.000") & " s", vbInformation
End Sub

' ---------- 3) Dictionary: consulta instantânea ----------
' Monta o índice UMA vez; depois cada consulta é imediata.
Public Sub BuscarComDicionario()
    Dim ws As Worksheet, ultima As Long, dados As Variant
    Dim dic As Object, i As Long, chave As String, t0 As Double

    Set ws = ActiveSheet
    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultima < 3 Then MsgBox "Gere os dados de teste primeiro.": Exit Sub

    t0 = Timer
    dados = ws.Range("A2:C" & ultima).Value
    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbTextCompare                   ' ignora maiúsc./minúsc.

    For i = 1 To UBound(dados, 1)
        If Not dic.Exists(CStr(dados(i, 1))) Then dic.Add CStr(dados(i, 1)), i
    Next i
    Dim tIndice As Double: tIndice = Timer - t0

    Do
        chave = InputBox("Código a buscar (ex.: COD00042)." & vbCrLf & _
                         "Vazio = sair.", "Dicionário")
        If Len(chave) = 0 Then Exit Do

        If dic.Exists(chave) Then
            i = dic(chave)
            MsgBox "Nome:  " & dados(i, 2) & vbCrLf & _
                   "Valor: " & dados(i, 3) & vbCrLf & vbCrLf & _
                   "(índice montado em " & Format(tIndice, "0.000") & " s)", vbInformation
        Else
            MsgBox "Código não encontrado.", vbExclamation
        End If
    Loop
End Sub

' ---------- 4) AutoFilter ----------
' Filtra a coluna B (campo 2) pelo termo; termo vazio remove o filtro.
Public Sub FiltrarPorTermo(ByVal termo As String)
    Dim ws As Worksheet
    Set ws = ActiveSheet

    If Len(termo) = 0 Then
        ws.AutoFilterMode = False
        Exit Sub
    End If
    ws.Range("A1").CurrentRegion.AutoFilter Field:=2, Criteria1:="*" & termo & "*"
End Sub

Public Sub DemoFiltro()
    FiltrarPorTermo InputBox("Filtrar nomes que contêm:", "Filtro", "Carla")
End Sub

' ---------- Duelo de velocidade: célula a célula x array ----------
Public Sub DueloDeVelocidade()
    Dim ws As Worksheet, ultima As Long, i As Long
    Dim t0 As Double, tLento As Double, tRapido As Double
    Dim dados As Variant, cont As Long

    Set ws = ActiveSheet
    ultima = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If ultima < 3 Then MsgBox "Gere os dados de teste primeiro.": Exit Sub

    t0 = Timer
    For i = 2 To ultima                                ' lento: toca a planilha 1x por linha
        If InStr(1, ws.Cells(i, 2).Value, "Ana", vbTextCompare) > 0 Then cont = cont + 1
    Next i
    tLento = Timer - t0

    cont = 0: t0 = Timer
    dados = ws.Range("B2:B" & ultima).Value            ' rápido: 1 leitura
    For i = 1 To UBound(dados, 1)
        If InStr(1, dados(i, 1), "Ana", vbTextCompare) > 0 Then cont = cont + 1
    Next i
    tRapido = Timer - t0

    MsgBox "Célula a célula: " & Format(tLento, "0.000") & " s" & vbCrLf & _
           "Com array:       " & Format(tRapido, "0.000") & " s" & vbCrLf & _
           "Ocorrências:     " & cont, vbInformation, "Duelo de velocidade"
End Sub
