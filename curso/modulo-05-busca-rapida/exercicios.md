# Exercícios — Módulo 05

(Use os dados de `GerarDadosTeste`.)

1. **Busca exata:** altere `BuscarTodasAsOcorrencias` para aceitar só correspondência exata (`xlWhole`).
2. **Soma por termo:** some os valores (coluna C) das linhas cujo Nome contém um termo digitado — usando array.
3. **Duplicados:** com `Dictionary`, liste quantas vezes cada **Nome** aparece (chave = nome, valor = contagem) e escreva o resultado a partir de `E1`.
4. **Busca em duas colunas:** filtre Nome **e** Valor > 500 ao mesmo tempo com `AutoFilter` (dois campos).
5. **Desafio:** faça a busca ao digitar também funcionar com **pausa**: só filtra quando o usuário pressiona Enter (já é o comportamento do `Change`) e mostra "X linhas encontradas" na barra de status.

---

## Gabarito (2, 3 e 5)

**2)**
```vba
Public Sub SomaPorTermo()
    Dim ws As Worksheet, dados As Variant, termo As String
    Dim i As Long, total As Double, ultima As Long
    Set ws = ActiveSheet
    ultima = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    termo = InputBox("Nome contém:")
    If Len(termo) = 0 Then Exit Sub
    dados = ws.Range("B2:C" & ultima).Value
    For i = 1 To UBound(dados, 1)
        If InStr(1, dados(i, 1), termo, vbTextCompare) > 0 Then total = total + dados(i, 2)
    Next i
    MsgBox "Total: " & Format(total, "#,##0.00")
End Sub
```

**3)**
```vba
Public Sub ContarNomes()
    Dim ws As Worksheet, dados As Variant, dic As Object
    Dim i As Long, k As Variant, linha As Long, ultima As Long
    Set ws = ActiveSheet
    ultima = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    dados = ws.Range("B2:B" & ultima).Value
    Set dic = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(dados, 1)
        dic(dados(i, 1)) = dic(dados(i, 1)) + 1     ' item inexistente vale 0 ao ser lido e criado
    Next i
    linha = 1
    For Each k In dic.Keys
        ws.Cells(linha, 5).Value = k
        ws.Cells(linha, 6).Value = dic(k)
        linha = linha + 1
    Next k
End Sub
```

**5)** Dentro de `FiltrarPorTermo`, após filtrar:
```vba
Dim vis As Long
vis = ws.Range("A2", ws.Cells(ws.Rows.Count, 1).End(xlUp)).SpecialCells(xlCellTypeVisible).Count
Application.StatusBar = vis & " linha(s) encontrada(s)"
```
(Trate o caso sem resultados com `On Error Resume Next`, pois `SpecialCells` gera erro quando não há células visíveis.)
