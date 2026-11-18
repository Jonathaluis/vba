# Exercícios — Módulo 04

1. **Data automática:** ao preencher qualquer célula da coluna B (linhas 2+), escreva automaticamente a data/hora na coluna C da mesma linha (somente se C estiver vazia).
2. **Validação:** na coluna D aceite só números de 0 a 100; se digitarem outra coisa, desfaça (`Application.Undo` ou volte ao valor anterior) e avise.
3. **Linha concluída:** duplo clique na coluna A marca ✔ **e** risca (`Font.Strikethrough`) o texto da coluna B da mesma linha.
4. **Contador de aberturas:** a cada abertura do arquivo, some 1 em uma célula escondida (planilha "Controle", célula A1).
5. **Integrando com o Módulo 03:** ao alterar `B2` (valor 0–100), atualize o semáforo.

---

## Gabarito

**1)**
```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim area As Range, c As Range
    Set area = Intersect(Target, Me.Range("B2:B1000"))
    If area Is Nothing Then Exit Sub
    On Error GoTo Fim
    Application.EnableEvents = False
    For Each c In area.Cells
        If Len(c.Value) > 0 And IsEmpty(c.Offset(0, 1).Value) Then
            c.Offset(0, 1).Value = Now
        End If
    Next c
Fim:
    Application.EnableEvents = True
End Sub
```

**2)**
```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim c As Range, area As Range
    Set area = Intersect(Target, Me.Range("D2:D1000"))
    If area Is Nothing Then Exit Sub
    On Error GoTo Fim
    Application.EnableEvents = False
    For Each c In area.Cells
        If Len(c.Value) > 0 Then
            If Not IsNumeric(c.Value) Then
                c.ClearContents
                MsgBox "Digite um número de 0 a 100.", vbExclamation
            ElseIf c.Value < 0 Or c.Value > 100 Then
                c.ClearContents
                MsgBox "Valor fora do intervalo 0–100.", vbExclamation
            End If
        End If
    Next c
Fim:
    Application.EnableEvents = True
End Sub
```

**4)** (em `ThisWorkbook`)
```vba
Private Sub Workbook_Open()
    With Me.Worksheets("Controle").Range("A1")
        .Value = .Value + 1
    End With
End Sub
```

**5)** (no código da planilha; requer `mod03_Formas` importado e semáforo criado)
```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Me.Range("B2")) Is Nothing Then Exit Sub
    If IsNumeric(Me.Range("B2").Value) Then AtualizarSemaforo CDbl(Me.Range("B2").Value)
End Sub
```
