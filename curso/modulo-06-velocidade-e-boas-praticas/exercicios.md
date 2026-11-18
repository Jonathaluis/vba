# Exercícios — Módulo 06

1. **Meça:** rode `TesteEscrita` com 5.000, 20.000 e 100.000 linhas e anote os tempos. O que cresce mais rápido?
2. **Quebre e conserte:** tire o tratamento de erro de uma macro com `LigarTurbo`, force um erro e observe o estado do Excel. Depois restaure com `DesligarTurbo` pela janela Imediata.
3. **Sem Select:** reescreva sem `Select`/`Activate`:
   ```vba
   Sheets("Dados").Select
   Range("A1").Select
   Selection.Value = "Total"
   Selection.Font.Bold = True
   ```
4. **Última linha:** escreva uma função `UltimaLinha(ws As Worksheet, col As Long) As Long` e use-a em outra macro.
5. **Desafio:** pegue uma macro sua com laço célula a célula (copiar valores de uma coluna para outra com regra) e converta para array. Meça o ganho.

---

## Gabarito

**3)**
```vba
With ThisWorkbook.Worksheets("Dados").Range("A1")
    .Value = "Total"
    .Font.Bold = True
End With
```

**4)**
```vba
Public Function UltimaLinha(ByVal ws As Worksheet, ByVal col As Long) As Long
    UltimaLinha = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function
```
