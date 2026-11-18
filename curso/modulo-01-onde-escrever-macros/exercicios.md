# Exercícios — Módulo 01

1. **Mapa:** crie um arquivo `.xlsm` com 2 módulos padrão (`modGeral` e `modUtil`), uma planilha chamada "Controle" e cole os códigos de exemplo em `ThisWorkbook` e `Plan1`. Faça uma lista (em papel) de qual código foi para onde e **por quê**.
2. **Public x Private:** crie `Sub Calcular()` (Public) que chama `Private Sub Registrar()`. Confirme que só `Calcular` aparece em `Alt+F8`.
3. **Function:** escreva `Function Desconto(valor As Double, pct As Double) As Double`, use na célula `=Desconto(A1; 10)` (separador pode ser `,` conforme o Excel).
4. **ThisWorkbook x ActiveWorkbook:** escreva uma macro que escreva "oi" na célula A1 da planilha ativa **de outro arquivo aberto** (use `ActiveWorkbook`) e outra que escreva na pasta onde a macro mora (`ThisWorkbook`).
5. **Desafio:** faça `Workbook_Open` mostrar uma saudação (`SaudacaoDoDia`) com o nome do usuário do Windows (`Application.UserName`).

---

## Gabarito

**2)**
```vba
Public Sub Calcular()
    Registrar
End Sub

Private Sub Registrar()
    Debug.Print "Calculado em " & Now
End Sub
```

**3)**
```vba
Public Function Desconto(ByVal valor As Double, ByVal pct As Double) As Double
    Desconto = valor * (1 - pct / 100)
End Function
```

**4)**
```vba
Public Sub EscreverNoArquivoAtivo()
    ActiveWorkbook.ActiveSheet.Range("A1").Value = "oi"
End Sub

Public Sub EscreverNoProprioArquivo()
    ThisWorkbook.Worksheets(1).Range("A1").Value = "oi"
End Sub
```

**5)** (em `ThisWorkbook`)
```vba
Private Sub Workbook_Open()
    MsgBox mod01_OndeEscrever.SaudacaoDoDia(Application.UserName)
End Sub
```
