# Exercícios — Módulo 03

1. **Bandeira:** com formas, desenhe uma bandeira simples (retângulo verde, losango `msoShapeDiamond` amarelo, círculo azul).
2. **Grade de bolinhas:** crie uma grade 5 × 5 de círculos (laços aninhados) e pinte aleatoriamente (`Rnd`) de verde ou vermelho.
3. **Pulso:** faça um círculo "pulsar" (aumentar e diminuir de tamanho mantendo o centro fixo).
4. **Status por linha:** para cada linha da tabela A2:B20 (coluna B com valores 0–100) coloque uma bolinha de status (verde/amarela/vermelha) na coluna C.
5. **Desafio:** botão "Sortear" (forma + `OnAction`) que escolhe aleatoriamente um dos 3 círculos do semáforo.

---

## Gabarito (2 e 3)

**2) Grade 5 × 5**
```vba
Public Sub GradeDeBolinhas()
    Dim ws As Worksheet, l As Long, c As Long, s As Shape
    Set ws = ActiveSheet
    Randomize
    For l = 0 To 4
        For c = 0 To 4
            Set s = ws.Shapes.AddShape(msoShapeOval, 40 + c * 45, 40 + l * 45, 36, 36)
            s.Name = "demo_grade_" & l & "_" & c
            s.Line.Visible = msoFalse
            If Rnd > 0.5 Then
                s.Fill.ForeColor.RGB = RGB(46, 204, 113)
            Else
                s.Fill.ForeColor.RGB = RGB(231, 76, 60)
            End If
        Next c
    Next l
End Sub
```

**3) Pulso (mantendo o centro)**
```vba
Public Sub Pulsar()
    Dim s As Shape, i As Long, cx As Double, cy As Double, d As Double
    Set s = ActiveSheet.Shapes("demo_circulo")
    cx = s.Left + s.Width / 2
    cy = s.Top + s.Height / 2
    For i = 0 To 360 Step 10
        d = 100 + 30 * Sin(i * 3.14159265 / 180)
        s.Width = d: s.Height = d
        s.Left = cx - d / 2
        s.Top = cy - d / 2
        DoEvents
    Next i
End Sub
```
