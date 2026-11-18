# Exercícios — Módulo 02

1. **Personalize:** altere `MontarBarra` para usar `#` e `-` em vez de blocos. Mude a largura para 40.
2. **Barra colorida:** na barra com formas, faça a cor mudar conforme o progresso (vermelho < 40%, amarelo < 80%, verde acima).
3. **Barra real:** escreva uma macro que percorre `A2:A5000`, limpa espaços extras (`Trim`) e mostra o progresso com o UserForm.
4. **Barra de tempo restante:** mostre no `lblMensagem` o tempo estimado restante (dica: `Timer`, regra de três com `pct`).
5. **Desafio:** adicione um botão "Cancelar" no UserForm que interrompe o laço (dica: variável `Public Cancelado As Boolean`, verificada em cada passo).

---

## Gabarito (2 e 4)

**2) Cor por faixa — troque o conteúdo de `AtualizarBarra`, depois de definir `larg`:**
```vba
    If pct < 0.4 Then
        ws.Shapes(NOME_BARRA).Fill.ForeColor.RGB = RGB(231, 76, 60)
    ElseIf pct < 0.8 Then
        ws.Shapes(NOME_BARRA).Fill.ForeColor.RGB = RGB(241, 196, 15)
    Else
        ws.Shapes(NOME_BARRA).Fill.ForeColor.RGB = RGB(39, 174, 96)
    End If
```

**4) Tempo restante (dentro do laço):**
```vba
Dim inicio As Double, restante As Double
inicio = Timer
' ... dentro do laço, com pct = i / TOTAL e pct > 0:
restante = (Timer - inicio) * (1 - pct) / pct
frmProgresso.Atualizar pct, "Restam ~" & Format(restante, "0") & " s"
```

**5) Cancelar:** no formulário:
```vba
Public Cancelado As Boolean
Private Sub btnCancelar_Click()
    Cancelado = True
End Sub
```
Na macro: `If frmProgresso.Cancelado Then Exit For` (o `DoEvents` dentro de `Atualizar` permite que o clique seja processado).
