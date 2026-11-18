# Módulo 01 — Onde escrever macros (Módulo, Pasta de Trabalho e Planilha)

**Duração alvo:** ~40 min · **Aulas:** 7 · **Nível:** zero absoluto · **Arquivos:** `codigo/mod01_OndeEscrever.bas`, `ThisWorkbook_exemplo.txt`, `Planilha_exemplo.txt`

**Promessa do módulo:** em 40 minutos você para de se perder no editor do VBA e sabe, para qualquer macro, **em qual lugar** ela deve morar.

**Gancho (frase de abertura):** "Já aconteceu de você copiar uma macro da internet, colar no lugar errado e nada acontecer? Hoje isso acaba."

---

## Aula 1.1 — O mapa do editor VBA (6 min) · *aula de pré-visualização gratuita*

**Na tela:** Excel aberto → aba Desenvolvedor (como ativá-la) → `Alt+F11`.
**Roteiro:**
1. Mostre como ativar a guia Desenvolvedor (Arquivo > Opções > Personalizar Faixa de Opções).
2. Abra o editor. Aponte 3 janelas: **Projeto** (esquerda), **Propriedades** (abaixo), **Código** (centro). Ative **Imediata** com `Ctrl+G`.
3. Mostre as "gavetas" do projeto: *Microsoft Excel Objects* (ThisWorkbook e planilhas), *Módulos*, *Formulários*, *Módulos de Classe*.

**Fala-chave:** "Pense no projeto VBA como um prédio. Cada tipo de componente é um andar. A macro certa no andar errado simplesmente não funciona."
**Salvar como .xlsm:** avise que arquivos com macro precisam de `.xlsm` (ou `.xlsb`).

## Aula 1.2 — Sua primeira macro em um Módulo padrão (6 min)

**Na tela:** Inserir > Módulo → digitar `OlaDoModulo` (do arquivo) → `F5`.
**Roteiro:**
- Explique `Sub … End Sub`, `Option Explicit` (por que sempre usar) e `MsgBox`.
- Mostre 3 formas de executar: `F5` no editor, `Alt+F8`, botão/forma na planilha.
- Mostre como atribuir atalho (`Alt+F8` > Opções).

**Fala-chave:** "Módulo padrão é a *caixa de ferramentas*: o que está aqui pode ser chamado de qualquer lugar."

## Aula 1.3 — Public x Private e o escopo (6 min)

**Na tela:** `MostrarContador` e `IncrementarContador`.
**Roteiro:**
- Rode `MostrarContador` várias vezes: o contador sobe.
- Mostre que `IncrementarContador` (Private) **não aparece** no `Alt+F8`.
- Explique variável de módulo (`Private contadorDeUso`) × variável dentro da macro (some quando ela termina).

**Regra de bolso (mostrar em slide):** *Public* = quem o usuário chama; *Private* = peças internas.

## Aula 1.4 — ThisWorkbook x ActiveWorkbook (5 min)

**Na tela:** abrir 2 arquivos, rodar `DiferencaEntreThisEActive` em cada situação.
**Roteiro:** mostre que `ThisWorkbook` = onde o código mora; `ActiveWorkbook` = o que está na frente. Erro clássico: macro "pessoal" que mexe no arquivo errado.

**Fala-chave:** "Se a macro precisa mexer *nela mesma*, use ThisWorkbook. Se é para mexer no arquivo que o usuário está olhando, ActiveWorkbook."

## Aula 1.5 — Function: criando suas próprias fórmulas (5 min)

**Na tela:** `=DobroDe(A1)` e `=SaudacaoDoDia("Maria")` direto em células.
**Roteiro:** diferença Sub (faz) × Function (devolve valor); por que Function em módulo padrão funciona na célula (e dentro de planilha/ThisWorkbook não).

## Aula 1.6 — O código da PASTA DE TRABALHO (ThisWorkbook) (6 min)

**Na tela:** duplo clique em *ThisWorkbook* → colar `ThisWorkbook_exemplo.txt`.
**Roteiro:** salve, feche, reabra: `Workbook_Open` roda sozinho. Mostre as listas suspensas de Objeto/Evento no topo do editor para descobrir outros eventos.

**Fala-chave:** "O código aqui não espera você chamar. O *Excel* chama, quando algo acontece com o arquivo inteiro."

## Aula 1.7 — O código da PLANILHA (e o resumo final) (6 min)

**Na tela:** duplo clique em *Plan1* → colar `Planilha_exemplo.txt` → alterar A1.
**Roteiro:**
- Mostre que o código da planilha só vale para ela (`Me` = a própria planilha).
- Feche com a **tabela de decisão** (slide):

| Quero… | Coloque em… |
|---|---|
| Macro para botão, atalho ou reuso geral | Módulo padrão |
| Fórmula personalizada para células | Function em módulo padrão |
| Reagir a algo do arquivo inteiro (abrir, salvar, fechar) | ThisWorkbook |
| Reagir a algo de UMA planilha (alterar célula, selecionar) | Código da planilha |
| Janela/tela para o usuário | UserForm |
| Objetos reutilizáveis com comportamento próprio | Módulo de classe |

**Chamada final:** apresente o Módulo 04 (Eventos) e o 02 (Barra de carregamento) como próximos passos.
