---
name: design-flexo
description: Regras e padroes de design do Console Flexo v2.0 / ProjetoFlexo. Use ao criar ou modificar qualquer form, frame, botao ou componente visual VBA neste projeto.
---

Ao trabalhar com qualquer form ou componente visual deste projeto, siga **obrigatoriamente** todas as regras abaixo. Elas derivam do `DesignSpec_ConsoleFlexo.pdf` e dos padroes aplicados em `src/ModuloFlexo/frmFlexo.frm`.

---

## 1. Paleta de Cores

Use sempre a funcao auxiliar `H(R, G, B)` para converter cores. Nunca use valores hexadecimais ou Long diretamente no codigo.

```vba
Private Function H(R As Long, G As Long, B As Long) As Long
    H = RGB(R, G, B)
End Function
```

| Token                | Hex       | RGB              | Uso                                              |
|----------------------|-----------|------------------|--------------------------------------------------|
| Fundo Formulario     | `#1A2030` | 26, 32, 48       | BackColor do form principal                      |
| Fundo Frame          | `#232D3F` | 35, 45, 63       | BackColor dos frames de secao                    |
| Botao Normal         | `#1E2A3A` | 30, 42, 58       | Estado padrao de todos os botoes (Label)         |
| Botao Hover          | `#243244` | 36, 50, 68       | MouseMove sobre botao                            |
| Botao Pressionado    | `#151C2B` | 21, 28, 43       | MouseDown sobre botao                            |
| Botao Concluido      | `#181F2C` | 24, 31, 44       | Apos acao executada com sucesso                  |
| Fundo Profundo       | `#111822` | 17, 24, 34       | Inputs (TextBox), sombras                        |
| Texto Normal         | `#9AB0C8` | 154, 176, 200    | Texto padrao de botoes e labels                  |
| Texto Hover          | `#C0D4E8` | 192, 212, 232    | Texto ao passar o mouse                          |
| Texto Concluido      | `#3A4E62` | 58, 78, 98       | Texto apos execucao (tom apagado)                |
| Texto Secundario     | `#6A7D96` | 106, 125, 150    | Labels de secao, tooltips, rodapes               |
| Texto Titulo         | `#E8EEF6` | 232, 238, 246    | Titulos e textos de destaque maximo              |
| Azul Acao Fundo      | `#1A3A5E` | 26, 58, 94       | Botao Desfazer — estado normal                   |
| Azul Acao Texto      | `#6AACE8` | 106, 172, 232    | Texto do Desfazer e destaques                    |
| Azul Acao Hover      | `#1E4672` | 30, 70, 114      | Hover do Desfazer                                |
| Azul Acao Press      | `#14304E` | 20, 48, 78       | Press do Desfazer                                |
| Borda Sutil          | `#4A5870` | 74, 88, 112      | Grade, separadores, BorderColor dos frames       |
| Vermelho Critico     | `#E05555` | 224, 85, 85      | Labels de itens criticos no PreFlight            |

---

## 2. Tipografia

- **Fonte unica:** `Segoe UI` em todos os elementos
- **Tamanho:** `8pt` em todos os casos
- **Nunca use outra fonte ou tamanho**

| Elemento             | Peso    | Alinhamento | Cor padrao              |
|----------------------|---------|-------------|-------------------------|
| Botoes de acao       | Bold    | Center      | Texto Normal `#9AB0C8`  |
| Titulo do Console    | Bold    | Center      | Texto Titulo `#E8EEF6`  |
| Botao Desfazer       | Bold    | Center      | Azul Acao Texto `#6AACE8` |
| Labels de secao      | Normal  | Left        | Texto Secundario `#6A7D96` |
| Labels de relatorio  | Normal  | Left        | Texto Normal `#9AB0C8`  |

---

## 3. Estados dos Botoes (Labels simulando botoes)

Todos os botoes sao `MSForms.Label` — nunca `CommandButton`. Cada botao implementa 4 estados via eventos:

| Estado       | Evento      | Fundo (`BackColor`) | Texto (`ForeColor`)  |
|--------------|-------------|---------------------|----------------------|
| Normal       | Padrao      | `H(30, 42, 58)`     | `H(154, 176, 200)`   |
| Hover        | MouseMove   | `H(36, 50, 68)`     | `H(192, 212, 232)`   |
| Pressionado  | MouseDown   | `H(21, 28, 43)`     | `H(192, 212, 232)`   |
| Concluido    | MouseUp     | `H(24, 31, 44)`     | `H(58, 78, 98)`      |

**Regras obrigatorias:**
- Hover e Press nao se aplicam se o botao ja estiver no estado Concluido (`ForeColor = H(58, 78, 98)`)
- Estado Concluido persiste ate o usuario clicar em Desfazer ou Reset
- Apos MouseUp bem-sucedido: chamar `MarcarConcluido()` com caption + checkmark `ChrW(10003)`
- Caption dos botoes usa `vbCrLf & texto` para centralizar verticalmente
- Botoes de selecao (indices 6, 7, 11) passam `apenasSelecao = True` — nao habilitam Desfazer

**Padrao de implementacao para cada botao:**
```vba
Private Sub btnX_MouseMove(...)
    AplicarHover Me.btnX
End Sub
Private Sub btnX_MouseDown(...)
    AplicarPress Me.btnX
End Sub
Private Sub btnX_MouseUp(...)
    ' chamar modulo de logica
    MarcarConcluido Me.btnX, "Caption Original", "Nome Acao", False
End Sub
```

---

## 4. Frames de Secao

Todos os frames seguem o mesmo padrao visual:

```vba
With Me.nomeFrame
    .BackColor  = H(26, 32, 48)      ' Fundo Formulario (nao Fundo Frame)
    .ForeColor  = H(106, 125, 150)   ' Texto Secundario
    .BorderColor = H(35, 45, 63)     ' Fundo Frame como borda
    .Font.Name  = "Segoe UI"
    .Font.Size  = 8
    .Font.Bold  = True
    .Caption    = " " & icone & "  " & TITULO_EM_MAIUSCULAS
End With
```

**Icones Unicode padrao para frames:**
- `ChrW(9679)` ● — Tratamento de Cores
- `ChrW(9998)` ✎ — Tratamento de Vetores
- `ChrW(9638)` ■ — Tratamento de Bitmaps
- `ChrW(9868)`    — Montagem
- `ChrW(9881)` ⚙ — Opcoes / Configuracoes
- `ChrW(9670)` ◆ — Dimensoes
- `ChrW(9632)` ■ — Espacamento
- `ChrW(8600)` ↘ — Reducao
- `ChrW(9733)` ★ — Resultados

**Frames com colapso:** se o form implementar frames colapsaveis, incluir setas no caption:
- `ChrW(9660)` ▼ — expandido
- `ChrW(9658)` ▶ — colapsado

---

## 5. Inputs (TextBox)

```vba
With Me.txtNome
    .BackColor     = H(17, 24, 34)      ' Fundo Profundo
    .ForeColor     = H(154, 176, 200)   ' Texto Normal
    .Font.Name     = "Segoe UI"
    .Font.Size     = 8
    .BorderStyle   = fmBorderStyleNone
    .SpecialEffect = fmSpecialEffectFlat
End With
```

Input desabilitado:
```vba
.BackColor = H(24, 31, 44)   ' Botao Concluido
.ForeColor = H(58, 78, 98)   ' Texto Concluido
.Enabled   = False
```

---

## 6. Layout e Estrutura

- **Form modeless:** `ShowModal = False` (Console permanece aberto enquanto o usuario edita no CorelDRAW)
- **Posicao inicial:** `Me.Left = 10 : Me.Top = 60` (canto superior esquerdo da tela)
- **Largura padrao frmFlexo:** ~230pt (4608 twips)
- **Estrutura de frames em ordem vertical:** Cores → Vetores → Bitmaps → Montagem → Rodape (Desfazer + Reset)
- **Separacao entre frames:** 3 twips
- Botao Reset posicionado dinamicamente: `btnReset.Left = btnDesfazer.Left + btnDesfazer.Width + 4`

---

## 7. Tooltips

Todo botao de acao **deve ter** `ControlTipText` configurado. Nunca deixar em branco.

---

## 8. Nomenclatura

| Elemento          | Prefixo   | Exemplo                     |
|-------------------|-----------|-----------------------------|
| Botao (Label)     | `btn`     | `btnBranco`, `btnMontar`    |
| Frame             | `frame`   | `frameCores`, `frameOpcoes` |
| TextBox           | `txt`     | `txtZ`, `txtLarguraFaca`    |
| Label resultado   | `lbl`     | `lblDesenvolvimento`        |
| Label radio sim.  | `lbl`     | `lbl114`, `lblPi314`        |
| CheckBox          | `chk`     | `chkCameron`, `chkRelatorio`|
| Form              | `frm`     | `frmFlexo`, `frmStepRepeat` |
| Modulo            | `ModNN_`  | `Mod02_Cores`, `Mod04_Montagem` |

---

## 9. Organizacao dos Modulos

```
src/
  ModuloFlexo/
    Mod00_Versao.bas
    Mod01_Setup.bas
    Mod02_Cores.bas
    Mod03_Vetores.bas
    Mod04_Montagem.bas
    Mod05_Imagens.bas
    Mod06_PadronizarLayers.bas
    Mod07_InserirMicropontos.bas
    frmFlexo.frm / .frx
  ModuloStepRepeat/
    Mod01_Config.bas
    Mod02_Montagem.bas
    Mod03_Cameron.bas
    Mod04_Relatorio.bas
    frmStepRepeat.frm / .frx
  ModuloPreflight/
    Mod01_AbrirScanner.bas
    Mod02_Scanner_Engine.bas
    frmPreFlight.frm / .frx
```

---

## 10. Ícones dos Botões de Ação

Todo botão de ação deve ter um ícone Unicode prefixando o caption. A fonte de verdade é `ObterCaptionOriginal()` em `frmFlexo.frm`. Padrão de caption: `ChrW(XXXX) & "  " & textoLegivel` — `AplicarEstiloLabelPadrao` adiciona o `vbCrLf` inicial automaticamente.

| Botão | Ícone | ChrW | Grupo |
|-------|-------|------|-------|
| btnBranco | ◎ | 9678 | Cores |
| btnPretoSujo | ◼ | 9724 | Cores |
| btnSpot | ◈ | 9672 | Cores |
| btnRGB | ⬡ | 11041 | Cores |
| btnCorRegistro | ✛ | 10011 | Cores |
| btnConverterPantone | ◉ | 9673 | Cores |
| btnSelPreenchimento | ▣ | 9635 | Cores |
| btnSelContorno | ▢ | 9634 | Cores |
| btnCorrigirBordaDura | ▤ | 9636 | Cores |
| btnLimparSujeira | ✖ | 10006 | Cores |
| btnTextosEmCurvas | ❧ | 10023 | Vetores |
| btnEspessuraMinima | ━ | 9473 | Vetores |
| btnCorrigirContornos | ⊟ | 8863 | Vetores |
| btnDesbloquear | ◓ | 9683 | Vetores |
| btnMicropontos | ⊕ | 8853 | Vetores |
| btnPadronizarImagens | ▨ | 9640 | Bitmaps |
| btnInserirTextos | ❐ | 10000 | Montagem |
| btnTrimBox | ⊞ | 8862 | Montagem |

---

## 11. Regras Criticas — Nunca Violar

1. **Nunca use `CommandButton`** — sempre `MSForms.Label` estilizado
2. **Nunca hardcode cores** como Long ou hex — sempre `H(R, G, B)`
3. **Nunca use outra fonte** alem de `Segoe UI 8pt`
4. **Nunca habilite Desfazer** para acoes de selecao (`apenasSelecao = True`)
5. **Nunca altere o `.frx`** diretamente — posicionamentos fixos ficam no designer; logica e cores ficam no `.frm`
6. **Sempre configure `ControlTipText`** em todos os botoes
7. **Sempre use `On Error Resume Next`** no inicio do `UserForm_Initialize`
8. **Sempre aplique tema via funcoes `AplicarTema*()`** chamadas no Initialize — nunca inline
