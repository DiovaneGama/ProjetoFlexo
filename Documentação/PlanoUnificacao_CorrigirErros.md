# Plano: Unificar "Corrigir Erros" com os Botões Individuais do frmFlexo

**Plataforma:** CorelDRAW 2026 | **Linguagem:** VBA | **Data:** Abril 2026

---

## Contexto

O botão "Corrigir Erros" do `frmPreFlight` chama `ExecutarCorrecoes` → `CrawlerCorrecoes`, uma rotina **privada e paralela** que reimplementa a lógica de correção de forma independente dos botões individuais do `frmFlexo`. Isso causou divergências críticas identificadas na análise:

| Item | Divergência |
|------|-------------|
| Preto Composto | Limiar K > 80 (PreFlight) vs K > 85 (frmFlexo) |
| Imagens | Eleva para 300 DPI (PreFlight) vs 600 DPI (frmFlexo) |
| Linhas Finas | PreFlight exclui intencionais; frmFlexo corrige todas |

O objetivo é fazer o "Corrigir Erros" chamar **exatamente as mesmas funções públicas** que cada botão individual usa, eliminando o `CrawlerCorrecoes` por completo.

---

## Abordagem: parâmetro `Optional silencioso As Boolean = False`

Cada função pública dos módulos de correção exibe diálogos de confirmação e mensagens de sucesso — necessários para uso manual, mas indesejados no fluxo automático do PreFlight. A solução é adicionar um parâmetro opcional `silencioso` a cada função:

- **`silencioso = False` (padrão):** comportamento atual dos botões individuais — diálogos, confirmações e mensagens de resultado mantidos
- **`silencioso = True`:** sem MsgBox de confirmação, sem MsgBox de resultado, `ConverterRGB` usa CMYK completo sem perguntar

Os botões individuais do `frmFlexo` **não mudam** — não passam o parâmetro, logo `False` é assumido.

---

## Arquivos Afetados

| Arquivo | O que muda |
|---------|------------|
| `src/ModuloFlexo/Mod02_Cores.bas` | `silencioso` em `CorrigirBrancoOverprint`, `DetectarPretoSujo`, `ConverterRGB`, `CorrigirBordaDuraGradientes` |
| `src/ModuloFlexo/Mod03_Vetores.bas` | `silencioso` em `PadronizarContornosFinos`, `ConverterTextosEmCurvas` |
| `src/ModuloFlexo/Mod05_Imagens.bas` | `silencioso` em `PadronizarImagensCMYK600` |
| `src/ModuloPreflight/Mod02_Scanner_Engine.bas` | `ExecutarCorrecoes` passa a chamar os módulos acima com `silencioso:=True`. `CrawlerCorrecoes` é removido. |

---

## Implementação Detalhada

### Passo 1 — `Mod02_Cores.bas`

**`CorrigirBrancoOverprint(Optional silencioso As Boolean = False)`**
- `silencioso = True`: sem MsgBox de confirmação, sem MsgBox de resultado

**`DetectarPretoSujo(Optional silencioso As Boolean = False)`**
- `silencioso = True`: sem MsgBox de confirmação, sem MsgBox de resultado

**`ConverterRGB(Optional silencioso As Boolean = False)`**
- `silencioso = True`: pula o diálogo CMY/CMYK e usa **CMYK completo**; sem MsgBox de resultado

**`CorrigirBordaDuraGradientes(minDot As Integer, Optional silencioso As Boolean = False)`**
- Adicionar `minDot` como primeiro parâmetro (hoje obtido via InputBox)
- `silencioso = True`: pula o InputBox, usa o `minDot` recebido; sem MsgBox de resultado
- `silencioso = False` (botão manual): mantém InputBox e mensagens existentes

### Passo 2 — `Mod03_Vetores.bas`

**`PadronizarContornosFinos(Optional silencioso As Boolean = False)`**
- `silencioso = True`: sem MsgBox "nenhum contorno encontrado", sem MsgBox de resultado

**`ConverterTextosEmCurvas(Optional silencioso As Boolean = False)`**
- `silencioso = True`: sem MsgBox de confirmação, sem MsgBox de resultado

### Passo 3 — `Mod05_Imagens.bas`

**`PadronizarImagensCMYK600(Optional silencioso As Boolean = False)`**
- `silencioso = True`: sem MsgBox de confirmação inicial, sem MsgBox de resultado

### Passo 4 — `Mod02_Scanner_Engine.bas`

**`ExecutarCorrecoes`** passa a orquestrar chamadas diretas aos módulos:

```vba
Public Sub ExecutarCorrecoes(ByVal minDot As Integer)
    ActiveDocument.BeginCommandGroup "Correcao Automatica PreFlight"
    On Error GoTo FimErro

    Call Mod02_Cores.CorrigirBrancoOverprint(silencioso:=True)
    Call Mod02_Cores.DetectarPretoSujo(silencioso:=True)
    Call Mod02_Cores.ConverterRGB(silencioso:=True)
    Call Mod03_Vetores.ConverterTextosEmCurvas(silencioso:=True)
    Call Mod03_Vetores.PadronizarContornosFinos(silencioso:=True)
    Call Mod05_Imagens.PadronizarImagensCMYK600(silencioso:=True)
    If minDot > 0 Then Call Mod02_Cores.CorrigirBordaDuraGradientes(minDot, silencioso:=True)

FimErro:
    ActiveDocument.EndCommandGroup
End Sub
```

**`CrawlerCorrecoes` é removido completamente** — zero lógica duplicada, incluindo a borda dura.

> **Nota sobre Borda Dura:** `CorrigirBordaDuraGradientes` em `Mod02_Cores.bas` tem botão equivalente no `frmFlexo` (`btnCorrigirBordaDura`). O algoritmo de 3 regras (CMYK, Pantone, Spot) é idêntico nos dois lados — a única diferença era o InputBox interativo, que no modo `silencioso:=True` é suprimido, usando o `minDot` passado como parâmetro.

---

## Ordem de Execução no "Corrigir Erros"

| # | Correção | Função | Módulo |
|---|----------|--------|--------|
| 1 | Branco Overprint | `CorrigirBrancoOverprint(silencioso:=True)` | `Mod02_Cores` |
| 2 | Preto Composto | `DetectarPretoSujo(silencioso:=True)` | `Mod02_Cores` |
| 3 | RGB → CMYK | `ConverterRGB(silencioso:=True)` — sempre CMYK completo | `Mod02_Cores` |
| 4 | Textos em Curvas | `ConverterTextosEmCurvas(silencioso:=True)` | `Mod03_Vetores` |
| 5 | Corrigir Contornos | `PadronizarContornosFinos(silencioso:=True)` | `Mod03_Vetores` |
| 6 | Padronizar Imagens | `PadronizarImagensCMYK600(silencioso:=True)` | `Mod05_Imagens` |
| 7 | Borda Dura | `CorrigirBordaDuraGradientes(minDot, silencioso:=True)` | `Mod02_Cores` |

---

## Critérios de Aceite

- Botões individuais do `frmFlexo`: comportamento **idêntico ao atual** — diálogos e mensagens continuam aparecendo
- "Corrigir Erros" do PreFlight: executa todas as correções em sequência sem diálogos, usando os mesmos algoritmos dos botões individuais
- `CrawlerCorrecoes` não existe mais no código

## Testes de Regressão Obrigatórios

| Bloco | Testes |
|-------|--------|
| Cores | T10–T25 |
| Vetores | T26–T34 |
| Bitmaps | T35–T36 |
| PreFlight | T43–T50 |

> **Critério geral:** nenhum dos 50 testes aprovados pode regredir após a implementação.

---

*Console Flexo v2.0 | CorelDRAW 2026 | Abril 2026*
