# Auditoria de Código — Correções Pendentes
**Data:** 2026-04-13

---

## 🔴 Crítico — Corrigir antes dos testes

### C1 — `CorrigirBordaDuraGradientes` nunca executa no modo silencioso
**Arquivo:** `src/ModuloFlexo/Mod02_Cores.bas` (~linha 433)

**Problema:** `minDotLong` é avaliado **antes** de receber o valor de `minDot`:
```vba
If silencioso Then
    If minDotLong <= 0 Then Exit Sub   ' ← minDotLong ainda é 0 aqui!
    minDotLong = CLng(minDot)          ' ← nunca chega aqui
End If
```
**Impacto:** O botão "Corrigir Erros" do PreFlight passa `minDot > 0` mas a função sai imediatamente — borda dura nunca é corrigida no modo em lote.

**Correção:**
```vba
If silencioso Then
    minDotLong = CLng(minDot)          ' ← primeiro atribui
    If minDotLong <= 0 Then Exit Sub   ' ← depois valida
End If
```

---

### C2 — `ErroHandler` comentado no Mod07_InserirMicropontos
**Arquivo:** `src/ModuloFlexo/Mod07_InserirMicropontos.bas` (linhas 164 e 338–344)

**Problema:** `On Error GoTo ErroHandler` e o bloco `ErroHandler:` estão comentados. Qualquer erro durante importação ou posicionamento encerra o macro sem mensagem e sem `EndCommandGroup`, deixando o documento em estado inconsistente.

**Correção:** Reativar o handler ou adicionar no mínimo `On Error GoTo Finalizar` na linha 164, aproveitando o bloco `Finalizar:` já existente que faz o cleanup correto.

---

## 🟡 Aviso — Corrigir antes da homologação

### A1 — `End If` possivelmente desbalanceado em `CorrigirBordaDuraGradientes`
**Arquivo:** `src/ModuloFlexo/Mod02_Cores.bas` (~linha 663)

**Problema:** O bloco `If srCorrigidos.Count > 0 Then` abre dentro de `If Not silencioso Then` mas o `End If` externo pode estar faltando — verificar se VBA compila sem erro de sintaxe.

**Correção:** Revisar o aninhamento e garantir dois `End If` fechando os dois `If` abertos.

---

### A5 — Lógica de reordenação de layers move apenas uma posição
**Arquivo:** `src/ModuloFlexo/Mod06_PadronizarLayers.bas` (linhas 83–86)

**Problema:** Se "Saida" estiver na posição 5, um único `MoveAbove lTop` não a coloca no topo.

**Correção:**
```vba
Do While pg.Layers.Item(1).Name <> names(1)
    lSaida.MoveAbove pg.Layers.Item(1)
Loop
```

---

## 🟢 Cosmético (baixa prioridade)

| ID | Arquivo | Descrição |
|----|---------|-----------|
| A2 | `Mod04_Montagem.bas` | `Application.Optimization = False` chamado duas vezes no path de sucesso |
| A3 | `Mod02_Scanner_Engine.bas` | `Dim p As Page` declarado mas nunca usado em `ExecutarScanner` |
| A4 | `Mod00_Versao.bas` | Caminhos hardcoded `X:\version.txt` e `X:\ConsoleFlexo.gms` — requer mapeamento de rede |
