# Plano de Ação — Pós-Lançamento v2.0
**Plataforma:** CorelDRAW 2026 | **Linguagem:** VBA | **Data:** Abril 2026

---

## Resultado dos Testes de Lançamento

| Testes Executados | Aprovados | Cancelados | Reprovados |
|:-----------------:|:---------:|:----------:|:----------:|
| **50** | **50** | **0** | **0** |

> Suite completa cobre 6 blocos: Interface (T01–T09), Cores (T10–T25), Vetores (T26–T34), Bitmaps (T35–T36), Montagem (T37–T42) e PreFlight (T43–T50).

---

## Correções Críticas Aplicadas pré-Lançamento

| ID | Módulo | Descrição | Status |
|----|--------|-----------|--------|
| P2 | Mod02_Scanner_Engine | Guard de documento e página aberta em ExecutarScanner e ExecutarCorrecoes | Aplicado |
| P3 | frmPreFlight | Validação de minDot com IsNumeric e range check (1–10%) | Aplicado |
| F1 | Mod02_Cores | On Error pontual no loop de escrita de gradientes | Aplicado |
| T8 | Mod02_Cores | Erro 80004005 ao aplicar Tint em CMYK — On Error Resume Next cobrindo FASE 2 | Aplicado |
| FP | Mod02_Scanner_Engine | Falso positivo de linha fina por imprecisão de ponto flutuante — Round(espMM,2) | Aplicado |

---

## Novas Funcionalidades Implementadas

| Funcionalidade | Módulo | Descrição |
|----------------|--------|-----------|
| Botão Desbloquear | Mod03_Vetores / frmFlexo | Desbloqueia todos os objetos da página ativa recursivamente via CrawlerDesbloquear |
| Detecção Ctrl+Shift+Q | Mod02_Scanner_Engine | Detecta contornos convertidos em objeto pela menor dimensão do bounding box |
| Filtro branco intencional | Mod02_Scanner_Engine | Exceção apenas para contornos brancos entre 0.02mm e 0.05mm |
| Aviso gradiente bloqueado | Mod02_Scanner_Engine / frmPreFlight | QtdGradBloqueado no relatório + MsgBox informativo após varredura |
| Abortar todos bloqueados | Mod02_Cores | Se todos os gradientes estiverem bloqueados, aborta e orienta o usuário |
| Botões de seleção sem Undo | frmFlexo | btnSelPreenchimento, btnSelContorno e btnEspessuraMinima não habilitam Desfazer |
| Hover em todos os botões | frmFlexo | MouseMove e MouseDown implementados em todos os 18 botões do Console |
| Reset visual | frmFlexo | Botão Reset restaura estado padrão de todos os botões sem alterar o arquivo |
| Varredura página ativa | Mod02_Scanner_Engine | Scanner varre apenas a página ativa, não o documento inteiro |
| Frames colapsáveis | frmFlexo | 4 seções colapsáveis com ícone de seta e reposicionamento dinâmico |

---

## Roadmap v1.1 — Cronograma (cadência a cada 2 dias)

**Início:** 14 de abril de 2026

| Dia | Data | ID | Item | Entrega | Testes de Regressão |
|-----|------|----|------|---------|---------------------|
| D1 | 14 abr (seg) | **FM3** | Centralizar botões | `ObterTodosBotoes()` em frmFlexo.frm | T01–T09, T07 |
| D3 | 16 abr (qua) | **F9** | Constantes nomeadas | Constantes em Mod05_Imagens.bas | T35–T36 |
| D5 | 18 abr (sex) | **P1** | On Error do Scanner | Error handling pontual em CrawlerMergulhoProfundo | T43–T47 |
| D7 | 22 abr (ter) | **P7** | Encapsular relatório | `relatorio` privado + `GetRelatorio()` | T43–T50 |
| D9 | 24 abr (qui) | **F2** | Unificar borda dura | `Mod06_Utils.bas` + refatorar Scanner e Cores | T19–T22, T43–T49 |
| D11 | 26 abr (sab) | — | **Homologação v1.1** | Bateria completa dos 50 testes | T01–T50 |

> Os dias D2, D4, D6, D8, D10 são reservados para testes manuais no CorelDRAW e ajustes.

### Sequência recomendada

```
FM3 → F9 → P1 → P7 → F2
(fácil → fácil → médio → médio → difícil)
```

---

## Análise de Viabilidade — v1.1

| ID | Item | Viabilidade | Esforço | Risco | Arquivos afetados |
|----|------|:-----------:|:-------:|:-----:|-------------------|
| FM3 | Centralizar botões | Alta | Baixo | Baixo | `frmFlexo.frm` |
| F9 | Constantes nomeadas | Alta | Baixo | Nenhum | `Mod05_Imagens.bas` |
| P1 | On Error do Scanner | Alta | Baixo | Baixo | `Mod02_Scanner_Engine.bas` |
| P7 | Encapsular relatório | Média | Médio | Médio | `Mod02_Scanner_Engine.bas`, `frmPreFlight.frm` |
| F2 | Unificar borda dura | Média | Alto | Médio | `Mod02_Scanner_Engine.bas`, `Mod02_Cores.bas`, novo `Mod06_Utils.bas` |

---

## Roadmap Futuro — Cronograma (Maio–Junho 2026)

| Dia | Data | ID | Item | Entrega |
|-----|------|----|------|---------|
| D1 | 05 mai (ter) | **M2** | Relatório por página | Campo `Pagina` na struct + exibição no PreFlight |
| D3 | 07 mai (qui) | **M2** | Testes e ajustes | Re-executar T43–T50 + validar campo Pagina |
| D5 | 12 mai (ter) | **M1** | Cores técnicas configurável (UI) | Form ou InputBox de configuração |
| D7 | 14 mai (qui) | **M1** | Persistência (arquivo INI) | Gravar/ler lista de cores técnicas entre sessões |
| D9 | 16 mai (sex) | **M1** | Testes e ajustes | Validar persistência entre sessões CorelDRAW |

## Análise de Viabilidade — Roadmap Futuro

| ID | Funcionalidade | Viabilidade | Esforço | Observação |
|----|----------------|:-----------:|:-------:|------------|
| M2 | Relatório por página | Alta | Médio | PreFlight já tem estrutura; adicionar campo `Pagina` no tipo |
| M1 | Cores técnicas configurável | Média | Alto | Requer UI + persistência via arquivo INI |
| FM2 | Persistência de estado | Baixa | Alto | Adiado para Agosto 2026 — complexidade alta, risco de corrupção de estado |

---

## Critérios de Aceite por Item

| Item | Testes de regressão obrigatórios |
|------|----------------------------------|
| **F2** Unificar borda dura | T19–T22, T43–T49, T46 |
| **FM3** Centralizar botões | T01–T09, T07 |
| **P7** Encapsular relatório | T43–T50 |
| **P1** On Error do Scanner | T43, T44, T45, T47 |
| **F9** Constantes nomeadas | T35–T36 |
| **M2** Relatório por página | T45 + validar campo Pagina no relatório |

> **Critério geral:** nenhum dos 50 testes aprovados pode regredir após qualquer mudança do roadmap.

---

## Padrões Técnicos de Flexografia

| Parâmetro | Valor | Observação |
|-----------|-------|------------|
| Resolução de imagens | 300 DPI mínimo | Scanner alerta abaixo disso |
| Espessura mínima de linha | ≤ 0.1mm | Padrão FIRST — detectado pelo Scanner e Inspetor |
| Contorno branco intencional | 0.02mm a 0.05mm | Exceção — não detectado como erro |
| Contorno padronizado | 0.2mm | Aplicado automaticamente pelo botão Corrigir Contornos |
| Ponto mínimo em gradientes | 2% ou 3% | Definido pelo operador no InputBox |
| Offset Banda Larga | 7mm cada lado | Opção no TrimBox |
| Offset Banda Estreita | 5mm cada lado | Opção no TrimBox |
| Preto Puro | C0 M0 Y0 K100 | Padrão para textos e traçados |
| Preto Rico | C100 M100 Y100 K100 | Único aceito além do puro |

---

*Console Flexo v2.0 | CorelDRAW 2026 | Abril 2026*
