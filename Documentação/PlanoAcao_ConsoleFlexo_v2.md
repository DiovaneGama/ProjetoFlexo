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

## Correções Pós-Homologação — 15/04/2026

| ID | Módulo | Descrição | Teste |
|----|--------|-----------|-------|
| H1 | Mod02_Cores | `CMYKAssign` não persistia em fill — substituído por atribuições individuais de canal (`CMYKCyan/Magenta/Yellow/Black =`) | T14b |
| H2 | Mod02_Scanner_Engine | Guard `ActiveDocument Is Nothing` em `ExecutarScanner` — previne Erro 91 ao abrir PreFlight sem documento | T08 |
| H3 | frmFlexo | Caption do botão "RGB p/ Pantone" corrigido para "Converter p/ Pantone" | T09 |
| H4 | Todos os módulos | Acentuação corrigida em todos os MsgBox/InputBox/ControlTipText — chars U+FFFD substituídos por `ChrW()` | — |

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

## Roadmap v1.1 — CONCLUÍDO ✓

**Concluído em:** 14 de abril de 2026

| ID | Item | Status |
|----|------|--------|
| FM3 | Centralizar botões — `ObterTodosBotoes()` em frmFlexo.frm | Concluído |
| F9 | Constantes nomeadas em Mod05_Imagens.bas | Concluído |
| P1 | On Error pontual em CrawlerMergulhoProfundo | Concluído |
| P7 | `relatorio` privado + `GetRelatorio()` | Concluído |
| F2 | Unificar borda dura em Mod08_Utils | Concluído |
| — | **Homologação v1.1** | **Concluído ✓ — 15/04/2026** |

---

## Roadmap Futuro — Boas Práticas Flexo (FIRST/FTA + ISO 12647-6)

### Regras de Negócio

**C1 — TAC (Total Area Coverage)**
- Scanner detecta objetos cuja soma dos canais CMYK ultrapassa o limite configurado (padrão: 280%)
- Limite baseado em FIRST/ISO 12647-6 — acima disso a tinta não seca, mancha e cola na bobina
- Exibe `QtdTAC` no relatório do PreFlight com contagem de objetos infratores
- Não há correção automática — operador deve ajustar manualmente os valores de cor

**C2 — Mínimo de Ponto (Min Dot)**
- Scanner detecta canais com valor entre 0,1% e 1,9% (abaixo do mínimo imprimível mas não zero)
- Valor zero é intencional e não deve ser sinalizado
- O tratamento de correção já existe no Console Flexo:
  - **Gradientes:** usar o botão **"Mínimas Degradê"** para aplicar o ponto mínimo
  - **Tons planos:** usar o botão **"Limpar Cores"** para remover sujeira de canal
- Scanner exibe `QtdMinDot` no relatório como alerta — sem correção automática neste módulo

**V1 — Fonte Mínima**
- Scanner detecta textos com corpo abaixo do mínimo imprimível em flexo:
  - Positivo (texto escuro em fundo claro): mínimo **6pt**
  - Negativo (texto claro em fundo escuro): mínimo **8pt**
- Abaixo desses limites as letras fecham ou somem na impressão
- Exibe `QtdFontesPequenas` no relatório

**I1 — DPI Efetivo**
- Scanner calcula o DPI real da imagem considerando a escala aplicada no CorelDRAW
- Ex: imagem 300 DPI escalada a 200% = 150 DPI efetivo — hoje o scanner lê apenas o DPI original
- Substitui ou complementa o campo `QtdImgBaixa` com o valor efetivo real

**D1 — Sangria (Bleed)**
- Verifica se os elementos de arte ultrapassam o TrimBox pelo mínimo exigido (3mm)
- Sangria insuficiente gera borda branca após o corte da embalagem
- Depende de TrimBox estar corretamente aplicado via botão **"Aplicar Trimbox"**

**D2 — Camadas Técnicas**
- Valida presença e nomenclatura das camadas de acabamento obrigatórias: Branco, Verniz, Corte, Vinco
- Extensão natural do `Mod06_PadronizarLayers` — apenas verificação, sem criação automática
- Exibe aviso no PreFlight se camada esperada estiver ausente

**D3 — Relatório por Página**
- Scanner exibe resultados separados por página do documento
- Útil para documentos com múltiplas variações de arte no mesmo arquivo
- Requer adição do campo `Pagina` na struct `RelatorioPreFlight`

---

## Análise de Viabilidade — Roadmap Futuro

| ID | Viabilidade | Esforço | Risco | Observação |
|----|:-----------:|:-------:|:-----:|------------|
| C1 | Alta | Baixo | Nenhum | Soma simples de CMYK por objeto no Scanner |
| C2 | Alta | Baixo | Nenhum | Correção já existe — apenas alerta no Scanner |
| V1 | Alta | Médio | Baixo | API CorelDRAW expõe tamanho de fonte via TextRange |
| I1 | Alta | Médio | Baixo | Requer cruzar ResolutionX com SizeWidth/OriginalWidth |
| D1 | Média | Médio | Médio | Depende de TrimBox estar aplicado corretamente |
| D2 | Alta | Baixo | Nenhum | Extensão natural do Mod06_PadronizarLayers |
| D3 | Alta | Médio | Baixo | PreFlight já tem estrutura — adicionar campo Pagina |

---

## Critérios de Aceite — Homologação v1.1

> Todos os itens do Roadmap v1.1 foram implementados. A bateria completa de 50 testes
> deve ser executada para homologar a versão antes de avançar para o Roadmap Futuro.

| Bloco | Testes | Cobertura |
|-------|--------|-----------|
| Interface | T01–T09 | FM3 |
| Cores | T10–T25 | F2, F9 |
| Vetores | T26–T34 | — |
| Bitmaps | T35–T36 | F9 |
| Montagem | T37–T42 | — |
| PreFlight | T43–T50 | P1, P7, F2 |

> **Critério geral:** nenhum dos 50 testes aprovados pode regredir.

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
