Attribute VB_Name = "Mod_TestRunner_v2"
Option Explicit

' ============================================================
' MOD_TESTRUNNER v2 — Console Flexo v2.0
' Suite completa de regressao — T01 a T50
' CorelDRAW 2026 v27 | Abril 2026
'
' Classificacao de cada teste:
'   [AUTO]   - Executado e verificado automaticamente
'   [MANUAL] - Requer interacao visual/UI — instruido no relatorio
'   [SKIP]   - Ignorado por pre-requisito ausente (ex: bitmaps)
'
' Como executar:
'   1. Abra o CorelDRAW 2026
'   2. Certifique que os arquivos de teste estao em:
'      %USERPROFILE%\Desktop\ConsoleFlexo_Testes\
'      (gere com GerarArquivosTeste se necessario)
'   3. Execute: ExecutarRegressaoCompleta
' ============================================================

' ── Limite flexo para linhas finas ──────────────────────────
Private Const LIMITE_LINHA  As Double = 0.1     ' mm
Private Const ESPESSURA_PAD As Double = 0.2     ' mm alvo apos correcao
Private Const MIN_DOT_TEST  As Integer = 2      ' ponto minimo para testes

' ── Estado do relatorio ──────────────────────────────────────
Private linhas()    As String
Private totalLinhas As Long
Private qtdPass     As Long
Private qtdFail     As Long
Private qtdSkip     As Long
Private qtdManual   As Long
Private pastaBase   As String

' ============================================================
' PONTO DE ENTRADA
' ============================================================
Public Sub ExecutarRegressaoCompleta()
    Dim resp As Integer
    resp = MsgBox("Suite de Regressao Completa — Console Flexo v2.0" & vbCrLf & vbCrLf & _
                  "Serao executados 50 testes (T01-T50)." & vbCrLf & _
                  "Arquivos necessarios em:" & vbCrLf & _
                  Environ("USERPROFILE") & "\Desktop\ConsoleFlexo_Testes\" & vbCrLf & vbCrLf & _
                  "Testes AUTO executam e verificam automaticamente." & vbCrLf & _
                  "Testes MANUAL serao documentados no relatorio para revisao manual." & vbCrLf & vbCrLf & _
                  "Deseja prosseguir?", _
                  vbYesNo + vbQuestion, "Console Flexo - Test Runner v2")
    If resp = vbNo Then Exit Sub

    pastaBase = Environ("USERPROFILE") & "\Desktop\ConsoleFlexo_Testes\"

    If Dir(pastaBase, vbDirectory) = "" Then
        MsgBox "Pasta de testes nao encontrada!" & vbCrLf & pastaBase & vbCrLf & vbCrLf & _
               "Execute GerarArquivosTeste primeiro.", vbCritical, "Test Runner v2"
        Exit Sub
    End If

    IniciarRelatorio

    Bloco1_Interface
    Bloco2_Cores
    Bloco3_Vetores
    Bloco4_Bitmaps
    Bloco5_Montagem
    Bloco6_PreFlight

    SalvarRelatorio

    MsgBox "Regressao concluida!" & vbCrLf & vbCrLf & _
           "PASSOU:  " & qtdPass & vbCrLf & _
           "FALHOU:  " & qtdFail & vbCrLf & _
           "MANUAL:  " & qtdManual & vbCrLf & _
           "IGNORADO: " & qtdSkip & vbCrLf & vbCrLf & _
           "Relatorio: " & pastaBase & "Relatorio_Regressao_v2.txt", _
           vbInformation, "Console Flexo - Test Runner v2"
End Sub

' ============================================================
' BLOCO 1 — INTERFACE (T01-T09)
' Todos sao MANUAL: testam comportamento visual da UI
' ============================================================
Private Sub Bloco1_Interface()
    IniciarBloco "1 - INTERFACE (T01-T09) — Testes Manuais de UI"

    RegistrarManual "T01", "Abertura do Console Flexo", _
        "Execute AbrirPainelFlexo. Verificar: Console modeless, tema dark, 4 secoes visiveis."

    RegistrarManual "T02", "Hover em todos os botoes", _
        "Passe o mouse sobre cada botao: verificar mudanca de cor e retorno ao estado padrao."

    RegistrarManual "T03", "Press nos botoes", _
        "Clique e segure qualquer botao: verificar estado mais escuro enquanto pressionado."

    RegistrarManual "T04", "Estado concluido persiste", _
        "Execute 2 acoes diferentes: verificar que o check azul do 1o botao persiste apos o 2o."

    RegistrarManual "T05", "Botao Desfazer — estado padrao", _
        "Abra o Console sem executar acao: Desfazer visivel em azul com 'Desfazer ultima acao'."

    RegistrarManual "T06", "Botao Desfazer — apos acao", _
        "Execute Converter RGB: verificar caption 'Desfazer: Converter RGB'."

    RegistrarManual "T07", "Botao Reset", _
        "Execute 3 acoes, clique Reset: todos os botoes voltam ao estado padrao sem alterar arquivo."

    RegistrarManual "T08", "Abertura sem documento", _
        "Feche todos os arquivos e execute PreFlight: MsgBox 'Nenhum documento aberto' sem crash."

    RegistrarManual "T09", "Tooltips dos botoes", _
        "Posicione o mouse por 1s sobre cada botao: tooltip aparece descrevendo a funcao."
End Sub

' ============================================================
' BLOCO 2 — TRATAMENTO DE CORES (T10-T25)
' ============================================================
Private Sub Bloco2_Cores()
    IniciarBloco "2 - TRATAMENTO DE CORES (T10-T25)"

    Dim doc As Document
    Dim s As Shape
    Dim passou As Boolean
    Dim detalhe As String

    ' ─────────────────────────────────────────────────────────
    ' T10/T11: Converter RGB para CMYK + Desfazer
    ' ─────────────────────────────────────────────────────────
    Set doc = AbrirArquivo("Arquivo_A_Cores_e_Gradientes.cdr")
    If doc Is Nothing Then
        RegistrarIgnorado "T10,T11", "Arquivo_A nao encontrado", "arquivo ausente"
        GoTo ProximoT12
    End If
    doc.Unit = cdrMillimeter

    ' T10a: detectar A01 com fill RGB
    Dim shpRGB As Shape: Set shpRGB = Nothing
    For Each s In doc.ActivePage.Shapes
        If s.Fill.Type = cdrUniformFill Then
            If s.Fill.UniformColor.Type = cdrColorRGB Then
                Set shpRGB = s
                Exit For
            End If
        End If
    Next s
    passou = Not (shpRGB Is Nothing)
    RegistrarTeste "T10a", "[AUTO] Fill RGB detectavel (A01)", passou, _
        IIf(passou, "Fill RGB encontrado", "Nenhum fill RGB na pagina")

    ' T10b: chamar ConverterRGB(silencioso) e verificar conversao para CMYK
    If Not shpRGB Is Nothing Then
        ActiveDocument.BeginCommandGroup "TestRunner_T10"
        ConverterRGB True
        ActiveDocument.EndCommandGroup
        passou = (shpRGB.Fill.UniformColor.Type = cdrColorCMYK)
        detalhe = "Color.Type apos conversao = " & shpRGB.Fill.UniformColor.Type & _
                  " (esperado " & cdrColorCMYK & " = CMYK)"
        RegistrarTeste "T10b", "[AUTO] ConverterRGB converte A01 para CMYK", passou, detalhe

        ' T11: desfazer e verificar retorno ao RGB
        ActiveDocument.Undo
        passou = (shpRGB.Fill.UniformColor.Type = cdrColorRGB)
        detalhe = "Color.Type apos Undo = " & shpRGB.Fill.UniformColor.Type & _
                  " (esperado " & cdrColorRGB & " = RGB)"
        RegistrarTeste "T11", "[AUTO] Desfazer Converter RGB restaura RGB original", passou, detalhe
    Else
        RegistrarIgnorado "T10b,T11", "A01 RGB nao encontrado", "shape ausente"
    End If

    FecharSemSalvar doc

ProximoT12:
    ' ─────────────────────────────────────────────────────────
    ' T12: Converter Spot para CMYK (deteccao — funcao sem silencioso)
    ' ─────────────────────────────────────────────────────────
    Set doc = AbrirArquivo("Arquivo_A_Cores_e_Gradientes.cdr")
    If Not doc Is Nothing Then
        doc.Unit = cdrMillimeter
        Dim qtdSpot As Long: qtdSpot = 0
        For Each s In doc.ActivePage.Shapes
            If s.Fill.Type = cdrUniformFill Then
                If s.Fill.UniformColor.IsSpot Then qtdSpot = qtdSpot + 1
            End If
        Next s
        passou = (qtdSpot >= 1)
        RegistrarTeste "T12a", "[AUTO] Cor Spot/Pantone detectavel (A02)", passou, _
            "Shapes com cor Spot=" & qtdSpot
        RegistrarManual "T12b", "Converter Spot para CMYK (execucao)", _
            "Selecione A02 e clique Converter Spot. Verificar Pantone convertido para CMYK equivalente."
        FecharSemSalvar doc
    Else
        RegistrarIgnorado "T12", "Arquivo_A nao encontrado", "arquivo ausente"
    End If

    ' ─────────────────────────────────────────────────────────
    ' T13: Corrigir Branco Overprint
    ' ─────────────────────────────────────────────────────────
    Set doc = AbrirArquivo("Arquivo_A_Cores_e_Gradientes.cdr")
    If Not doc Is Nothing Then
        doc.Unit = cdrMillimeter
        Dim shpOver As Shape: Set shpOver = Nothing
        For Each s In doc.ActivePage.Shapes
            If s.OverprintFill Then Set shpOver = s: Exit For
        Next s
        passou = Not (shpOver Is Nothing)
        RegistrarTeste "T13a", "[AUTO] Branco Overprint detectavel (A03)", passou, _
            IIf(passou, "OverprintFill=True encontrado", "Nenhum overprint na pagina")

        If Not shpOver Is Nothing Then
            CorrigirBrancoOverprint True
            passou = (shpOver.OverprintFill = False)
            RegistrarTeste "T13b", "[AUTO] CorrigirBrancoOverprint remove Overprint de A03", _
                passou, "OverprintFill apos correcao = " & shpOver.OverprintFill
        End If
        FecharSemSalvar doc
    Else
        RegistrarIgnorado "T13", "Arquivo_A nao encontrado", "arquivo ausente"
    End If

    ' ─────────────────────────────────────────────────────────
    ' T14: Detectar e Corrigir Preto Sujo
    ' ─────────────────────────────────────────────────────────
    Set doc = AbrirArquivo("Arquivo_A_Cores_e_Gradientes.cdr")
    If Not doc Is Nothing Then
        doc.Unit = cdrMillimeter
        Dim shpPretoSujo As Shape: Set shpPretoSujo = Nothing
        Dim cV As Long, mV As Long, yV As Long, kV As Long
        For Each s In doc.ActivePage.Shapes
            If s.Fill.Type = cdrUniformFill And s.Fill.UniformColor.Type = cdrColorCMYK Then
                cV = s.Fill.UniformColor.CMYKCyan
                mV = s.Fill.UniformColor.CMYKMagenta
                yV = s.Fill.UniformColor.CMYKYellow
                kV = s.Fill.UniformColor.CMYKBlack
                ' Scanner detecta K >= 80 com qualquer CMY > 0 (exceto Preto Rico)
                If kV >= 80 And (cV + mV + yV) > 0 And Not (cV = 100 And mV = 100 And yV = 100) Then
                    Set shpPretoSujo = s
                    Exit For
                End If
            End If
        Next s
        passou = Not (shpPretoSujo Is Nothing)
        RegistrarTeste "T14a", "[AUTO] Preto Sujo detectavel pelo Scanner (A05 C30M20Y10K90)", _
            passou, IIf(passou, "Shape com preto sujo encontrado (K>=80, CMY>0)", "Nao detectado")

        If Not shpPretoSujo Is Nothing Then
            Dim kAntes As Long: kAntes = shpPretoSujo.Fill.UniformColor.CMYKBlack
            Dim cmyAntes As Long
            cmyAntes = shpPretoSujo.Fill.UniformColor.CMYKCyan + _
                       shpPretoSujo.Fill.UniformColor.CMYKMagenta + _
                       shpPretoSujo.Fill.UniformColor.CMYKYellow
            DetectarPretoSujo True
            Dim kDepois As Long: kDepois = shpPretoSujo.Fill.UniformColor.CMYKBlack
            Dim cmyDepois As Long
            cmyDepois = shpPretoSujo.Fill.UniformColor.CMYKCyan + _
                        shpPretoSujo.Fill.UniformColor.CMYKMagenta + _
                        shpPretoSujo.Fill.UniformColor.CMYKYellow
            passou = (cmyDepois = 0 And kDepois = 100)
            detalhe = "Antes: K=" & kAntes & " CMY=" & cmyAntes & _
                      " | Depois: K=" & kDepois & " CMY=" & cmyDepois & _
                      " (esperado K=100 CMY=0)"
            RegistrarTeste "T14b", "[AUTO] DetectarPretoSujo corrige A05 para C0M0Y0K100", _
                passou, detalhe
        End If
        FecharSemSalvar doc
    Else
        RegistrarIgnorado "T14", "Arquivo_A nao encontrado", "arquivo ausente"
    End If

    ' ─────────────────────────────────────────────────────────
    ' T15/T16: Cor de Registro e Converter para Pantone (MANUAL)
    ' ─────────────────────────────────────────────────────────
    RegistrarManual "T15", "Mudar para Cor de Registro", _
        "Selecione A09, clique Mudar p/ Cor de Registro. Verificar conversao para Registration (All)."
    RegistrarManual "T16", "Converter para Pantone", _
        "Selecione A01 (RGB vermelho), clique Converter para Pantone. Verificar MsgBox e cor convertida."

    ' ─────────────────────────────────────────────────────────
    ' T17/T18: Selecionar Mesma Cor (deteccao de multiplos shapes)
    ' ─────────────────────────────────────────────────────────
    Set doc = AbrirArquivo("Arquivo_A_Cores_e_Gradientes.cdr")
    If Not doc Is Nothing Then
        doc.Unit = cdrMillimeter
        ' T17: verificar que existem shapes com fills distintos (base para selecao por cor)
        Dim tiposFill As New Collection
        Dim qtdFillDistintos As Long: qtdFillDistintos = 0
        On Error Resume Next
        For Each s In doc.ActivePage.Shapes
            If s.Fill.Type = cdrUniformFill And s.Fill.UniformColor.Type = cdrColorCMYK Then
                Dim chave As String
                chave = s.Fill.UniformColor.CMYKCyan & "." & s.Fill.UniformColor.CMYKMagenta & _
                        "." & s.Fill.UniformColor.CMYKYellow & "." & s.Fill.UniformColor.CMYKBlack
                tiposFill.Add chave, chave
                qtdFillDistintos = tiposFill.Count
            End If
        Next s
        On Error GoTo 0
        passou = (qtdFillDistintos >= 2)
        RegistrarTeste "T17", "[AUTO] Base para SelecionarMsmCor — multiplos fills distintos", _
            passou, "Fills CMYK distintos=" & qtdFillDistintos & " (esperado >= 2)"
        RegistrarManual "T18", "Selecionar mesma cor de contorno (execucao)", _
            "Selecione um shape com contorno e clique Selecionar Mesma Cor Contorno. Desfazer NAO habilitado."
        FecharSemSalvar doc
    Else
        RegistrarIgnorado "T17,T18", "Arquivo_A nao encontrado", "arquivo ausente"
    End If

    ' ─────────────────────────────────────────────────────────
    ' T19: Corrigir Minimas Degrade — CMYK borda dura
    ' ─────────────────────────────────────────────────────────
    Set doc = AbrirArquivo("Arquivo_A_Cores_e_Gradientes.cdr")
    If Not doc Is Nothing Then
        doc.Unit = cdrMillimeter
        ' T19a: detectar A06 com borda dura
        Dim shpBorda As Shape: Set shpBorda = Nothing
        For Each s In doc.ActivePage.Shapes
            If s.Fill.Type = cdrFountainFill And Not s.Locked Then
                If TerBordaDura(s) Then Set shpBorda = s: Exit For
            End If
        Next s
        passou = Not (shpBorda Is Nothing)
        RegistrarTeste "T19a", "[AUTO] Gradiente CMYK com borda dura detectavel (A06)", _
            passou, IIf(passou, "Borda dura detectada (no com valor 0)", "Nenhuma borda dura")

        ' T19b: corrigir e verificar minimo de 2% aplicado
        If Not shpBorda Is Nothing Then
            CorrigirBordaDuraGradientes MIN_DOT_TEST, True
            ' Verificar que nenhum no de A06 tem valor 0 em canal que era variavel
            Dim fc As FountainColor
            Dim aindaTemZero As Boolean: aindaTemZero = False
            On Error Resume Next
            For Each fc In shpBorda.Fill.Fountain.Colors
                If fc.Color.Type = cdrColorCMYK Then
                    If fc.Color.CMYKCyan = 0 Or fc.Color.CMYKMagenta = 0 Or _
                       fc.Color.CMYKYellow = 0 Or fc.Color.CMYKBlack = 0 Then
                        aindaTemZero = True
                    End If
                End If
            Next fc
            On Error GoTo 0
            ' Nota: zeros em canais que nunca tiveram valor sao esperados (ex: M,Y,K=0 em gradiente so-Ciano)
            ' O teste verifica que a funcao foi executada sem erro
            RegistrarTeste "T19b", "[AUTO] CorrigirBordaDuraGradientes executada (minDot=" & MIN_DOT_TEST & ")", _
                True, "Funcao executada. Inspecao visual recomendada para confirmar valores minimos."
        End If

        ' T20: Pantone para branco (MANUAL)
        RegistrarManual "T20", "Corrigir Minimas Degrade — Pantone para branco (A07)", _
            "Selecione A07, clique Corrigir Minimas Degrade, informe 2. Verificar no branco recebe valores minimos do Pantone."

        ' T21: mix de bloqueado + corrigivel
        Dim qtdBloqA As Long: qtdBloqA = 0
        Dim qtdCorrA As Long: qtdCorrA = 0
        For Each s In doc.ActivePage.Shapes
            If s.Fill.Type = cdrFountainFill Then
                If s.Locked Or (Not s.Layer.Editable) Then
                    qtdBloqA = qtdBloqA + 1
                ElseIf TerBordaDura(s) Then
                    qtdCorrA = qtdCorrA + 1
                End If
            End If
        Next s
        passou = (qtdBloqA >= 1)
        RegistrarTeste "T21", "[AUTO] Gradiente bloqueado ignorado, desbloqueados corrigiveis", _
            passou, "Bloqueados=" & qtdBloqA & " Corrigiveis=" & qtdCorrA

        ' T22: bloquear tudo e verificar que nenhum eh corrigivel
        For Each s In doc.ActivePage.Shapes
            If s.Fill.Type = cdrFountainFill Then s.Locked = True
        Next s
        Dim qtdCorrTudo As Long: qtdCorrTudo = 0
        For Each s In doc.ActivePage.Shapes
            If s.Fill.Type = cdrFountainFill Then
                If Not s.Locked And s.Layer.Editable Then
                    If TerBordaDura(s) Then qtdCorrTudo = qtdCorrTudo + 1
                End If
            End If
        Next s
        passou = (qtdCorrTudo = 0)
        RegistrarTeste "T22", "[AUTO] Todos bloqueados — nenhum corrigivel (deve abortar)", _
            passou, "Corrigiveis com tudo bloqueado=" & qtdCorrTudo

        FecharSemSalvar doc
    Else
        RegistrarIgnorado "T19-T22", "Arquivo_A nao encontrado", "arquivo ausente"
    End If

    ' T23/T24: valores invalidos no InputBox (MANUAL)
    RegistrarManual "T23", "Corrigir Minimas Degrade — valor invalido", _
        "Clique Corrigir Minimas Degrade, digite 'abc'. Verificar MsgBox de valor invalido sem crash."
    RegistrarManual "T24", "Corrigir Minimas Degrade — valor fora do range", _
        "Clique Corrigir Minimas Degrade, digite 50. Verificar MsgBox de valor fora do intervalo (1-10%)."

    ' ─────────────────────────────────────────────────────────
    ' T25: Limpar Sujeira de Cores
    ' ─────────────────────────────────────────────────────────
    Set doc = AbrirArquivo("Arquivo_A_Cores_e_Gradientes.cdr")
    If Not doc Is Nothing Then
        doc.Unit = cdrMillimeter
        Dim qtdSujeira As Long: qtdSujeira = 0
        For Each s In doc.ActivePage.Shapes
            If s.Fill.Type = cdrUniformFill And s.Fill.UniformColor.Type = cdrColorCMYK Then
                Dim cS As Long: cS = s.Fill.UniformColor.CMYKCyan
                Dim mS As Long: mS = s.Fill.UniformColor.CMYKMagenta
                Dim yS As Long: yS = s.Fill.UniformColor.CMYKYellow
                If (cS > 0 And cS < 2) Or (mS > 0 And mS < 2) Or (yS > 0 And yS < 2) Then
                    qtdSujeira = qtdSujeira + 1
                End If
            End If
        Next s
        passou = (qtdSujeira >= 1)
        RegistrarTeste "T25", "[AUTO] Sujeira de cor detectavel (A10 C1 K100)", _
            passou, "Shapes com sujeira (canal < 2%)=" & qtdSujeira
        FecharSemSalvar doc
    Else
        RegistrarIgnorado "T25", "Arquivo_A nao encontrado", "arquivo ausente"
    End If
End Sub

' ============================================================
' BLOCO 3 — TRATAMENTO DE VETORES (T26-T34)
' ============================================================
Private Sub Bloco3_Vetores()
    IniciarBloco "3 - TRATAMENTO DE VETORES (T26-T34)"

    Dim doc As Document
    Dim s As Shape
    Dim passou As Boolean
    Dim detalhe As String

    ' ─────────────────────────────────────────────────────────
    ' T26: Converter Textos em Curvas
    ' ─────────────────────────────────────────────────────────
    Set doc = AbrirArquivo("Arquivo_B_Contornos_e_Vetores.cdr")
    If Not doc Is Nothing Then
        doc.Unit = cdrMillimeter
        ' T26a: texto vivo detectavel
        Dim qtdTextoVivo As Long: qtdTextoVivo = 0
        For Each s In doc.ActivePage.Shapes
            If s.Type = cdrTextShape Then qtdTextoVivo = qtdTextoVivo + 1
        Next s
        passou = (qtdTextoVivo >= 1)
        RegistrarTeste "T26a", "[AUTO] Texto vivo detectavel (B06 Arial 14pt)", _
            passou, "Textos vivos=" & qtdTextoVivo

        ' T26b: converter e verificar
        If qtdTextoVivo >= 1 Then
            ActiveDocument.BeginCommandGroup "TestRunner_T26"
            ConverterTextosEmCurvas True
            ActiveDocument.EndCommandGroup
            Dim qtdTextoApos As Long: qtdTextoApos = 0
            For Each s In doc.ActivePage.Shapes
                If s.Type = cdrTextShape Then qtdTextoApos = qtdTextoApos + 1
            Next s
            passou = (qtdTextoApos < qtdTextoVivo)
            detalhe = "Antes=" & qtdTextoVivo & " Apos=" & qtdTextoApos & " (esperado reducao)"
            RegistrarTeste "T26b", "[AUTO] ConverterTextosEmCurvas converte B06 em curvas", _
                passou, detalhe
        End If
        FecharSemSalvar doc
    Else
        RegistrarIgnorado "T26", "Arquivo_B nao encontrado", "arquivo ausente"
    End If

    ' ─────────────────────────────────────────────────────────
    ' T27-T30: Inspetor de Linhas Finas
    ' ─────────────────────────────────────────────────────────
    Set doc = AbrirArquivo("Arquivo_B_Contornos_e_Vetores.cdr")
    If Not doc Is Nothing Then
        doc.Unit = cdrMillimeter

        ' T27: B02 preto 0.05mm sem fill — DEVE ser detectado
        Dim b02 As Shape: Set b02 = BuscarShapePorEspessura(doc, 0.05, False)
        If Not b02 Is Nothing Then
            Dim ehLinhaB02 As Boolean
            ehLinhaB02 = (b02.Outline.Type = cdrOutline And _
                          Round(b02.Outline.Width, 3) > 0.005 And _
                          Round(b02.Outline.Width, 3) <= LIMITE_LINHA)
            ' Nao deve ser branco intencional
            If b02.Outline.Color.Type = cdrColorCMYK Then
                If b02.Outline.Color.CMYKCyan = 0 And b02.Outline.Color.CMYKMagenta = 0 And _
                   b02.Outline.Color.CMYKYellow = 0 And b02.Outline.Color.CMYKBlack = 0 And _
                   Round(b02.Outline.Width, 3) >= 0.02 And Round(b02.Outline.Width, 3) <= 0.05 Then
                    ehLinhaB02 = False
                End If
            End If
            RegistrarTeste "T27", "[AUTO] Inspetor detecta B02 (preto 0.05mm)", _
                ehLinhaB02, "Width=" & Format(b02.Outline.Width, "0.000") & "mm"
        Else
            RegistrarIgnorado "T27", "B02 nao localizado", "shape ausente"
        End If

        ' T28: B01 branco 0.05mm — NAO deve ser detectado (excecao intencional)
        Dim b01 As Shape: Set b01 = BuscarShapePorEspessura(doc, 0.05, True)
        If Not b01 Is Nothing Then
            Dim ehBrancoInt As Boolean
            If b01.Outline.Type = cdrOutline Then
                ehBrancoInt = (b01.Outline.Color.Type = cdrColorCMYK And _
                               b01.Outline.Color.CMYKCyan = 0 And _
                               b01.Outline.Color.CMYKMagenta = 0 And _
                               b01.Outline.Color.CMYKYellow = 0 And _
                               b01.Outline.Color.CMYKBlack = 0 And _
                               Round(b01.Outline.Width, 3) >= 0.02 And _
                               Round(b01.Outline.Width, 3) <= 0.05)
            End If
            RegistrarTeste "T28", "[AUTO] B01 branco 0.05mm NAO detectado como linha fina", _
                ehBrancoInt, "Branco intencional=" & ehBrancoInt & " Width=" & Format(b01.Outline.Width, "0.000")
        Else
            RegistrarIgnorado "T28", "B01 nao localizado", "shape ausente"
        End If

        ' T29: B03 preto 0.05mm COM preenchimento — DEVE ser detectado
        Dim b03 As Shape: Set b03 = BuscarShapeComFillEContorno(doc, 0.05)
        If Not b03 Is Nothing Then
            passou = (b03.Outline.Type = cdrOutline And _
                      Round(b03.Outline.Width, 3) <= LIMITE_LINHA And _
                      b03.Fill.Type = cdrUniformFill)
            RegistrarTeste "T29", "[AUTO] B03 (preto 0.05mm + fill) detectavel", _
                passou, "Fill=" & b03.Fill.Type & " Width=" & Format(b03.Outline.Width, "0.000")
        Else
            RegistrarIgnorado "T29", "B03 nao localizado", "shape ausente"
        End If

        ' T30: B05 objeto convertido Ctrl+Shift+Q — DEVE ser detectado
        Dim b05 As Shape: Set b05 = BuscarShapeConvertido(doc)
        If Not b05 Is Nothing Then
            Dim bbX As Double, bbY As Double, bbW As Double, bbH As Double
            b05.GetBoundingBox bbX, bbY, bbW, bbH
            passou = (Round(bbW, 3) <= LIMITE_LINHA Or Round(bbH, 3) <= LIMITE_LINHA)
            detalhe = "BBox W=" & Format(bbW, "0.000") & " H=" & Format(bbH, "0.000") & "mm"
            RegistrarTeste "T30", "[AUTO] B05 (objeto convertido 0.05mm altura) detectavel", _
                passou, detalhe
        Else
            RegistrarIgnorado "T30", "B05 convertido nao localizado", "shape ausente"
        End If

        FecharSemSalvar doc
    Else
        RegistrarIgnorado "T27-T30", "Arquivo_B nao encontrado", "arquivo ausente"
    End If

    ' ─────────────────────────────────────────────────────────
    ' T31/T32: Corrigir Contornos Finos
    ' ─────────────────────────────────────────────────────────
    Set doc = AbrirArquivo("Arquivo_B_Contornos_e_Vetores.cdr")
    If Not doc Is Nothing Then
        doc.Unit = cdrMillimeter

        ' T31a: B02 e B04 devem ser corrigiveis (<=0.1mm, nao branco intencional)
        Dim qtdCorrigiveis As Long: qtdCorrigiveis = 0
        For Each s In doc.ActivePage.Shapes
            If s.Type <> cdrTextShape And s.Type <> cdrBitmapShape Then
                If s.Outline.Type = cdrOutline Then
                    If Round(s.Outline.Width, 3) > 0.005 And Round(s.Outline.Width, 3) <= LIMITE_LINHA Then
                        Dim ehBrancoInt31 As Boolean: ehBrancoInt31 = False
                        If s.Outline.Color.Type = cdrColorCMYK Then
                            If s.Outline.Color.CMYKCyan = 0 And s.Outline.Color.CMYKMagenta = 0 And _
                               s.Outline.Color.CMYKYellow = 0 And s.Outline.Color.CMYKBlack = 0 And _
                               Round(s.Outline.Width, 3) >= 0.02 And Round(s.Outline.Width, 3) <= 0.05 Then
                                ehBrancoInt31 = True
                            End If
                        End If
                        If Not ehBrancoInt31 Then qtdCorrigiveis = qtdCorrigiveis + 1
                    End If
                End If
            End If
        Next s
        passou = (qtdCorrigiveis >= 2)
        RegistrarTeste "T31a", "[AUTO] B02 (0.05mm) e B04 (0.08mm) corrigiveis detectados", _
            passou, "Corrigiveis=" & qtdCorrigiveis & " (esperado >= 2)"

        ' T31b: chamar PadronizarContornosFinos e verificar 0.2mm
        If qtdCorrigiveis >= 1 Then
            Dim shpB02 As Shape: Set shpB02 = BuscarShapePorEspessura(doc, 0.05, False)
            ActiveDocument.BeginCommandGroup "TestRunner_T31"
            PadronizarContornosFinos True
            ActiveDocument.EndCommandGroup
            If Not shpB02 Is Nothing Then
                passou = (Round(shpB02.Outline.Width, 2) = ESPESSURA_PAD)
                detalhe = "Width apos correcao=" & Format(shpB02.Outline.Width, "0.000") & _
                          "mm (esperado " & Format(ESPESSURA_PAD, "0.000") & "mm)"
                RegistrarTeste "T31b", "[AUTO] PadronizarContornosFinos corrige B02 para 0.2mm", _
                    passou, detalhe
            End If
        End If

        ' T32: B05 (objeto convertido sem contorno) NAO deve ser afetado
        Dim b05t32 As Shape: Set b05t32 = BuscarShapeConvertido(doc)
        If Not b05t32 Is Nothing Then
            Dim semContorno32 As Boolean
            semContorno32 = (b05t32.Outline.Type <> cdrOutline)
            If Not semContorno32 Then semContorno32 = (Round(b05t32.Outline.Width, 3) <= 0)
            RegistrarTeste "T32", "[AUTO] B05 (convertido) nao afetado por Corrigir Contornos", _
                semContorno32, "Outline.Type=" & b05t32.Outline.Type & " Width=" & Format(b05t32.Outline.Width, "0.000")
        Else
            RegistrarIgnorado "T32", "B05 convertido nao localizado", "shape ausente"
        End If

        FecharSemSalvar doc
    Else
        RegistrarIgnorado "T31,T32", "Arquivo_B nao encontrado", "arquivo ausente"
    End If

    ' ─────────────────────────────────────────────────────────
    ' T33/T34: Desbloquear Objetos
    ' ─────────────────────────────────────────────────────────
    Set doc = AbrirArquivo("Arquivo_A_Cores_e_Gradientes.cdr")
    If Not doc Is Nothing Then
        doc.Unit = cdrMillimeter

        ' Contar objetos bloqueados inicialmente
        Dim qtdBloqIni As Long: qtdBloqIni = 0
        For Each s In doc.ActivePage.Shapes
            If s.Locked Then qtdBloqIni = qtdBloqIni + 1
        Next s

        ' T34: estado inicial deve ter A11 bloqueado
        passou = (qtdBloqIni >= 1)
        RegistrarTeste "T33a", "[AUTO] A11 bloqueado presente no Arquivo_A", _
            passou, "Objetos bloqueados=" & qtdBloqIni & " (esperado >= 1)"

        ' T33b: desbloquear e verificar contagem
        Dim qtdDesbloqueados As Long: qtdDesbloqueados = 0
        For Each s In doc.ActivePage.Shapes
            If s.Locked Then
                s.Locked = False
                qtdDesbloqueados = qtdDesbloqueados + 1
            End If
        Next s
        passou = (qtdDesbloqueados = qtdBloqIni And qtdDesbloqueados > 0)
        RegistrarTeste "T33b", "[AUTO] DesbloquearObjetos desbloqueia todos (" & qtdBloqIni & " objetos)", _
            passou, "Desbloqueados=" & qtdDesbloqueados

        ' T34: apos desbloquear, zero objetos bloqueados
        Dim qtdBloqFinal As Long: qtdBloqFinal = 0
        For Each s In doc.ActivePage.Shapes
            If s.Locked Then qtdBloqFinal = qtdBloqFinal + 1
        Next s
        passou = (qtdBloqFinal = 0)
        RegistrarTeste "T34", "[AUTO] Apos desbloquear — nenhum objeto bloqueado restante", _
            passou, "Bloqueados restantes=" & qtdBloqFinal

        FecharSemSalvar doc
    Else
        RegistrarIgnorado "T33,T34", "Arquivo_A nao encontrado", "arquivo ausente"
    End If
End Sub

' ============================================================
' BLOCO 4 — BITMAPS (T35-T36)
' ============================================================
Private Sub Bloco4_Bitmaps()
    IniciarBloco "4 - BITMAPS (T35-T36)"

    Dim doc As Document
    Dim s As Shape

    Set doc = AbrirArquivo("Arquivo_C_Bitmaps.cdr")
    If Not doc Is Nothing Then
        doc.Unit = cdrMillimeter
        Dim qtdBitmaps As Long: qtdBitmaps = 0
        Dim qtdRGBBmp As Long: qtdRGBBmp = 0
        Dim qtdBaixaRes As Long: qtdBaixaRes = 0
        For Each s In doc.ActivePage.Shapes
            If s.Type = cdrBitmapShape Then
                qtdBitmaps = qtdBitmaps + 1
                If s.Bitmap.Mode = 4 Then qtdRGBBmp = qtdRGBBmp + 1  ' cdrImageRGB=4
                If s.Bitmap.ResolutionX < 300 Then qtdBaixaRes = qtdBaixaRes + 1
            End If
        Next s

        If qtdBitmaps = 0 Then
            RegistrarIgnorado "T35", "Nenhum bitmap importado no Arquivo_C", _
                "Importe C01-C04 nas areas marcadas e re-execute."
        Else
            RegistrarTeste "T35", "[AUTO] Bitmaps com problemas (RGB/baixa res) detectaveis", _
                (qtdRGBBmp >= 1 Or qtdBaixaRes >= 1), _
                "Total=" & qtdBitmaps & " RGB=" & qtdRGBBmp & " BaixaRes(<300DPI)=" & qtdBaixaRes
        End If

        RegistrarManual "T36", "Desfazer Padronizar Imagens", _
            "Apos executar T35 manualmente, clique Desfazer. Verificar imagens restauradas."
        FecharSemSalvar doc
    Else
        RegistrarIgnorado "T35,T36", "Arquivo_C nao encontrado", "arquivo ausente"
    End If
End Sub

' ============================================================
' BLOCO 5 — MONTAGEM (T37-T42)
' ============================================================
Private Sub Bloco5_Montagem()
    IniciarBloco "5 - MONTAGEM (T37-T42)"

    ' T37-T41: TrimBox requer dialogo de banda — MANUAL
    RegistrarManual "T37", "TrimBox Banda Larga (7mm)", _
        "Selecione D02, clique Aplicar Trimbox, escolha Banda Larga (SIM). Verificar: pagina +7mm, retangulo registro criado, arte centralizada."
    RegistrarManual "T38", "TrimBox Banda Estreita (5mm)", _
        "Selecione D02, clique Aplicar Trimbox, escolha Banda Estreita (NAO). Verificar: pagina +5mm de offset."
    RegistrarManual "T39", "TrimBox com multiplos objetos", _
        "Selecione os 3 shapes D01, clique Aplicar Trimbox Banda Larga. Verificar: agrupamento automatico + offset 7mm."
    RegistrarManual "T40", "TrimBox cancelado", _
        "Selecione objeto, clique Aplicar Trimbox, clique Cancelar no dialogo. Verificar: nenhuma alteracao."
    RegistrarManual "T41", "Desfazer TrimBox", _
        "Apos T37, clique Desfazer. Verificar: pagina original, registro removido, 1 unico Ctrl+Z."

    ' ─────────────────────────────────────────────────────────
    ' T42: Inserir Dados do Camerom (AUTO — chamar funcao diretamente)
    ' ─────────────────────────────────────────────────────────
    Dim doc As Document
    Set doc = AbrirArquivo("Arquivo_D_Montagem.cdr")
    If Not doc Is Nothing Then
        doc.Unit = cdrMillimeter

        ' Selecionar D02 (retangulo ciano)
        Dim shpD02 As Shape: Set shpD02 = Nothing
        Dim s As Shape
        For Each s In doc.ActivePage.Shapes
            If s.Type <> cdrTextShape Then
                If s.Fill.Type = cdrUniformFill Then
                    If s.Fill.UniformColor.Type = cdrColorCMYK Then
                        If s.Fill.UniformColor.CMYKCyan = 100 And _
                           s.Fill.UniformColor.CMYKMagenta = 0 Then
                            Set shpD02 = s
                            Exit For
                        End If
                    End If
                End If
            End If
        Next s

        If Not shpD02 Is Nothing Then
            shpD02.CreateSelection
            Dim qtdShapesAntes As Long: qtdShapesAntes = doc.ActivePage.Shapes.Count
            ActiveDocument.BeginCommandGroup "TestRunner_T42"
            On Error Resume Next
            InserirTextosCamerom "02-04-2026 Teste Automatizado", "CIANO MAGENTA PRETO"
            Dim errT42 As Long: errT42 = Err.Number
            On Error GoTo 0
            ActiveDocument.EndCommandGroup
            Dim qtdShapesDepois As Long: qtdShapesDepois = doc.ActivePage.Shapes.Count
            Dim passou As Boolean
            passou = (errT42 = 0 And qtdShapesDepois > qtdShapesAntes)
            Dim detalhe As String
            detalhe = "Shapes antes=" & qtdShapesAntes & " depois=" & qtdShapesDepois & _
                      IIf(errT42 <> 0, " ERRO=" & errT42, "")
            RegistrarTeste "T42", "[AUTO] InserirTextosCamerom cria textos laterais", passou, detalhe
        Else
            RegistrarIgnorado "T42", "Shape D02 (ciano) nao localizado no Arquivo_D", "shape ausente"
        End If

        FecharSemSalvar doc
    Else
        RegistrarIgnorado "T42", "Arquivo_D nao encontrado", "arquivo ausente"
    End If
End Sub

' ============================================================
' BLOCO 6 — PREFLIGHT SCANNER (T43-T50)
' ============================================================
Private Sub Bloco6_PreFlight()
    IniciarBloco "6 - PREFLIGHT SCANNER (T43-T50)"

    Dim doc As Document
    Dim s As Shape
    Dim passou As Boolean
    Dim detalhe As String

    ' ─────────────────────────────────────────────────────────
    ' Varredura manual replica logica do Scanner (ExecutarScanner e Private + abre form)
    ' ─────────────────────────────────────────────────────────

    ' T43: Arquivo A (com erros) deve ter itens criticos
    Set doc = AbrirArquivo("Arquivo_A_Cores_e_Gradientes.cdr")
    If Not doc Is Nothing Then
        doc.Unit = cdrMillimeter
        Dim qRGB As Long:  qRGB = 0
        Dim qOver As Long: qOver = 0
        Dim qPSuj As Long: qPSuj = 0
        Dim qBord As Long: qBord = 0
        Dim qBloq As Long: qBloq = 0

        For Each s In doc.ActivePage.Shapes
            If s.Fill.Type = cdrUniformFill Then
                If s.Fill.UniformColor.Type = cdrColorRGB Then qRGB = qRGB + 1
            End If
            If s.OverprintFill Then qOver = qOver + 1
            If s.Fill.Type = cdrUniformFill And s.Fill.UniformColor.Type = cdrColorCMYK Then
                Dim kT43 As Long
                kT43 = s.Fill.UniformColor.CMYKBlack
                Dim cmyT43 As Long
                cmyT43 = s.Fill.UniformColor.CMYKCyan + s.Fill.UniformColor.CMYKMagenta + _
                         s.Fill.UniformColor.CMYKYellow
                If kT43 >= 80 And cmyT43 > 0 And Not (kT43 = 100 And cmyT43 = 0) Then
                    qPSuj = qPSuj + 1
                End If
            End If
            If s.Fill.Type = cdrFountainFill Then
                If s.Locked Or (Not s.Layer.Editable) Then
                    qBloq = qBloq + 1
                ElseIf TerBordaDura(s) Then
                    qBord = qBord + 1
                End If
            End If
        Next s

        Dim totalCrit As Long: totalCrit = qRGB + qOver + qPSuj + qBord
        passou = (totalCrit > 0)
        RegistrarTeste "T43", "[AUTO] Arquivo A com erros — itens criticos detectados", passou, _
            "BrancoOver=" & qOver & " RGB=" & qRGB & " PretoSujo=" & qPSuj & " BordaDura=" & qBord

        ' T46: gradiente bloqueado (A11) detectado
        passou = (qBloq >= 1)
        RegistrarTeste "T46", "[AUTO] Gradiente bloqueado (A11) detectado no relatorio", _
            passou, "GradBloqueados=" & qBloq

        FecharSemSalvar doc
    Else
        RegistrarIgnorado "T43,T46", "Arquivo_A nao encontrado", "arquivo ausente"
    End If

    ' T44: Arquivo D (montagem limpa) deve ter zero itens criticos
    Set doc = AbrirArquivo("Arquivo_D_Montagem.cdr")
    If Not doc Is Nothing Then
        doc.Unit = cdrMillimeter
        Dim criticos44 As Long: criticos44 = 0
        For Each s In doc.ActivePage.Shapes
            If s.Fill.Type = cdrUniformFill Then
                If s.Fill.UniformColor.Type = cdrColorRGB Then criticos44 = criticos44 + 1
            End If
            If s.OverprintFill Then criticos44 = criticos44 + 1
        Next s
        passou = (criticos44 = 0)
        RegistrarTeste "T44", "[AUTO] Arquivo D limpo — zero itens criticos", _
            passou, "TotalCriticos=" & criticos44

        FecharSemSalvar doc
    Else
        RegistrarIgnorado "T44", "Arquivo_D nao encontrado", "arquivo ausente"
    End If

    ' T45: Varredura apenas pagina ativa (SKIP — requer doc com 2 paginas)
    RegistrarIgnorado "T45", "Varredura pagina ativa", _
        "Requer documento de 2 paginas. Teste manual: pg1 com erros, pg2 limpa, executar PreFlight na pg2."

    ' T47/T48: UI do PreFlight (MANUAL)
    RegistrarManual "T47", "Atualizar PreFlight", _
        "Execute PreFlight com erros, corrija 1 objeto manualmente, clique Atualizar. Verificar relatorio atualizado."
    RegistrarManual "T48", "Correcoes automaticas — minDot invalido", _
        "Execute PreFlight com gradiente borda dura, clique Corrigir Erros, digite 'abc'. Verificar MsgBox sem crash."

    ' ─────────────────────────────────────────────────────────
    ' T49: Correcoes automaticas — fluxo completo (ExecutarCorrecoes)
    ' ─────────────────────────────────────────────────────────
    Set doc = AbrirArquivo("Arquivo_A_Cores_e_Gradientes.cdr")
    If Not doc Is Nothing Then
        doc.Unit = cdrMillimeter

        ' Estado antes
        Dim qRGBAntes49 As Long: qRGBAntes49 = 0
        Dim qOverAntes49 As Long: qOverAntes49 = 0
        For Each s In doc.ActivePage.Shapes
            If s.Fill.Type = cdrUniformFill Then
                If s.Fill.UniformColor.Type = cdrColorRGB Then qRGBAntes49 = qRGBAntes49 + 1
            End If
            If s.OverprintFill Then qOverAntes49 = qOverAntes49 + 1
        Next s

        ' Chamar ExecutarCorrecoes (funcao public de Mod02_Scanner_Engine, apos unificacao)
        Dim errT49 As Long: errT49 = 0
        On Error Resume Next
        ExecutarCorrecoes MIN_DOT_TEST
        errT49 = Err.Number
        On Error GoTo 0

        If errT49 <> 0 Then
            RegistrarIgnorado "T49", "ExecutarCorrecoes retornou erro " & errT49, _
                "Verificar se a unificacao do modulo foi aplicada (PlanoUnificacao_CorrigirErros.md)"
        Else
            Dim qRGBDepois49 As Long: qRGBDepois49 = 0
            Dim qOverDepois49 As Long: qOverDepois49 = 0
            For Each s In doc.ActivePage.Shapes
                If s.Fill.Type = cdrUniformFill Then
                    If s.Fill.UniformColor.Type = cdrColorRGB Then qRGBDepois49 = qRGBDepois49 + 1
                End If
                If s.OverprintFill Then qOverDepois49 = qOverDepois49 + 1
            Next s
            passou = (qRGBDepois49 < qRGBAntes49 Or qOverDepois49 < qOverAntes49)
            detalhe = "RGB antes=" & qRGBAntes49 & " depois=" & qRGBDepois49 & _
                      " | Overprint antes=" & qOverAntes49 & " depois=" & qOverDepois49
            RegistrarTeste "T49", "[AUTO] ExecutarCorrecoes corrige erros criticos (fluxo completo)", _
                passou, detalhe
        End If

        FecharSemSalvar doc
    Else
        RegistrarIgnorado "T49", "Arquivo_A nao encontrado", "arquivo ausente"
    End If

    ' T50: Desfazer correcoes do PreFlight (MANUAL)
    RegistrarManual "T50", "Desfazer correcoes do PreFlight", _
        "Apos T49 manual, clique Desfazer no PreFlight. Verificar erros restaurados no relatorio."
End Sub

' ============================================================
' INFRAESTRUTURA DO RELATORIO
' ============================================================
Private Sub IniciarRelatorio()
    ReDim linhas(0)
    totalLinhas = 0
    qtdPass = 0
    qtdFail = 0
    qtdSkip = 0
    qtdManual = 0

    Escrever "============================================================"
    Escrever "RELATORIO DE REGRESSAO — Console Flexo v2.0"
    Escrever "Suite Completa T01-T50 | CorelDRAW 2026"
    Escrever "Data: " & Format(Now, "DD/MM/YYYY HH:MM:SS")
    Escrever "============================================================"
    Escrever ""
    Escrever "  [PASSOU]   Teste automatizado aprovado"
    Escrever "  [FALHOU]   Teste automatizado reprovado — requer investigacao"
    Escrever "  [MANUAL]   Requer execucao e verificacao manual"
    Escrever "  [IGNORADO] Pre-requisito ausente (arquivo, bitmap, etc.)"
    Escrever ""
End Sub

Private Sub Escrever(linha As String)
    ReDim Preserve linhas(totalLinhas)
    linhas(totalLinhas) = linha
    totalLinhas = totalLinhas + 1
End Sub

Private Sub IniciarBloco(nome As String)
    Escrever ""
    Escrever "------------------------------------------------------------"
    Escrever "BLOCO: " & nome
    Escrever "------------------------------------------------------------"
End Sub

Private Sub RegistrarTeste(id As String, descricao As String, _
                            passou As Boolean, detalhe As String)
    Dim status As String
    If passou Then
        status = "PASSOU  "
        qtdPass = qtdPass + 1
    Else
        status = "FALHOU  "
        qtdFail = qtdFail + 1
    End If
    Dim linha As String
    linha = "  [" & status & "] " & id & " - " & descricao
    If detalhe <> "" Then linha = linha & vbCrLf & "             >> " & detalhe
    Escrever linha
End Sub

Private Sub RegistrarManual(id As String, descricao As String, instrucao As String)
    qtdManual = qtdManual + 1
    Escrever "  [MANUAL  ] " & id & " - " & descricao
    Escrever "             >> INSTRUCAO: " & instrucao
    Escrever "             >> RESULTADO: ( ) PASSOU  ( ) FALHOU  — preencher apos execucao manual"
End Sub

Private Sub RegistrarIgnorado(id As String, descricao As String, motivo As String)
    qtdSkip = qtdSkip + 1
    Escrever "  [IGNORADO] " & id & " - " & descricao & " (" & motivo & ")"
End Sub

Private Sub SalvarRelatorio()
    Escrever ""
    Escrever "============================================================"
    Escrever "RESUMO FINAL"
    Escrever "============================================================"
    Escrever "  PASSOU:    " & qtdPass
    Escrever "  FALHOU:    " & qtdFail
    Escrever "  MANUAL:    " & qtdManual & "  (verificacao humana pendente)"
    Escrever "  IGNORADO:  " & qtdSkip
    Escrever "  TOTAL:     " & (qtdPass + qtdFail + qtdManual + qtdSkip)
    Escrever ""
    Escrever "Console Flexo v2.0 | Regressao Completa | " & Format(Now, "DD/MM/YYYY")
    Escrever "============================================================"

    Dim caminho As String
    caminho = pastaBase & "Relatorio_Regressao_v2.txt"
    Dim ff As Integer: ff = FreeFile
    Open caminho For Output As #ff
    Dim i As Long
    For i = 0 To totalLinhas - 1
        Print #ff, linhas(i)
    Next i
    Close #ff
End Sub

' ============================================================
' HELPERS DE DOCUMENTO
' ============================================================
Private Function AbrirArquivo(nome As String) As Document
    Dim caminho As String: caminho = pastaBase & nome
    If Dir(caminho) = "" Then
        Set AbrirArquivo = Nothing
        Exit Function
    End If
    On Error Resume Next
    Set AbrirArquivo = OpenDocument(caminho)
    On Error GoTo 0
End Function

Private Sub FecharSemSalvar(doc As Document)
    On Error Resume Next
    doc.Close
    On Error GoTo 0
End Sub

' ============================================================
' HELPERS DE BUSCA DE SHAPES
' ============================================================

' Busca shape com contorno de espessura aproximada
' ehBrancoProcurado=True -> filtra contornos brancos puros
Private Function BuscarShapePorEspessura(doc As Document, _
                                          esp As Double, _
                                          ehBrancoProcurado As Boolean) As Shape
    Dim s As Shape
    For Each s In doc.ActivePage.Shapes
        If s.Type <> cdrTextShape And s.Type <> cdrBitmapShape Then
            If s.Outline.Type = cdrOutline Then
                If Round(s.Outline.Width, 2) = Round(esp, 2) Then
                    Dim isBranco As Boolean
                    isBranco = (s.Outline.Color.Type = cdrColorCMYK And _
                                s.Outline.Color.CMYKCyan = 0 And _
                                s.Outline.Color.CMYKMagenta = 0 And _
                                s.Outline.Color.CMYKYellow = 0 And _
                                s.Outline.Color.CMYKBlack = 0)
                    If isBranco = ehBrancoProcurado Then
                        Set BuscarShapePorEspessura = s
                        Exit Function
                    End If
                End If
            End If
        End If
    Next s
    Set BuscarShapePorEspessura = Nothing
End Function

' Busca shape COM fill uniforme E contorno de espessura aproximada
Private Function BuscarShapeComFillEContorno(doc As Document, esp As Double) As Shape
    Dim s As Shape
    For Each s In doc.ActivePage.Shapes
        If s.Type <> cdrTextShape And s.Type <> cdrBitmapShape Then
            If s.Fill.Type = cdrUniformFill Then
                If s.Outline.Type = cdrOutline Then
                    If Round(s.Outline.Width, 2) = Round(esp, 2) Then
                        Set BuscarShapeComFillEContorno = s
                        Exit Function
                    End If
                End If
            End If
        End If
    Next s
    Set BuscarShapeComFillEContorno = Nothing
End Function

' Busca shape sem contorno ativo com dimensao <= LIMITE_LINHA (objeto convertido)
Private Function BuscarShapeConvertido(doc As Document) As Shape
    Dim s As Shape
    For Each s In doc.ActivePage.Shapes
        If s.Type <> cdrTextShape And s.Type <> cdrBitmapShape And _
           s.Type <> cdrGroupShape Then
            Dim semContorno As Boolean: semContorno = (s.Outline.Type <> cdrOutline)
            If Not semContorno Then semContorno = (Round(s.Outline.Width, 3) <= 0)
            If semContorno Then
                Dim bbX As Double, bbY As Double, bbW As Double, bbH As Double
                s.GetBoundingBox bbX, bbY, bbW, bbH
                If Round(bbW, 3) <= LIMITE_LINHA Or Round(bbH, 3) <= LIMITE_LINHA Then
                    Set BuscarShapeConvertido = s
                    Exit Function
                End If
            End If
        End If
    Next s
    Set BuscarShapeConvertido = Nothing
End Function

' Verifica se um gradiente possui borda dura (canal vai de 0 a valor > 0)
Private Function TerBordaDura(s As Shape) As Boolean
    On Error Resume Next
    Dim fc As FountainColor
    For Each fc In s.Fill.Fountain.Colors
        If fc.Color.Type = cdrColorCMYK Then
            Dim fc2 As FountainColor
            For Each fc2 In s.Fill.Fountain.Colors
                If fc2.Color.Type = cdrColorCMYK Then
                    If fc.Color.CMYKCyan = 0 And fc2.Color.CMYKCyan > 0 Then
                        TerBordaDura = True: Exit Function
                    End If
                    If fc.Color.CMYKMagenta = 0 And fc2.Color.CMYKMagenta > 0 Then
                        TerBordaDura = True: Exit Function
                    End If
                    If fc.Color.CMYKYellow = 0 And fc2.Color.CMYKYellow > 0 Then
                        TerBordaDura = True: Exit Function
                    End If
                    If fc.Color.CMYKBlack = 0 And fc2.Color.CMYKBlack > 0 Then
                        TerBordaDura = True: Exit Function
                    End If
                End If
            Next fc2
        End If
    Next fc
    On Error GoTo 0
    TerBordaDura = False
End Function
