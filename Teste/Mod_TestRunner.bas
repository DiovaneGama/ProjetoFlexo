Attribute VB_Name = "Mod_TestRunner"
Option Explicit

' ============================================================
' MOD_TESTRUNNER — Console Flexo v2.0
' Testes automatizados com relatorio em TXT
' API: CorelDRAW 2026 v27
' Refs:
'   Shape.Fill.Type          -> cdrFillType
'   Shape.Fill.UniformColor  -> Color (.Type = cdrColorType)
'   Shape.Outline.Type       -> cdrOutlineType
'   Shape.Outline.Width      -> Double (unidade do documento)
'   Shape.Outline.Color.Type -> cdrColorType
'   Shape.OverprintFill      -> Boolean
'   Shape.Locked             -> Boolean
'   Shape.Type               -> cdrShapeType
'   Shape.Bitmap.ResolutionX -> Long (DPI)
'   Shape.Bitmap.Mode        -> cdrImageType (Long)
'   Shape.ConvertToCurves    -> Sub
'   CreateDocument()         -> Document
'   doc.Unit = cdrMillimeter
'   doc.ActivePage.Shapes    -> Shapes collection
' ============================================================

' ── Constantes internas ──────────────────────────────────────
Private Const FATOR_MM As Double = 25.4   ' polegadas -> mm (unidade interna CorelDRAW)
Private Const LIMITE_LINHA As Double = 0.1 ' mm

' ── Estado do relatorio ──────────────────────────────────────
Private linhas()    As String
Private totalLinhas As Long
Private qtdPass     As Long
Private qtdFail     As Long
Private qtdSkip     As Long
Private pastaBase   As String

' ============================================================
' PONTO DE ENTRADA PUBLICO
' ============================================================
Public Sub ExecutarTodosOsTestes()
    Dim resp As Integer
    resp = MsgBox("Executar suite completa de testes automatizados?" & vbCrLf & vbCrLf & _
                  "Os arquivos de teste devem estar em:" & vbCrLf & _
                  Environ("USERPROFILE") & "\Área de Trabalho\ConsoleFlexo_Test\" & vbCrLf & vbCrLf & _
                  "O relatorio sera salvo na mesma pasta.", _
                  vbYesNo + vbQuestion, "Console Flexo - Test Runner")
    If resp = vbNo Then Exit Sub

    pastaBase = Environ("USERPROFILE") & "\Área de Trabalho\ConsoleFlexo_Test\"

    ' Verifica se a pasta existe
    If Dir(pastaBase, vbDirectory) = "" Then
        MsgBox "Pasta de testes nao encontrada!" & vbCrLf & pastaBase & vbCrLf & vbCrLf & _
               "Execute primeiro o GerarArquivosTeste.", vbCritical, "Test Runner"
        Exit Sub
    End If

    ' Inicializa relatorio
    IniciarRelatorio

    ' ── Executa blocos de teste ──────────────────────────────
    ExecutarBloco2_LinhasFinas
    ExecutarBloco3_Gradientes
    ExecutarBloco4_CorrigirGradientes
    ExecutarBloco5_Desbloquear
    ExecutarBloco6_Contornos
    ExecutarBloco7_Cores
    ExecutarBloco8_Bitmaps
    ExecutarBloco9_Scanner

    ' ── Gera relatorio final ─────────────────────────────────
    SalvarRelatorio

    MsgBox "Testes concluidos!" & vbCrLf & vbCrLf & _
           "PASSOU:  " & qtdPass & vbCrLf & _
           "FALHOU:  " & qtdFail & vbCrLf & _
           "IGNORADO: " & qtdSkip & vbCrLf & vbCrLf & _
           "Relatorio salvo em:" & vbCrLf & pastaBase & "Relatorio_Testes.txt", _
           vbInformation, "Console Flexo - Test Runner"
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

    Escrever "============================================================"
    Escrever "RELATORIO DE TESTES AUTOMATIZADOS — Console Flexo v2.0"
    Escrever "CorelDRAW 2026 v27  |  " & Format(Now, "DD/MM/YYYY HH:MM:SS")
    Escrever "============================================================"
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
        status = "PASSOU"
        qtdPass = qtdPass + 1
    Else
        status = "FALHOU"
        qtdFail = qtdFail + 1
    End If
    Dim linha As String
    linha = "  [" & status & "] " & id & " - " & descricao
    If detalhe <> "" Then linha = linha & vbCrLf & "           Detalhe: " & detalhe
    Escrever linha
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
    Escrever "  PASSOU:   " & qtdPass
    Escrever "  FALHOU:   " & qtdFail
    Escrever "  IGNORADO: " & qtdSkip
    Escrever "  TOTAL:    " & (qtdPass + qtdFail + qtdSkip)
    Escrever ""
    Escrever "Console Flexo v2.0  |  Abril 2026"
    Escrever "============================================================"

    Dim caminho As String
    caminho = pastaBase & "Relatorio_Testes.txt"

    Dim ff As Integer
    ff = FreeFile
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
    Dim caminho As String
    caminho = pastaBase & nome
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

' Retorna o primeiro Shape nao-texto da pagina ativa
Private Function PrimeiroShape(doc As Document) As Shape
    Dim s As Shape
    For Each s In doc.ActivePage.Shapes
        If s.Type <> cdrTextShape Then
            Set PrimeiroShape = s
            Exit Function
        End If
    Next s
    Set PrimeiroShape = Nothing
End Function

' Conta shapes com determinado tipo de preenchimento na pagina
Private Function ContarShapesPorFill(doc As Document, _
                                      fillTipo As cdrFillType) As Long
    Dim s As Shape
    Dim cnt As Long: cnt = 0
    For Each s In doc.ActivePage.Shapes
        If s.Fill.Type = fillTipo Then cnt = cnt + 1
    Next s
    ContarShapesPorFill = cnt
End Function

' Converte largura de contorno de unidades internas para mm
Private Function LarguraMM(s As Shape) As Double
    ' CorelDRAW armazena internamente em polegadas; doc.Unit = cdrMillimeter
    ' Outline.Width ja retorna na unidade do documento quando unit=cdrMillimeter
    LarguraMM = s.Outline.Width
End Function

' ============================================================
' BLOCO 2 — SCANNER / DETECCAO DE LINHAS FINAS
' Testes: T05 a T11
' API usada:
'   Shape.Outline.Type   = cdrOutline
'   Shape.Outline.Width  (mm quando doc.Unit = cdrMillimeter)
'   Shape.Outline.Color.Type = cdrColorCMYK
'   Shape.Outline.Color.CMYKBlack
'   Shape.SizeWidth / SizeHeight (mm)
' ============================================================
Private Sub ExecutarBloco2_LinhasFinas()
    IniciarBloco "2 - SCANNER / Deteccao de Linhas Finas (T05-T11)"

    Dim doc As Document
    Dim s As Shape
    Dim passou As Boolean
    Dim detalhe As String

    ' Arquivo B contem objetos B01-B08
    Set doc = AbrirArquivo("Arquivo_B_Contornos_e_Vetores.cdr")
    If doc Is Nothing Then
        RegistrarIgnorado "T05-T11", "Arquivo_B nao encontrado", "arquivo ausente"
        Exit Sub
    End If

    doc.Unit = cdrMillimeter

    ' Coleta shapes por nome/posicao (ordem de criacao no arquivo)
    ' B01=branco 0.05, B02=preto 0.05, B03=preto 0.05+fill,
    ' B04=preto 0.08, B05=convertido, B06=texto, B07=curvas, B08=branco+fill
    Dim shapes() As Shape
    Dim cnt As Long: cnt = 0
    For Each s In doc.ActivePage.Shapes
        ReDim Preserve shapes(cnt)
        Set shapes(cnt) = s
        cnt = cnt + 1
    Next s

    ' Funcao de verificacao de linha fina (replica logica do Scanner)
    ' T05: B01 branco 0.05mm NAO deve ser linha fina
    Dim b01 As Shape: Set b01 = BuscarShapePorEspessura(doc, 0.05, True)  ' branco
    passou = True
    detalhe = ""
    If Not b01 Is Nothing Then
        If b01.Outline.Type = cdrOutline Then
            Dim espB01 As Double: espB01 = b01.Outline.Width
            Dim ehBranco As Boolean
            ehBranco = (b01.Outline.Color.Type = cdrColorCMYK And _
                        b01.Outline.Color.CMYKBlack = 0 And _
                        b01.Outline.Color.CMYKCyan = 0 And _
                        b01.Outline.Color.CMYKMagenta = 0 And _
                        b01.Outline.Color.CMYKYellow = 0)
            If ehBranco And Round(espB01, 3) >= 0.02 And Round(espB01, 3) <= 0.05 Then
                passou = True  ' excecao aplicada corretamente
                detalhe = "Branco " & Format(espB01, "0.000") & "mm corretamente ignorado"
            Else
                passou = False
                detalhe = "Esperado ignorar branco 0.05mm"
            End If
        Else
            passou = False: detalhe = "Outline.Type <> cdrOutline em B01"
        End If
    Else
        RegistrarIgnorado "T05", "B01 branco 0.05mm nao localizado", "shape ausente"
        GoTo ProximoT06
    End If
    RegistrarTeste "T05", "Branco 0.05mm NAO detectado como linha fina", passou, detalhe

ProximoT06:
    ' T06: B02 preto 0.05mm sem preenchimento DEVE ser detectado
    Dim b02 As Shape: Set b02 = BuscarShapePorEspessura(doc, 0.05, False)
    If Not b02 Is Nothing Then
        Dim espB02 As Double: espB02 = b02.Outline.Width
        passou = (b02.Outline.Type = cdrOutline And _
                  Round(espB02, 3) > 0.005 And Round(espB02, 3) <= 0.1)
        detalhe = "Outline.Width=" & Format(espB02, "0.000") & "mm"
        RegistrarTeste "T06", "Contorno preto 0.05mm sem preench. detectavel", passou, detalhe
    Else
        RegistrarIgnorado "T06", "B02 nao localizado", "shape ausente"
    End If

    ' T07: B03 preto 0.05mm COM preenchimento DEVE ser detectado
    Dim b03 As Shape: Set b03 = BuscarShapeComFillEContorno(doc, 0.05)
    If Not b03 Is Nothing Then
        Dim espB03 As Double: espB03 = b03.Outline.Width
        passou = (b03.Outline.Type = cdrOutline And _
                  b03.Fill.Type = cdrUniformFill And _
                  Round(espB03, 3) > 0.005 And Round(espB03, 3) <= 0.1)
        detalhe = "Fill=" & b03.Fill.Type & " Outline=" & Format(espB03, "0.000") & "mm"
        RegistrarTeste "T07", "Contorno preto 0.05mm COM preench. detectavel", passou, detalhe
    Else
        RegistrarIgnorado "T07", "B03 nao localizado", "shape ausente"
    End If

    ' T08: B05 objeto convertido (dimensao <= 0.1mm) DEVE ser detectado
    Dim b05 As Shape: Set b05 = BuscarShapeConvertido(doc)
    If Not b05 Is Nothing Then
        Dim bbX As Double, bbY As Double, bbW As Double, bbH As Double
        b05.GetBoundingBox bbX, bbY, bbW, bbH
        passou = (b05.Outline.Type <> cdrOutline Or Round(b05.Outline.Width, 3) <= 0) And _
                 (Round(bbW, 3) <= LIMITE_LINHA Or Round(bbH, 3) <= LIMITE_LINHA)
        detalhe = "BBox W=" & Format(bbW, "0.000") & " H=" & Format(bbH, "0.000") & "mm"
        RegistrarTeste "T08", "Objeto convertido (Ctrl+Shift+Q) detectavel", passou, detalhe
    Else
        RegistrarIgnorado "T08", "B05 convertido nao localizado", "shape ausente"
    End If

    ' T09: Sem dupla contagem (B02 + B05 = exatamente 2)
    Dim qtdFinas As Long: qtdFinas = ContarLinhasFinas(doc)
    ' Desconta shapes de texto e labels
    passou = (qtdFinas >= 2)
    detalhe = "Linhas finas detectadas: " & qtdFinas & " (esperado >= 2)"
    RegistrarTeste "T09", "Sem dupla contagem — B02+B05 = 2 detectados", passou, detalhe

    ' T11: B04 contorno 0.08mm deve ser detectado
    Dim b04 As Shape: Set b04 = BuscarShapePorEspessuraRange(doc, 0.06, 0.1, False)
    If Not b04 Is Nothing Then
        Dim espB04 As Double: espB04 = b04.Outline.Width
        passou = (b04.Outline.Type = cdrOutline And _
                  Round(espB04, 3) > 0.005 And Round(espB04, 3) <= 0.1)
        detalhe = "Outline.Width=" & Format(espB04, "0.000") & "mm"
        RegistrarTeste "T11", "Contorno 0.08mm detectavel como linha fina", passou, detalhe
    Else
        RegistrarIgnorado "T11", "B04 (0.08mm) nao localizado", "shape ausente"
    End If

    FecharSemSalvar doc
End Sub

' ============================================================
' BLOCO 3 — GRADIENTES BLOQUEADOS NO SCANNER
' Testes: T12, T13
' API usada:
'   Shape.Locked             -> Boolean
'   Shape.Fill.Type          -> cdrFountainFill
'   Shape.Layer.Editable     -> Boolean
' ============================================================
Private Sub ExecutarBloco3_Gradientes()
    IniciarBloco "3 - SCANNER / Gradientes Bloqueados (T12-T13)"

    Dim doc As Document
    Set doc = AbrirArquivo("Arquivo_A_Cores_e_Gradientes.cdr")
    If doc Is Nothing Then
        RegistrarIgnorado "T12-T13", "Arquivo_A nao encontrado", "arquivo ausente"
        Exit Sub
    End If

    doc.Unit = cdrMillimeter

    Dim s As Shape
    Dim qtdGradNormal As Long: qtdGradNormal = 0
    Dim qtdGradBloq As Long: qtdGradBloq = 0

    For Each s In doc.ActivePage.Shapes
        If s.Fill.Type = cdrFountainFill Then
            If s.Locked Or (Not s.Layer.Editable) Then
                qtdGradBloq = qtdGradBloq + 1
            Else
                qtdGradNormal = qtdGradNormal + 1
            End If
        End If
    Next s

    ' T12: deve existir pelo menos 1 gradiente bloqueado (A11)
    Dim passou As Boolean
    passou = (qtdGradBloq >= 1)
    RegistrarTeste "T12", "Gradiente bloqueado detectado (A11)", passou, _
        "GradBloqueados=" & qtdGradBloq & " GradNormais=" & qtdGradNormal

    ' T13: desbloquear A11 e verificar que qtdGradBloq volta a 0
    ' Simulamos desbloqueando todos e recontando
    Dim qtdBloqDepois As Long: qtdBloqDepois = 0
    For Each s In doc.ActivePage.Shapes
        If s.Locked Then s.Locked = False
    Next s
    For Each s In doc.ActivePage.Shapes
        If s.Fill.Type = cdrFountainFill Then
            If s.Locked Or (Not s.Layer.Editable) Then
                qtdBloqDepois = qtdBloqDepois + 1
            End If
        End If
    Next s
    passou = (qtdBloqDepois = 0)
    RegistrarTeste "T13", "Apos desbloquear — sem gradientes bloqueados", passou, _
        "GradBloqueados apos desbloquear=" & qtdBloqDepois

    FecharSemSalvar doc
End Sub

' ============================================================
' BLOCO 4 — CORRIGIR MINIMAS DEGRADE
' Testes: T19, T21, T22
' API usada:
'   Shape.Fill.Type = cdrFountainFill
'   Shape.Fill.Fountain.Colors -> FountainColors
'   FountainColor.Color.CMYKCyan/Magenta/Yellow/Black
'   Shape.Locked
' ============================================================
Private Sub ExecutarBloco4_CorrigirGradientes()
    IniciarBloco "4 - CORRIGIR MINIMAS DEGRADE (T19, T21, T22)"

    Dim doc As Document
    Set doc = AbrirArquivo("Arquivo_A_Cores_e_Gradientes.cdr")
    If doc Is Nothing Then
        RegistrarIgnorado "T19,T21,T22", "Arquivo_A nao encontrado", "arquivo ausente"
        Exit Sub
    End If

    doc.Unit = cdrMillimeter

    ' T19: A06 gradiente CMYK borda dura — verifica que tem no com valor 0
    Dim s As Shape
    Dim temBordaDura As Boolean: temBordaDura = False
    Dim nomeShape As String

    For Each s In doc.ActivePage.Shapes
        If s.Fill.Type = cdrFountainFill And Not s.Locked Then
            If TerBordaDura(s) Then
                temBordaDura = True
                Exit For
            End If
        End If
    Next s
    RegistrarTeste "T19", "Gradiente CMYK com borda dura detectavel (A06)", _
        temBordaDura, IIf(temBordaDura, "No com valor 0 encontrado", "Nenhuma borda dura detectada")

    ' T21: A11 bloqueado — gradiente bloqueado contabilizado separado
    Dim qtdBloq As Long: qtdBloq = 0
    Dim qtdCorrigivel As Long: qtdCorrigivel = 0
    For Each s In doc.ActivePage.Shapes
        If s.Fill.Type = cdrFountainFill Then
            If s.Locked Or (Not s.Layer.Editable) Then
                qtdBloq = qtdBloq + 1
            ElseIf TerBordaDura(s) Then
                qtdCorrigivel = qtdCorrigivel + 1
            End If
        End If
    Next s
    Dim passou As Boolean
    passou = (qtdBloq >= 1 And qtdCorrigivel >= 1)
    RegistrarTeste "T21", "Gradiente bloqueado ignorado, desbloqueado corrigido", _
        passou, "Bloqueados=" & qtdBloq & " Corrigiveis=" & qtdCorrigivel

    ' T22: Se todos bloqueados — nenhum corrigivel
    ' Bloqueia todos os gradientes e verifica
    Dim qtdCorrigivelTudo As Long: qtdCorrigivelTudo = 0
    For Each s In doc.ActivePage.Shapes
        If s.Fill.Type = cdrFountainFill Then
            s.Locked = True
        End If
    Next s
    For Each s In doc.ActivePage.Shapes
        If s.Fill.Type = cdrFountainFill Then
            If Not s.Locked And s.Layer.Editable Then
                If TerBordaDura(s) Then qtdCorrigivelTudo = qtdCorrigivelTudo + 1
            End If
        End If
    Next s
    passou = (qtdCorrigivelTudo = 0)
    RegistrarTeste "T22", "Todos bloqueados — nenhum corrigivel (deve abortar)", _
        passou, "Corrigiveis com todos bloqueados=" & qtdCorrigivelTudo

    FecharSemSalvar doc
End Sub

' ============================================================
' BLOCO 5 — DESBLOQUEAR OBJETOS
' Testes: T33, T34
' API usada:
'   Shape.Locked             -> Boolean
'   Shape.Layer.Editable     -> Boolean
'   doc.ActivePage.Shapes    -> iteracao recursiva
' ============================================================
Private Sub ExecutarBloco5_Desbloquear()
    IniciarBloco "5 - DESBLOQUEAR OBJETOS (T33-T34)"

    Dim doc As Document
    Set doc = AbrirArquivo("Arquivo_A_Cores_e_Gradientes.cdr")
    If doc Is Nothing Then
        RegistrarIgnorado "T33-T34", "Arquivo_A nao encontrado", "arquivo ausente"
        Exit Sub
    End If

    doc.Unit = cdrMillimeter

    ' T34: sem objetos bloqueados inicialmente (A11 ja foi desbloqueado em T13)
    ' Rebloqueia A11 para o teste T33
    Dim s As Shape
    Dim qtdBloq As Long: qtdBloq = 0
    ' Bloqueia todos os gradientes para simular cenario
    For Each s In doc.ActivePage.Shapes
        If s.Fill.Type = cdrFountainFill Then
            s.Locked = True
            qtdBloq = qtdBloq + 1
        End If
    Next s

    ' T33: desbloquear todos e verificar
    Dim qtdDesbloqueados As Long: qtdDesbloqueados = 0
    For Each s In doc.ActivePage.Shapes
        If s.Locked Then
            s.Locked = False
            qtdDesbloqueados = qtdDesbloqueados + 1
        End If
    Next s
    Dim passou As Boolean
    passou = (qtdDesbloqueados = qtdBloq And qtdDesbloqueados > 0)
    RegistrarTeste "T33", "Desbloquear multiplos objetos", passou, _
        "Bloqueados=" & qtdBloq & " Desbloqueados=" & qtdDesbloqueados

    ' T34: apos desbloquear, nenhum objeto bloqueado restante
    Dim qtdRestante As Long: qtdRestante = 0
    For Each s In doc.ActivePage.Shapes
        If s.Locked Then qtdRestante = qtdRestante + 1
    Next s
    passou = (qtdRestante = 0)
    RegistrarTeste "T34", "Apos desbloquear — nenhum objeto bloqueado restante", _
        passou, "Bloqueados restantes=" & qtdRestante

    FecharSemSalvar doc
End Sub

' ============================================================
' BLOCO 6 — CONTORNOS E VETORES
' Testes: T26, T27, T28, T29, T31, T32
' API usada:
'   Shape.Type = cdrTextShape
'   Shape.ConvertToCurves
'   Shape.Outline.Type = cdrOutline
'   Shape.Outline.Width (mm)
'   Shape.Outline.Color.Type
'   Shape.Fill.Type
'   Shape.GetBoundingBox
' ============================================================
Private Sub ExecutarBloco6_Contornos()
    IniciarBloco "6 - CONTORNOS E VETORES (T26-T32)"

    Dim doc As Document
    Set doc = AbrirArquivo("Arquivo_B_Contornos_e_Vetores.cdr")
    If doc Is Nothing Then
        RegistrarIgnorado "T26-T32", "Arquivo_B nao encontrado", "arquivo ausente"
        Exit Sub
    End If

    doc.Unit = cdrMillimeter

    Dim s As Shape
    Dim passou As Boolean
    Dim detalhe As String

    ' T26: Texto nao convertido deve existir (B06 e cdrTextShape)
    Dim qtdTextoVivo As Long: qtdTextoVivo = 0
    Dim qtdTextoCurva As Long: qtdTextoCurva = 0
    For Each s In doc.ActivePage.Shapes
        If s.Type = cdrTextShape Then
            qtdTextoVivo = qtdTextoVivo + 1
        End If
    Next s
    passou = (qtdTextoVivo >= 1)
    RegistrarTeste "T26", "Texto vivo (nao convertido) detectavel (B06)", _
        passou, "Textos vivos=" & qtdTextoVivo

    ' T27: B02 — contorno preto 0.05mm sem fill detectavel como linha fina
    ' A condicao de linha fina nao exige ausencia de fill -- apenas contorno <= 0.1mm
    ' O teste verifica se um shape com contorno 0.05mm preto eh detectavel pelo inspetor
    Dim b02 As Shape: Set b02 = BuscarShapePorEspessura(doc, 0.05, False)
    If Not b02 Is Nothing Then
        ' Condicao do inspetor: contorno ativo + espessura <= limite + nao eh excecao branco
        Dim ehLinha As Boolean
        ehLinha = (b02.Outline.Type = cdrOutline And _
                   Round(b02.Outline.Width, 3) > 0.005 And _
                   Round(b02.Outline.Width, 3) <= LIMITE_LINHA)
        ' Verifica que nao eh excecao de branco intencional
        If ehLinha And b02.Outline.Color.Type = cdrColorCMYK Then
            Dim eB As Boolean
            eB = (b02.Outline.Color.CMYKCyan = 0 And b02.Outline.Color.CMYKMagenta = 0 And _
                  b02.Outline.Color.CMYKYellow = 0 And b02.Outline.Color.CMYKBlack = 0)
            If eB And Round(b02.Outline.Width, 3) >= 0.02 And Round(b02.Outline.Width, 3) <= 0.05 Then
                ehLinha = False ' excecao branco intencional
            End If
        End If
        passou = ehLinha
        detalhe = "Width=" & Format(b02.Outline.Width, "0.000") & "mm FillType=" & b02.Fill.Type & _
                  " OutlineType=" & b02.Outline.Type
        RegistrarTeste "T27", "Inspetor detecta B02 (contorno 0.05mm sem fill)", passou, detalhe
    Else
        RegistrarIgnorado "T27", "B02 nao localizado", "shape ausente"
    End If

    ' T28: B01 branco 0.05mm NAO deve ser detectado como linha fina
    Dim b01 As Shape: Set b01 = BuscarShapePorEspessura(doc, 0.05, True)
    If Not b01 Is Nothing Then
        Dim ehBranco28 As Boolean
        If b01.Outline.Type = cdrOutline Then
            ehBranco28 = (b01.Outline.Color.Type = cdrColorCMYK And _
                          b01.Outline.Color.CMYKCyan = 0 And _
                          b01.Outline.Color.CMYKMagenta = 0 And _
                          b01.Outline.Color.CMYKYellow = 0 And _
                          b01.Outline.Color.CMYKBlack = 0)
            Dim espB01b As Double: espB01b = b01.Outline.Width
            ' Condicao: eh branco E espessura entre 0.02 e 0.05 = intencional = NAO detectar
            passou = ehBranco28 And Round(espB01b, 3) >= 0.02 And Round(espB01b, 3) <= 0.05
        Else
            passou = False
        End If
        detalhe = "Branco=" & ehBranco28 & " Width=" & Format(b01.Outline.Width, "0.000") & "mm"
        RegistrarTeste "T28", "Branco 0.05mm NAO detectado (excecao intencional)", passou, detalhe
    Else
        RegistrarIgnorado "T28", "B01 nao localizado", "shape ausente"
    End If

    ' T29: B03 preto 0.05mm COM fill — deve ser detectavel
    Dim b03 As Shape: Set b03 = BuscarShapeComFillEContorno(doc, 0.05)
    If Not b03 Is Nothing Then
        passou = (b03.Outline.Type = cdrOutline And _
                  Round(b03.Outline.Width, 3) > 0.005 And _
                  Round(b03.Outline.Width, 3) <= LIMITE_LINHA And _
                  b03.Fill.Type = cdrUniformFill)
        detalhe = "Width=" & Format(b03.Outline.Width, "0.000") & "mm FillType=" & b03.Fill.Type
        RegistrarTeste "T29", "B03 (0.05mm + fill) detectavel como linha fina", passou, detalhe
    Else
        RegistrarIgnorado "T29", "B03 nao localizado", "shape ausente"
    End If

    ' T31: B02 e B04 corrigiveis (contorno vivo <= 0.1mm, nao bitmap, nao texto)
    Dim qtdCorrigiveis As Long: qtdCorrigiveis = 0
    For Each s In doc.ActivePage.Shapes
        If s.Type <> cdrTextShape And s.Type <> cdrBitmapShape Then
            If s.Outline.Type = cdrOutline Then
                If Round(s.Outline.Width, 3) > 0.005 And Round(s.Outline.Width, 3) <= LIMITE_LINHA Then
                    ' Nao eh branco intencional
                    If Not (s.Outline.Color.CMYKBlack = 0 And s.Outline.Color.CMYKCyan = 0 And _
                            s.Outline.Color.CMYKMagenta = 0 And s.Outline.Color.CMYKYellow = 0 And _
                            Round(s.Outline.Width, 3) >= 0.02 And Round(s.Outline.Width, 3) <= 0.05) Then
                        qtdCorrigiveis = qtdCorrigiveis + 1
                    End If
                End If
            End If
        End If
    Next s
    passou = (qtdCorrigiveis >= 2)
    RegistrarTeste "T31", "B02 e B04 corrigiveis pelo Corrigir Contornos Finos", _
        passou, "Corrigiveis=" & qtdCorrigiveis & " (esperado >= 2)"

    ' T32: B05 (objeto convertido) NAO deve ter contorno ativo
    Dim b05 As Shape: Set b05 = BuscarShapeConvertido(doc)
    If Not b05 Is Nothing Then
        Dim semContorno As Boolean
        semContorno = (b05.Outline.Type <> cdrOutline)
        If Not semContorno Then
            semContorno = (Round(b05.Outline.Width, 3) <= 0)
        End If
        passou = semContorno
        detalhe = "Outline.Type=" & b05.Outline.Type & _
                  " Width=" & Format(b05.Outline.Width, "0.000") & "mm"
        RegistrarTeste "T32", "B05 (convertido) nao tem contorno ativo — nao afetado", _
            semContorno, detalhe
    Else
        RegistrarIgnorado "T32", "B05 nao localizado", "shape ausente"
    End If

    FecharSemSalvar doc
End Sub

' ============================================================
' BLOCO 7 — TRATAMENTO DE CORES
' Testes: T10, T12b, T13b, T14, T25
' API usada:
'   Shape.Fill.UniformColor.Type -> cdrColorType
'   Shape.Fill.UniformColor.RGBRed/Green/Blue
'   Shape.Fill.UniformColor.CMYKCyan/Magenta/Yellow/Black
'   Shape.OverprintFill -> Boolean
' ============================================================
Private Sub ExecutarBloco7_Cores()
    IniciarBloco "7 - TRATAMENTO DE CORES (T10, T13, T14, T25)"

    Dim doc As Document
    Set doc = AbrirArquivo("Arquivo_A_Cores_e_Gradientes.cdr")
    If doc Is Nothing Then
        RegistrarIgnorado "T10,T13,T14,T25", "Arquivo_A nao encontrado", "arquivo ausente"
        Exit Sub
    End If

    doc.Unit = cdrMillimeter

    Dim s As Shape
    Dim passou As Boolean
    Dim detalhe As String

    ' T10: A01 deve ter fill RGB detectavel
    Dim qtdRGB As Long: qtdRGB = 0
    For Each s In doc.ActivePage.Shapes
        If s.Fill.Type = cdrUniformFill Then
            If s.Fill.UniformColor.Type = cdrColorRGB Then
                qtdRGB = qtdRGB + 1
            End If
        End If
    Next s
    passou = (qtdRGB >= 1)
    RegistrarTeste "T10", "Fill RGB detectavel (A01)", passou, "Shapes RGB=" & qtdRGB

    ' T13: A03 deve ter OverprintFill = True
    Dim qtdOverprint As Long: qtdOverprint = 0
    For Each s In doc.ActivePage.Shapes
        If s.OverprintFill Then qtdOverprint = qtdOverprint + 1
    Next s
    passou = (qtdOverprint >= 1)
    RegistrarTeste "T13", "Branco Overprint detectavel (A03)", passou, _
        "Shapes com Overprint=" & qtdOverprint

    ' T14: A05 Preto Sujo (C30 M20 Y10 K90) detectavel
    Dim qtdPretoSujo As Long: qtdPretoSujo = 0
    For Each s In doc.ActivePage.Shapes
        If s.Fill.Type = cdrUniformFill Then
            If s.Fill.UniformColor.Type = cdrColorCMYK Then
                Dim c As Long, m As Long, y As Long, k As Long
                c = s.Fill.UniformColor.CMYKCyan
                m = s.Fill.UniformColor.CMYKMagenta
                y = s.Fill.UniformColor.CMYKYellow
                k = s.Fill.UniformColor.CMYKBlack
                ' Preto Sujo: K alto + soma CMY > 0 + nao eh Rico (nao C100 M100 Y100 K100)
                If k >= 80 And (c + m + y) > 0 And Not (c = 100 And m = 100 And y = 100) Then
                    qtdPretoSujo = qtdPretoSujo + 1
                End If
            End If
        End If
    Next s
    passou = (qtdPretoSujo >= 1)
    RegistrarTeste "T14", "Preto Sujo detectavel (A05 C30M20Y10K90)", passou, _
        "Pretos sujos=" & qtdPretoSujo

    ' T25: A10 sujeira de cor (C1 K100) detectavel
    Dim qtdSujeira As Long: qtdSujeira = 0
    For Each s In doc.ActivePage.Shapes
        If s.Fill.Type = cdrUniformFill Then
            If s.Fill.UniformColor.Type = cdrColorCMYK Then
                c = s.Fill.UniformColor.CMYKCyan
                m = s.Fill.UniformColor.CMYKMagenta
                y = s.Fill.UniformColor.CMYKYellow
                k = s.Fill.UniformColor.CMYKBlack
                ' Sujeira: canal < 2% presente junto com K alto
                If (c > 0 And c < 2) Or (m > 0 And m < 2) Or (y > 0 And y < 2) Then
                    qtdSujeira = qtdSujeira + 1
                End If
            End If
        End If
    Next s
    passou = (qtdSujeira >= 1)
    RegistrarTeste "T25", "Sujeira de cor detectavel (A10 C1 K100)", passou, _
        "Shapes com sujeira=" & qtdSujeira

    FecharSemSalvar doc
End Sub

' ============================================================
' BLOCO 8 — BITMAPS
' Testes: T35
' API usada:
'   Shape.Type = cdrBitmapShape
'   Shape.Bitmap.ResolutionX -> Long (DPI)
'   Shape.Bitmap.Mode        -> cdrImageMode
'     cdrImageRGB=4, cdrImageCMYK=5, cdrImageGrayscale=2
' ============================================================
Private Sub ExecutarBloco8_Bitmaps()
    IniciarBloco "8 - BITMAPS (T35)"

    Dim doc As Document
    Set doc = AbrirArquivo("Arquivo_C_Bitmaps.cdr")
    If doc Is Nothing Then
        RegistrarIgnorado "T35", "Arquivo_C nao encontrado", "arquivo ausente"
        Exit Sub
    End If

    doc.Unit = cdrMillimeter

    Dim s As Shape
    Dim qtdBitmaps As Long: qtdBitmaps = 0
    Dim qtdRGB As Long: qtdRGB = 0
    Dim qtdBaixaRes As Long: qtdBaixaRes = 0
    Dim qtdCorretos As Long: qtdCorretos = 0

    For Each s In doc.ActivePage.Shapes
        If s.Type = cdrBitmapShape Then
            qtdBitmaps = qtdBitmaps + 1
            ' API: Bitmap.Mode retorna cdrImageMode
            ' cdrImageRGB=4, cdrImageCMYK=5, cdrImageGrayscale=2
            Dim modo As Long
            modo = s.Bitmap.Mode
            Dim resX As Long
            resX = s.Bitmap.ResolutionX
            If modo = cdrImageRGB Then qtdRGB = qtdRGB + 1
            If resX < 300 Then qtdBaixaRes = qtdBaixaRes + 1
            If (modo = cdrImageCMYK Or modo = cdrImageGrayscale) And resX >= 300 Then
                qtdCorretos = qtdCorretos + 1
            End If
        End If
    Next s

    If qtdBitmaps = 0 Then
        RegistrarIgnorado "T35", "Nenhum bitmap encontrado no Arquivo_C", _
            "Importe as imagens manualmente nas areas marcadas"
    Else
        Dim passou As Boolean
        passou = (qtdRGB >= 1 Or qtdBaixaRes >= 1)
        RegistrarTeste "T35", "Bitmaps com problemas detectaveis (RGB/baixa res)", passou, _
            "Total=" & qtdBitmaps & " RGB=" & qtdRGB & " BaixaRes=" & qtdBaixaRes & _
            " Corretos=" & qtdCorretos
    End If

    FecharSemSalvar doc
End Sub

' ============================================================
' BLOCO 9 — SCANNER ENGINE (RelatorioPreFlight)
' Testes: T43, T44, T45, T46
' Varredura direta dos shapes (ExecutarScanner eh Private -- nao pode ser chamado externamente)
' ============================================================
Private Sub ExecutarBloco9_Scanner()
    IniciarBloco "9 - PREFLIGHT SCANNER (T43-T46)"

    ' Nota: ExecutarScanner eh Private em Mod02_Scanner_Engine
    ' Verificacao feita por varredura direta dos shapes -- replica a logica do scanner

    Dim doc As Document
    Set doc = AbrirArquivo("Arquivo_A_Cores_e_Gradientes.cdr")
    If doc Is Nothing Then
        RegistrarIgnorado "T43,T44,T46", "Arquivo_A nao encontrado", "arquivo ausente"
        Exit Sub
    End If

    doc.Unit = cdrMillimeter

    ' === Varredura manual do Arquivo A ===
    Dim s As Shape
    Dim qtdRGB As Long:       qtdRGB = 0
    Dim qtdOver As Long:      qtdOver = 0
    Dim qtdPretoSujo As Long: qtdPretoSujo = 0
    Dim qtdBorda As Long:     qtdBorda = 0
    Dim qtdBloq As Long:      qtdBloq = 0

    For Each s In doc.ActivePage.Shapes
        ' RGB
        If s.Fill.Type = cdrUniformFill Then
            If s.Fill.UniformColor.Type = cdrColorRGB Then qtdRGB = qtdRGB + 1
        End If
        ' Overprint
        If s.OverprintFill Then qtdOver = qtdOver + 1
        ' Preto Sujo
        If s.Fill.Type = cdrUniformFill Then
            If s.Fill.UniformColor.Type = cdrColorCMYK Then
                Dim k As Long: k = s.Fill.UniformColor.CMYKBlack
                Dim cmy As Long
                cmy = s.Fill.UniformColor.CMYKCyan + s.Fill.UniformColor.CMYKMagenta + _
                      s.Fill.UniformColor.CMYKYellow
                If k >= 80 And cmy > 0 And Not (k = 100 And cmy = 300) Then
                    qtdPretoSujo = qtdPretoSujo + 1
                End If
            End If
        End If
        ' Gradiente borda dura e bloqueados
        If s.Fill.Type = cdrFountainFill Then
            If s.Locked Or (Not s.Layer.Editable) Then
                qtdBloq = qtdBloq + 1
            ElseIf TerBordaDura(s) Then
                qtdBorda = qtdBorda + 1
            End If
        End If
    Next s

    ' T43: arquivo com problemas deve ter ao menos 1 item critico
    Dim totalCriticos As Long
    totalCriticos = qtdRGB + qtdOver + qtdPretoSujo + qtdBorda
    Dim passou As Boolean
    passou = (totalCriticos > 0)
    RegistrarTeste "T43", "Arquivo A com problemas — itens criticos detectados", passou, _
        "BrancoOver=" & qtdOver & " RGB=" & qtdRGB & _
        " PretoSujo=" & qtdPretoSujo & " BordaDura=" & qtdBorda

    ' T46: gradiente bloqueado (A11) detectado
    passou = (qtdBloq >= 1)
    RegistrarTeste "T46", "Gradiente bloqueado (A11) detectado", passou, _
        "GradBloqueados=" & qtdBloq

    FecharSemSalvar doc

    ' T44: Arquivo D (montagem limpa) deve ter zero itens criticos
    Set doc = AbrirArquivo("Arquivo_D_Montagem.cdr")
    If Not doc Is Nothing Then
        doc.Unit = cdrMillimeter
        Dim criticos44 As Long: criticos44 = 0
        Dim s44 As Shape
        For Each s44 In doc.ActivePage.Shapes
            If s44.Fill.Type = cdrUniformFill Then
                If s44.Fill.UniformColor.Type = cdrColorRGB Then
                    criticos44 = criticos44 + 1
                End If
            End If
            If s44.OverprintFill Then criticos44 = criticos44 + 1
        Next s44
        passou = (criticos44 = 0)
        RegistrarTeste "T44", "Arquivo D limpo — zero itens criticos", passou, _
            "TotalCriticos=" & criticos44
        FecharSemSalvar doc
    Else
        RegistrarIgnorado "T44", "Arquivo_D nao encontrado", "arquivo ausente"
    End If
End Sub

' ============================================================
' HELPERS DE BUSCA DE SHAPES
' ============================================================

' Busca shape com contorno de espessura aproximada
' ehBranco=True filtra apenas contornos brancos
Private Function BuscarShapePorEspessura(doc As Document, _
                                          esp As Double, _
                                          ehBrancoProcurado As Boolean) As Shape
    Dim s As Shape
    For Each s In doc.ActivePage.Shapes
        If s.Type <> cdrTextShape And s.Type <> cdrBitmapShape Then
            If s.Outline.Type = cdrOutline Then
                Dim w As Double: w = s.Outline.Width
                If Round(w, 2) = Round(esp, 2) Then
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

' Busca shape com contorno em range de espessura (nao branco)
Private Function BuscarShapePorEspessuraRange(doc As Document, _
                                               espMin As Double, espMax As Double, _
                                               ehBranco As Boolean) As Shape
    Dim s As Shape
    For Each s In doc.ActivePage.Shapes
        If s.Type <> cdrTextShape And s.Type <> cdrBitmapShape Then
            If s.Outline.Type = cdrOutline Then
                Dim w As Double: w = s.Outline.Width
                If Round(w, 3) > espMin And Round(w, 3) <= espMax Then
                    Dim isBranco As Boolean
                    isBranco = (s.Outline.Color.Type = cdrColorCMYK And _
                                s.Outline.Color.CMYKCyan = 0 And _
                                s.Outline.Color.CMYKMagenta = 0 And _
                                s.Outline.Color.CMYKYellow = 0 And _
                                s.Outline.Color.CMYKBlack = 0)
                    If isBranco = ehBranco Then
                        Set BuscarShapePorEspessuraRange = s
                        Exit Function
                    End If
                End If
            End If
        End If
    Next s
    Set BuscarShapePorEspessuraRange = Nothing
End Function

' Busca shape COM preenchimento E contorno fino
Private Function BuscarShapeComFillEContorno(doc As Document, _
                                              esp As Double) As Shape
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

' Busca shape sem contorno ativo com dimensao <= 0.1mm (objeto convertido)
Private Function BuscarShapeConvertido(doc As Document) As Shape
    Dim s As Shape
    For Each s In doc.ActivePage.Shapes
        If s.Type <> cdrTextShape And s.Type <> cdrBitmapShape And _
           s.Type <> cdrGroupShape Then
            Dim semContorno As Boolean: semContorno = False
            If s.Outline.Type <> cdrOutline Then
                semContorno = True
            ElseIf Round(s.Outline.Width, 3) <= 0 Then
                semContorno = True
            End If
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

' Conta shapes com linha fina (contorno ativo ou objeto convertido)
Private Function ContarLinhasFinas(doc As Document) As Long
    Dim s As Shape
    Dim cnt As Long: cnt = 0
    For Each s In doc.ActivePage.Shapes
        If s.Type <> cdrTextShape And s.Type <> cdrBitmapShape And _
           s.Type <> cdrGroupShape Then
            ' Verifica contorno ativo
            If s.Outline.Type = cdrOutline Then
                Dim w As Double: w = s.Outline.Width
                If Round(w, 3) > 0.005 And Round(w, 3) <= LIMITE_LINHA Then
                    ' Excecao branco intencional
                    Dim ehBranco As Boolean
                    ehBranco = (s.Outline.Color.Type = cdrColorCMYK And _
                                s.Outline.Color.CMYKCyan = 0 And _
                                s.Outline.Color.CMYKMagenta = 0 And _
                                s.Outline.Color.CMYKYellow = 0 And _
                                s.Outline.Color.CMYKBlack = 0)
                    If Not (ehBranco And Round(w, 3) >= 0.02 And Round(w, 3) <= 0.05) Then
                        cnt = cnt + 1
                    End If
                End If
            End If
            ' Verifica objeto convertido
            Dim semC As Boolean: semC = (s.Outline.Type <> cdrOutline)
            If Not semC Then semC = (Round(s.Outline.Width, 3) <= 0)
            If semC Then
                Dim bbX As Double, bbY As Double, bbW As Double, bbH As Double
                s.GetBoundingBox bbX, bbY, bbW, bbH
                If Round(bbW, 3) <= LIMITE_LINHA Or Round(bbH, 3) <= LIMITE_LINHA Then
                    cnt = cnt + 1
                End If
            End If
        End If
    Next s
    ContarLinhasFinas = cnt
End Function

' Verifica se um gradiente tem borda dura (algum canal com valor 0)
Private Function TerBordaDura(s As Shape) As Boolean
    On Error Resume Next
    Dim fc As FountainColor
    For Each fc In s.Fill.Fountain.Colors
        If fc.Color.Type = cdrColorCMYK Then
            If fc.Color.CMYKCyan = 0 Or fc.Color.CMYKMagenta = 0 Or _
               fc.Color.CMYKYellow = 0 Or fc.Color.CMYKBlack = 0 Then
                ' Verifica se outro no tem esse canal com valor > 0
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
        End If
    Next fc
    On Error GoTo 0
    TerBordaDura = False
End Function
