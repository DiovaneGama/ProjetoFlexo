Attribute VB_Name = "Mod06_TrapMaker"
' =============================================================================
' TRAPMAKER v1.9.0
' CorelDRAW 2026 v27 - API: https://community.coreldraw.com/sdk/api/draw/27
'
' HISTORICO:
'   v1.1.0  Versao original (outra IA)
'   v1.2.0  Correcao ConvertToObject + ApplyUniformFill + CalcLuma unificada
'   v1.3.x  Correcao Overprint (Fill.Overprint -> Shape.OverprintFill)
'           Correcao Move (Shape.Move -> Shape.Layer = trapLayer)
'   v1.4.0  Reescrita estrategia ConvertToObject
'   v1.5.0  Revisao completa contra API v27 (confirmada em docs oficiais):
'           - ConvertToObject() As Shape: retorna o Shape diretamente
'           - ConvertToCurves() obrigatorio ANTES de SetProperties
'             (retangulos/elipses nativos ignoram SetProperties silenciosamente)
'           - Validacao de Outline.Type apos SetProperties detecta falha
'           - Shape.OverprintFill (Boolean r/w) confirmado API v27
'           - Shape.Layer (Layer r/w) confirmado API v27
'           - Fill.ApplyUniformFill(Color) confirmado API v27
'           - Outline.SetNoOutline() confirmado API v27
'           - AplicarTrap agora retorna Boolean indicando sucesso/falha
'           - RunTrapMaker registra falhas de geracao no relatorio
' =============================================================================

Option Explicit

Private Const TRAP_LAYER_NAME  As String = "TRAP"
Private Const LUMA_THRESHOLD   As Double = 0.1
Private Const LOG_OK           As String = "[OK]  "
Private Const LOG_SKIP         As String = "[---] "

' cdrFillType (API v27 enum)
Private Const CDR_UNIFORM_FILL As Integer = 1
Private Const CDR_BITMAP_TYPE  As Integer = 9
Private Const CDR_OLE_TYPE     As Integer = 25

' cdrTriState (API v27 enum)
Private Const CDR_TRUE         As Integer = 1
Private Const CDR_FALSE        As Integer = 0
Private Const CDR_UNDEFINED    As Integer = -1

' cdrOutlineType: 1 = cdrNoOutline (sem contorno)
Private Const CDR_NO_OUTLINE   As Integer = 1

Private Type TrapResult
    TipoTrap    As String
    CorTrapC    As Double
    CorTrapM    As Double
    CorTrapY    As Double
    CorTrapK    As Double
    ValorPt     As Double
    Mensagem    As String
End Type


' =============================================================================
Public Sub DiagnosticarSelecao()
' =============================================================================
    Dim sel As ShapeRange
    Dim s   As Shape
    Dim msg As String
    Dim i   As Integer

    If ActiveDocument Is Nothing Then
        MsgBox "Nenhum documento aberto.", vbExclamation, "TrapMaker"
        Exit Sub
    End If

    Set sel = ActiveSelectionRange
    If sel.Count = 0 Then
        MsgBox "Nenhum objeto selecionado.", vbInformation, "TrapMaker"
        Exit Sub
    End If

    msg = "Objetos selecionados: " & sel.Count & vbCrLf
    msg = msg & String(38, "-") & vbCrLf

    For i = 1 To sel.Count
        Set s = sel(i)
        msg = msg & "Obj " & i & ":"
        msg = msg & "  Type=" & s.Type
        msg = msg & "  CanHaveOutline=" & s.CanHaveOutline
        On Error Resume Next
        msg = msg & "  Fill=" & s.Fill.Type
        msg = msg & "  C=" & Format(s.Fill.UniformColor.CMYKCyan, "0")
        msg = msg & "  M=" & Format(s.Fill.UniformColor.CMYKMagenta, "0")
        msg = msg & "  Y=" & Format(s.Fill.UniformColor.CMYKYellow, "0")
        msg = msg & "  K=" & Format(s.Fill.UniformColor.CMYKBlack, "0")
        On Error GoTo 0
        msg = msg & vbCrLf
    Next i

    MsgBox msg, vbInformation, "TrapMaker - Diagnostico"
End Sub


' =============================================================================
Public Sub RunTrapMaker()
' =============================================================================
    Dim sel              As ShapeRange
    Dim trapLayer        As Layer
    Dim trapValorMm      As Double
    Dim trapValorPt      As Double
    Dim sInput           As String
    Dim resposta         As Integer
    Dim log()            As String
    Dim nLog             As Integer
    Dim totalProcessados As Integer
    Dim totalIgnorados   As Integer
    Dim i                As Integer
    Dim objFrente        As Shape
    Dim objFundo         As Shape
    Dim resultado        As TrapResult
    Dim relatorio        As String
    Dim j                As Integer
    Dim ok               As Boolean

    If ActiveDocument Is Nothing Then
        MsgBox "Nenhum documento aberto.", vbExclamation, "TrapMaker"
        Exit Sub
    End If

    Set sel = ActiveSelectionRange
    If sel.Count < 2 Then
        MsgBox "Selecione pelo menos 2 objetos.", vbInformation, "TrapMaker"
        Exit Sub
    End If

    sInput = InputBox( _
        "TrapMaker v1.9.0" & vbCrLf & vbCrLf & _
        "Valor de trapping (mm):" & vbCrLf & _
        "  Minimo : 0,05 mm" & vbCrLf & _
        "  Padrao : 0,10 mm" & vbCrLf & _
        "  Maximo : 0,50 mm" & vbCrLf & vbCrLf & _
        "Deixe em branco para usar 0,10 mm.", _
        "TrapMaker - Valor de Trap", "0,10")

    If sInput = "" Then
        resposta = MsgBox("Nenhum valor digitado." & vbCrLf & vbCrLf & _
                          "Deseja usar o valor padrao (0,10 mm)?", _
                          vbQuestion + vbYesNo, "TrapMaker")
        If resposta = vbNo Then Exit Sub
        sInput = "0,10"
    End If

    sInput = Trim(Replace(sInput, ",", "."))

    If Not IsNumeric(sInput) Then
        MsgBox "Valor invalido. Ex: 0,15", vbExclamation, "TrapMaker"
        Exit Sub
    End If

    trapValorMm = Val(sInput)

    If trapValorMm < 0.05 Or trapValorMm > 0.5 Then
        MsgBox "Valor fora do intervalo (0,05 a 0,50 mm).", vbExclamation, "TrapMaker"
        Exit Sub
    End If

    trapValorPt = trapValorMm * 2.8346

    Set trapLayer = ObterOuCriarLayerTrap()
    If trapLayer Is Nothing Then
        MsgBox "Nao foi possivel criar a layer TRAP.", vbExclamation, "TrapMaker"
        Exit Sub
    End If

    nLog = 0
    totalProcessados = 0
    totalIgnorados = 0
    ReDim log(0)

    For i = 1 To sel.Count - 1
        Set objFrente = sel(i)
        Set objFundo = sel(i + 1)

        If Not ObjetoSuportado(objFrente) Or Not ObjetoSuportado(objFundo) Then
            nLog = nLog + 1
            ReDim Preserve log(nLog - 1)
            log(nLog - 1) = LOG_SKIP & "Par " & i & ": tipo nao suportado" & _
                            " (Type=" & objFrente.Type & "/" & objFundo.Type & ")."
            totalIgnorados = totalIgnorados + 1
        Else
            resultado = CalcularTrap(objFrente, objFundo, trapValorPt)

            If resultado.TipoTrap = "IGNORADO" Then
                nLog = nLog + 1
                ReDim Preserve log(nLog - 1)
                log(nLog - 1) = LOG_SKIP & "Par " & i & ": " & resultado.Mensagem
                totalIgnorados = totalIgnorados + 1
            Else
                ok = AplicarTrap(objFrente, objFundo, resultado, trapLayer)
                nLog = nLog + 1
                ReDim Preserve log(nLog - 1)
                If ok Then
                    log(nLog - 1) = LOG_OK & "Par " & i & _
                        " [" & resultado.TipoTrap & "]" & _
                        " | " & Format(trapValorMm, "0.00") & "mm" & _
                        " | C=" & Format(resultado.CorTrapC, "0") & _
                        " M=" & Format(resultado.CorTrapM, "0") & _
                        " Y=" & Format(resultado.CorTrapY, "0") & _
                        " K=" & Format(resultado.CorTrapK, "0") & _
                        " | " & resultado.Mensagem
                    totalProcessados = totalProcessados + 1
                Else
                    log(nLog - 1) = LOG_SKIP & "Par " & i & _
                        " [" & resultado.TipoTrap & _
                        "] - falha ao gerar objeto de trap (ver DiagnosticarConvertToObject)."
                    totalIgnorados = totalIgnorados + 1
                End If
            End If
        End If
    Next i

    relatorio = "TrapMaker v1.9.0 - Relatorio" & vbCrLf
    relatorio = relatorio & String(40, "-") & vbCrLf
    relatorio = relatorio & "Valor    : " & Format(trapValorMm, "0.00") & " mm" & vbCrLf
    relatorio = relatorio & "Gerados  : " & totalProcessados & vbCrLf
    relatorio = relatorio & "Ignorados: " & totalIgnorados & vbCrLf
    relatorio = relatorio & String(40, "-") & vbCrLf & vbCrLf

    For j = 0 To nLog - 1
        relatorio = relatorio & log(j) & vbCrLf
    Next j

    relatorio = relatorio & vbCrLf & "Revise a layer TRAP no Gerenciador de Objetos."
    MsgBox relatorio, vbInformation, "TrapMaker - Concluido"
End Sub


' =============================================================================
Private Function ObjetoSuportado(obj As Shape) As Boolean
    Dim t As Integer
    t = obj.Type
    If t = CDR_BITMAP_TYPE Then ObjetoSuportado = False: Exit Function
    If t = CDR_OLE_TYPE    Then ObjetoSuportado = False: Exit Function
    ObjetoSuportado = True
End Function


' =============================================================================
Private Function CalcularTrap(objFrente As Shape, _
                               objFundo  As Shape, _
                               trapPt    As Double) As TrapResult
    Dim res       As TrapResult
    Dim fC As Double, fM As Double, fY As Double, fK As Double
    Dim bC As Double, bM As Double, bY As Double, bK As Double
    Dim lumaF     As Double
    Dim lumaB     As Double
    Dim deltaLuma As Double

    res.ValorPt = trapPt

    If Not LerCorCMYK(objFrente, fC, fM, fY, fK) Then
        res.TipoTrap = "IGNORADO"
        res.Mensagem = "Frente: Fill=" & objFrente.Fill.Type & " - nao e CMYK solido."
        CalcularTrap = res
        Exit Function
    End If

    If Not LerCorCMYK(objFundo, bC, bM, bY, bK) Then
        res.TipoTrap = "IGNORADO"
        res.Mensagem = "Fundo: Fill=" & objFundo.Fill.Type & " - nao e CMYK solido."
        CalcularTrap = res
        Exit Function
    End If

    lumaF = CalcLuma(fC, fM, fY, fK)
    lumaB = CalcLuma(bC, bM, bY, bK)
    deltaLuma = lumaF - lumaB

    If Abs(deltaLuma) < LUMA_THRESHOLD Then
        res.TipoTrap = "NEUTRO"
        res.Mensagem = "Luminancias similares - media das cores."
    ElseIf deltaLuma > 0 Then
        res.TipoTrap = "SPREAD"
        res.Mensagem = "Frente clara (" & Format(lumaF, "0.00") & _
                       ") x fundo escuro (" & Format(lumaB, "0.00") & ")."
    Else
        res.TipoTrap = "CHOKE"
        res.Mensagem = "Fundo claro (" & Format(lumaB, "0.00") & _
                       ") x objeto escuro (" & Format(lumaF, "0.00") & ")."
    End If

    Select Case res.TipoTrap
        Case "SPREAD"
            res.CorTrapC = fC: res.CorTrapM = fM
            res.CorTrapY = fY: res.CorTrapK = fK
        Case "CHOKE"
            res.CorTrapC = bC: res.CorTrapM = bM
            res.CorTrapY = bY: res.CorTrapK = bK
        Case "NEUTRO"
            res.CorTrapC = (fC + bC) / 2: res.CorTrapM = (fM + bM) / 2
            res.CorTrapY = (fY + bY) / 2: res.CorTrapK = (fK + bK) / 2
    End Select

    CalcularTrap = res
End Function


' =============================================================================
' =============================================================================
' AplicarTrap v1.9.0
'
' FLUXO BASEADO NO PROCESSO MANUAL CONFIRMADO:
'   1. CreateContour(direcao, offset_mm, etapas=1, cantos arredondados, cor)
'      Equivalente a: Ferramenta Contorno no menu lateral
'   2. BreakApart() — equivalente ao Ctrl+K (Separar Grupo de Contorno)
'      Apos separacao, shapes ficam em ActiveSelectionRange
'   3. Identificar trap (maior=SPREAD, menor=CHOKE), configurar, mover para TRAP
'
' API v27 (community.coreldraw.com/sdk/api/draw/27):
'   Shape.CreateContour(Direction, Offset, Steps, ..., CornerType) -> Effect
'   Shape.BreakApart() -> Sub  [Ctrl+K]
'   cdrContourDirection: cdrContourOutside=2, cdrContourInside=0
'   cdrContourCornerType: arredondado=1
'
' Retorna True se o objeto trap foi criado com sucesso.
' =============================================================================
Private Function AplicarTrap(objFrente As Shape, _
                              objFundo  As Shape, _
                              res       As TrapResult, _
                              trapLayer As Layer) As Boolean
    Dim baseShape As Shape
    Dim trapShape As Shape
    Dim origShape As Shape
    Dim corTrap   As Color
    Dim trapMm    As Double
    Dim direcao   As Long
    Dim selApos   As ShapeRange
    Dim area1     As Double
    Dim area2     As Double
    Dim k         As Integer

    AplicarTrap = False

    ' 1 pt = 0.352778 mm
    trapMm = res.ValorPt * 0.352778

    ' Shape base e direcao
    Select Case res.TipoTrap
        Case "SPREAD", "NEUTRO"
            Set baseShape = objFrente.Duplicate()
            direcao = 2   ' cdrContourOutside
        Case "CHOKE"
            Set baseShape = objFundo.Duplicate()
            direcao = 0   ' cdrContourInside
        Case Else
            Exit Function
    End Select

    If baseShape Is Nothing Then Exit Function

    Set corTrap = CreateCMYKColor(res.CorTrapC, res.CorTrapM, res.CorTrapY, res.CorTrapK)

    ' -----------------------------------------------------------------------
    ' PASSO 1: CreateContour — mesmo fluxo da ferramenta manual
    ' Direction, Offset(mm), Steps=1, BlendType, OutlineColor, FillColor,
    ' FillColor2, SpacingAccel, ColorAccel, EndCapType, CornerType=1(Round)
    ' -----------------------------------------------------------------------
    On Error Resume Next
    baseShape.CreateContour direcao, trapMm, 1, , , corTrap, , , , , 1
    On Error GoTo 0

    ' -----------------------------------------------------------------------
    ' PASSO 2: BreakApart — equivalente ao Ctrl+K
    ' Separa o grupo. Os shapes resultantes ficam em ActiveSelectionRange.
    ' -----------------------------------------------------------------------
    On Error Resume Next
    baseShape.BreakApart
    Set selApos = ActiveSelectionRange
    On Error GoTo 0

    If selApos Is Nothing Then
        On Error Resume Next: baseShape.Delete: On Error GoTo 0
        Exit Function
    End If

    If selApos.Count < 2 Then
        On Error Resume Next
        For k = 1 To selApos.Count
            selApos(k).Delete
        Next k
        On Error GoTo 0
        Exit Function
    End If

    ' -----------------------------------------------------------------------
    ' PASSO 3: Identificar trap pela area da bounding box
    ' Outside (SPREAD/NEUTRO): trap = shape MAIOR
    ' Inside  (CHOKE)        : trap = shape MENOR
    ' -----------------------------------------------------------------------
    area1 = selApos(1).SizeWidth * selApos(1).SizeHeight
    area2 = selApos(2).SizeWidth * selApos(2).SizeHeight

    If direcao = 2 Then
        If area1 >= area2 Then
            Set trapShape = selApos(1): Set origShape = selApos(2)
        Else
            Set trapShape = selApos(2): Set origShape = selApos(1)
        End If
    Else
        If area1 <= area2 Then
            Set trapShape = selApos(1): Set origShape = selApos(2)
        Else
            Set trapShape = selApos(2): Set origShape = selApos(1)
        End If
    End If

    On Error Resume Next: origShape.Delete: On Error GoTo 0

    If trapShape Is Nothing Then Exit Function

    ' Configurar objeto trap final
    On Error Resume Next
    trapShape.Fill.ApplyUniformFill corTrap
    trapShape.Outline.SetNoOutline
    trapShape.OverprintFill = True
    trapShape.Layer = trapLayer
    trapShape.Name = "TRAP_" & res.TipoTrap & "_" & _
                     Format(res.ValorPt / 2.8346, "0.00") & "mm"
    On Error GoTo 0

    AplicarTrap = True
End Function



' =============================================================================
' DiagnosticarContour v1.8
' Testa o pipeline completo de CreateContour em um objeto selecionado.
' Selecione 1 retangulo original e execute esta macro.
' =============================================================================
Public Sub DiagnosticarContour()
    Dim sel       As ShapeRange
    Dim s         As Shape
    Dim base      As Shape
    Dim efeito    As Effect
    Dim trapShape As Shape
    Dim corTeste  As Color
    Dim msg       As String
    Dim errNum    As Long
    Dim errDesc   As String
    Dim trapMm    As Double

    If ActiveDocument Is Nothing Then
        MsgBox "Nenhum documento aberto.", vbExclamation
        Exit Sub
    End If

    Set sel = ActiveSelectionRange
    If sel.Count <> 1 Then
        MsgBox "Selecione exatamente 1 objeto original.", vbInformation
        Exit Sub
    End If

    Set s = sel(1)
    trapMm = 0.1  ' 0,10mm de teste

    msg = "=== PIPELINE CreateContour v1.8 ===" & vbCrLf
    msg = msg & "Shape: Type=" & s.Type & " Layer=" & s.Layer.Name & vbCrLf & vbCrLf

    ' Etapa 1: Duplicate
    On Error Resume Next
    Set base = s.Duplicate()
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo 0
    msg = msg & "E1 Duplicate: "
    If errNum <> 0 Or base Is Nothing Then
        msg = msg & "FALHOU " & errNum & " " & errDesc & vbCrLf
        MsgBox msg, vbExclamation: Exit Sub
    End If
    msg = msg & "OK" & vbCrLf

    ' Etapa 2: ConvertToCurves
    On Error Resume Next
    base.ConvertToCurves
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo 0
    msg = msg & "E2 ConvertToCurves: "
    If errNum <> 0 Then
        msg = msg & "ERRO " & errNum & " " & errDesc & vbCrLf
    Else
        msg = msg & "OK | Type=" & base.Type & vbCrLf
    End If

    ' Etapa 3: CreateContour (Inside=0 para simular CHOKE)
    Set corTeste = CreateCMYKColor(0, 50, 50, 0)
    On Error Resume Next
    Set efeito = base.CreateContour(0, trapMm, 1, , , corTeste)
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo 0
    msg = msg & "E3 CreateContour(Inside, " & trapMm & "mm): "
    If errNum <> 0 Then
        msg = msg & "ERRO " & errNum & " " & errDesc & vbCrLf
    ElseIf efeito Is Nothing Then
        msg = msg & "RETORNOU Nothing" & vbCrLf
    Else
        msg = msg & "OK | efeito.Shape.Shapes.Count=" & efeito.Shape.Shapes.Count & vbCrLf
    End If

    If efeito Is Nothing Then
        On Error Resume Next: base.Delete: On Error GoTo 0
        MsgBox msg, vbExclamation: Exit Sub
    End If

    ' Etapa 4: Capturar Shapes(1)
    On Error Resume Next
    Set trapShape = efeito.Shape.Shapes(1)
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo 0
    msg = msg & "E4 efeito.Shape.Shapes(1): "
    If errNum <> 0 Then
        msg = msg & "ERRO " & errNum & " " & errDesc & vbCrLf
    ElseIf trapShape Is Nothing Then
        msg = msg & "RETORNOU Nothing" & vbCrLf
    Else
        msg = msg & "OK | Name='" & trapShape.Name & "'" & vbCrLf
        msg = msg & "     Layer=" & trapShape.Layer.Name & vbCrLf
        msg = msg & "     W=" & Format(trapShape.SizeWidth, "0.000") & vbCrLf
        msg = msg & "     H=" & Format(trapShape.SizeHeight, "0.000") & vbCrLf
    End If

    ' Etapa 5: ClearEffect cdrContour
    On Error Resume Next
    base.ClearEffect cdrContour
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo 0
    msg = msg & "E5 ClearEffect cdrContour: "
    If errNum <> 0 Then
        msg = msg & "ERRO " & errNum & " " & errDesc & vbCrLf
    Else
        msg = msg & "OK" & vbCrLf
    End If

    ' Verificar se trapShape ainda e valido
    msg = msg & "E5 trapShape apos ClearEffect: "
    On Error Resume Next
    Dim testName As String
    testName = trapShape.Name
    errNum = Err.Number: Err.Clear
    On Error GoTo 0
    If errNum <> 0 Then
        msg = msg & "INVALIDO (referencia perdida)" & vbCrLf
    Else
        msg = msg & "OK | Name='" & testName & "'" & vbCrLf
        msg = msg & "     Layer=" & trapShape.Layer.Name & vbCrLf
        msg = msg & "     Fill.Type=" & trapShape.Fill.Type & vbCrLf
    End If

    ' Cleanup
    On Error Resume Next
    base.Delete
    On Error GoTo 0

    MsgBox msg, vbInformation, "Diagnostico Contour v1.8"
End Sub



' =============================================================================
' LerCorCMYK
' Le valores CMYK de um Shape com preenchimento uniforme (cdrUniformFill = 1).
' Ref: Fill.Type (cdrFillType), Fill.UniformColor (Color), API v27.
' Retorna False se Fill.Type <> 1 (nao e preenchimento CMYK solido).
' =============================================================================
Private Function LerCorCMYK(obj As Shape, _
                              ByRef C As Double, ByRef M As Double, _
                              ByRef Y As Double, ByRef K As Double) As Boolean
    C = 0: M = 0: Y = 0: K = 0
    LerCorCMYK = False
    On Error GoTo SemCor
    If obj.Fill.Type <> CDR_UNIFORM_FILL Then GoTo SemCor
    C = obj.Fill.UniformColor.CMYKCyan
    M = obj.Fill.UniformColor.CMYKMagenta
    Y = obj.Fill.UniformColor.CMYKYellow
    K = obj.Fill.UniformColor.CMYKBlack
    LerCorCMYK = True
    Exit Function
SemCor:
    LerCorCMYK = False
End Function


' =============================================================================
' CalcLuma - luminancia perceptual CMYK, retorna [0,1]
' Pesos: C=0.30, M=0.59, Y=0.11, K=1.00 (normalizados /100)
' =============================================================================
Private Function CalcLuma(C As Double, M As Double, _
                           Y As Double, K As Double) As Double
    Dim v As Double
    v = 1 - ((0.3 * C) + (0.59 * M) + (0.11 * Y) + K) / 100
    CalcLuma = ClampDouble(v, 0, 1)
End Function


' =============================================================================
Private Function TipoTrapLuma(lumaF As Double, lumaB As Double) As String
    Dim d As Double
    d = lumaF - lumaB
    If Abs(d) < LUMA_THRESHOLD Then
        TipoTrapLuma = "NEUTRO"
    ElseIf d > 0 Then
        TipoTrapLuma = "SPREAD"
    Else
        TipoTrapLuma = "CHOKE"
    End If
End Function


' =============================================================================
Private Function ClampDouble(valor As Double, _
                              minVal As Double, _
                              maxVal As Double) As Double
    If valor < minVal Then
        ClampDouble = minVal
    ElseIf valor > maxVal Then
        ClampDouble = maxVal
    Else
        ClampDouble = valor
    End If
End Function


' =============================================================================
' ObterMaiorObjeto - [RESERVA Fase 2A] Shape de maior area na selecao atual
' =============================================================================
Private Function ObterMaiorObjeto() As Shape
    Dim selAtual  As ShapeRange
    Dim s         As Shape
    Dim maiorArea As Double
    Dim melhor    As Shape
    Dim area      As Double

    Set selAtual = ActiveSelectionRange
    maiorArea = 0
    For Each s In selAtual
        area = s.SizeWidth * s.SizeHeight
        If area > maiorArea Then
            maiorArea = area
            Set melhor = s
        End If
    Next s
    Set ObterMaiorObjeto = melhor
End Function


' =============================================================================
Private Function ObterOuCriarLayerTrap() As Layer
    Dim doc As Document
    Dim pg  As Page
    Dim lyr As Layer
    Dim idx As Integer

    Set doc = ActiveDocument
    Set pg = doc.ActivePage

    For idx = 1 To pg.Layers.Count
        If pg.Layers(idx).Name = TRAP_LAYER_NAME Then
            Set lyr = pg.Layers(idx)
            Exit For
        End If
    Next idx

    If lyr Is Nothing Then
        On Error GoTo ErroLayer
        Set lyr = pg.CreateLayer(TRAP_LAYER_NAME)
        lyr.Visible   = True
        lyr.Printable = True
        lyr.Editable  = True
        On Error Resume Next
        lyr.Color.RGBAssign 83, 74, 183
        On Error GoTo 0
    End If

    Set ObterOuCriarLayerTrap = lyr
    Exit Function

ErroLayer:
    On Error Resume Next
    Set lyr = doc.MasterPage.CreateLayer(TRAP_LAYER_NAME)
    On Error GoTo 0
    Set ObterOuCriarLayerTrap = lyr
End Function


' =============================================================================
' DiagnosticarConvertToObject v2
' Testa cada etapa do pipeline de geracao de trap em um objeto selecionado.
' Execute com 1 retangulo ou curva (sem bitmap, sem grupo).
' =============================================================================
Public Sub DiagnosticarConvertToObject()
    Dim sel       As ShapeRange
    Dim s         As Shape
    Dim dupShape  As Shape
    Dim trapShape As Shape
    Dim corTeste  As Color
    Dim msg       As String
    Dim errNum    As Long
    Dim errDesc   As String

    If ActiveDocument Is Nothing Then
        MsgBox "Nenhum documento aberto.", vbExclamation, "Diagnostico"
        Exit Sub
    End If

    Set sel = ActiveSelectionRange
    If sel.Count <> 1 Then
        MsgBox "Selecione exatamente 1 objeto (retangulo ou curva).", _
               vbInformation, "Diagnostico"
        Exit Sub
    End If

    Set s = sel(1)
    msg = "=== ETAPA 1: Shape original ===" & vbCrLf
    msg = msg & "Type=" & s.Type & " | Layer=" & s.Layer.Name & vbCrLf
    msg = msg & "CanHaveOutline=" & s.CanHaveOutline & vbCrLf
    msg = msg & "Outline.Type=" & s.Outline.Type & vbCrLf & vbCrLf

    ' Etapa 2: Duplicate
    On Error Resume Next
    Set dupShape = s.Duplicate()
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo 0
    msg = msg & "=== ETAPA 2: Duplicate ===" & vbCrLf
    If errNum <> 0 Then
        msg = msg & "ERRO " & errNum & ": " & errDesc & vbCrLf
        MsgBox msg, vbExclamation, "Diagnostico": Exit Sub
    End If
    msg = msg & "OK" & vbCrLf & vbCrLf

    ' Etapa 3: ConvertToCurves
    On Error Resume Next
    dupShape.ConvertToCurves
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo 0
    msg = msg & "=== ETAPA 3: ConvertToCurves ===" & vbCrLf
    If errNum <> 0 Then
        msg = msg & "ERRO " & errNum & ": " & errDesc & vbCrLf
    Else
        msg = msg & "OK | Type apos=" & dupShape.Type & vbCrLf
    End If
    msg = msg & vbCrLf

    ' Etapa 4: Propriedades individuais de Outline (em vez de SetProperties)
    ' Outline.Width e Outline.Color funcionam quando Type=0 (cdrOutlineNone)
    Set corTeste = CreateCMYKColor(0, 100, 0, 0)
    On Error Resume Next
    dupShape.Outline.Width          = 2
    dupShape.Outline.Color          = corTeste
    dupShape.Outline.BehindFill     = CDR_FALSE
    dupShape.Outline.ScaleWithShape = CDR_FALSE
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo 0
    msg = msg & "=== ETAPA 4: Outline.Width/Color/BehindFill ===" & vbCrLf
    If errNum <> 0 Then
        msg = msg & "ERRO " & errNum & ": " & errDesc & vbCrLf
    Else
        msg = msg & "OK" & vbCrLf
    End If
    msg = msg & "Outline.Type apos=" & dupShape.Outline.Type & vbCrLf
    msg = msg & "Outline.Width apos=" & Format(dupShape.Outline.Width, "0.000") & "pt" & vbCrLf
    If dupShape.Outline.Type = 0 Or dupShape.Outline.Type = 1 Then
        msg = msg & "** ATENCAO: Outline ainda inativo (Type=" & dupShape.Outline.Type & ") **" & vbCrLf
    Else
        msg = msg & "Outline ATIVO - pronto para ConvertToObject" & vbCrLf
    End If
    msg = msg & vbCrLf

    ' Etapa 5: ApplyNoFill
    On Error Resume Next
    dupShape.Fill.ApplyNoFill
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo 0
    msg = msg & "=== ETAPA 5: ApplyNoFill ===" & vbCrLf
    If errNum <> 0 Then
        msg = msg & "ERRO " & errNum & ": " & errDesc & vbCrLf
    Else
        msg = msg & "OK" & vbCrLf
    End If
    msg = msg & vbCrLf

    ' Etapa 6: ConvertToObject (retorna Shape, API v27)
    On Error Resume Next
    Set trapShape = dupShape.Outline.ConvertToObject
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo 0
    msg = msg & "=== ETAPA 6: ConvertToObject ===" & vbCrLf
    If errNum <> 0 Then
        msg = msg & "ERRO " & errNum & ": " & errDesc & vbCrLf
    ElseIf trapShape Is Nothing Then
        msg = msg & "RETORNOU Nothing" & vbCrLf
        msg = msg & "(outline ausente ou falha interna)" & vbCrLf
    Else
        msg = msg & "OK - Shape retornado com sucesso!" & vbCrLf
        msg = msg & "  Layer=" & trapShape.Layer.Name & vbCrLf
        msg = msg & "  Fill.Type=" & trapShape.Fill.Type & vbCrLf
        msg = msg & "  Outline.Type=" & trapShape.Outline.Type & vbCrLf
    End If

    MsgBox msg, vbInformation, "Diagnostico ConvertToObject v2"
End Sub


' =============================================================================
Public Sub TestarCalculoLuminancia()
' =============================================================================
    Dim sRes As String
    Dim l1f As Double, l1b As Double
    Dim l2f As Double, l2b As Double
    Dim l3f As Double, l3b As Double
    Dim l4f As Double, l4b As Double

    sRes = "TESTE - TrapMaker v1.9.0" & vbCrLf
    sRes = sRes & String(40, "=") & vbCrLf & vbCrLf

    l1f = CalcLuma(0, 0, 100, 0)
    l1b = CalcLuma(100, 100, 0, 0)
    sRes = sRes & "Caso 1 Amarelo x Azul" & vbCrLf
    sRes = sRes & "  Esperado: SPREAD | " & TipoTrapLuma(l1f, l1b) & vbCrLf & vbCrLf

    l2f = CalcLuma(0, 0, 0, 100)
    l2b = CalcLuma(0, 0, 0, 0)
    sRes = sRes & "Caso 2 Preto x Branco" & vbCrLf
    sRes = sRes & "  Esperado: CHOKE  | " & TipoTrapLuma(l2f, l2b) & vbCrLf & vbCrLf

    l3f = CalcLuma(0, 0, 0, 50)
    l3b = CalcLuma(0, 0, 0, 52)
    sRes = sRes & "Caso 3 Cinza50 x Cinza52" & vbCrLf
    sRes = sRes & "  Esperado: NEUTRO | " & TipoTrapLuma(l3f, l3b) & vbCrLf & vbCrLf

    l4f = CalcLuma(100, 0, 0, 0)
    l4b = CalcLuma(0, 100, 0, 0)
    sRes = sRes & "Caso 4 Ciano x Magenta" & vbCrLf
    sRes = sRes & "  Calculado: " & TipoTrapLuma(l4f, l4b)

    MsgBox sRes, vbInformation, "TrapMaker - Teste Luminancia"
End Sub


' =============================================================================
Public Sub SobreOTrapMaker()
' =============================================================================
    Dim s As String
    s = "TrapMaker v1.9.0" & vbCrLf
    s = s & String(38, "-") & vbCrLf
    s = s & "Macro de trapping semi-automatico" & vbCrLf
    s = s & "para CorelDRAW 2026 v27 ou superior." & vbCrLf & vbCrLf
    s = s & "COMO USAR:" & vbCrLf
    s = s & "1. Selecione 2 ou mais objetos com fill CMYK solido" & vbCrLf
    s = s & "2. Execute RunTrapMaker" & vbCrLf
    s = s & "3. Informe o valor de trap em mm" & vbCrLf
    s = s & "4. Revise a layer TRAP no Gerenciador de Objetos" & vbCrLf & vbCrLf
    s = s & "TIPOS:" & vbCrLf
    s = s & "  SPREAD : objeto claro expande sobre fundo escuro" & vbCrLf
    s = s & "  CHOKE  : fundo claro expande sob objeto escuro" & vbCrLf
    s = s & "  NEUTRO : luminancias similares, trap medio" & vbCrLf & vbCrLf
    s = s & "DIAGNOSTICO:" & vbCrLf
    s = s & "  DiagnosticarSelecao: verifica Fill/CMYK dos objetos" & vbCrLf
    s = s & "  DiagnosticarConvertToObject: testa pipeline de geracao"
    MsgBox s, vbInformation, "Sobre o TrapMaker"
End Sub

' === FIM TrapMaker v1.9.0 ===
