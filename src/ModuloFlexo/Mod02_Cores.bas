Attribute VB_Name = "Mod02_Cores"
' ============================================================
' M�DULO: Mod02_Cores (VERS�O CONSOLIDADA DEFINITIVA V1.5)
' DESCRI��O: Motor de busca recursiva absoluta para Flexografia
' ============================================================

Option Explicit

Private srRes As ShapeRange
Private escolhaMetodo As VbMsgBoxResult

' ============================================================
' BLOCO 1: FERRAMENTAS DE CONVERS�O E CORRE��O
' ============================================================

Public Sub ConverterRGB(Optional silencioso As Boolean = False)
    If silencioso Then
        escolhaMetodo = vbYes   ' sempre CMYK completo no modo autom�tico
    Else
        escolhaMetodo = MsgBox("Converter RGB para:" & vbCrLf & vbCrLf & _
                            "[ SIM ] -> CMYK (Completo)" & vbCrLf & _
                            "[ N" & ChrW(195) & "O ] -> CMY (Sem Preto - Ideal para Flexo)", _
                            vbYesNoCancel + vbQuestion, "Convers" & ChrW(227) & "o RGB")
        If escolhaMetodo = vbCancel Then Exit Sub
    End If
    ChamarMotor "RGB", silencioso
    If Not silencioso Then Finalizar "Objetos RGB convertidos"
End Sub

Public Sub ConverterSpotParaCMYK()
    escolhaMetodo = MsgBox("Converter PANTONE/SPOT para:" & vbCrLf & vbCrLf & _
                        "[ SIM ] -> CMYK (Completo)" & vbCrLf & _
                        "[ N" & ChrW(195) & "O ] -> CMY (Sem Preto - Evita sujar a cor)", _
                        vbYesNoCancel + vbQuestion, "Convers" & ChrW(227) & "o Pantone")
    If escolhaMetodo = vbCancel Then Exit Sub
    ChamarMotor "Spot"
    Finalizar "Cores Spot/Pantone convertidas"
End Sub

Public Sub CorrigirBrancoOverprint(Optional silencioso As Boolean = False)
    ChamarMotor "Branco", silencioso
    If Not silencioso Then Finalizar "Brancos corrigidos"
End Sub

Public Sub DetectarPretoSujo(Optional silencioso As Boolean = False)
    ChamarMotor "PretoSujo", silencioso
    If Not silencioso Then Finalizar "Pretos Sujos limpos"
End Sub

' --- GERENCIADORES DE FLUXO ---

Private Sub ChamarMotor(acao As String, Optional silencioso As Boolean = False)
    Dim s As Shape
    Set srRes = CreateShapeRange

    If Not silencioso Then ActiveDocument.BeginCommandGroup "Console Flexo - " & acao
    Application.Optimization = True
    Application.EventsEnabled = False
    On Error GoTo FimErro

    For Each s In ActivePage.shapes
        ExecutarCrawler s, acao
    Next s

FimErro:
    If Not silencioso Then ActiveDocument.EndCommandGroup
    Application.EventsEnabled = True
    Application.Optimization = False
    If Not silencioso Then Application.Refresh
End Sub

Private Sub Finalizar(texto As String)
    If srRes.Count > 0 Then
        srRes.CreateSelection
        MsgBox texto & " e selecionados: " & srRes.Count, vbInformation, "Console Flexo"
    Else
        MsgBox "Nenhum objeto com esse crit" & ChrW(233) & "rio foi encontrado.", vbInformation, "Console Flexo"
    End If
End Sub

Private Sub ExecutarCrawler(s As Shape, tipoAcao As String)
    Dim subS As Shape
    Dim C As Integer, M As Integer, Y As Integer, K As Integer
    Dim sofreuAlteracao As Boolean
    On Error Resume Next
    
    If s.Type <> cdrGroupShape And s.Type <> cdrGuidelineShape Then
        sofreuAlteracao = False
        Select Case tipoAcao
            Case "Branco"
                If s.Fill.Type = cdrUniformFill Then
                    If s.Fill.UniformColor.IsWhite And s.OverprintFill Then
                        s.OverprintFill = False: sofreuAlteracao = True
                    End If
                End If
                If s.Outline.Type = cdrOutline Then
                    If s.Outline.Color.IsWhite And s.OverprintOutline Then
                        s.OverprintOutline = False: sofreuAlteracao = True
                    End If
                End If
            Case "RGB", "Spot"
                If s.Fill.Type = cdrUniformFill Then
                    Dim targetF As Boolean: targetF = False
                    If tipoAcao = "RGB" And s.Fill.UniformColor.Type = cdrColorRGB Then targetF = True
                    If tipoAcao = "Spot" And s.Fill.UniformColor.IsSpot Then targetF = True
                    If targetF Then
                        s.Fill.UniformColor.ConvertToCMYK
                        If escolhaMetodo = vbNo Then s.Fill.UniformColor.CMYKBlack = 0
                        sofreuAlteracao = True
                    End If
                End If
                If s.Fill.Type = cdrFountainFill Then
                    Dim nG As Integer
                    Dim totalG As Integer: totalG = 2 + s.Fill.Fountain.Colors.Count
                    For nG = 1 To totalG
                        Dim cG As Color
                        On Error Resume Next
                        If nG = 1 Then Set cG = s.Fill.Fountain.StartColor
                        If nG = 2 Then Set cG = s.Fill.Fountain.EndColor
                        If nG > 2  Then Set cG = s.Fill.Fountain.Colors(nG - 3).Color
                        On Error GoTo 0
                        If Not cG Is Nothing Then
                            Dim deveConverterG As Boolean: deveConverterG = False
                            If tipoAcao = "RGB"  And cG.Type = cdrColorRGB Then deveConverterG = True
                            If tipoAcao = "Spot" And cG.IsSpot             Then deveConverterG = True
                            If deveConverterG Then
                                cG.ConvertToCMYK
                                If escolhaMetodo = vbNo Then cG.CMYKBlack = 0
                                sofreuAlteracao = True
                            End If
                        End If
                        Set cG = Nothing
                    Next nG
                End If
                If s.Outline.Type = cdrOutline Then
                    Dim targetO As Boolean: targetO = False
                    If tipoAcao = "RGB" And s.Outline.Color.Type = cdrColorRGB Then targetO = True
                    If tipoAcao = "Spot" And s.Outline.Color.IsSpot Then targetO = True
                    If targetO Then
                        s.Outline.Color.ConvertToCMYK
                        If escolhaMetodo = vbNo Then s.Outline.Color.CMYKBlack = 0
                        sofreuAlteracao = True
                    End If
                End If
            Case "PretoSujo"
                ' Limiar alinhado com o Scanner: K > 85 com qualquer contamina��o CMY > 0
                ' Exceto Preto Puro (K100) e Preto Rico (C100 M100 Y100 K100)
                If s.Fill.Type = cdrUniformFill And s.Fill.UniformColor.Type = cdrColorCMYK Then
                    C = s.Fill.UniformColor.CMYKCyan: M = s.Fill.UniformColor.CMYKMagenta
                    Y = s.Fill.UniformColor.CMYKYellow: K = s.Fill.UniformColor.CMYKBlack
                    If (C + M + Y) > 0 And K > 85 Then
                        If Not (C = 0 And M = 0 And Y = 0 And K = 100) Then
                            s.Fill.UniformColor.CMYKCyan = 0
                            s.Fill.UniformColor.CMYKMagenta = 0
                            s.Fill.UniformColor.CMYKYellow = 0
                            s.Fill.UniformColor.CMYKBlack = 100
                            sofreuAlteracao = True
                        End If
                    End If
                End If
                If s.Outline.Type = cdrOutline And s.Outline.Color.Type = cdrColorCMYK Then
                    C = s.Outline.Color.CMYKCyan: M = s.Outline.Color.CMYKMagenta
                    Y = s.Outline.Color.CMYKYellow: K = s.Outline.Color.CMYKBlack
                    If (C + M + Y) > 0 And K > 85 Then
                        If Not (C = 0 And M = 0 And Y = 0 And K = 100) Then
                            s.Outline.Color.CMYKCyan = 0
                            s.Outline.Color.CMYKMagenta = 0
                            s.Outline.Color.CMYKYellow = 0
                            s.Outline.Color.CMYKBlack = 100
                            sofreuAlteracao = True
                        End If
                    End If
                End If
        End Select
        If sofreuAlteracao Then srRes.Add s
    End If
    On Error GoTo 0

    If s.Type = cdrGroupShape Then
        For Each subS In s.Shapes: ExecutarCrawler subS, tipoAcao: Next subS
    End If
    If Not s.PowerClip Is Nothing Then
        For Each subS In s.PowerClip.Shapes: ExecutarCrawler subS, tipoAcao: Next subS
    End If
End Sub

' ============================================================
' FERRAMENTA: SELECIONAR MESMA COR
' ============================================================
Public Sub SelecionarMsmCor(modoBusca As Integer)
    Dim refShape As Shape
    Dim s As Shape
    Dim srCoresIguais As ShapeRange
    
    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Selecione um objeto de refer" & ChrW(234) & "ncia para eu saber qual cor procurar!", vbExclamation, "Console Flexo"
        Exit Sub
    End If
    
    Set refShape = ActiveSelection.Shapes(1)
    Dim temPre As Boolean: temPre = (refShape.Fill.Type = cdrUniformFill)
    Dim temCon As Boolean: temCon = (refShape.Outline.Type = cdrOutline)
    
    If modoBusca = 1 Then temCon = False
    If modoBusca = 2 Then temPre = False
    
    If Not temPre And Not temCon Then
        MsgBox "O objeto de refer" & ChrW(234) & "ncia n" & ChrW(227) & "o possui a propriedade de cor que voc" & ChrW(234) & " quer buscar.", vbInformation, "Console Flexo"
        Exit Sub
    End If
    
    Dim corBuscaFill As Color, corBuscaOutline As Color
    If temPre Then Set corBuscaFill = refShape.Fill.UniformColor
    If temCon Then Set corBuscaOutline = refShape.Outline.Color
    
    Set srCoresIguais = CreateShapeRange
    For Each s In ActivePage.shapes
        CrawlerBuscaCor s, corBuscaFill, corBuscaOutline, temPre, temCon, srCoresIguais
    Next s
    
    If srCoresIguais.Count > 0 Then
        srCoresIguais.CreateSelection
        MsgBox srCoresIguais.Count & " objetos encontrados e selecionados!", vbInformation, "Console Flexo"
    Else
        MsgBox "Nenhum outro objeto com essa cor foi encontrado.", vbInformation, "Console Flexo"
    End If
End Sub

Private Sub CrawlerBuscaCor(s As Shape, cFill As Color, cOut As Color, chkF As Boolean, chkO As Boolean, ByRef sacola As ShapeRange)
    ' Iterativo: usa pilha para evitar Stack Overflow em grupos/PowerClips profundos.
    Dim pilha As New Collection
    pilha.Add s

    Do While pilha.Count > 0
        Dim atual As Shape
        Set atual = pilha.Item(pilha.Count)
        pilha.Remove pilha.Count

        On Error Resume Next
        If atual.Type <> cdrGroupShape And atual.Type <> cdrGuidelineShape Then
            ' Filtra camadas: somente shapes em camadas imprimiveis, visiveis e nao-especiais
            If Not atual.Layer Is Nothing Then
                If atual.Layer.IsSpecialLayer = False And atual.Layer.Printable = True And atual.Layer.Visible = True Then
                    Dim ehIgual As Boolean: ehIgual = False
                    If chkF And atual.Fill.Type = cdrUniformFill Then
                        If CompararCoresSeguro(atual.Fill.UniformColor, cFill) Then ehIgual = True
                    End If
                    If chkO And atual.Outline.Type = cdrOutline Then
                        If CompararCoresSeguro(atual.Outline.Color, cOut) Then ehIgual = True
                    End If
                    If ehIgual Then sacola.Add atual
                End If
            End If
        End If
        On Error GoTo 0

        Dim subS As Shape
        If atual.Type = cdrGroupShape Then
            For Each subS In atual.shapes: pilha.Add subS: Next subS
        End If
        If Not atual.PowerClip Is Nothing Then
            For Each subS In atual.PowerClip.shapes: pilha.Add subS: Next subS
        End If
    Loop
End Sub

Private Function CompararCoresSeguro(c1 As Color, c2 As Color) As Boolean
    CompararCoresSeguro = False
    If c1.Type <> c2.Type Then Exit Function
    Select Case c1.Type
        Case cdrColorCMYK
            If c1.CMYKCyan = c2.CMYKCyan And c1.CMYKMagenta = c2.CMYKMagenta And _
               c1.CMYKYellow = c2.CMYKYellow And c1.CMYKBlack = c2.CMYKBlack Then
                CompararCoresSeguro = True
            End If
        Case cdrColorRGB
            If c1.RGBRed = c2.RGBRed And c1.RGBGreen = c2.RGBGreen And c1.RGBBlue = c2.RGBBlue Then
                CompararCoresSeguro = True
            End If
        Case cdrColorSpot
            If c1.SpotName = c2.SpotName And c1.Tint = c2.Tint Then
                CompararCoresSeguro = True
            End If
        Case Else
            CompararCoresSeguro = c1.IsSame(c2)
    End Select
End Function

' ============================================================
' FERRAMENTA: MUDAR PARA COR DE REGISTRO (Filtro Seguro)
' ============================================================
Public Sub MudarParaCorDeRegistro()
    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Selecione os objetos primeiro!", vbExclamation, "Console Flexo"
        Exit Sub
    End If

    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Tem certeza que deseja converter as cores PRETAS para COR DE REGISTRO?", _
                      vbYesNo + vbQuestion, "Mudar para Cor de Registro")
    If resposta = vbNo Then Exit Sub

    ActiveDocument.BeginCommandGroup "Console Flexo - Cor de Registro"
    On Error GoTo FimErro

    Dim corRegistro As New Color
    corRegistro.RegistrationAssign
    Dim s As Shape
    Dim alteracoes As Integer: alteracoes = 0
    For Each s In ActiveSelection.Shapes
        CrawlerAplicarRegistro s, corRegistro, alteracoes
    Next s

    ActiveDocument.EndCommandGroup

    If alteracoes > 0 Then
        MsgBox "Sucesso! " & alteracoes & " propriedades alteradas para Cor de Registro.", vbInformation, "Console Flexo"
    Else
        MsgBox "Nenhum objeto com Preto Puro ou Preto Rico encontrado.", vbInformation, "Console Flexo"
    End If
    Exit Sub

FimErro:
    ActiveDocument.EndCommandGroup
    MsgBox "Erro: " & Err.Description, vbCritical, "Console Flexo"
End Sub

Private Sub CrawlerAplicarRegistro(s As Shape, corReg As Color, ByRef contador As Integer)
    Dim subS As Shape
    On Error Resume Next
    
    If s.Type <> cdrGroupShape And s.Type <> cdrGuidelineShape Then
        ' 1. PREENCHIMENTO - s� converte se for preto puro ou preto rico
        If s.Fill.Type = cdrUniformFill Then
            If EhPretoParaRegistro(s.Fill.UniformColor) Then
                s.Fill.UniformColor.CopyAssign corReg
                contador = contador + 1
            End If
        End If
        ' 2. CONTORNO - mesma regra
        If s.Outline.Type = cdrOutline Then
            If EhPretoParaRegistro(s.Outline.Color) Then
                s.Outline.Color.CopyAssign corReg
                contador = contador + 1
            End If
        End If
    End If
    On Error GoTo 0
    
    ' Mergulha nos Grupos
    If s.Type = cdrGroupShape Then
        For Each subS In s.Shapes
            CrawlerAplicarRegistro subS, corReg, contador
        Next subS
    End If
    ' Mergulha nos PowerClips
    If Not s.PowerClip Is Nothing Then
        For Each subS In s.PowerClip.Shapes
            CrawlerAplicarRegistro subS, corReg, contador
        Next subS
    End If
End Sub

Private Function EhPretoParaRegistro(C As Color) As Boolean
    EhPretoParaRegistro = False
    If C.Type <> cdrColorCMYK Then Exit Function
    Dim cy As Long, mg As Long, ye As Long, bK As Long
    cy = C.CMYKCyan: mg = C.CMYKMagenta: ye = C.CMYKYellow: bK = C.CMYKBlack
    ' Preto Puro (0,0,0,100)
    If cy = 0 And mg = 0 And ye = 0 And bK = 100 Then EhPretoParaRegistro = True: Exit Function
    ' Preto Rico (100,100,100,100)
    If cy = 100 And mg = 100 And ye = 100 And bK = 100 Then EhPretoParaRegistro = True: Exit Function
End Function

' ============================================================
' FERRAMENTA: APROXIMAR PARA PANTONE
' ============================================================
Public Sub ConverterParaPantone()
    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Selecione os objetos primeiro!", vbExclamation, "Console Flexo"
        Exit Sub
    End If

    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Deseja converter para PANTONE Solid Coated?", vbYesNo + vbQuestion, "Aproximar para Pantone")
    If resposta = vbNo Then Exit Sub

    Dim palPantone As Palette
    On Error Resume Next
    Set palPantone = Palettes.OpenFixed(cdrPANTONECoated)
    On Error GoTo 0

    If palPantone Is Nothing Then
        MsgBox "Biblioteca PANTONE n" & ChrW(227) & "o encontrada.", vbCritical, "Erro de Paleta"
        Exit Sub
    End If

    ActiveDocument.BeginCommandGroup "Console Flexo - Converter Pantone"
    On Error GoTo FimErro

    Dim s As Shape
    Dim alteracoes As Integer: alteracoes = 0
    For Each s In ActiveSelection.Shapes
        CrawlerConverterPantone s, palPantone, alteracoes
    Next s

    ActiveDocument.EndCommandGroup

    If alteracoes > 0 Then
        MsgBox "Sucesso! " & alteracoes & " propriedades convertidas para Pantone.", vbInformation, "Console Flexo"
    Else
        MsgBox "Nenhuma convers" & ChrW(227) & "o feita.", vbInformation, "Console Flexo"
    End If
    Exit Sub

FimErro:
    ActiveDocument.EndCommandGroup
    MsgBox "Erro: " & Err.Description, vbCritical, "Console Flexo"
End Sub

Private Sub CrawlerConverterPantone(s As Shape, pal As Palette, ByRef contador As Integer)
    Dim subS As Shape
    Dim indice As Long
    On Error Resume Next
    If s.Type <> cdrGroupShape And s.Type <> cdrGuidelineShape Then
        If s.Fill.Type = cdrUniformFill Then
            If Not s.Fill.UniformColor.IsSpot Then
                indice = pal.MatchColor(s.Fill.UniformColor)
                s.Fill.UniformColor.CopyAssign pal.Color(indice)
                contador = contador + 1
            End If
        End If
        If s.Outline.Type = cdrOutline Then
            If Not s.Outline.Color.IsSpot Then
                indice = pal.MatchColor(s.Outline.Color)
                s.Outline.Color.CopyAssign pal.Color(indice)
                contador = contador + 1
            End If
        End If
    End If
    On Error GoTo 0
    If s.Type = cdrGroupShape Then
        For Each subS In s.Shapes: CrawlerConverterPantone subS, pal, contador: Next subS
    End If
    If Not s.PowerClip Is Nothing Then
        For Each subS In s.PowerClip.Shapes: CrawlerConverterPantone subS, pal, contador: Next subS
    End If
End Sub

' ============================================================
' FERRAMENTA: CORRETOR DE BORDA DURA EM DEGRAD�S
' ============================================================
Public Sub CorrigirBordaDuraGradientes(Optional minDot As Integer = 0, Optional silencioso As Boolean = False)
    Dim minDotLong As Long

    If silencioso Then
        ' Modo autom�tico: usa o valor passado por par�metro
        minDotLong = CLng(minDot)
        If minDotLong <= 0 Then Exit Sub
    Else
        ' Modo manual: obt�m o valor via InputBox (ignora o par�metro)
        Dim inputMin As String
        inputMin = InputBox("RESOLVER BORDA DURA (HARD EDGE)" & vbCrLf & vbCrLf & _
                            "Digite a porcentagem m" & ChrW(237) & "nima de ponto a ser inserida nos valores zero do degrad" & ChrW(234) & ":" & vbCrLf & _
                            "(Escolha 2 ou 3)", "Ponto M" & ChrW(237) & "nimo Flexo", "2")
        If inputMin = "" Then Exit Sub  ' usuario cancelou
        If inputMin <> "2" And inputMin <> "3" Then
            MsgBox "Valor fora do intervalo!" & vbCrLf & _
                   "Digite apenas 2 ou 3.", vbExclamation, "Console Flexo"
            Exit Sub
        End If
        minDotLong = CLng(inputMin)
    End If
    Dim srProblemas As ShapeRange: Set srProblemas = CreateShapeRange
    Dim srCorrigidos As ShapeRange: Set srCorrigidos = CreateShapeRange
    Dim s As Shape

    For Each s In ActivePage.shapes
        CrawlerBuscaGradientes s, srProblemas
    Next s

    If srProblemas.Count = 0 Then
        If Not silencioso Then MsgBox "Nenhum preenchimento gradiente (degrad" & ChrW(234) & ") encontrado na p" & ChrW(225) & "gina.", vbInformation, "Console Flexo"
        Exit Sub
    End If

    ' Abre a paleta PANTONE para uso na reconstru��o de cores Spot
    Dim palPantone As Palette
    On Error Resume Next
    Set palPantone = Palettes.OpenFixed(cdrPANTONECoated)
    On Error GoTo 0

    If Not silencioso Then ActiveDocument.BeginCommandGroup "Corrigir Borda Dura Gradientes"

    Dim obj As Shape
    Dim maxC As Long, maxM As Long, maxY As Long, maxK As Long, maxTint As Long
    Dim newC As Long, newM As Long, newY As Long, newK As Long
    Dim mudou As Boolean
    Dim mudouCor As Boolean
    Dim ehBrancoPantone As Boolean
    Dim K As Integer
    Dim temSpot As Boolean
    Dim temBrancoCMYK As Boolean
    Dim nomePantone As String
    Dim idxPantone As Long
    Dim cores() As Color
    Dim totalCores As Integer

    Dim qtdBloqueados As Integer: qtdBloqueados = 0
    For Each obj In srProblemas
        ' [Fix T8] Pula objetos bloqueados -- nao e possivel editar
        If obj.Locked Or (Not obj.Layer.Editable) Then
            qtdBloqueados = qtdBloqueados + 1
            GoTo ProximoObj
        End If
        maxC = 0: maxM = 0: maxY = 0: maxK = 0: maxTint = 0
        temSpot = False: temBrancoCMYK = False
        nomePantone = "": idxPantone = -1
        mudou = False

        On Error Resume Next
        totalCores = 2 + obj.Fill.Fountain.Colors.Count
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo ProximoObj
        End If
        On Error GoTo 0

        ReDim cores(1 To totalCores)
        On Error Resume Next
        Set cores(1) = obj.Fill.Fountain.StartColor
        Set cores(2) = obj.Fill.Fountain.EndColor
        For K = 0 To obj.Fill.Fountain.Colors.Count - 1
            Set cores(3 + K) = obj.Fill.Fountain.Colors(K).Color
        Next K
        On Error GoTo 0

        ' === FASE 1: L� o DNA do gradiente ===
        On Error Resume Next
        For K = 1 To totalCores
            If cores(K).Type = cdrColorCMYK Then
                If cores(K).CMYKCyan > maxC Then maxC = cores(K).CMYKCyan
                If cores(K).CMYKMagenta > maxM Then maxM = cores(K).CMYKMagenta
                If cores(K).CMYKYellow > maxY Then maxY = cores(K).CMYKYellow
                If cores(K).CMYKBlack > maxK Then maxK = cores(K).CMYKBlack
                If (cores(K).CMYKCyan + cores(K).CMYKMagenta + cores(K).CMYKYellow + cores(K).CMYKBlack) = 0 Then
                    temBrancoCMYK = True
                End If
            ElseIf cores(K).Type = cdrColorSpot Then
                If cores(K).Tint > maxTint Then maxTint = cores(K).Tint
                If cores(K).Tint > 0 Then
                    temSpot = True
                    If nomePantone = "" Then
                        nomePantone = cores(K).Name
                        ' Busca o �ndice na paleta para uso posterior
                        If Not palPantone Is Nothing Then
                            Dim p As Long
                            For p = 1 To palPantone.ColorCount
                                If InStr(1, palPantone.Color(p).Name, nomePantone, vbTextCompare) > 0 Then
                                    idxPantone = p
                                    Exit For
                                End If
                            Next p
                        End If
                    End If
                End If
            End If
        Next K
        On Error GoTo 0

        ' === FASE 2: Aplica corre��o n� a n� ===
        ' [T14/T15] On Error Resume Next cobre toda a fase 2
        ' para proteger contra objetos com camada bloqueada ou propriedades somente-leitura
        On Error Resume Next
        For K = 1 To totalCores
            Dim tipoK As Long: tipoK = cores(K).Type
            If Err.Number <> 0 Then Err.Clear: GoTo ProximoNo

            If tipoK = cdrColorCMYK Then
                ' Detecta se este n� � um branco CMYK que � Pantone 0% disfar�ado
                ehBrancoPantone = False
                If temSpot And temBrancoCMYK Then
                    If (cores(K).CMYKCyan + cores(K).CMYKMagenta + cores(K).CMYKYellow + cores(K).CMYKBlack) = 0 Then
                        ehBrancoPantone = True
                    End If
                End If

                If ehBrancoPantone And idxPantone > 0 And Not palPantone Is Nothing Then
                    ' REGRA 2: Reconstr�i n� como Spot usando cor da paleta + Tint
                    If K = 1 Then
                        obj.Fill.Fountain.StartColor.CopyAssign palPantone.Color(idxPantone)
                        obj.Fill.Fountain.StartColor.Tint = minDotLong
                    ElseIf K = 2 Then
                        obj.Fill.Fountain.EndColor.CopyAssign palPantone.Color(idxPantone)
                        obj.Fill.Fountain.EndColor.Tint = minDotLong
                    Else
                        obj.Fill.Fountain.Colors(K - 3).Color.CopyAssign palPantone.Color(idxPantone)
                        obj.Fill.Fountain.Colors(K - 3).Color.Tint = minDotLong
                    End If
                    mudou = True
                ElseIf Not ehBrancoPantone Then
                    ' REGRA 1: Gradiente CMYK puro � zero em canal ativo = borda dura
                    newC = cores(K).CMYKCyan
                    newM = cores(K).CMYKMagenta
                    newY = cores(K).CMYKYellow
                    newK = cores(K).CMYKBlack
                    mudouCor = False
                    ' Corrige zeros em canais ativos (>= 2% em algum n�)
                    If maxC >= 2 And newC = 0 Then newC = minDotLong: mudouCor = True
                    If maxM >= 2 And newM = 0 Then newM = minDotLong: mudouCor = True
                    If maxY >= 2 And newY = 0 Then newY = minDotLong: mudouCor = True
                    If maxK >= 2 And newK = 0 Then newK = minDotLong: mudouCor = True
                    ' Tamb�m corrige valores residuais entre 1 e minDot
                    If maxC >= 2 And newC > 0 And newC < minDotLong Then newC = minDotLong: mudouCor = True
                    If maxM >= 2 And newM > 0 And newM < minDotLong Then newM = minDotLong: mudouCor = True
                    If maxY >= 2 And newY > 0 And newY < minDotLong Then newY = minDotLong: mudouCor = True
                    If maxK >= 2 And newK > 0 And newK < minDotLong Then newK = minDotLong: mudouCor = True
                    If mudouCor Then
                        If K = 1 Then
                            obj.Fill.Fountain.StartColor.CMYKAssign newC, newM, newY, newK
                        ElseIf K = 2 Then
                            obj.Fill.Fountain.EndColor.CMYKAssign newC, newM, newY, newK
                        Else
                            obj.Fill.Fountain.Colors(K - 3).Color.CMYKAssign newC, newM, newY, newK
                        End If
                        mudou = True
                    End If
                End If

            ElseIf cores(K).Type = cdrColorSpot Then
                ' REGRA 3: Spot com Tint abaixo do m�nimo = borda dura
                ' [Fix T8] Tint so funciona em nos Spot
                ' On Error Resume Next pontual protege contra no CMYK disfar�ado
                If cores(K).Tint < minDotLong Then
                    If K = 1 Then
                        On Error Resume Next
                        obj.Fill.Fountain.StartColor.Tint = minDotLong
                        On Error GoTo 0
                    ElseIf K = 2 Then
                        On Error Resume Next
                        obj.Fill.Fountain.EndColor.Tint = minDotLong
                        On Error GoTo 0
                    Else
                        On Error Resume Next
                        obj.Fill.Fountain.Colors(K - 3).Color.Tint = minDotLong
                        On Error GoTo 0
                    End If
                    mudou = True
                End If
            End If

            On Error GoTo 0
ProximoNo:
        Next K
        On Error GoTo 0

        If mudou Then srCorrigidos.Add obj

ProximoObj:
    Next obj

    ' [T14] Se todos os gradientes estavam bloqueados -- aborta sem continuar
    If qtdBloqueados > 0 And srCorrigidos.Count = 0 Then
        If Not silencioso Then ActiveDocument.EndCommandGroup
        If Not silencioso Then Application.Refresh
        If Not silencioso Then
            MsgBox "Aten" & ChrW(231) & ChrW(227) & "o: todos os " & qtdBloqueados & _
                   " gradiente(s) encontrados est" & ChrW(227) & "o BLOQUEADOS." & vbCrLf & vbCrLf & _
                   "Use o bot" & ChrW(227) & "o 'Desbloquear Objetos' no Console Flexo" & _
                   " e execute novamente.", vbExclamation, "Console Flexo"
        End If
        Exit Sub
    End If

    ' Avisa se havia objetos bloqueados misturados com desbloqueados
    If qtdBloqueados > 0 And Not silencioso Then
        MsgBox "Aten" & ChrW(231) & ChrW(227) & "o: " & qtdBloqueados & " gradiente(s) bloqueado(s) foram ignorados." & vbCrLf & _
               "Desbloqueie os objetos e execute novamente.", vbExclamation, "Console Flexo"
    End If

    If Not silencioso Then ActiveDocument.EndCommandGroup
    If Not silencioso Then Application.Refresh

    If Not silencioso Then
        If srCorrigidos.Count > 0 Then
            srCorrigidos.CreateSelection
            MsgBox "Sucesso! " & srCorrigidos.Count & " gradiente(s) tiveram sua quebra m" & ChrW(237) & "nima ajustada para " & minDotLong & "% e est" & ChrW(227) & "o selecionados.", vbInformation, "Console Flexo"
        Else
            MsgBox "Varredura conclu" & ChrW(237) & "da. Todos os degrad" & ChrW(234) & "s j" & ChrW(225) & " possuem pontos m" & ChrW(237) & "nimos suficientes (acima de " & minDotLong & "%).", vbInformation, "Console Flexo"
        End If
    End If
    Exit Sub
End Sub

Private Sub CrawlerBuscaGradientes(s As Shape, ByRef sacola As ShapeRange)
    Dim subS As Shape
    On Error Resume Next
    If s.Fill.Type = cdrFountainFill Then sacola.Add s
    On Error GoTo 0
    If s.Type = cdrGroupShape Then
        For Each subS In s.Shapes: CrawlerBuscaGradientes subS, sacola: Next subS
    End If
    If Not s.PowerClip Is Nothing Then
        For Each subS In s.PowerClip.Shapes: CrawlerBuscaGradientes subS, sacola: Next subS
    End If
End Sub

' ============================================================
' FERRAMENTA: FAXINA DE CORES
' ============================================================
Public Sub LimparSujeiraCores()
    Dim s As Shape
    Dim srProblemas As ShapeRange: Set srProblemas = CreateShapeRange
    
    If MsgBox("Deseja rastrear o documento e ZERAR qualquer canal de cor (CMYK) que esteja com menos de 2%?", _
              vbYesNo + vbQuestion, "Console Flexo") = vbNo Then Exit Sub

    ' ? CORRE��O: agrupa todas as altera��es em 1 �nico Undo
    ActiveDocument.BeginCommandGroup "Limpar Sujeira de Cores"
    On Error GoTo FimErro

    Application.Optimization = True
    For Each s In ActivePage.shapes
        CrawlerFaxinaCores s, srProblemas
    Next s
    Application.Optimization = False
    Application.Refresh

    ActiveDocument.EndCommandGroup

    If srProblemas.Count > 0 Then
        srProblemas.CreateSelection
        MsgBox "Faxina conclu" & ChrW(237) & "da! " & srProblemas.Count & " objeto(s) tiveram sujeiras de cor removidas e est" & ChrW(227) & "o selecionados.", vbInformation, "Console Flexo"
    Else
        MsgBox "Nenhuma sujeira encontrada. As cores j" & ChrW(225) & " est" & ChrW(227) & "o limpas!", vbInformation, "Console Flexo"
    End If
    Exit Sub

FimErro:
    ActiveDocument.EndCommandGroup
    Application.Optimization = False
    Application.Refresh
    MsgBox "Erro ao limpar cores: " & Err.Description, vbCritical, "Console Flexo"
End Sub

Private Sub CrawlerFaxinaCores(s As Shape, ByRef sacola As ShapeRange)
    Dim subS As Shape
    Dim mudouObj As Boolean: mudouObj = False
    On Error Resume Next
    If s.Type <> cdrBitmapShape And s.Type <> cdrGroupShape Then
        If s.Fill.Type = cdrUniformFill Then
            If LimparCanalCMYK(s.Fill.UniformColor) Then mudouObj = True
        End If
        If s.Fill.Type = cdrFountainFill Then
            Dim fC As FountainColor
            For Each fC In s.Fill.Fountain.Colors
                If LimparCanalCMYK(fC.Color) Then mudouObj = True
            Next fC
        End If
        If s.Outline.Type = cdrOutline Then
            If LimparCanalCMYK(s.Outline.Color) Then mudouObj = True
        End If
        If mudouObj Then sacola.Add s
    End If
    If s.Type = cdrGroupShape Then
        For Each subS In s.Shapes: CrawlerFaxinaCores subS, sacola: Next subS
    End If
    If Not s.PowerClip Is Nothing Then
        For Each subS In s.PowerClip.Shapes: CrawlerFaxinaCores subS, sacola: Next subS
    End If
    On Error GoTo 0
End Sub

Private Function LimparCanalCMYK(C As Color) As Boolean
    Dim limpou As Boolean: limpou = False
    If C.Type = cdrColorCMYK Then
        Dim cyan As Long, mag As Long, yel As Long, blk As Long
        cyan = C.CMYKCyan: mag = C.CMYKMagenta: yel = C.CMYKYellow: blk = C.CMYKBlack
        If cyan > 0 And cyan < 2 Then cyan = 0: limpou = True
        If mag > 0 And mag < 2 Then mag = 0: limpou = True
        If yel > 0 And yel < 2 Then yel = 0: limpou = True
        If blk > 0 And blk < 2 Then blk = 0: limpou = True
        If limpou Then C.CMYKAssign cyan, mag, yel, blk
    End If
    LimparCanalCMYK = limpou
End Function
