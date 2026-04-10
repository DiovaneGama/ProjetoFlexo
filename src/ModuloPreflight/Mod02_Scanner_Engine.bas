Attribute VB_Name = "Mod02_Scanner_Engine"
Option Explicit

Public Type RelatorioPreFlight
    QtdBrancoOver As Integer
    QtdPretoSujo As Integer
    QtdRGB As Integer
    QtdPantone As Integer
    BibliotecasPantone As String
    QtdBordaDura As Integer
    QtdRegistro As Integer
    QtdTecnicas As Integer
    BibliotecasTecnicas As String
    QtdLinhasFinas As Integer
    QtdBloqueados As Integer
    QtdInvisiveis As Integer
    QtdImgBaixa As Integer
    QtdImgRGB As Integer
    QtdFontesVivas As Integer
    QtdGradBloqueado As Integer
End Type

Public relatorio As RelatorioPreFlight

Public Sub ExecutarScanner()
    Dim p As Page
    With relatorio
        .QtdBrancoOver = 0: .QtdPretoSujo = 0: .QtdRGB = 0: .QtdPantone = 0: .BibliotecasPantone = ""
        .QtdBordaDura = 0: .QtdRegistro = 0: .QtdTecnicas = 0: .BibliotecasTecnicas = ""
        .QtdLinhasFinas = 0: .QtdBloqueados = 0: .QtdInvisiveis = 0: .QtdImgBaixa = 0
        .QtdImgRGB = 0: .QtdFontesVivas = 0: .QtdGradBloqueado = 0
    End With

    ' ? Varre apenas a página ativa — não todo o documento
    CrawlerMergulhoProfundo ActivePage.shapes

    frmPreFlight.Show vbModeless
End Sub

Private Sub CrawlerMergulhoProfundo(shps As shapes)
    Dim s As Shape
    On Error Resume Next
    For Each s In shps.FindShapes(Recursive:=True)
        If s.Type <> cdrGuidelineShape Then
            If Not s.Layer Is Nothing Then
                If s.Layer.IsSpecialLayer = False And s.Layer.Printable = True Then
                    If s.Locked Or (Not s.Layer.Editable) Then relatorio.QtdBloqueados = relatorio.QtdBloqueados + 1
                    If (Not s.Visible) Or (Not s.Layer.Visible) Then relatorio.QtdInvisiveis = relatorio.QtdInvisiveis + 1

                    If s.Outline.Type = cdrOutline Then
                        Dim espMM As Double
                        Dim fator As Double
                        Select Case ActiveDocument.Unit
                            Case cdrMillimeter: fator = 1
                            Case cdrCentimeter: fator = 10
                            Case cdrInch: fator = 25.4
                            Case Else: fator = 1
                        End Select
                        espMM = Round(s.Outline.Width * fator, 3)

                        If espMM > 0.005 And espMM <= 0.1 Then
                            Dim ehIntencional As Boolean: ehIntencional = False
                            If EhBranco(s.Outline.Color) Then ehIntencional = True
                            If espMM <= 0.05 Then
                                If s.Fill.Type = cdrUniformFill Or _
                                   s.Fill.Type = cdrFountainFill Or _
                                   s.Fill.Type = cdrTextureFill Or _
                                   s.Fill.Type = cdrPatternFill Then
                                    ehIntencional = True
                                End If
                            End If
                            If Not ehIntencional Then relatorio.QtdLinhasFinas = relatorio.QtdLinhasFinas + 1
                        End If

                        AnalisarCor s.Outline.Color, s, True
                    End If

                    If s.Fill.Type = cdrUniformFill Then
                        AnalisarCor s.Fill.UniformColor, s, False
                    ElseIf s.Fill.Type = cdrFountainFill Then
                        AnalisarGradiente s
                    End If

                    If s.Type = cdrBitmapShape Then
                        If s.Bitmap.ResolutionX < 300 Or s.Bitmap.ResolutionY < 300 Then relatorio.QtdImgBaixa = relatorio.QtdImgBaixa + 1
                        If s.Bitmap.Mode <> cdrCMYKColorImage And s.Bitmap.Mode <> cdrGrayscaleImage And s.Bitmap.Mode <> cdrBlackAndWhiteImage Then
                            relatorio.QtdImgRGB = relatorio.QtdImgRGB + 1
                        End If
                    End If

                    If s.Type = cdrTextShape Then relatorio.QtdFontesVivas = relatorio.QtdFontesVivas + 1
                    If Not s.PowerClip Is Nothing Then CrawlerMergulhoProfundo s.PowerClip.shapes
                End If
            End If
        End If
    Next s
    On Error GoTo 0
End Sub

Private Sub AnalisarCor(C As Color, s As Shape, isOutline As Boolean)
    On Error Resume Next
    Dim ehRegistro As Boolean: ehRegistro = False
    If C.Type = cdrColorRegistration Then
        ehRegistro = True
    Else
        Dim nCor As String: nCor = LCase(Trim(C.Name))
        If InStr(nCor, "registration") > 0 Or InStr(nCor, "registro") > 0 Or nCor = "all" Then ehRegistro = True
    End If
    If ehRegistro Then
        relatorio.QtdRegistro = relatorio.QtdRegistro + 1
        Exit Sub
    End If

    Dim ehBrancoCheck As Boolean: ehBrancoCheck = False
    If C.Type = cdrColorCMYK Then
        If (C.CMYKCyan + C.CMYKMagenta + C.CMYKYellow + C.CMYKBlack) = 0 Then ehBrancoCheck = True
    ElseIf C.Type = cdrColorRGB Then
        If C.RGBRed = 255 And C.RGBGreen = 255 And C.RGBBlue = 255 Then ehBrancoCheck = True
    ElseIf C.IsSpot Then
        If C.Tint = 0 Then ehBrancoCheck = True
    End If
    If ehBrancoCheck Then
        If (Not isOutline And s.OverprintFill) Or (isOutline And s.OverprintOutline) Then
            relatorio.QtdBrancoOver = relatorio.QtdBrancoOver + 1
        End If
    End If

    ' PANTONE - conta e lista cores únicas, excluindo cores técnicas
    If C.IsSpot Then
        Dim nomeCor As String: nomeCor = Trim(C.Name)
        If nomeCor <> "" Then
            If Not EhCorTecnica(LCase(nomeCor)) Then
                If InStr(1, relatorio.BibliotecasPantone, nomeCor, vbTextCompare) = 0 Then
                    relatorio.QtdPantone = relatorio.QtdPantone + 1
                    relatorio.BibliotecasPantone = relatorio.BibliotecasPantone & nomeCor & "|"
                End If
            End If
        End If
    End If

    ' CORES TÉCNICAS - conta e lista cores únicas
    Dim nTec As String: nTec = LCase(Trim(C.Name))
    If EhCorTecnica(nTec) Then
        Dim nomeTec As String: nomeTec = Trim(C.Name)
        If nomeTec <> "" Then
            If InStr(1, relatorio.BibliotecasTecnicas, nomeTec, vbTextCompare) = 0 Then
                relatorio.QtdTecnicas = relatorio.QtdTecnicas + 1
                relatorio.BibliotecasTecnicas = relatorio.BibliotecasTecnicas & nomeTec & "|"
            End If
        End If
    End If

    If C.Type = cdrColorCMYK Then
        If (DimSomaCMY(C) > 200) And C.CMYKBlack > 80 Then relatorio.QtdPretoSujo = relatorio.QtdPretoSujo + 1
    ElseIf C.Type = cdrColorRGB Then
        relatorio.QtdRGB = relatorio.QtdRGB + 1
    End If
    On Error GoTo 0
End Sub

Private Sub AnalisarGradiente(s As Shape)
    Dim K As Integer
    Dim temBordaDura As Boolean: temBordaDura = False
    On Error Resume Next
    Dim maxC As Long, maxM As Long, maxY As Long, maxK As Long, maxTint As Long
    maxC = 0: maxM = 0: maxY = 0: maxK = 0: maxTint = 0
    Dim temSpot As Boolean: temSpot = False
    Dim temBrancoCMYK As Boolean: temBrancoCMYK = False
    Dim cores() As Color
    Dim totalCores As Integer
    totalCores = 2 + s.Fill.Fountain.Colors.Count
    ReDim cores(1 To totalCores)
    Set cores(1) = s.Fill.Fountain.StartColor
    Set cores(2) = s.Fill.Fountain.EndColor
    For K = 0 To s.Fill.Fountain.Colors.Count - 1
        Set cores(3 + K) = s.Fill.Fountain.Colors(K).Color
    Next K
    For K = 1 To totalCores
        If cores(K).Type = cdrColorCMYK Then
            If cores(K).CMYKCyan > maxC Then maxC = cores(K).CMYKCyan
            If cores(K).CMYKMagenta > maxM Then maxM = cores(K).CMYKMagenta
            If cores(K).CMYKYellow > maxY Then maxY = cores(K).CMYKYellow
            If cores(K).CMYKBlack > maxK Then maxK = cores(K).CMYKBlack
            If (cores(K).CMYKCyan + cores(K).CMYKMagenta + cores(K).CMYKYellow + cores(K).CMYKBlack) = 0 Then temBrancoCMYK = True
        ElseIf cores(K).Type = cdrColorSpot Then
            If cores(K).Tint > maxTint Then maxTint = cores(K).Tint
            If cores(K).Tint > 0 Then temSpot = True
        End If
    Next K
    For K = 1 To totalCores
        If cores(K).Type = cdrColorCMYK Then
            If maxC > 0 And cores(K).CMYKCyan = 0 Then temBordaDura = True: Exit For
            If maxM > 0 And cores(K).CMYKMagenta = 0 Then temBordaDura = True: Exit For
            If maxY > 0 And cores(K).CMYKYellow = 0 Then temBordaDura = True: Exit For
            If maxK > 0 And cores(K).CMYKBlack = 0 Then temBordaDura = True: Exit For
        ElseIf cores(K).Type = cdrColorSpot Then
            If maxTint > 0 And cores(K).Tint = 0 Then temBordaDura = True: Exit For
        End If
    Next K
    If Not temBordaDura Then
        If temSpot And temBrancoCMYK Then temBordaDura = True
    End If
    If temBordaDura Then relatorio.QtdBordaDura = relatorio.QtdBordaDura + 1
    On Error GoTo 0
End Sub

Private Function DimSomaCMY(C As Color) As Long
    DimSomaCMY = CLng(C.CMYKCyan) + CLng(C.CMYKMagenta) + CLng(C.CMYKYellow)
End Function

Private Function EhBranco(C As Color) As Boolean
    EhBranco = False
    If C.Type = cdrColorCMYK Then
        If (C.CMYKCyan + C.CMYKMagenta + C.CMYKYellow + C.CMYKBlack) = 0 Then EhBranco = True
    ElseIf C.Type = cdrColorRGB Then
        If C.RGBRed = 255 And C.RGBGreen = 255 And C.RGBBlue = 255 Then EhBranco = True
    ElseIf C.IsSpot Then
        If C.Tint = 0 Then EhBranco = True
    End If
End Function

Private Function EhCorTecnica(nCorLower As String) As Boolean
    EhCorTecnica = False
    If InStr(nCorLower, "faca") > 0 Or _
       InStr(nCorLower, "corte") > 0 Or _
       InStr(nCorLower, "cutting") > 0 Or _
       InStr(nCorLower, "creasing") > 0 Or _
       InStr(nCorLower, "verniz") > 0 Or _
       InStr(nCorLower, "varnish") > 0 Or _
       InStr(nCorLower, "foil") > 0 Or _
       InStr(nCorLower, "embossing") > 0 Or _
       InStr(nCorLower, "debossing") > 0 Or _
       InStr(nCorLower, "braille") > 0 Or _
       InStr(nCorLower, "bleed") > 0 Or _
       InStr(nCorLower, "punching") > 0 Or _
       InStr(nCorLower, "perforating") > 0 Or _
       InStr(nCorLower, "folding") > 0 Or _
       InStr(nCorLower, "gluing") > 0 Or _
       InStr(nCorLower, "stapling") > 0 Or _
       InStr(nCorLower, "drilling") > 0 Or _
       InStr(nCorLower, "hologram") > 0 Or _
       (nCorLower = "white") Then
        EhCorTecnica = True
    End If
End Function

' ============================================================
' Função auxiliar: busca índice de uma cor Spot na paleta PANTONE
' ============================================================
Private Function BuscarIndicePaleta(paleta As Palette, nomeCor As String) As Long
    BuscarIndicePaleta = -1
    If paleta Is Nothing Then Exit Function
    Dim i As Long
    On Error Resume Next
    For i = 1 To paleta.ColorCount
        If InStr(1, paleta.Color(i).Name, nomeCor, vbTextCompare) > 0 Then
            BuscarIndicePaleta = i
            Exit For
        End If
    Next i
    On Error GoTo 0
End Function

Public Sub ExecutarCorrecoes(ByVal minDot As Integer)
    ActiveDocument.BeginCommandGroup "Corre" & ChrW(231) & ChrW(227) & "o Autom" & ChrW(225) & "tica PreFlight"
    On Error GoTo FimErro

    ' ? Corrige apenas a página ativa
    CrawlerCorrecoes ActivePage.shapes, minDot

FimErro:
    ActiveDocument.EndCommandGroup
End Sub

Private Sub CrawlerCorrecoes(shps As shapes, ByVal minDot As Integer)
    Dim s As Shape
    On Error Resume Next
    For Each s In shps.FindShapes(Recursive:=True)
        If s.Type = cdrGroupShape Then GoTo ProximoShape
        If s.Type <> cdrGuidelineShape And Not s.Layer Is Nothing Then
            If s.Layer.IsSpecialLayer = False And s.Layer.Printable = True And Not s.Locked And s.Layer.Editable Then

                ' 1. CORREÇÃO DE PREENCHIMENTOS UNIFORMES
                If s.Fill.Type = cdrUniformFill Then
                    If s.Fill.UniformColor.Type = cdrColorRGB Then s.Fill.UniformColor.ConvertToCMYK
                    If EhBranco(s.Fill.UniformColor) And s.OverprintFill = True Then s.OverprintFill = False
                    If s.Fill.UniformColor.Type = cdrColorCMYK Then
                        If (DimSomaCMY(s.Fill.UniformColor) > 200) And s.Fill.UniformColor.CMYKBlack > 80 Then
                            s.Fill.UniformColor.CMYKCyan = 0: s.Fill.UniformColor.CMYKMagenta = 0
                            s.Fill.UniformColor.CMYKYellow = 0: s.Fill.UniformColor.CMYKBlack = 100
                        End If
                    End If
                End If

                ' 2. CORREÇÃO DE GRADIENTES (BORDA DURA)
                If s.Fill.Type = cdrFountainFill And minDot > 0 Then
                    Dim maxC As Long, maxM As Long, maxY As Long, maxK As Long, maxTint As Long
                    Dim newC As Long, newM As Long, newY As Long, newK As Long
                    Dim K As Integer
                    Dim mudouCor As Boolean
                    Dim ehBrancoPantone As Boolean
                    Dim temSpot As Boolean: temSpot = False
                    Dim temBrancoCMYK As Boolean: temBrancoCMYK = False
                    Dim nomePantone As String: nomePantone = ""
                    Dim idxPantone As Long: idxPantone = -1
                    maxC = 0: maxM = 0: maxY = 0: maxK = 0: maxTint = 0
                    Dim cores() As Color
                    Dim totalCores As Integer
                    totalCores = 2 + s.Fill.Fountain.Colors.Count
                    ReDim cores(1 To totalCores)
                    Set cores(1) = s.Fill.Fountain.StartColor
                    Set cores(2) = s.Fill.Fountain.EndColor
                    For K = 0 To s.Fill.Fountain.Colors.Count - 1
                        Set cores(3 + K) = s.Fill.Fountain.Colors(K).Color
                    Next K

                    ' FASE 1: Lê o DNA do gradiente
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
                                    ' Busca índice na paleta PANTONE para reconstrução Spot
                                    Dim palPantone As Palette
                                    Set palPantone = Palettes.OpenFixed(cdrPANTONECoated)
                                    idxPantone = BuscarIndicePaleta(palPantone, nomePantone)
                                End If
                            End If
                        End If
                    Next K

                    ' FASE 2: Aplica correção nó a nó
                    For K = 1 To totalCores
                        If cores(K).Type = cdrColorCMYK Then
                            ' Detecta branco CMYK que é Pantone 0% disfarçado
                            ehBrancoPantone = False
                            If temSpot And temBrancoCMYK Then
                                If (cores(K).CMYKCyan + cores(K).CMYKMagenta + cores(K).CMYKYellow + cores(K).CMYKBlack) = 0 Then
                                    ehBrancoPantone = True
                                End If
                            End If

                            If ehBrancoPantone And idxPantone > 0 And Not palPantone Is Nothing Then
                                ' REGRA 2: Reconstrói nó como Spot via paleta + Tint
                                If K = 1 Then
                                    s.Fill.Fountain.StartColor.CopyAssign palPantone.Color(idxPantone)
                                    s.Fill.Fountain.StartColor.Tint = minDot
                                ElseIf K = 2 Then
                                    s.Fill.Fountain.EndColor.CopyAssign palPantone.Color(idxPantone)
                                    s.Fill.Fountain.EndColor.Tint = minDot
                                Else
                                    s.Fill.Fountain.Colors(K - 3).Color.CopyAssign palPantone.Color(idxPantone)
                                    s.Fill.Fountain.Colors(K - 3).Color.Tint = minDot
                                End If
                            ElseIf Not ehBrancoPantone Then
                                ' REGRA 1: Gradiente CMYK — zero em canal ativo = borda dura
                                newC = cores(K).CMYKCyan
                                newM = cores(K).CMYKMagenta
                                newY = cores(K).CMYKYellow
                                newK = cores(K).CMYKBlack
                                mudouCor = False
                                If maxC >= 2 And newC = 0 Then newC = minDot: mudouCor = True
                                If maxM >= 2 And newM = 0 Then newM = minDot: mudouCor = True
                                If maxY >= 2 And newY = 0 Then newY = minDot: mudouCor = True
                                If maxK >= 2 And newK = 0 Then newK = minDot: mudouCor = True
                                If maxC >= 2 And newC > 0 And newC < minDot Then newC = minDot: mudouCor = True
                                If maxM >= 2 And newM > 0 And newM < minDot Then newM = minDot: mudouCor = True
                                If maxY >= 2 And newY > 0 And newY < minDot Then newY = minDot: mudouCor = True
                                If maxK >= 2 And newK > 0 And newK < minDot Then newK = minDot: mudouCor = True
                                If mudouCor Then
                                    If K = 1 Then
                                        s.Fill.Fountain.StartColor.CMYKAssign newC, newM, newY, newK
                                    ElseIf K = 2 Then
                                        s.Fill.Fountain.EndColor.CMYKAssign newC, newM, newY, newK
                                    Else
                                        s.Fill.Fountain.Colors(K - 3).Color.CMYKAssign newC, newM, newY, newK
                                    End If
                                End If
                            End If

                        ElseIf cores(K).Type = cdrColorSpot Then
                            ' REGRA 3: Spot com Tint abaixo do mínimo = borda dura
                            If cores(K).Tint < minDot Then
                                If K = 1 Then
                                    s.Fill.Fountain.StartColor.Tint = minDot
                                ElseIf K = 2 Then
                                    s.Fill.Fountain.EndColor.Tint = minDot
                                Else
                                    s.Fill.Fountain.Colors(K - 3).Color.Tint = minDot
                                End If
                            End If
                        End If
                    Next K
                End If

                ' 3. CORREÇÃO DE CONTORNOS E LINHAS FINAS
                If s.Outline.Type = cdrOutline Then
                    If s.Outline.Color.Type = cdrColorRGB Then s.Outline.Color.ConvertToCMYK
                    If EhBranco(s.Outline.Color) And s.OverprintOutline = True Then s.OverprintOutline = False
                    If s.Outline.Color.Type = cdrColorCMYK Then
                        If (DimSomaCMY(s.Outline.Color) > 200) And s.Outline.Color.CMYKBlack > 80 Then
                            s.Outline.Color.CMYKCyan = 0: s.Outline.Color.CMYKMagenta = 0
                            s.Outline.Color.CMYKYellow = 0: s.Outline.Color.CMYKBlack = 100
                        End If
                    End If
                    Dim espMM As Double: Dim fator As Double
                    Select Case ActiveDocument.Unit
                        Case cdrMillimeter: fator = 1
                        Case cdrCentimeter: fator = 10
                        Case cdrInch: fator = 25.4
                        Case Else: fator = 1
                    End Select
                    espMM = Round(s.Outline.Width * fator, 3)
                    If espMM > 0.005 And espMM <= 0.1 Then
                        Dim corrigirLinha As Boolean: corrigirLinha = True
                        If EhBranco(s.Outline.Color) Then corrigirLinha = False
                        If espMM <= 0.05 Then
                            If s.Fill.Type = cdrUniformFill Or _
                               s.Fill.Type = cdrFountainFill Or _
                               s.Fill.Type = cdrTextureFill Or _
                               s.Fill.Type = cdrPatternFill Then
                                corrigirLinha = False
                            End If
                        End If
                        If corrigirLinha Then s.Outline.Width = 0.2 / fator
                    End If
                End If

                ' 4. CORREÇÃO DE BITMAPS E IMAGENS
                If s.Type = cdrBitmapShape Then
                    If s.Bitmap.Mode <> cdrCMYKColorImage And s.Bitmap.Mode <> cdrGrayscaleImage And s.Bitmap.Mode <> cdrBlackAndWhiteImage Then
                        s.Bitmap.ConvertTo cdrCMYKColorImage
                    End If
                    If s.Bitmap.ResolutionX < 299 Or s.Bitmap.ResolutionY < 299 Then
                        Dim imgX As Double, imgY As Double, imgW As Double, imgH As Double
                        s.GetBoundingBox imgX, imgY, imgW, imgH
                        Dim pxW As Long, pxH As Long
                        pxW = Round(s.Bitmap.Width * (300 / s.Bitmap.ResolutionX))
                        pxH = Round(s.Bitmap.Height * (300 / s.Bitmap.ResolutionY))
                        s.Bitmap.Resample pxW, pxH, False, 300, 300
                        s.SetBoundingBox imgX, imgY, imgW, imgH
                    End If
                End If

                ' 5. FONTES VIVAS
                If s.Type = cdrTextShape Then s.ConvertToCurves

                ' 6. MERGULHO
                If Not s.PowerClip Is Nothing Then CrawlerCorrecoes s.PowerClip.shapes, minDot

            End If
        End If
ProximoShape:
    Next s
    On Error GoTo 0
End Sub

