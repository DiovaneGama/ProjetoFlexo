Attribute VB_Name = "GerarArquivosTeste"
Option Explicit

' ============================================================
' GERADOR DE ARQUIVOS DE TESTE — Console Flexo v2.0
' Execute: GerarTodosArquivos
' Sintaxe ApplyFountainFill baseada na API oficial CorelDRAW:
' ApplyFountainFill(StartColor, EndColor, Type, Angle, Steps,
'                  EdgePad, MidPoint, BlendType, CenterX, CenterY)
' ============================================================

Public Sub GerarTodosArquivos()
    Dim resp As Integer
    resp = MsgBox("Gerar arquivos de teste para o Console Flexo v2.0?" & vbCrLf & vbCrLf & _
                  "Serao criados 4 arquivos na sua Area de Trabalho:" & vbCrLf & _
                  "  - Arquivo_A_Cores_e_Gradientes.cdr" & vbCrLf & _
                  "  - Arquivo_B_Contornos_e_Vetores.cdr" & vbCrLf & _
                  "  - Arquivo_C_Bitmaps.cdr" & vbCrLf & _
                  "  - Arquivo_D_Montagem.cdr", _
                  vbYesNo + vbQuestion, "Console Flexo - Gerar Testes")

    If resp = vbNo Then Exit Sub

    Dim desktop As String
    desktop = Environ("USERPROFILE") & "\Desktop\ConsoleFlexo_Testes\"
    On Error Resume Next
    MkDir desktop
    On Error GoTo 0

    GerarArquivoA desktop
    GerarArquivoB desktop
    GerarArquivoC desktop
    GerarArquivoD desktop

    MsgBox "Arquivos de teste criados com sucesso!" & vbCrLf & _
           "Local: " & desktop, vbInformation, "Console Flexo"
End Sub

' ============================================================
' HELPERS DE COR
' ============================================================
Private Function CorCMYK(C As Long, M As Long, Y As Long, K As Long) As Color
    Dim cor As New Color
    cor.CMYKAssign C, M, Y, K
    Set CorCMYK = cor
End Function

Private Function CorSpot(nome As String) As Color
    Dim cor As New Color
    On Error Resume Next
    Dim pal As Palette
    Set pal = Palettes.OpenFixed(cdrPANTONECoated)
    If Not pal Is Nothing Then
        Dim idx As Long
        idx = pal.FindColor(nome)
        If idx > 0 Then
            cor.CopyAssign pal.Color(idx)
        Else
            cor.CMYKAssign 0, 100, 100, 0 ' Fallback vermelho CMYK
        End If
    Else
        cor.CMYKAssign 0, 100, 100, 0
    End If
    On Error GoTo 0
    Set CorSpot = cor
End Function

Private Sub Label(ly As Layer, x As Double, y As Double, texto As String)
    On Error Resume Next
    Dim lbl As Shape
    Dim corAzul As New Color
    corAzul.CMYKAssign 100, 0, 0, 0
    Set lbl = ly.CreateArtisticText(x, y, texto, cdrLanguageNone, , "Arial", 6)
    lbl.Fill.ApplyUniformFill corAzul
    lbl.Outline.SetNoOutline
    On Error GoTo 0
End Sub

' ============================================================
' ARQUIVO A — Cores e Gradientes
' ============================================================
Private Sub GerarArquivoA(pasta As String)
    Dim doc As Document
    Set doc = CreateDocument()
    doc.Unit = cdrMillimeter
    doc.ActivePage.SetSize 297, 210

    Dim ly As Layer
    Set ly = doc.ActivePage.ActiveLayer

    Dim corPreta As New Color: corPreta.CMYKAssign 0, 0, 0, 100
    Dim corBranca As New Color: corBranca.CMYKAssign 0, 0, 0, 0
    Dim s As Shape

    ' A01 - RGB vermelho
    Set s = ly.CreateRectangle2(10, 155, 40, 40)
    Dim corRGB As New Color: corRGB.RGBAssign 255, 0, 0
    s.Fill.ApplyUniformFill corRGB
    s.Outline.SetNoOutline
    Label ly, 10, 152, "A01 - RGB Vermelho"

    ' A02 - Pantone 485 C
    Set s = ly.CreateRectangle2(60, 155, 40, 40)
    s.Fill.ApplyUniformFill CorSpot("PANTONE 485 C")
    s.Outline.SetNoOutline
    Label ly, 60, 152, "A02 - Pantone 485 C"

    ' A03 - Branco CMYK com Overprint
    Set s = ly.CreateRectangle2(110, 155, 40, 40)
    s.Fill.ApplyUniformFill corBranca
    s.OverprintFill = True
    s.Outline.SetProperties 0.5, , corPreta
    Label ly, 110, 152, "A03 - Branco Overprint"

    ' A04 - Preto Rico
    Set s = ly.CreateRectangle2(160, 155, 40, 40)
    s.Fill.ApplyUniformFill CorCMYK(100, 100, 100, 100)
    s.Outline.SetNoOutline
    Label ly, 160, 152, "A04 - Preto Rico"

    ' A05 - Preto Sujo
    Set s = ly.CreateRectangle2(210, 155, 40, 40)
    s.Fill.ApplyUniformFill CorCMYK(30, 20, 10, 90)
    s.Outline.SetNoOutline
    Label ly, 210, 152, "A05 - Preto Sujo"

    ' A06 - Gradiente CMYK borda dura C100 -> C0
    ' ApplyFountainFill(StartColor, EndColor, Type, Angle, Steps, EdgePad, MidPoint, BlendType)
    Set s = ly.CreateRectangle2(10, 95, 70, 45)
    s.Fill.ApplyFountainFill CorCMYK(100, 0, 0, 0), CorCMYK(0, 0, 0, 0), _
        cdrLinearFountainFill, 0, 0, 0, 50, cdrDirectFountainFillBlend
    s.Outline.SetNoOutline
    Label ly, 10, 92, "A06 - Gradiente Borda Dura CMYK (C100->C0)"

    ' A07 - Gradiente Pantone para branco CMYK
    Set s = ly.CreateRectangle2(90, 95, 70, 45)
    s.Fill.ApplyFountainFill CorSpot("PANTONE 485 C"), CorCMYK(0, 0, 0, 0), _
        cdrLinearFountainFill, 0, 0, 0, 50, cdrDirectFountainFillBlend
    s.Outline.SetNoOutline
    Label ly, 90, 92, "A07 - Pantone -> Branco CMYK"

    ' A08 - Gradiente CMYK correto C100 -> C2
    Set s = ly.CreateRectangle2(170, 95, 70, 45)
    s.Fill.ApplyFountainFill CorCMYK(100, 0, 0, 0), CorCMYK(2, 0, 0, 0), _
        cdrLinearFountainFill, 0, 0, 0, 50, cdrDirectFountainFillBlend
    s.Outline.SetNoOutline
    Label ly, 170, 92, "A08 - Gradiente Correto (C100->C2)"

    ' A09 - Cor de Registro
    Set s = ly.CreateRectangle2(10, 40, 70, 45)
    Dim corReg As New Color: corReg.RegistrationAssign
    s.Fill.ApplyUniformFill corReg
    s.Outline.SetNoOutline
    Label ly, 10, 37, "A09 - Cor de Registro"

    ' A10 - Sujeira de cor C1 K100
    Set s = ly.CreateRectangle2(90, 40, 70, 45)
    s.Fill.ApplyUniformFill CorCMYK(1, 0, 0, 100)
    s.Outline.SetNoOutline
    Label ly, 90, 37, "A10 - Sujeira C1 K100"

    ' A11 - Gradiente CMYK borda dura BLOQUEADO
    Set s = ly.CreateRectangle2(170, 40, 70, 45)
    s.Fill.ApplyFountainFill CorCMYK(100, 0, 0, 0), CorCMYK(0, 0, 0, 0), _
        cdrLinearFountainFill, 0, 0, 0, 50, cdrDirectFountainFillBlend
    s.Outline.SetNoOutline
    s.Locked = True
    Label ly, 170, 37, "A11 - Gradiente BLOQUEADO"

    ' Titulo
    Dim tit As Shape
    Set tit = ly.CreateArtisticText(10, 207, "ARQUIVO A - Cores e Gradientes | Console Flexo v2.0", _
        cdrLanguageNone, , "Arial", 9, cdrTrue)
    tit.Fill.ApplyUniformFill corPreta
    tit.Outline.SetNoOutline

    doc.SaveAs pasta & "Arquivo_A_Cores_e_Gradientes.cdr"
    doc.Close
    MsgBox "Arquivo A gerado!", vbInformation, "Console Flexo"
End Sub

' ============================================================
' ARQUIVO B — Contornos e Vetores
' ============================================================
Private Sub GerarArquivoB(pasta As String)
    Dim doc As Document
    Set doc = CreateDocument()
    doc.Unit = cdrMillimeter
    doc.ActivePage.SetSize 297, 210

    Dim ly As Layer
    Set ly = doc.ActivePage.ActiveLayer

    Dim corPreta As New Color: corPreta.CMYKAssign 0, 0, 0, 100
    Dim corBranca As New Color: corBranca.CMYKAssign 0, 0, 0, 0
    Dim s As Shape

    ' B01 - Contorno BRANCO 0.05mm sem preenchimento
    Set s = ly.CreateRectangle2(10, 155, 50, 40)
    s.Fill.ApplyNoFill
    s.Outline.SetProperties 0.05, , corBranca
    Label ly, 10, 152, "B01 - Contorno Branco 0.05mm (intencional)"

    ' B02 - Contorno PRETO 0.05mm sem preenchimento
    Set s = ly.CreateRectangle2(70, 155, 50, 40)
    s.Fill.ApplyNoFill
    s.Outline.SetProperties 0.05, , corPreta
    Label ly, 70, 152, "B02 - Contorno Preto 0.05mm"

    ' B03 - Contorno PRETO 0.05mm COM preenchimento
    Set s = ly.CreateRectangle2(130, 155, 50, 40)
    s.Fill.ApplyUniformFill CorCMYK(0, 100, 100, 0)
    s.Outline.SetProperties 0.05, , corPreta
    Label ly, 130, 152, "B03 - Preto 0.05 + Preenchimento"

    ' B04 - Contorno 0.08mm
    Set s = ly.CreateRectangle2(190, 155, 50, 40)
    s.Fill.ApplyNoFill
    s.Outline.SetProperties 0.08, , corPreta
    Label ly, 190, 152, "B04 - Contorno Preto 0.08mm"

    ' B05 - Objeto simulando Ctrl+Shift+Q (retangulo muito fino)
    Set s = ly.CreateRectangle2(10, 105, 120, 0.05)
    s.Fill.ApplyUniformFill corPreta
    s.Outline.SetNoOutline
    Label ly, 10, 102, "B05 - Objeto convertido (0.05mm altura = Ctrl+Shift+Q)"

    ' B06 - Texto nao convertido (fonte viva)
    Set s = ly.CreateArtisticText(10, 75, "TEXTO NAO CONVERTIDO - B06", _
        cdrLanguageNone, , "Arial", 14, cdrTrue)
    s.Fill.ApplyUniformFill corPreta
    s.Outline.SetNoOutline
    Label ly, 10, 60, "B06 - Fonte Viva (nao converter em curvas)"

    ' B07 - Texto JA convertido em curvas
    Set s = ly.CreateArtisticText(10, 45, "TEXTO EM CURVAS - B07", _
        cdrLanguageNone, , "Arial", 14, cdrTrue)
    s.Fill.ApplyUniformFill CorCMYK(100, 0, 0, 0)
    s.Outline.SetNoOutline
    s.ConvertToCurves
    Label ly, 10, 30, "B07 - Ja convertido em curvas"

    ' B08 - Branco 0.05mm COM preenchimento (NAO deve detectar)
    Set s = ly.CreateRectangle2(190, 105, 50, 40)
    s.Fill.ApplyUniformFill CorCMYK(0, 100, 100, 0)
    s.Outline.SetProperties 0.05, , corBranca
    Label ly, 190, 102, "B08 - Branco 0.05 + Preench. (NAO detectar)"

    ' Titulo
    Dim tit As Shape
    Set tit = ly.CreateArtisticText(10, 207, "ARQUIVO B - Contornos e Vetores | Console Flexo v2.0", _
        cdrLanguageNone, , "Arial", 9, cdrTrue)
    tit.Fill.ApplyUniformFill corPreta
    tit.Outline.SetNoOutline

    doc.SaveAs pasta & "Arquivo_B_Contornos_e_Vetores.cdr"
    doc.Close
    MsgBox "Arquivo B gerado!", vbInformation, "Console Flexo"
End Sub

' ============================================================
' ARQUIVO C — Bitmaps (areas marcadas para importacao manual)
' ============================================================
Private Sub GerarArquivoC(pasta As String)
    Dim doc As Document
    Set doc = CreateDocument()
    doc.Unit = cdrMillimeter
    doc.ActivePage.SetSize 297, 210

    Dim ly As Layer
    Set ly = doc.ActivePage.ActiveLayer

    Dim corPreta As New Color: corPreta.CMYKAssign 0, 0, 0, 100
    Dim corCinza As New Color: corCinza.CMYKAssign 0, 0, 0, 20
    Dim corAzul As New Color: corAzul.CMYKAssign 100, 0, 0, 0

    ' Titulo
    Dim tit As Shape
    Set tit = ly.CreateArtisticText(10, 207, "ARQUIVO C - Bitmaps | Console Flexo v2.0", _
        cdrLanguageNone, , "Arial", 9, cdrTrue)
    tit.Fill.ApplyUniformFill corPreta
    tit.Outline.SetNoOutline

    ' Instrucao
    Dim inst As Shape
    Set inst = ly.CreateArtisticText(10, 195, _
        "INSTRUCAO: Importe manualmente as imagens nas areas indicadas abaixo.", _
        cdrLanguageNone, , "Arial", 8)
    inst.Fill.ApplyUniformFill corAzul
    inst.Outline.SetNoOutline

    ' 4 areas para importacao
    Dim areas(3) As String
    Dim descricoes(3) As String
    areas(0) = "C01"
    descricoes(0) = "C01 - Importe aqui: Imagem CMYK 600 DPI (padrao correto - NAO deve alterar)"
    areas(1) = "C02"
    descricoes(1) = "C02 - Importe aqui: Imagem RGB 72 DPI (deve converter para CMYK 600 DPI)"
    areas(2) = "C03"
    descricoes(2) = "C03 - Importe aqui: Imagem Grayscale 300 DPI (padrao correto - NAO deve alterar)"
    areas(3) = "C04"
    descricoes(3) = "C04 - Importe aqui: Imagem CMYK 150 DPI (deve reamostrar para 600 DPI)"

    Dim i As Integer
    For i = 0 To 3
        Dim col As Double: col = 10 + (i Mod 2) * 148
        Dim row As Double: row = 155 - (i \ 2) * 85

        Dim box As Shape
        Set box = ly.CreateRectangle2(col, row, 135, 70)
        box.Fill.ApplyNoFill
        box.Outline.SetProperties 0.5, , corCinza

        Dim lbl As Shape
        Set lbl = ly.CreateArtisticText(col + 3, row + 66, descricoes(i), _
            cdrLanguageNone, , "Arial", 6)
        lbl.Fill.ApplyUniformFill corPreta
        lbl.Outline.SetNoOutline

        Dim id As Shape
        Set id = ly.CreateArtisticText(col + 55, row + 35, areas(i), _
            cdrLanguageNone, , "Arial", 16, cdrTrue)
        id.Fill.ApplyUniformFill corCinza
        id.Outline.SetNoOutline
    Next i

    doc.SaveAs pasta & "Arquivo_C_Bitmaps.cdr"
    doc.Close
    MsgBox "Arquivo C gerado!", vbInformation, "Console Flexo"
End Sub

' ============================================================
' ARQUIVO D — Montagem (TrimBox e Camerom)
' ============================================================
Private Sub GerarArquivoD(pasta As String)
    Dim doc As Document
    Set doc = CreateDocument()
    doc.Unit = cdrMillimeter
    doc.ActivePage.SetSize 297, 210

    Dim ly As Layer
    Set ly = doc.ActivePage.ActiveLayer

    Dim corPreta As New Color: corPreta.CMYKAssign 0, 0, 0, 100
    Dim corCiano As New Color: corCiano.CMYKAssign 100, 0, 0, 0
    Dim corMag As New Color: corMag.CMYKAssign 0, 100, 0, 0
    Dim s As Shape

    ' Titulo
    Dim tit As Shape
    Set tit = ly.CreateArtisticText(10, 207, "ARQUIVO D - Montagem | Console Flexo v2.0", _
        cdrLanguageNone, , "Arial", 9, cdrTrue)
    tit.Fill.ApplyUniformFill corPreta
    tit.Outline.SetNoOutline

    ' --- D01: Arte com multiplos objetos ---
    Dim inst1 As Shape
    Set inst1 = ly.CreateArtisticText(10, 185, _
        "D01 - Selecione os 3 objetos abaixo para testar TrimBox com multiplos elementos:", _
        cdrLanguageNone, , "Arial", 7)
    inst1.Fill.ApplyUniformFill corCiano
    inst1.Outline.SetNoOutline

    ' Objeto 1 do grupo D01
    Set s = ly.CreateRectangle2(10, 130, 80, 50)
    s.Fill.ApplyUniformFill CorCMYK(0, 100, 100, 0)
    s.Outline.SetNoOutline

    ' Objeto 2 do grupo D01
    Set s = ly.CreateRectangle2(30, 140, 40, 20)
    s.Fill.ApplyUniformFill corPreta
    s.Outline.SetNoOutline

    ' Objeto 3 do grupo D01 (texto)
    Set s = ly.CreateArtisticText(15, 170, "ARTE D01", cdrLanguageNone, , "Arial", 8, cdrTrue)
    s.Fill.ApplyUniformFill corPreta
    s.Outline.SetNoOutline
    s.ConvertToCurves

    ' --- D02: Arte unica ---
    Dim inst2 As Shape
    Set inst2 = ly.CreateArtisticText(160, 185, _
        "D02 - Selecione apenas este retangulo para TrimBox simples:", _
        cdrLanguageNone, , "Arial", 7)
    inst2.Fill.ApplyUniformFill corCiano
    inst2.Outline.SetNoOutline

    Set s = ly.CreateRectangle2(170, 130, 100, 50)
    s.Fill.ApplyUniformFill corCiano
    s.Outline.SetNoOutline

    ' --- Instrucao T42 Camerom ---
    Dim instCam As Shape
    Set instCam = ly.CreateArtisticText(10, 100, _
        "T42 - CAMEROM: Selecione qualquer arte acima, preencha os campos no Console e clique em Inserir Dados.", _
        cdrLanguageNone, , "Arial", 7)
    instCam.Fill.ApplyUniformFill corMag
    instCam.Outline.SetNoOutline

    Dim instCam2 As Shape
    Set instCam2 = ly.CreateArtisticText(10, 90, _
        "Exemplo Dados: '02-04-2026 Teste Flexo'  |  Exemplo Cores: 'CIANO MAGENTA PRETO'", _
        cdrLanguageNone, , "Arial", 7)
    instCam2.Fill.ApplyUniformFill corPreta
    instCam2.Outline.SetNoOutline

    doc.SaveAs pasta & "Arquivo_D_Montagem.cdr"
    doc.Close
    MsgBox "Arquivo D gerado!", vbInformation, "Console Flexo"
End Sub
