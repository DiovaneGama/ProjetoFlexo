Attribute VB_Name = "Mod05_Imagens"
' ============================================================
' MODULO: Mod05_Imagens (AUDITORIA E TRATAMENTO DE BITMAPS)
' DESCRICAO: Padronizacao Absoluta (CMYK + 600 DPI em 1 Clique)
'
' ENUMS IMPORTANTES -- nao confundir:
'   Bitmap.Mode retorna  cdrImageMode:  cdrImageCMYK=5, cdrImageGrayscale=2, cdrImageBlackWhite=0
'   Bitmap.ConvertTo usa cdrImageType:  cdrCMYKColorImage=5, cdrGrayscaleImage=2, cdrBlackAndWhiteImage=0
'
' Bitmap.ConvertTo e Bitmap.Resample operam IN-PLACE:
' o shape permanece no lugar -- inclusive dentro de PowerClips.
' ============================================================
Option Explicit

' cdrImageMode (Bitmap.Mode) — usado para leitura
Private Const MODO_CMYK       As Long = 5
Private Const MODO_GRAYSCALE  As Long = 2
Private Const MODO_PB         As Long = 0

' Limiar e alvo de resolucao
Private Const DPI_MINIMO      As Long = 599   ' abaixo disso = problema
Private Const DPI_ALVO        As Long = 600   ' resolucao padrao de saida

Public Sub PadronizarImagensCMYK600()

    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Deseja padronizar automaticamente todas as imagens" & _
                      " da p" & ChrW(225) & "gina para CMYK e 600 DPI?", _
                      vbYesNo + vbQuestion, "Console Flexo")
    If resposta = vbNo Then Exit Sub

    ' ── Varredura ───────────────────────────────────────────
    ' Usa valores numericos para evitar ambiguidade entre enums
    ' cdrImageMode: CMYK=5, Grayscale=2, BlackWhite=0
    Dim sacola As ShapeRange
    Set sacola = CreateShapeRange

    Dim s As Shape
    For Each s In ActivePage.shapes
        CrawlerPadronizar s, sacola
    Next s

    If sacola.Count = 0 Then
        MsgBox "Aprovado! Todas as imagens j" & ChrW(225) & _
               " est" & ChrW(227) & "o no padr" & ChrW(227) & _
               "o (CMYK/Grayscale/1-bit e 600 DPI).", _
               vbInformation, "Console Flexo"
        Exit Sub
    End If

    ' ── Declaracoes ANTES de qualquer On Error ──────────────
    Dim alteracoes As Integer
    Dim i As Integer
    Dim img As Shape
    Dim modoAtual As Long
    Dim resX As Long
    Dim pxW As Long
    Dim pxH As Long
    Dim okConvert As Boolean
    Dim okResample As Boolean

    alteracoes = 0

    ' ── Conversao ───────────────────────────────────────────
    ActiveDocument.BeginCommandGroup "Padronizar Imagens CMYK 600"
    On Error GoTo FimErro

    For i = 1 To sacola.Count
        Set img = sacola(i)
        okConvert = True
        okResample = True

        ' Passo 1: converter modo de cor IN-PLACE
        ' Bitmap.Mode usa cdrImageMode -- comparamos com valores numericos
        ' para evitar confusao com cdrImageType usado no ConvertTo
        '   cdrImageMode: CMYK=5, Grayscale=2, BlackWhite=0
        On Error Resume Next
        modoAtual = img.Bitmap.Mode

        If modoAtual <> MODO_CMYK And modoAtual <> MODO_GRAYSCALE And modoAtual <> MODO_PB Then
            ' ConvertTo usa cdrImageType: cdrCMYKColorImage = 5
            img.Bitmap.ConvertTo cdrCMYKColorImage
            okConvert = (Err.Number = 0)
            Err.Clear
        End If

        ' Passo 2: resamplear para 600 DPI IN-PLACE
        resX = img.Bitmap.ResolutionX
        If resX < DPI_MINIMO Then
            pxW = CLng(img.Bitmap.Width * (DPI_ALVO / resX))
            pxH = CLng(img.Bitmap.Height * (DPI_ALVO / img.Bitmap.ResolutionY))
            Err.Clear
            img.Bitmap.Resample pxW, pxH, False, DPI_ALVO, DPI_ALVO
            okResample = (Err.Number = 0)
            Err.Clear
        End If

        ' Conta se ao menos uma operacao ocorreu sem erro
        If okConvert And okResample Then
            alteracoes = alteracoes + 1
            img.Selected = True
        End If

        On Error GoTo FimErro
    Next i

    ActiveDocument.EndCommandGroup
    Application.Refresh

    MsgBox "Sucesso! " & alteracoes & " imagem(ns) padronizada(s)" & _
           " para CMYK @ 600 DPI.", vbInformation, "Console Flexo"
    Exit Sub

FimErro:
    ActiveDocument.EndCommandGroup
    Application.Refresh
    MsgBox "Erro ao padronizar imagens: " & Err.Description, _
           vbCritical, "Console Flexo"
End Sub

' ============================================================
' CRAWLER: coleta bitmaps com problemas recursivamente
' Usa valores numericos de cdrImageMode para Bitmap.Mode:
'   CMYK=5, Grayscale=2, BlackWhite=0
' ============================================================
Private Sub CrawlerPadronizar(s As Shape, ByRef sacola As ShapeRange)
    Dim subS As Shape
    On Error Resume Next

    If s.Type = cdrBitmapShape Then
        Dim modo As Long: modo = s.Bitmap.Mode
        Dim dX As Long:   dX = s.Bitmap.ResolutionX

        ' Problema: nao e CMYK(5), Grayscale(2) ou BlackWhite(0)
        '           OU resolucao abaixo de 599 DPI
        If (modo <> MODO_CMYK And modo <> MODO_GRAYSCALE And modo <> MODO_PB) Or dX < DPI_MINIMO Then
            sacola.Add s
        End If
    End If

    ' Mergulho em Grupos
    If s.Type = cdrGroupShape Then
        For Each subS In s.shapes
            CrawlerPadronizar subS, sacola
        Next subS
    End If

    ' Mergulho em PowerClips
    If Not s.PowerClip Is Nothing Then
        For Each subS In s.PowerClip.shapes
            CrawlerPadronizar subS, sacola
        Next subS
    End If

    On Error GoTo 0
End Sub
