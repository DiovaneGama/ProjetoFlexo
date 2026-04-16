Attribute VB_Name = "Mod09_Diagnostico"
' ============================================================
' MODULO: Mod09_Diagnostico (RELATORIO DE DIAGNOSTICO)
' DESCRICAO: Exporta relatorio estrutural do documento ativo
'            para um arquivo .txt. Somente leitura -- nao
'            modifica o documento em nenhum momento.
'            Util para debugar crashes antes de rodar o scanner.
' ============================================================
Option Explicit

' ============================================================
' ENTRADA PUBLICA
' ============================================================
Public Sub ExportarDiagnostico()
    If ActiveDocument Is Nothing Then
        MsgBox "Nenhum documento aberto.", vbExclamation, "Console Flexo"
        Exit Sub
    End If

    ' Determina o caminho do arquivo de saida
    Dim caminho As String
    Dim base As String
    base = ActiveDocument.FullFileName
    If base = "" Then
        ' Documento nao salvo -- pede caminho ao usuario
        base = Application.GetSaveFileName("Salvar relatorio de diagnostico", "Arquivo de Texto (*.txt)|*.txt")
        If base = "" Then Exit Sub
        caminho = base
    Else
        ' Mesmo diretorio do .cdr, com sufixo de data/hora
        Dim dir As String
        Dim sem As String
        dir = Left(base, InStrRev(base, "\"))
        sem = Left(base, InStrRev(base, ".") - 1)
        sem = Mid(sem, InStrRev(sem, "\") + 1) ' apenas nome do arquivo sem extensao
        caminho = dir & sem & "_diagnostico_" & Format(Now, "YYYYMMDD_HHMMSS") & ".txt"
    End If

    Dim f As Integer
    f = FreeFile
    On Error GoTo ErrDiag

    Open caminho For Output As #f

    EscreverCabecalho f
    EscreverResumoDocumento f
    EscreverProfundidadePowerClip f
    EscreverLayers f
    EscreverContornosFinos f
    EscreverTextosVivos f
    EscreverImagensBaixaRes f
    EscreverCamposPreflight f

    Print #f, "================================================================================"
    Print #f, "FIM DO RELATORIO"
    Print #f, "================================================================================"

    Close #f
    MsgBox "Relatorio gerado com sucesso!" & vbCrLf & caminho, vbInformation, "Console Flexo"
    Exit Sub

ErrDiag:
    On Error Resume Next
    Close #f
    MsgBox "Erro ao gerar relatorio: " & Err.Description, vbCritical, "Console Flexo"
End Sub

' ============================================================
' HELPERS PRIVADOS
' ============================================================

Private Sub EscreverCabecalho(f As Integer)
    Dim unidNome As String
    Select Case ActiveDocument.Unit
        Case cdrMillimeter:  unidNome = "Milimetros"
        Case cdrCentimeter:  unidNome = "Centimetros"
        Case cdrInch:        unidNome = "Polegadas"
        Case cdrPoint:       unidNome = "Pontos"
        Case Else:           unidNome = "Outro"
    End Select

    Print #f, "================================================================================"
    Print #f, "CONSOLE FLEXO - RELATORIO DE DIAGNOSTICO"
    Print #f, "================================================================================"
    Print #f, "Arquivo  : " & ActiveDocument.FullFileName
    Print #f, "Gerado   : " & Format(Now, "DD/MM/YYYY HH:MM:SS")
    Print #f, "CorelDRAW: " & Application.ProductVersion
    Print #f, "Unidade  : " & unidNome
    Print #f, ""
End Sub

Private Sub EscreverResumoDocumento(f As Integer)
    Print #f, "--------------------------------------------------------------------------------"
    Print #f, "[1] RESUMO DO DOCUMENTO"
    Print #f, "--------------------------------------------------------------------------------"

    Dim totalPaginas As Long: totalPaginas = ActiveDocument.Pages.Count
    Print #f, "Paginas         : " & totalPaginas

    ' Conta shapes por tipo usando traversal iterativo em todas as paginas
    Dim cCurvas As Long, cBitmaps As Long, cTextos As Long
    Dim cGrupos As Long, cOutros As Long, cTotal As Long

    Dim pg As Page
    For Each pg In ActiveDocument.Pages
        Dim pilha As New Collection
        Dim s As Shape
        For Each s In pg.shapes: pilha.Add s: Next s

        Do While pilha.Count > 0
            Dim atual As Shape
            Set atual = pilha.Item(pilha.Count)
            pilha.Remove pilha.Count
            On Error Resume Next

            cTotal = cTotal + 1
            Select Case atual.Type
                Case cdrCurveShape:   cCurvas = cCurvas + 1
                Case cdrBitmapShape:  cBitmaps = cBitmaps + 1
                Case cdrTextShape:    cTextos = cTextos + 1
                Case cdrGroupShape:   cGrupos = cGrupos + 1
                Case Else:            cOutros = cOutros + 1
            End Select

            Dim subS As Shape
            If atual.Type = cdrGroupShape Then
                For Each subS In atual.shapes: pilha.Add subS: Next subS
            End If
            If Not atual.PowerClip Is Nothing Then
                For Each subS In atual.PowerClip.shapes: pilha.Add subS: Next subS
            End If
            On Error GoTo 0
        Loop
    Next pg

    Print #f, "Total shapes    : " & cTotal
    Print #f, "  Curvas        : " & cCurvas
    Print #f, "  Bitmaps       : " & cBitmaps
    Print #f, "  Textos vivos  : " & cTextos
    Print #f, "  Grupos        : " & cGrupos
    Print #f, "  Outros        : " & cOutros
    Print #f, ""
End Sub

Private Sub EscreverProfundidadePowerClip(f As Integer)
    Print #f, "--------------------------------------------------------------------------------"
    Print #f, "[2] PROFUNDIDADE DE POWERCLIP"
    Print #f, "--------------------------------------------------------------------------------"

    ' Pilha guarda pares: Shape + profundidade atual (como dois Collections sincronizados)
    Dim maxProf As Long: maxProf = 0
    Dim maxPagina As Long: maxPagina = 0
    Dim pgIdx As Long: pgIdx = 0

    Dim pg As Page
    For Each pg In ActiveDocument.Pages
        pgIdx = pgIdx + 1
        Dim pilhaS As New Collection
        Dim pilhaD As New Collection
        Dim s As Shape
        For Each s In pg.shapes
            pilhaS.Add s
            pilhaD.Add 0
        Next s

        Do While pilhaS.Count > 0
            Dim atual As Shape
            Set atual = pilhaS.Item(pilhaS.Count)
            Dim profAtual As Long
            profAtual = pilhaD.Item(pilhaD.Count)
            pilhaS.Remove pilhaS.Count
            pilhaD.Remove pilhaD.Count

            On Error Resume Next
            Dim subS As Shape
            If atual.Type = cdrGroupShape Then
                For Each subS In atual.shapes
                    pilhaS.Add subS
                    pilhaD.Add profAtual
                Next subS
            End If
            If Not atual.PowerClip Is Nothing Then
                Dim novaProfPC As Long: novaProfPC = profAtual + 1
                If novaProfPC > maxProf Then
                    maxProf = novaProfPC
                    maxPagina = pgIdx
                End If
                For Each subS In atual.PowerClip.shapes
                    pilhaS.Add subS
                    pilhaD.Add novaProfPC
                Next subS
            End If
            On Error GoTo 0
        Loop
    Next pg

    Print #f, "Profundidade maxima encontrada: " & maxProf
    If maxProf > 0 Then Print #f, "(encontrada na pagina " & maxPagina & ")"
    Print #f, ""
End Sub

Private Sub EscreverLayers(f As Integer)
    Print #f, "--------------------------------------------------------------------------------"
    Print #f, "[3] LAYERS (pagina ativa)"
    Print #f, "--------------------------------------------------------------------------------"
    Print #f, PadDir("Layer", 24) & PadDir("Imprimir", 10) & PadDir("Visivel", 9) & PadDir("Editavel", 10) & PadDir("Especial", 10) & "Shapes"
    Print #f, String(24, "-") & String(10, "-") & String(9, "-") & String(10, "-") & String(10, "-") & "------"

    Dim ly As Layer
    For Each ly In ActivePage.Layers
        On Error Resume Next
        ' Conta shapes na layer via traversal iterativo
        Dim cShapes As Long: cShapes = 0
        Dim pilha As New Collection
        Dim s As Shape
        For Each s In ly.shapes: pilha.Add s: Next s

        Do While pilha.Count > 0
            Dim atual As Shape
            Set atual = pilha.Item(pilha.Count)
            pilha.Remove pilha.Count
            cShapes = cShapes + 1
            Dim subS As Shape
            If atual.Type = cdrGroupShape Then
                For Each subS In atual.shapes: pilha.Add subS: Next subS
            End If
            If Not atual.PowerClip Is Nothing Then
                For Each subS In atual.PowerClip.shapes: pilha.Add subS: Next subS
            End If
        Loop

        Print #f, PadDir(ly.Name, 24) & _
                  PadDir(IIf(ly.Printable, "Sim", "Nao"), 10) & _
                  PadDir(IIf(ly.Visible, "Sim", "Nao"), 9) & _
                  PadDir(IIf(ly.Editable, "Sim", "Nao"), 10) & _
                  PadDir(IIf(ly.IsSpecialLayer, "Sim", "Nao"), 10) & _
                  cShapes
        On Error GoTo 0
    Next ly
    Print #f, ""
End Sub

Private Sub EscreverContornosFinos(f As Integer)
    Print #f, "--------------------------------------------------------------------------------"
    Print #f, "[4] CONTORNOS FINOS (<= 0.101mm)"
    Print #f, "--------------------------------------------------------------------------------"

    ' Salva e padroniza unidade para milimetros
    Dim unidOrig As cdrUnit: unidOrig = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter

    Dim cFinos As Long: cFinos = 0
    Dim pilha As New Collection
    Dim s As Shape
    For Each s In ActivePage.shapes: pilha.Add s: Next s

    Do While pilha.Count > 0
        Dim atual As Shape
        Set atual = pilha.Item(pilha.Count)
        pilha.Remove pilha.Count

        On Error Resume Next
        If atual.Type <> cdrBitmapShape And atual.Type <> cdrGroupShape Then
            If atual.Outline.Type = cdrOutline Then
                Dim espW As Double: espW = atual.Outline.Width
                If espW > 0 And espW <= 0.101 Then
                    Dim ehInt As Boolean: ehInt = False
                    If atual.Outline.Color.Type = cdrColorCMYK Then
                        If (atual.Outline.Color.CMYKCyan + atual.Outline.Color.CMYKMagenta + _
                            atual.Outline.Color.CMYKYellow + atual.Outline.Color.CMYKBlack) = 0 Then
                            If atual.Fill.Type = cdrUniformFill Then
                                If atual.Fill.UniformColor.Type = cdrColorCMYK Then
                                    If (atual.Fill.UniformColor.CMYKCyan + atual.Fill.UniformColor.CMYKMagenta + _
                                        atual.Fill.UniformColor.CMYKYellow + atual.Fill.UniformColor.CMYKBlack) = 0 Then
                                        ehInt = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If Not ehInt Then
                        If atual.Fill.Type = cdrUniformFill Then
                            If Mod08_Utils.CompararCoresSeguro(atual.Outline.Color, atual.Fill.UniformColor) Then
                                ehInt = True
                            End If
                        End If
                    End If
                    If Not ehInt Then cFinos = cFinos + 1
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

    ActiveDocument.Unit = unidOrig

    Print #f, "Shapes com contorno fino detectado: " & cFinos
    Print #f, "(contornos intencionais excluidos -- mesma logica do scanner)"
    Print #f, ""
End Sub

Private Sub EscreverTextosVivos(f As Integer)
    Print #f, "--------------------------------------------------------------------------------"
    Print #f, "[5] TEXTOS VIVOS"
    Print #f, "--------------------------------------------------------------------------------"

    Dim cTextos As Long: cTextos = 0
    Dim pilha As New Collection
    Dim s As Shape
    For Each s In ActivePage.shapes: pilha.Add s: Next s

    Do While pilha.Count > 0
        Dim atual As Shape
        Set atual = pilha.Item(pilha.Count)
        pilha.Remove pilha.Count

        On Error Resume Next
        If atual.Type = cdrTextShape Then cTextos = cTextos + 1
        On Error GoTo 0

        Dim subS As Shape
        If atual.Type = cdrGroupShape Then
            For Each subS In atual.shapes: pilha.Add subS: Next subS
        End If
        If Not atual.PowerClip Is Nothing Then
            For Each subS In atual.PowerClip.shapes: pilha.Add subS: Next subS
        End If
    Loop

    Print #f, "Shapes de texto vivo encontrados: " & cTextos
    Print #f, ""
End Sub

Private Sub EscreverImagensBaixaRes(f As Integer)
    Print #f, "--------------------------------------------------------------------------------"
    Print #f, "[6] IMAGENS ABAIXO DE 300 DPI"
    Print #f, "--------------------------------------------------------------------------------"

    Dim cBaixo As Long: cBaixo = 0
    Dim pilha As New Collection
    Dim s As Shape
    For Each s In ActivePage.shapes: pilha.Add s: Next s

    Do While pilha.Count > 0
        Dim atual As Shape
        Set atual = pilha.Item(pilha.Count)
        pilha.Remove pilha.Count

        On Error Resume Next
        If atual.Type = cdrBitmapShape Then
            If atual.Bitmap.ResolutionX < 300 Or atual.Bitmap.ResolutionY < 300 Then
                cBaixo = cBaixo + 1
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

    Print #f, "Bitmaps com resolucao baixa: " & cBaixo
    Print #f, ""
End Sub

Private Sub EscreverCamposPreflight(f As Integer)
    Print #f, "--------------------------------------------------------------------------------"
    Print #f, "[7] CAMPOS RELATORIO PREFLIGHT (referencia -- preenchidos pelo scanner)"
    Print #f, "--------------------------------------------------------------------------------"
    Print #f, "QtdBrancoOver  | QtdPretoSujo    | QtdRGB          | QtdPantone"
    Print #f, "BibliotecasPantone (lista pipe-delimitada de nomes Pantone unicos)"
    Print #f, "QtdBordaDura   | QtdRegistro     | QtdTecnicas"
    Print #f, "BibliotecasTecnicas (lista pipe-delimitada de cores tecnicas unicas)"
    Print #f, "QtdLinhasFinas | QtdBloqueados   | QtdInvisiveis"
    Print #f, "QtdImgBaixa    | QtdImgRGB       | QtdFontesVivas  | QtdGradBloqueado"
    Print #f, ""
End Sub

' ============================================================
' UTILITARIO: padding de string a comprimento fixo
' ============================================================
Private Function PadDir(texto As String, tamanho As Integer) As String
    PadDir = Left(texto & Space(tamanho), tamanho)
End Function
