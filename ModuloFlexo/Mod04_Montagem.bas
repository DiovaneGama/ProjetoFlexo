Attribute VB_Name = "Mod04_Montagem"
' ============================================================
' MÓDULO: Mod04_Montagem (PREPARAÇÃO E TRIMBOX)
' DESCRIÇÃO: Inserção de Dados e Cores com Pintura Automática
' ============================================================

Option Explicit

' ------------------------------------------------------------
' 1. INSERIR TEXTOS CAMEROM (Rotacionados, Recuados e Coloridos)
' ------------------------------------------------------------
Public Sub InserirTextosCamerom(textoDados As String, textoCores As String)
    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Selecione a arte primeiro para ter uma referência de posicionamento.", vbExclamation, "Console Flexo"
        Exit Sub
    End If
    If Trim(textoDados) = "" Or Trim(textoCores) = "" Then
        MsgBox "Preencha as duas caixas de texto (Dados e Cores) antes de executar.", vbExclamation, "Console Flexo"
        Exit Sub
    End If

    ' ? CORREÇÃO: agrupa todas as alterações em 1 único Undo
    ActiveDocument.BeginCommandGroup "Inserir Textos Camerom"
    On Error GoTo FimErro

    Dim unOriginal As cdrUnit: unOriginal = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter
    
    Dim srOriginal As ShapeRange
    Set srOriginal = ActiveSelectionRange
    
    Dim X As Double, Y As Double, w As Double, H As Double
    srOriginal.GetBoundingBox X, Y, w, H
    
    Application.Optimization = True
    Dim corReg As New Color: corReg.RegistrationAssign
    
    Dim txtDados As Shape
    Set txtDados = ActiveLayer.CreateArtisticText(0, 0, textoDados, cdrLanguageNone, , "Arial", 5, cdrTrue)
    txtDados.Fill.ApplyUniformFill corReg
    txtDados.Outline.SetNoOutline
    txtDados.Rotate 90
    txtDados.LeftX = X - txtDados.SizeWidth - 0.5
    txtDados.TopY = (Y + H) - 5
    
    Dim txtCores As Shape
    Set txtCores = ActiveLayer.CreateArtisticText(0, 0, textoCores, cdrLanguageNone, , "Arial", 5)
    txtCores.Outline.SetNoOutline
    txtCores.Rotate 90
    txtCores.LeftX = X - txtCores.SizeWidth - 0.5
    txtCores.BottomY = Y + 5
    
    ColorirPalavrasDaLegenda txtCores
    
    srOriginal.CreateSelection
    
    Application.Optimization = False
    Application.Refresh
    ActiveDocument.Unit = unOriginal

    ActiveDocument.EndCommandGroup

    MsgBox "Informações do Camerom ancoradas e coloridas com sucesso!", vbInformation, "Console Flexo"
    Exit Sub

FimErro:
    ActiveDocument.EndCommandGroup
    Application.Optimization = False
    ActiveDocument.Unit = unOriginal
    Application.Refresh
    MsgBox "Erro ao inserir textos: " & Err.Description, vbCritical, "Console Flexo"
End Sub

' ------------------------------------------------------------
' FUNÇÃO INTELIGENTE: PINTA AS PALAVRAS DO TEXTO COM SUAS CORES
' ------------------------------------------------------------
Private Sub ColorirPalavrasDaLegenda(txtShape As Shape)
    Dim w As TextRange
    Dim txt As String
    Dim C As New Color
    Dim pularProxima As Boolean: pularProxima = False
    Dim corAplicada As Boolean
    
    ' 0. Cria as cores puras na memória
    Dim colCiano As New Color: colCiano.CMYKAssign 100, 0, 0, 0
    Dim colMagenta As New Color: colMagenta.CMYKAssign 0, 100, 0, 0
    Dim colAmarelo As New Color: colAmarelo.CMYKAssign 0, 0, 100, 0
    Dim colPreto As New Color: colPreto.CMYKAssign 0, 0, 0, 100
    Dim colReg As New Color: colReg.RegistrationAssign
    
    Dim i As Integer
    For i = 1 To txtShape.Text.Story.Words.Count
        If pularProxima Then
            pularProxima = False
            GoTo ProximaPalavra
        End If
        
        Set w = txtShape.Text.Story.Words(i)
        txt = UCase(Trim(w.Text))
        corAplicada = False
        
        ' 1. TRATAMENTO CMYK
        If InStr(txt, "CIANO") > 0 Or InStr(txt, "CYAN") > 0 Then
            w.Fill.ApplyUniformFill colCiano
            corAplicada = True
        ElseIf InStr(txt, "MAGENTA") > 0 Then
            w.Fill.ApplyUniformFill colMagenta
            corAplicada = True
        ElseIf InStr(txt, "AMARELO") > 0 Or InStr(txt, "YELLOW") > 0 Then
            w.Fill.ApplyUniformFill colAmarelo
            corAplicada = True
        ElseIf InStr(txt, "PRETO") > 0 Or InStr(txt, "BLACK") > 0 Then
            w.Fill.ApplyUniformFill colPreto
            corAplicada = True
            
        ' 2. TRATAMENTO PANTONE (NOVA ESTRATÉGIA BLINDADA)
        ElseIf InStr(txt, "P") > 0 Then
            Dim numPantone As String: numPantone = ""
            Dim j As Integer
            
            ' Extrai os números da palavra
            For j = 1 To Len(txt)
                If IsNumeric(Mid(txt, j, 1)) Then numPantone = numPantone & Mid(txt, j, 1)
            Next j
            
            ' Se o número estiver na próxima palavra (Ex: "P 485")
            If numPantone = "" And i < txtShape.Text.Story.Words.Count Then
                Dim proxTxt As String
                proxTxt = UCase(Trim(txtShape.Text.Story.Words(i + 1).Text))
                For j = 1 To Len(proxTxt)
                    If IsNumeric(Mid(proxTxt, j, 1)) Then numPantone = numPantone & Mid(proxTxt, j, 1)
                Next j
                If numPantone <> "" Then pularProxima = True
            End If
            
            If numPantone <> "" Then
                Dim idCor As Long: idCor = 0
                Dim pal As Palette
                
                ' ESTRATÉGIA A: Busca em TODAS as paletas abertas (Incluindo Paleta do Documento)
                For Each pal In Palettes
                    idCor = pal.FindColor("PANTONE " & numPantone & " C")
                    If idCor = 0 Then idCor = pal.FindColor("PANTONE " & numPantone & "C")
                    If idCor = 0 Then idCor = pal.FindColor("PANTONE " & numPantone)
                    
                    If idCor > 0 Then
                        C.CopyAssign pal.Color(idCor)
                        Exit For
                    End If
                Next pal
                
                ' ESTRATÉGIA B: Se não achou em nenhuma aberta, tenta forçar a biblioteca oficial
                If idCor = 0 Then
                    Dim palPantone As Palette
                    On Error Resume Next
                    Set palPantone = Palettes.OpenFixed(14) ' Plus Solid Coated
                    If palPantone Is Nothing Then Set palPantone = Palettes.OpenFixed(7) ' Solid Coated Clássica
                    On Error GoTo 0
                    
                    If Not palPantone Is Nothing Then
                        idCor = palPantone.FindColor("PANTONE " & numPantone & " C")
                        If idCor = 0 Then idCor = palPantone.FindColor("PANTONE " & numPantone & "C")
                        If idCor > 0 Then C.CopyAssign palPantone.Color(idCor)
                    End If
                End If
                
                ' SE ACHOU O PANTONE, APLICA!
                If idCor > 0 Then
                    w.Fill.ApplyUniformFill C
                    If pularProxima Then txtShape.Text.Story.Words(i + 1).Fill.ApplyUniformFill C
                    corAplicada = True
                Else
                    ' ESTRATÉGIA C (Trava de Segurança): Se o Corel não tiver essa cor de jeito nenhum,
                    ' pinta de Laranja para tirar do "Registro" e não sujar as outras chapas do cliente.
                    C.CMYKAssign 0, 70, 80, 0
                    w.Fill.ApplyUniformFill C
                    If pularProxima Then txtShape.Text.Story.Words(i + 1).Fill.ApplyUniformFill C
                    corAplicada = True
                End If
            End If
        End If
        
        ' 3. REGRA DE OURO: O que sobrar (Dados do cliente, traços, etc) ganha Cor de Registro!
        If Not corAplicada Then
            w.Fill.ApplyUniformFill colReg
        End If
        
ProximaPalavra:
    Next i
End Sub

Public Sub AjustarTrimBoxEBorda()
    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Selecione a arte que definirá o tamanho final do TrimBox.", vbExclamation, "Console Flexo"
        Exit Sub
    End If

    ' =====================================================
    ' ESCOLHA DO TIPO DE BANDA
    ' =====================================================
    Dim respBanda As VbMsgBoxResult
    respBanda = MsgBox("Qual o tipo de banda do arquivo?" & vbCrLf & vbCrLf & _
                       "[ SIM ]  ?  Banda Larga  (7mm de offset)" & vbCrLf & _
                       "[ NÃO ]  ?  Banda Estreita  (5mm de offset)", _
                       vbYesNoCancel + vbQuestion, "Tipo de Banda")

    If respBanda = vbCancel Then Exit Sub

    Dim margem As Double
    If respBanda = vbYes Then
        margem = 7
    Else
        margem = 5
    End If

    Application.Optimization = True

    ' *** ABRE O GRUPO DE COMANDOS — tudo vira 1 único Undo ***
    ActiveDocument.BeginCommandGroup "Aplicar TrimBox"

    On Error GoTo ErroHandler

    Dim unOriginal As cdrUnit: unOriginal = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter

    ' =====================================================
    ' PRÉ-REQUISITO 1: GARANTIR QUE ESTÁ AGRUPADO
    ' =====================================================
    Dim srSelecao As ShapeRange
    Set srSelecao = ActiveSelectionRange

    Dim grpArte As Shape
    If srSelecao.Count > 1 Then
        Set grpArte = srSelecao.Group
    ElseIf srSelecao.Count = 1 Then
        Set grpArte = srSelecao(1)
    End If

    ' =====================================================
    ' PROCESSAMENTO PRINCIPAL
    ' =====================================================
    Dim sr As ShapeRange
    Set sr = CreateShapeRange
    sr.Add grpArte

    Dim X As Double, Y As Double, w As Double, H As Double
    sr.GetBoundingBox X, Y, w, H

    Dim novaLargura As Double: novaLargura = w + (margem * 2)
    Dim novaAltura As Double: novaAltura = H + (margem * 2)

    ' [T37/T38/T39] Redimensiona a pagina primeiro
    ActivePage.SetSize novaLargura, novaAltura
    Application.Refresh

    ' Centraliza a arte na nova pagina usando AlignRangeToPage
    ' Metodo correto conforme API CorelDRAW 2026
    sr.AlignRangeToPage cdrAlignHCenter + cdrAlignVCenter

    Dim rectBorda As Shape
    Set rectBorda = ActiveLayer.CreateRectangle2(0, 0, novaLargura, novaAltura)

    Dim corReg As New Color: corReg.RegistrationAssign
    rectBorda.Fill.ApplyNoFill
    rectBorda.Outline.SetProperties Width:=0.35, Color:=corReg
    rectBorda.AlignToPageCenter cdrAlignHCenter + cdrAlignVCenter

    ' *** FECHA O GRUPO — 1 único Undo a partir daqui ***
    ActiveDocument.EndCommandGroup

    ActiveDocument.Unit = unOriginal
    Application.Optimization = False
    Application.Refresh

    Dim tipoBanda As String
    If respBanda = vbYes Then
        tipoBanda = "Banda Larga (7mm)"
    Else
        tipoBanda = "Banda Estreita (5mm)"
    End If

    MsgBox "TrimBox ajustado com sucesso!" & vbCrLf & _
           "Tipo: " & tipoBanda & vbCrLf & _
           "Tamanho: " & Format(novaLargura, "0.00") & " x " & Format(novaAltura, "0.00") & " mm", _
           vbInformation, "Console Flexo"
    Exit Sub

ErroHandler:
    ActiveDocument.EndCommandGroup
    ActiveDocument.Unit = unOriginal
    Application.Optimization = False
    Application.Refresh
    MsgBox "Erro ao ajustar TrimBox: " & Err.Description, vbCritical, "Console Flexo"
End Sub
