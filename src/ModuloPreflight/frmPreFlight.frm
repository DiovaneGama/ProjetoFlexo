VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPreFlight 
   Caption         =   "Preflight"
   ClientHeight    =   9564.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4536
   OleObjectBlob   =   "frmPreFlight.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPreFlight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ==============================================================================
' REF. API OFICIAL CORELDRAW OBJECT MODEL - Consulte sempre a documentacao:
' https://community.coreldraw.com/sdk/api/draw/27
' ==============================================================================
Option Explicit

' ------------------------------------------------------------------------------
Private Function H(R As Long, G As Long, B As Long) As Long
    H = RGB(R, G, B)
End Function

' ------------------------------------------------------------------------------
' Estados normais dos botoes (restaurados no MouseUp)
Private Sub RestaurarBtnAtualizar()
    Me.btnAtualizar.BackColor = H(30, 42, 58)
    Me.btnAtualizar.ForeColor = H(154, 176, 200)
End Sub

Private Sub RestaurarBtnDesfazer()
    Me.btnDesfazer.BackColor = H(30, 42, 58)
    Me.btnDesfazer.ForeColor = H(154, 176, 200)
End Sub

Private Sub RestaurarBtnCorrigir()
    If Me.btnCorrigir.Enabled Then
        Me.btnCorrigir.BackColor = H(26, 58, 94)
        Me.btnCorrigir.ForeColor = H(106, 172, 232)
    End If
End Sub

' ------------------------------------------------------------------------------
' Hover / Press - btnAtualizar
Private Sub btnAtualizar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btnAtualizar.BackColor = H(36, 50, 68)
    Me.btnAtualizar.ForeColor = H(192, 212, 232)
    RestaurarBtnDesfazer
    RestaurarBtnCorrigir
End Sub

Private Sub btnAtualizar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btnAtualizar.BackColor = H(21, 28, 43)
    Me.btnAtualizar.ForeColor = H(192, 212, 232)
End Sub

Private Sub btnAtualizar_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RestaurarBtnAtualizar
End Sub

' Hover / Press - btnDesfazer
Private Sub btnDesfazer_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btnDesfazer.BackColor = H(36, 50, 68)
    Me.btnDesfazer.ForeColor = H(192, 212, 232)
    RestaurarBtnAtualizar
    RestaurarBtnCorrigir
End Sub

Private Sub btnDesfazer_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btnDesfazer.BackColor = H(21, 28, 43)
    Me.btnDesfazer.ForeColor = H(192, 212, 232)
End Sub

Private Sub btnDesfazer_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RestaurarBtnDesfazer
End Sub

' Hover / Press - btnCorrigir (azul acao)
Private Sub btnCorrigir_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not Me.btnCorrigir.Enabled Then Exit Sub
    Me.btnCorrigir.BackColor = H(30, 70, 114)
    Me.btnCorrigir.ForeColor = H(192, 212, 232)
    RestaurarBtnAtualizar
    RestaurarBtnDesfazer
End Sub

Private Sub btnCorrigir_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not Me.btnCorrigir.Enabled Then Exit Sub
    Me.btnCorrigir.BackColor = H(20, 48, 78)
    Me.btnCorrigir.ForeColor = H(192, 212, 232)
End Sub

Private Sub btnCorrigir_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RestaurarBtnCorrigir
End Sub

' Restaura todos ao mover o mouse sobre o formulario (fora dos botoes)
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RestaurarBtnAtualizar
    RestaurarBtnDesfazer
    RestaurarBtnCorrigir
End Sub

' ------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    On Error Resume Next
    Me.StartUpPosition = 0
    Me.Left = Application.Window.Left + Application.Window.Width - Me.Width - 2000
    Me.Top = Application.Window.Top + 20

    ' =============================================
    ' TEMA DARK - COREL 2026
    ' =============================================
    Dim corFundo As Long
    Dim corFundoSecundario As Long
    Dim corFundoStatus As Long
    Dim corBorda As Long
    Dim corTexto As Long
    Dim corTextoDim As Long

    corFundo = RGB(30, 30, 30)
    corFundoSecundario = RGB(38, 38, 38)
    corBorda = RGB(58, 58, 58)
    corTexto = RGB(204, 204, 204)
    corTextoDim = RGB(120, 120, 120)

    ' FORMUL�RIO
    Me.BackColor = corFundo
    Me.Width = 240
    Me.Height = 450

    ' LABEL STATUS GERAL
    With Me.lblStatusGeral
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .BackColor = RGB(45, 45, 45)
        .ForeColor = RGB(224, 85, 85)
        .Top = 4
        .Left = 0
        .Width = Me.Width - 4
        .Height = 24
        .TextAlign = 2
    End With

    ' SE��O LABELS - aplica estilo base em todos
    Dim lbls(14) As MSForms.Label
    Set lbls(0) = Me.lblBrancoOver
    Set lbls(1) = Me.lblPretoSujo
    Set lbls(2) = Me.lblRGB
    Set lbls(3) = Me.lblRegistro
    Set lbls(4) = Me.lblBordaDura
    Set lbls(5) = Me.lblBloqueados
    Set lbls(6) = Me.lblLinhasFinas
    Set lbls(7) = Me.lblInvisiveis
    Set lbls(8) = Me.lblImgBaixa
    Set lbls(9) = Me.lblImgRGB
    Set lbls(10) = Me.lblFontesVivas
    Set lbls(11) = Me.lblPantone
    Set lbls(12) = Me.lblListaPantones
    Set lbls(13) = Me.lblTecnicas

    Dim i As Integer
    For i = 0 To 13
        If Not lbls(i) Is Nothing Then
            lbls(i).BackColor = corFundo
            lbls(i).ForeColor = corTextoDim
            lbls(i).Font.Name = "Segoe UI"
            lbls(i).Font.Size = 9
            lbls(i).Font.Bold = False
            lbls(i).Left = 8
            lbls(i).Width = Me.Width - 16
        End If
    Next i

    ' POSICIONAMENTO VERTICAL DOS ITENS
    ' Se��o Erros Cr�ticos
    Dim yPos As Integer: yPos = 32

    ' Separador de se��o
    Me.lblBrancoOver.Top = yPos: yPos = yPos + 18
    Me.lblPretoSujo.Top = yPos: yPos = yPos + 18
    Me.lblRGB.Top = yPos: yPos = yPos + 18
    Me.lblBordaDura.Top = yPos: yPos = yPos + 18
    Me.lblLinhasFinas.Top = yPos: yPos = yPos + 18
    Me.lblImgBaixa.Top = yPos: yPos = yPos + 18
    Me.lblImgRGB.Top = yPos: yPos = yPos + 18
    Me.lblFontesVivas.Top = yPos: yPos = yPos + 22

    ' Se��o Informa��es
    Me.lblRegistro.Top = yPos: yPos = yPos + 18
    Me.lblBloqueados.Top = yPos: yPos = yPos + 18
    Me.lblInvisiveis.Top = yPos: yPos = yPos + 22

    ' Se��o Cores Especiais
    Me.lblPantone.Top = yPos: yPos = yPos + 18
    Me.lblListaPantones.Top = yPos
    Me.lblListaPantones.Height = 54
    Me.lblListaPantones.Left = 16
    Me.lblListaPantones.Width = Me.Width - 24
    Me.lblListaPantones.Font.Size = 8
    Me.lblListaPantones.ForeColor = RGB(91, 155, 213)
    yPos = yPos + 58

    Me.lblTecnicas.Top = yPos
    Me.lblTecnicas.Height = 54
    Me.lblTecnicas.Left = 8
    Me.lblTecnicas.Width = Me.Width - 16

    ' BOT�ES
    Dim btnH As Integer: btnH = 22
    Dim btnY As Integer: btnY = Me.Height - 75
    Dim btnW As Integer: btnW = 70

    With Me.btnAtualizar
        .Top = btnY: .Left = 4: .Width = btnW: .Height = btnH
        .Font.Name = "Segoe UI": .Font.Size = 8: .Font.Bold = True
        .BackColor = H(30, 42, 58)
        .ForeColor = H(154, 176, 200)
        .TextAlign = fmTextAlignCenter
        .WordWrap = True
        .BorderStyle = fmBorderStyleNone
        .Caption = vbCrLf & "Atualizar"
        .ControlTipText = "Executar nova varredura PreFlight na p" & ChrW(225) & "gina ativa"
    End With

    With Me.btnDesfazer
        .Top = btnY: .Left = 78: .Width = btnW: .Height = btnH
        .Font.Name = "Segoe UI": .Font.Size = 8: .Font.Bold = True
        .BackColor = H(30, 42, 58)
        .ForeColor = H(154, 176, 200)
        .TextAlign = fmTextAlignCenter
        .WordWrap = True
        .BorderStyle = fmBorderStyleNone
        .Caption = vbCrLf & "Desfazer"
        .ControlTipText = "Desfazer as corre" & ChrW(231) & ChrW(245) & "es aplicadas"
    End With

    With Me.btnCorrigir
        .Top = btnY: .Left = 152: .Width = btnW: .Height = btnH
        .Font.Name = "Segoe UI": .Font.Size = 8: .Font.Bold = True
        .BackColor = H(26, 58, 94)
        .ForeColor = H(106, 172, 232)
        .TextAlign = fmTextAlignCenter
        .WordWrap = True
        .BorderStyle = fmBorderStyleNone
        .Caption = vbCrLf & "Corrigir Erros"
        .ControlTipText = "Aplicar todas as corre" & ChrW(231) & ChrW(245) & "es cr" & ChrW(237) & "ticas automaticamente"
    End With
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next

    ' 1. Atualiza��o Visual das Labels
    AtualizarStatus Me.lblBrancoOver, "Brancos com Overprint: ", relatorio.QtdBrancoOver, True
    AtualizarStatus Me.lblPretoSujo, "Pretos Compostos/Sujos: ", relatorio.QtdPretoSujo, True
    AtualizarStatus Me.lblRGB, "Cores RGB (Vetor): ", relatorio.QtdRGB, True
    AtualizarStatus Me.lblRegistro, "Cores de Registro: ", relatorio.QtdRegistro, False
    AtualizarStatus Me.lblBordaDura, "Gradientes 0% (Borda Dura): ", relatorio.QtdBordaDura, True
    AtualizarStatus Me.lblBloqueados, "Objetos Bloqueados: ", relatorio.QtdBloqueados, False
    AtualizarStatus Me.lblLinhasFinas, "Linhas <= 0.1mm: ", relatorio.QtdLinhasFinas, True
    AtualizarStatus Me.lblInvisiveis, "Objetos Ocultos: ", relatorio.QtdInvisiveis, False
    AtualizarStatus Me.lblImgBaixa, "Imagens < 300 DPI: ", relatorio.QtdImgBaixa, True
    AtualizarStatus Me.lblImgRGB, "Imagens em RGB: ", relatorio.QtdImgRGB, True
    AtualizarStatus Me.lblFontesVivas, "Fontes Vivas: ", relatorio.QtdFontesVivas, True
    AtualizarStatus Me.lblPantone, "Cores Pantone: ", relatorio.QtdPantone, False
    AtualizarStatus Me.lblTecnicas, "Cores Tecnicas (Faca/Vz): ", relatorio.QtdTecnicas, False

    ' 2. T�tulo Din�mico
    Dim totalCritico As Integer
    totalCritico = relatorio.QtdBrancoOver + relatorio.QtdPretoSujo + relatorio.QtdBordaDura + _
                   relatorio.QtdLinhasFinas + relatorio.QtdImgBaixa + relatorio.QtdRGB + _
                   relatorio.QtdFontesVivas

    If totalCritico = 0 Then
        Me.lblStatusGeral.Caption = "ARQUIVO OK PARA PRODUCAO"
        Me.lblStatusGeral.ForeColor = RGB(76, 175, 130)
        Me.lblStatusGeral.BackColor = RGB(26, 58, 42)
        Me.btnCorrigir.Enabled = False
        Me.btnCorrigir.BackColor = H(24, 31, 44)
        Me.btnCorrigir.ForeColor = H(58, 78, 98)
    Else
        Me.lblStatusGeral.Caption = "REVISAR: " & totalCritico & " ITENS CRITICOS"
        Me.lblStatusGeral.ForeColor = RGB(224, 85, 85)
        Me.lblStatusGeral.BackColor = RGB(58, 26, 26)
        Me.btnCorrigir.Enabled = True
        Me.btnCorrigir.BackColor = H(26, 58, 94)
        Me.btnCorrigir.ForeColor = H(106, 172, 232)
    End If

    ' 3. Lista de Pantones
    If relatorio.QtdPantone > 0 Then
        Dim strLista As String
        strLista = Trim(relatorio.BibliotecasPantone)
        If Right(strLista, 1) = "|" Then strLista = Left(strLista, Len(strLista) - 1)
        If Len(strLista) > 0 Then
            Dim arrNomes() As String
            arrNomes = Split(strLista, "|")
            Dim listaFinal As String: listaFinal = ""
            Dim i As Integer
            For i = 0 To UBound(arrNomes)
                If Trim(arrNomes(i)) <> "" Then
                    listaFinal = listaFinal & Chr(149) & " " & Trim(arrNomes(i)) & vbCrLf
                End If
            Next i
            Me.lblListaPantones.Caption = listaFinal
            Me.lblListaPantones.ForeColor = RGB(91, 155, 213)
        Else
            Me.lblListaPantones.Caption = "  Cores Spot customizadas"
            Me.lblListaPantones.ForeColor = RGB(91, 155, 213)
        End If
    Else
        Me.lblListaPantones.Caption = "  Nenhuma cor Pantone detectada."
        Me.lblListaPantones.ForeColor = RGB(80, 80, 80)
    End If

    ' 4. Lista de Cores T�cnicas
    If relatorio.QtdTecnicas > 0 Then
        Dim strTec As String
        strTec = Trim(relatorio.BibliotecasTecnicas)
        If Right(strTec, 1) = "|" Then strTec = Left(strTec, Len(strTec) - 1)
        If Len(strTec) > 0 Then
            Dim arrTec() As String
            arrTec = Split(strTec, "|")
            Dim listaFinalTec As String: listaFinalTec = ""
            Dim iTec As Integer
            For iTec = 0 To UBound(arrTec)
                If Trim(arrTec(iTec)) <> "" Then
                    listaFinalTec = listaFinalTec & Chr(149) & " " & Trim(arrTec(iTec)) & vbCrLf
                End If
            Next iTec
            Me.lblTecnicas.Caption = "Cores Tecnicas (Faca/Vz): " & relatorio.QtdTecnicas & vbCrLf & listaFinalTec
            Me.lblTecnicas.ForeColor = RGB(91, 155, 213)
        End If
    End If
End Sub

Private Sub AtualizarStatus(lbl As MSForms.Label, txt As String, val As Integer, ehErro As Boolean)
    If lbl Is Nothing Then Exit Sub
    lbl.Caption = txt & val
    If val > 0 Then
        If ehErro Then
            lbl.ForeColor = RGB(224, 85, 85)
        Else
            lbl.ForeColor = RGB(91, 155, 213)
        End If
        lbl.Font.Bold = True
    Else
        lbl.ForeColor = RGB(100, 100, 100)
        lbl.Font.Bold = False
    End If
End Sub

Private Sub btnAtualizar_Click()
    Me.MousePointer = fmMousePointerHourGlass
    ExecutarScanner
    Call UserForm_Activate
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub btnCorrigir_Click()
    Dim resposta As Integer
    Dim minDot As String
    Dim minDotVal As Integer

    resposta = MsgBox("Deseja aplicar todas as correcoes criticas de forma automatica na arte?", vbYesNo + vbQuestion, "Correcao Automatica")
    If resposta <> vbYes Then Exit Sub

    If relatorio.QtdBordaDura > 0 Then
        minDot = InputBox("Foram detectados Gradientes indo a 0% (Borda Dura)." & vbCrLf & vbCrLf & _
                          "Qual a porcentagem minima do cliche que devo aplicar?" & vbCrLf & _
                          "(Ex: digite 2 para 2%, 3 para 3%)", "Ponto Minimo de Flexografia", "2")
        If minDot = "" Then Exit Sub
        minDotVal = CInt(minDot)
    Else
        minDotVal = 0
    End If

    Me.MousePointer = fmMousePointerHourGlass
    ExecutarCorrecoes minDotVal
    ExecutarScanner
    Call UserForm_Activate
    Me.MousePointer = fmMousePointerDefault
    MsgBox "Correcoes aplicadas com sucesso!", vbInformation, "PreFlight"
End Sub

Private Sub btnDesfazer_Click()
    On Error Resume Next
    ActiveDocument.Undo
    Me.MousePointer = fmMousePointerHourGlass
    ExecutarScanner
    Call UserForm_Activate
    Me.MousePointer = fmMousePointerDefault
    MsgBox "Correcoes desfeitas! O arquivo voltou ao estado original.", vbInformation, "Desfazer"
    On Error GoTo 0
End Sub

