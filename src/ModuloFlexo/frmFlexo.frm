VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFlexo 
   Caption         =   "Console Flexo v2.0"
   ClientHeight    =   10848
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4608
   OleObjectBlob   =   "frmFlexo.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFlexo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

' ============================================================
' ESTADO DO SISTEMA
' ============================================================
Private ultimoLabelAtivo As MSForms.Label
Private ultimaCaptionOriginal As String
Private ultimaAcao As String
Private ultimaAcaoEhSelecao As Boolean

' ============================================================
' FRAMES COLAPS�VEIS
' ============================================================
Private mFrameNomes(3)              As String
Private mFrameIcones(3)             As String
Private mFrameTitulos(3)            As String
Private mFrameAlturasOriginais(3)   As Double
Private mFrameTopInicial            As Double
Private mFrameColapsado(3)          As Boolean
Private mPaddingRodape              As Double   ' altura abaixo do btnDesfazer (capturado no init)
Private Const ALTURA_CAPTION        As Double = 16   ' twips -- so o titulo
Private Const ESPACO_FRAMES         As Double = 3    ' gap entre frames

' ============================================================
' CORES
' ============================================================
Private Const C_FUNDO_FORM      As Long = 3416624   ' #341F30 ? #1A2030
Private Const C_FUNDO_BTN       As Long = 3814942   ' #1E2A3A
Private Const C_FUNDO_HOVER     As Long = 4539972   ' #243244
Private Const C_FUNDO_PRESS     As Long = 2763291   ' #151C2B
Private Const C_FUNDO_DONE      As Long = 3149596   ' #181F2C
Private Const C_FUNDO_INPUT     As Long = 1839634   ' #111822
Private Const C_FUNDO_DESFAZER  As Long = 6174234   ' #1A3A5E ? ajustado
Private Const C_TEXTO_BTN       As Long = 13287594  ' #9AB0C8 ? ajustado
Private Const C_TEXTO_DONE      As Long = 4016706   ' #3A4E62
Private Const C_TEXTO_LABEL     As Long = 4872266   ' #4A5870
Private Const C_TEXTO_SEC       As Long = 6322070   ' #6A7D96 ? ajustado
Private Const C_AZUL            As Long = 15253610  ' #6AACE8 ? ajustado
Private Const C_BORDA_SEC       As Long = 2367011   ' #232D3F ? ajustado

' ============================================================
' HELPER: converte #RRGGBB para Long do VBA (BGR)
' ============================================================
Private Function H(R As Long, G As Long, B As Long) As Long
    H = RGB(R, G, B)
End Function

' ============================================================
' INICIALIZA��O
' ============================================================
Private Sub UserForm_Initialize()
    On Error Resume Next
    Me.StartUpPosition = 0
    Me.Left = 10
    Me.Top = 60
    Me.Width = 230
    Me.Height = 600
    Me.BackColor = H(26, 32, 48)
    Me.Caption = "Console Flexo v2.0"

    AplicarTemaFrames
    AplicarTemaLabels
    AplicarTemaInputs
    ResetarDesfazer
    AplicarTooltips
    InicializarFramesColapsaveis
End Sub

' ============================================================
' TEMA - FRAMES
' ============================================================
Private Sub AplicarTemaFrames()
    With Me.frameTratamentoDeCores
        .BackColor = H(26, 32, 48)
        .ForeColor = H(106, 125, 150)
        .BorderColor = H(35, 45, 63)
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .Caption = " " & ChrW(9679) & "  TRATAMENTO DE CORES"
    End With
    With Me.frameVetores
        .BackColor = H(26, 32, 48)
        .ForeColor = H(106, 125, 150)
        .BorderColor = H(35, 45, 63)
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .Caption = " " & ChrW(9998) & "  TRATAMENTO DE VETORES"
    End With
    With Me.FrameBitmaps
        .BackColor = H(26, 32, 48)
        .ForeColor = H(106, 125, 150)
        .BorderColor = H(35, 45, 63)
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .Caption = " " & ChrW(9638) & "  TRATAMENTO DE BITMAPS"
    End With
    With Me.FrameMontagem
        .BackColor = H(26, 32, 48)
        .ForeColor = H(106, 125, 150)
        .BorderColor = H(35, 45, 63)
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .Caption = " " & ChrW(9868) & "  MONTAGEM"
    End With
End Sub

' ============================================================
' TEMA - TODAS AS LABELS/BOT�ES
' ============================================================
Private Sub AplicarTemaLabels()
    Dim lbls(17) As MSForms.Label
    Set lbls(0) = Me.btnBranco
    Set lbls(1) = Me.btnPretoSujo
    Set lbls(2) = Me.btnSpot
    Set lbls(3) = Me.btnRGB
    Set lbls(4) = Me.btnCorRegistro
    Set lbls(5) = Me.btnConverterPantone
    Set lbls(6) = Me.btnSelPreenchimento
    Set lbls(7) = Me.btnSelContorno
    Set lbls(8) = Me.btnCorrigirBordaDura
    Set lbls(9) = Me.btnLimparSujeira
    Set lbls(10) = Me.btnTextosEmCurvas
    Set lbls(11) = Me.btnEspessuraMinima
    Set lbls(12) = Me.btnCorrigirContornos
    Set lbls(13) = Me.btnDesbloquear
    Set lbls(14) = Me.btnPadronizarImagens
    Set lbls(15) = Me.btnInserirTextos
    Set lbls(16) = Me.btnTrimBox
    Set lbls(17) = Me.btnMicropontos

    Dim i As Integer
    For i = 0 To 17
        ' Injeta o caption com icone antes de aplicar o estilo
        ' (AplicarEstiloLabelPadrao le .Caption do designer — precisa ser sobrescrito primeiro)
        lbls(i).Caption = ObterCaptionOriginal(i)
        AplicarEstiloLabelPadrao lbls(i)
    Next i

    ' Botao Desfazer -- inicia desabilitado
    With Me.btnDesfazer
        .Enabled = False
        .BackColor = H(30, 42, 58)
        .ForeColor = H(30, 55, 95)
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .BorderStyle = fmBorderStyleNone
        .Caption = vbCrLf & ChrW(8635) & "  Desfazer " & ChrW(250) & "ltima a" & ChrW(231) & ChrW(227) & "o"
    End With

    ' Bot�o Reset
    With Me.btnReset
        .BackColor = H(30, 42, 58)
        .ForeColor = H(106, 125, 150)
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .BorderStyle = fmBorderStyleNone
        .Caption = vbCrLf & ChrW(8635)
        .ControlTipText = "Resetar status dos bot" & ChrW(245) & "es"
        .Left = Me.btnDesfazer.Left + Me.btnDesfazer.Width + 4
        .Top = Me.btnDesfazer.Top
        .Width = 24
        .Height = Me.btnDesfazer.Height
    End With
End Sub

Private Sub AplicarTooltips()
    Me.btnBranco.ControlTipText = "Esse Bot" & ChrW(227) & "o Remove a propriedade Overprint de objetos brancos"
    Me.btnPretoSujo.ControlTipText = "Esse Bot" & ChrW(227) & "o converte Pretos sujos/Ricos para Preto Puro"
    Me.btnSpot.ControlTipText = "Esse bot" & ChrW(227) & "o converte cores pantone/spot para CMYK"
    Me.btnRGB.ControlTipText = "Esse bot" & ChrW(227) & "o converte cores RGB para CMYK"
    Me.btnCorRegistro.ControlTipText = "Mudar Objetos como Camerom e Micropontos para cor de Registro"
    Me.btnConverterPantone.ControlTipText = "Esse bot" & ChrW(227) & "o Converte cores RGB para a escala Pantone mais pr" & ChrW(243) & "xima"
    Me.btnSelPreenchimento.ControlTipText = "Esse bot" & ChrW(227) & "o seleciona objetos com o mesma cor de preenchimento"
    Me.btnSelContorno.ControlTipText = "Esse bot" & ChrW(227) & "o seleciona contornos com a mesma cor"
    Me.btnCorrigirBordaDura.ControlTipText = "Esse bot" & ChrW(227) & "o corrige degrad" & ChrW(234) & "s sem ponto m" & ChrW(237) & "nimo" & "(escolha o ajuste entre 2 e 3%)"
    Me.btnLimparSujeira.ControlTipText = "Esse bot" & ChrW(227) & "o limpa poss" & ChrW(237) & "veis sujeiras de cor (Cores abaixo de 2% ele derruba)"
    Me.btnTextosEmCurvas.ControlTipText = "Localiza e converte textos em curvas"
    Me.btnEspessuraMinima.ControlTipText = "Localiza contornos e objetos menores que 0,1mm"
    Me.btnCorrigirContornos.ControlTipText = "Esse bot" & ChrW(227) & "o corrige contornos abaixo de 0,1mm"
    Me.btnPadronizarImagens.ControlTipText = "Esse bot" & ChrW(227) & "o localiza e converte imagens para CMYK 600dpi"
    Me.btnInserirTextos.ControlTipText = "Esse bot" & ChrW(227) & "o insere os dados do camerom na arte"
    Me.btnDesbloquear.ControlTipText = "Desbloqueia todos os objetos bloqueados da p�gina ativa"
    Me.btnTrimBox.ControlTipText = "Esse bot" & ChrW(227) & "o aplica o offset de cada lado da arte e cria o trimbox(Escolha entre 5mm e 7mm)"
    Me.btnMicropontos.ControlTipText = "Insere 4 micropontos ao redor do objeto selecionado (offset 1,5 mm)"
End Sub

Private Sub AplicarEstiloLabelPadrao(lbl As MSForms.Label)
    If lbl Is Nothing Then Exit Sub
    With lbl
        .BackColor = H(30, 42, 58)
        .ForeColor = H(154, 176, 200)
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = False
        .TextAlign = fmTextAlignCenter
        .WordWrap = True
        .BorderStyle = fmBorderStyleNone
        .MousePointer = fmMousePointerDefault
        ' Centraliza verticalmente com padding top
        .Caption = vbCrLf & .Caption
    End With
End Sub

' ============================================================
' TEMA - INPUTS
' ============================================================
Private Sub AplicarTemaInputs()
    With Me.lblDadosCamerom
        .BackColor = H(26, 32, 48)
        .ForeColor = H(52, 82, 118)
        .Font.Name = "Segoe UI"
        .Font.Size = 8
    End With
    With Me.lbsCores
        .BackColor = H(26, 32, 48)
        .ForeColor = H(52, 82, 118)
        .Font.Name = "Segoe UI"
        .Font.Size = 8
    End With
    With Me.txtDados
        .BackColor = H(17, 24, 34)
        .ForeColor = H(154, 176, 200)
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With
    With Me.txtCores
        .BackColor = H(17, 24, 34)
        .ForeColor = H(154, 176, 200)
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

' ============================================================
' SISTEMA DE ESTADO � MARCAR CONCLU�DO
' ============================================================
Private Sub MarcarConcluido(lbl As MSForms.Label, captionOrig As String, nomeAcao As String, apenasSelecao As Boolean)
    With lbl
        .BackColor = H(24, 31, 44)
        .ForeColor = H(58, 78, 98)
        .Caption = vbCrLf & captionOrig & "  " & ChrW(10003)
    End With

    Set ultimoLabelAtivo = lbl
    ultimaCaptionOriginal = captionOrig
    ultimaAcao = nomeAcao
    ultimaAcaoEhSelecao = apenasSelecao  ' ? GRAVA O FLAG

    If Not apenasSelecao Then
        With Me.btnDesfazer
            .Enabled = True
            .BackColor = H(26, 58, 94)
            .ForeColor = H(106, 172, 232)
            .Caption = vbCrLf & ChrW(8635) & "  Desfazer: " & nomeAcao
        End With
    Else
        ' Acao de selecao -- mantem estado atual do Desfazer sem alterar
    End If
End Sub

' ============================================================
' RESETAR DESFAZER � reseta APENAS o �ltimo bot�o clicado
' ============================================================
Private Sub ResetarDesfazer()
    If Not ultimoLabelAtivo Is Nothing Then
        With ultimoLabelAtivo
            .Caption = vbCrLf & ultimaCaptionOriginal
            .BackColor = H(30, 42, 58)
            .ForeColor = H(154, 176, 200)
        End With
        Set ultimoLabelAtivo = Nothing
    End If

    ultimaCaptionOriginal = ""
    ultimaAcao = ""
    ultimaAcaoEhSelecao = False  ' ? LIMPA O FLAG

    With Me.btnDesfazer
        .Enabled = False
        .BackColor = H(30, 42, 58)
        .ForeColor = H(52, 82, 118)
        .Caption = vbCrLf & ChrW(8635) & "  Desfazer " & ChrW(250) & "ltima a" & ChrW(231) & ChrW(227) & "o"
    End With
End Sub

' ============================================================
' HOVER � MOUSE ENTER / LEAVE
' ============================================================
Private Sub AplicarHover(lbl As MSForms.Label)
    If lbl.ForeColor = H(58, 78, 98) Then Exit Sub ' N�o aplica hover em done
    lbl.BackColor = H(36, 50, 68)
    lbl.ForeColor = H(192, 212, 232)
End Sub

Private Sub RemoverHover(lbl As MSForms.Label)
    If lbl.ForeColor = H(58, 78, 98) Then Exit Sub ' N�o remove em done
    lbl.BackColor = H(30, 42, 58)
    lbl.ForeColor = H(154, 176, 200)
End Sub

Private Sub AplicarPress(lbl As MSForms.Label)
    If lbl.ForeColor = H(58, 78, 98) Then Exit Sub
    lbl.BackColor = H(17, 24, 34)  ' fundo bem mais escuro
    lbl.ForeColor = H(230, 240, 252) ' texto quase branco para maximo contraste
End Sub

' ============================================================
' EVENTOS � btnBranco
' ============================================================
Private Sub btnBranco_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarHover Me.btnBranco
End Sub
Private Sub btnBranco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarPress Me.btnBranco
End Sub
Private Sub btnPretoSujo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarHover Me.btnPretoSujo
End Sub
Private Sub btnPretoSujo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarPress Me.btnPretoSujo
End Sub
Private Sub btnSpot_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarHover Me.btnSpot
End Sub
Private Sub btnSpot_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarPress Me.btnSpot
End Sub
Private Sub btnRGB_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarHover Me.btnRGB
End Sub
Private Sub btnRGB_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarPress Me.btnRGB
End Sub
Private Sub btnCorRegistro_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarHover Me.btnCorRegistro
End Sub
Private Sub btnCorRegistro_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarPress Me.btnCorRegistro
End Sub
Private Sub btnConverterPantone_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarHover Me.btnConverterPantone
End Sub
Private Sub btnConverterPantone_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarPress Me.btnConverterPantone
End Sub
Private Sub btnSelPreenchimento_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarHover Me.btnSelPreenchimento
End Sub
Private Sub btnSelPreenchimento_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarPress Me.btnSelPreenchimento
End Sub
Private Sub btnSelContorno_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarHover Me.btnSelContorno
End Sub
Private Sub btnSelContorno_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarPress Me.btnSelContorno
End Sub
Private Sub btnCorrigirBordaDura_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarHover Me.btnCorrigirBordaDura
End Sub
Private Sub btnCorrigirBordaDura_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarPress Me.btnCorrigirBordaDura
End Sub
Private Sub btnLimparSujeira_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarHover Me.btnLimparSujeira
End Sub
Private Sub btnLimparSujeira_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarPress Me.btnLimparSujeira
End Sub
Private Sub btnTextosEmCurvas_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarHover Me.btnTextosEmCurvas
End Sub
Private Sub btnTextosEmCurvas_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarPress Me.btnTextosEmCurvas
End Sub
Private Sub btnEspessuraMinima_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarHover Me.btnEspessuraMinima
End Sub
Private Sub btnEspessuraMinima_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarPress Me.btnEspessuraMinima
End Sub
Private Sub btnCorrigirContornos_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarHover Me.btnCorrigirContornos
End Sub
Private Sub btnCorrigirContornos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarPress Me.btnCorrigirContornos
End Sub
Private Sub btnPadronizarImagens_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarHover Me.btnPadronizarImagens
End Sub
Private Sub btnPadronizarImagens_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarPress Me.btnPadronizarImagens
End Sub
Private Sub btnInserirTextos_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarHover Me.btnInserirTextos
End Sub
Private Sub btnInserirTextos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarPress Me.btnInserirTextos
End Sub
Private Sub btnTrimBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarHover Me.btnTrimBox
End Sub
Private Sub btnTrimBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarPress Me.btnTrimBox
End Sub
' ============================================================
' BOT�ES � TRATAMENTO DE CORES
' ============================================================
Private Sub btnBranco_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Mod02_Cores.CorrigirBrancoOverprint
    MarcarConcluido Me.btnBranco, "Branco Overprint", "Branco Overprint", False
End Sub

Private Sub btnPretoSujo_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Mod02_Cores.DetectarPretoSujo
    MarcarConcluido Me.btnPretoSujo, "Preto Composto", "Preto Composto", False
End Sub

Private Sub btnSpot_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Mod02_Cores.ConverterSpotParaCMYK
    MarcarConcluido Me.btnSpot, "Converter Spot p/ CMYK", "Converter Spot", False
End Sub

Private Sub btnRGB_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Mod02_Cores.ConverterRGB
    MarcarConcluido Me.btnRGB, "Converter RGB p/ CMYK", "Converter RGB", False
End Sub

Private Sub btnCorRegistro_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Mod02_Cores.MudarParaCorDeRegistro
    MarcarConcluido Me.btnCorRegistro, "Mudar p/ Cor de Registro", "Cor de Registro", False
End Sub

Private Sub btnConverterPantone_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Mod02_Cores.ConverterParaPantone
    MarcarConcluido Me.btnConverterPantone, "Converter para Pantone", "Converter Pantone", False
End Sub

' ? apenasSelecao = True � n�o habilita Desfazer
Private Sub btnSelPreenchimento_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Mod02_Cores.SelecionarMsmCor(1)
    MarcarConcluido Me.btnSelPreenchimento, "Seleciona Msm Cor Preenchimento", "Sel. Preenchimento", True
End Sub

' ? apenasSelecao = True � n�o habilita Desfazer
Private Sub btnSelContorno_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Mod02_Cores.SelecionarMsmCor(2)
    MarcarConcluido Me.btnSelContorno, "Seleciona Msm Cor Contorno", "Sel. Contorno", True
End Sub

Private Sub btnCorrigirBordaDura_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Mod02_Cores.CorrigirBordaDuraGradientes
    MarcarConcluido Me.btnCorrigirBordaDura, "Corrigir Minimas Degrade", "Corrigir Degrade", False
End Sub

Private Sub btnLimparSujeira_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Mod02_Cores.LimparSujeiraCores
    MarcarConcluido Me.btnLimparSujeira, "Limpar Cores", "Limpar Cores", False
End Sub

' ============================================================
' BOT�ES � TRATAMENTO DE VETORES
' ============================================================
Private Sub btnTextosEmCurvas_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Mod03_Vetores.ConverterTextosEmCurvas
    MarcarConcluido Me.btnTextosEmCurvas, "Textos em Curvas", "Textos em Curvas", False
End Sub

' ? apenasSelecao = True � n�o habilita Desfazer
Private Sub btnEspessuraMinima_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Mod03_Vetores.InspecionarEspessuraMinima
    MarcarConcluido Me.btnEspessuraMinima, "Inspetor de Linhas Finas", "Linhas Finas", True
End Sub

Private Sub btnCorrigirContornos_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Mod03_Vetores.PadronizarContornosFinos
    MarcarConcluido Me.btnCorrigirContornos, "Corrigir Contornos Finos", "Contornos Finos", False
End Sub

' ============================================================
' EVENTOS -- btnDesbloquear
' ============================================================
Private Sub btnDesbloquear_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarHover Me.btnDesbloquear
End Sub
Private Sub btnDesbloquear_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarPress Me.btnDesbloquear
End Sub
Private Sub btnDesbloquear_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Mod03_Vetores.DesbloquearObjetos
    MarcarConcluido Me.btnDesbloquear, "Desbloquear Objetos", "Desbloquear", False
End Sub

' ============================================================
' BOT�ES � TRATAMENTO DE BITMAPS
' ============================================================
Private Sub btnPadronizarImagens_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Mod05_Imagens.PadronizarImagensCMYK600
    MarcarConcluido Me.btnPadronizarImagens, "Padronizar Imagens", "Padronizar Imagens", False
End Sub

' ============================================================
' BOT�ES � MONTAGEM
' ============================================================
Private Sub btnInserirTextos_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Mod04_Montagem.InserirTextosCamerom(Me.txtDados.Text, Me.txtCores.Text)
    Me.txtDados.Text = ""
    Me.txtCores.Text = ""
    Me.txtDados.SetFocus
    MarcarConcluido Me.btnInserirTextos, "Inserir Dados", "Inserir Dados", False
End Sub

Private Sub btnTrimBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Mod04_Montagem.AjustarTrimBoxEBorda
    MarcarConcluido Me.btnTrimBox, "Aplicar Trimbox", "Aplicar Trimbox", False
End Sub

Private Sub btnMicropontos_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarHover Me.btnMicropontos
End Sub
Private Sub btnMicropontos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarPress Me.btnMicropontos
End Sub
Private Sub btnMicropontos_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Mod07_InserirMicropontos.InserirMicropontos
    MarcarConcluido Me.btnMicropontos, ObterCaptionOriginal(17), "Micropontos", False
End Sub

' ============================================================
' LEAVE - restaura hover quando mouse sai do btnMicropontos
' ============================================================

' ============================================================
' EVENTOS � btnDesfazer
' ============================================================
Private Sub btnDesfazer_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not Me.btnDesfazer.Enabled Then Exit Sub
    Me.btnDesfazer.BackColor = H(30, 70, 114)
End Sub
Private Sub btnDesfazer_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not Me.btnDesfazer.Enabled Then Exit Sub
    Me.btnDesfazer.BackColor = H(20, 48, 78)
End Sub
Private Sub btnDesfazer_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not Me.btnDesfazer.Enabled Then Exit Sub
    If ultimaAcaoEhSelecao Then
        ResetarDesfazer
        Exit Sub
    End If

    On Error Resume Next
    ActiveDocument.Undo  ' ? 1 �nico Undo � BeginCommandGroup garante lote
    On Error GoTo 0
    ResetarDesfazer
End Sub

' ============================================================
' LEAVE � restaura hover quando mouse sai dos bot�es
' ============================================================
' ============================================================
' FRAMES COLAPS�VEIS -- Inicializar
' ============================================================
Private Sub InicializarFramesColapsaveis()
    mFrameNomes(0) = "frameTratamentoDeCores"
    mFrameNomes(1) = "frameVetores"
    mFrameNomes(2) = "FrameBitmaps"
    mFrameNomes(3) = "FrameMontagem"

    mFrameIcones(0) = ChrW(9679)    ' ?
    mFrameIcones(1) = ChrW(9998)    ' ?
    mFrameIcones(2) = ChrW(9638)    ' ?
    mFrameIcones(3) = ChrW(9868)    ' montagem

    mFrameTitulos(0) = "TRATAMENTO DE CORES"
    mFrameTitulos(1) = "TRATAMENTO DE VETORES"
    mFrameTitulos(2) = "TRATAMENTO DE BITMAPS"
    mFrameTitulos(3) = "MONTAGEM"

    mFrameTopInicial = Me.Controls(mFrameNomes(0)).Top

    ' Captura o padding real do rodape (distancia do btn ate o fim do form)
    ' Inclui automaticamente barra de titulo + bordas, independente de DPI
    mPaddingRodape = Me.Height - Me.btnDesfazer.Top - Me.btnDesfazer.Height

    Dim i As Integer
    For i = 0 To 3
        mFrameAlturasOriginais(i) = Me.Controls(mFrameNomes(i)).Height
        mFrameColapsado(i) = False
    Next i

    For i = 0 To 3
        AtualizarCaptionFrame i
    Next i
End Sub

' ============================================================
' FRAMES COLAPS�VEIS -- Caption com seta
' ============================================================
Private Sub AtualizarCaptionFrame(idx As Integer)
    Dim seta As String
    If mFrameColapsado(idx) Then
        seta = ChrW(9658)   ' ?
    Else
        seta = ChrW(9660)   ' ?
    End If
    Me.Controls(mFrameNomes(idx)).Caption = _
        " " & mFrameIcones(idx) & "  " & mFrameTitulos(idx) & "  " & seta
End Sub

' ============================================================
' FRAMES COLAPS�VEIS -- Toggle
' ============================================================
Private Sub ToggleFrame(idx As Integer)
    Dim frm As MSForms.Frame
    Set frm = Me.Controls(mFrameNomes(idx))

    mFrameColapsado(idx) = Not mFrameColapsado(idx)

    If mFrameColapsado(idx) Then
        frm.Height = ALTURA_CAPTION
    Else
        frm.Height = mFrameAlturasOriginais(idx)
    End If

    AtualizarCaptionFrame idx
    ReposicionarFrames
End Sub

' ============================================================
' FRAMES COLAPS�VEIS -- Reposicionar e ajustar form
' ============================================================
Private Sub ReposicionarFrames()
    Dim topAtual As Double
    topAtual = mFrameTopInicial

    Dim i As Integer
    For i = 0 To 3
        Dim frm As MSForms.Frame
        Set frm = Me.Controls(mFrameNomes(i))
        frm.Top = topAtual
        topAtual = topAtual + frm.Height + ESPACO_FRAMES
    Next i

    Me.btnDesfazer.Top = topAtual + 4
    Me.btnReset.Top = Me.btnDesfazer.Top
    Me.Height = Me.btnDesfazer.Top + Me.btnDesfazer.Height + mPaddingRodape
End Sub

' ============================================================
' FRAMES COLAPS�VEIS -- Eventos Click
' ============================================================
Private Sub frameTratamentoDeCores_Click()
    ToggleFrame 0
End Sub
Private Sub frameVetores_Click()
    ToggleFrame 1
End Sub
Private Sub FrameBitmaps_Click()
    ToggleFrame 2
End Sub
Private Sub FrameMontagem_Click()
    ToggleFrame 3
End Sub

Private Sub frameTratamentoDeCores_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RemoverHoverTodos
End Sub
Private Sub frameVetores_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RemoverHoverTodos
End Sub
Private Sub FrameBitmaps_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RemoverHoverTodos
End Sub
Private Sub FrameMontagem_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RemoverHoverTodos
End Sub

Private Sub RemoverHoverTodos()
    Dim lbls(17) As MSForms.Label
    Set lbls(0) = Me.btnBranco
    Set lbls(1) = Me.btnPretoSujo
    Set lbls(2) = Me.btnSpot
    Set lbls(3) = Me.btnRGB
    Set lbls(4) = Me.btnCorRegistro
    Set lbls(5) = Me.btnConverterPantone
    Set lbls(6) = Me.btnSelPreenchimento
    Set lbls(7) = Me.btnSelContorno
    Set lbls(8) = Me.btnCorrigirBordaDura
    Set lbls(9) = Me.btnLimparSujeira
    Set lbls(10) = Me.btnTextosEmCurvas
    Set lbls(11) = Me.btnEspessuraMinima
    Set lbls(12) = Me.btnCorrigirContornos
    Set lbls(13) = Me.btnDesbloquear
    Set lbls(14) = Me.btnPadronizarImagens
    Set lbls(15) = Me.btnInserirTextos
    Set lbls(16) = Me.btnTrimBox
    Set lbls(17) = Me.btnMicropontos

    Dim i As Integer
    For i = 0 To 17
        RemoverHover lbls(i)
    Next i

    ' So restaura hover do Desfazer se ele estiver habilitado
    If Me.btnDesfazer.Enabled Then
        Me.btnDesfazer.BackColor = H(26, 58, 94)
        Me.btnDesfazer.ForeColor = H(106, 172, 232)
    End If
    Me.btnReset.BackColor = H(30, 42, 58)
    Me.btnReset.ForeColor = H(106, 125, 150)
End Sub

Private Sub ResetarTodosBotoes()
    Dim lbls(17) As MSForms.Label
    Set lbls(0) = Me.btnBranco
    Set lbls(1) = Me.btnPretoSujo
    Set lbls(2) = Me.btnSpot
    Set lbls(3) = Me.btnRGB
    Set lbls(4) = Me.btnCorRegistro
    Set lbls(5) = Me.btnConverterPantone
    Set lbls(6) = Me.btnSelPreenchimento
    Set lbls(7) = Me.btnSelContorno
    Set lbls(8) = Me.btnCorrigirBordaDura
    Set lbls(9) = Me.btnLimparSujeira
    Set lbls(10) = Me.btnTextosEmCurvas
    Set lbls(11) = Me.btnEspessuraMinima
    Set lbls(12) = Me.btnCorrigirContornos
    Set lbls(13) = Me.btnDesbloquear
    Set lbls(14) = Me.btnPadronizarImagens
    Set lbls(15) = Me.btnInserirTextos
    Set lbls(16) = Me.btnTrimBox
    Set lbls(17) = Me.btnMicropontos

    Dim i As Integer
    For i = 0 To 17
        With lbls(i)
            .BackColor = H(30, 42, 58)
            .ForeColor = H(154, 176, 200)
            .Caption = vbCrLf & ObterCaptionOriginal(i)
        End With
    Next i

    Set ultimoLabelAtivo = Nothing
    ultimaCaptionOriginal = ""
    ultimaAcao = ""
    ultimaAcaoEhSelecao = False

    ' Desabilita e esmaece o Desfazer igual ao estado inicial
    ResetarDesfazer
End Sub

Private Function ObterCaptionOriginal(index As Integer) As String
    Select Case index
        ' -- TRATAMENTO DE CORES --
        Case 0:  ObterCaptionOriginal = ChrW(9678) & "  Branco Overprint"
        Case 1:  ObterCaptionOriginal = ChrW(9724) & "  Preto Composto"
        Case 2:  ObterCaptionOriginal = ChrW(9672) & "  Converter Spot p/ CMYK"
        Case 3:  ObterCaptionOriginal = ChrW(11041) & "  Converter RGB p/ CMYK"
        Case 4:  ObterCaptionOriginal = ChrW(10011) & "  Cor de Registro"
        Case 5:  ObterCaptionOriginal = ChrW(9673) & "  Pantone p/ CMYK"
        Case 6:  ObterCaptionOriginal = ChrW(9635) & "  Mesmo Preenchimento"
        Case 7:  ObterCaptionOriginal = ChrW(9634) & "  Mesmo Contorno"
        Case 8:  ObterCaptionOriginal = ChrW(9636) & "  Minimas Degrade"
        Case 9:  ObterCaptionOriginal = ChrW(10006) & "  Limpar Cores"
        ' -- TRATAMENTO DE VETORES --
        Case 10: ObterCaptionOriginal = ChrW(10023) & "  Textos em Curvas"
        Case 11: ObterCaptionOriginal = ChrW(9473) & "  Inspetor de Linhas"
        Case 12: ObterCaptionOriginal = ChrW(8863) & "  Corrigir Contornos"
        Case 13: ObterCaptionOriginal = ChrW(9683) & "  Desbloquear Objetos"
        ' -- TRATAMENTO DE BITMAPS --
        Case 14: ObterCaptionOriginal = ChrW(9640) & "  Padronizar Imagens"
        ' -- MONTAGEM --
        Case 15: ObterCaptionOriginal = ChrW(10000) & "  Inserir Dados"
        Case 16: ObterCaptionOriginal = ChrW(8862) & "  Aplicar Trimbox"
        ' -- VETORES (cont.) --
        Case 17: ObterCaptionOriginal = ChrW(8853) & "  Inserir Micropontos"    ' ⊕
        Case Else: ObterCaptionOriginal = ""
    End Select
End Function

Private Sub btnReset_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btnReset.BackColor = H(36, 50, 68)
    Me.btnReset.ForeColor = H(192, 212, 232)
End Sub

Private Sub btnReset_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btnReset.BackColor = H(21, 28, 43)
End Sub

Private Sub btnReset_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btnReset.BackColor = H(30, 42, 58)
    Me.btnReset.ForeColor = H(106, 125, 150)
    ResetarTodosBotoes
End Sub
