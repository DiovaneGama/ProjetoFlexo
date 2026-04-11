VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStepRepeat 
   Caption         =   "Step&Repeat"
   ClientHeight    =   10440
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmStepRepeat.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStepRepeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





' ============================================================
' frmStepRepeat � Step & Repeat v1.0
' Design baseado no frmFlexo (Console Flexo v2.0)
' Labels como botoes, hover/press/done, Segoe UI 8pt
' Docker Modeless � Banda Estreita
' ============================================================
Option Explicit

' ============================================================
' ESTADO
' ============================================================
Private ultimoLabelAtivo As MSForms.Label
Private ultimaCaptionOriginal As String
Private mCameronFilePath As String

Private Sub chkCameronCenter_Click()

End Sub

' ============================================================
' CONTROLES ESPERADOS NO FORM (.frm designer):
' ============================================================
' Frames:
'   frameEspessura, frameDimensoes, frameEspacamento,
'   frameReducao, frameOpcoes, frameResultados
'
' Labels (Radio simulado � espessura):
'   lbl114, lbl170
'
' Labels (Radio simulado � Pi):
'   lblPi314, lblPi3175
'
' Labels (Radio simulado � Reducao 1,14):
'   lblRed638, lblRed622
'
' Labels (Radio simulado � Reducao 1,70):
'   lblRed9, lblRed95, lblRed10
'
' TextBoxes:
'   txtZ, txtLarguraFaca, txtAlturaFaca, txtLarguraMaterial,
'   txtPistas, txtRepeticoes, txtGapPistas
'
' CheckBoxes:
'   chkCameron, chkCameronCenter, chkRelatorio
'
' Labels de Resultado (somente leitura):
'   lblDesenvolvimento, lblGapReps, lblGapPistas,
'   lblReducao, lblPasso
'
' Labels de Acao (botoes):
'   btnMontar, btnReset

' ============================================================
' INICIALIZACAO
' ============================================================
Private Sub UserForm_Initialize()
    On Error Resume Next
    Me.StartUpPosition = 0
    Me.Left = 10
    Me.Top = 60
    Me.Width = 240
    Me.Height = 580
    Me.BackColor = H(26, 32, 48)
    Me.Caption = "Step & Repeat v1.0"
    
    AplicarTemaFrames
    AplicarTemaInputs
    AplicarTemaRadios
    AplicarTemaResultados
    AplicarTemaBotoes
    AplicarTooltips
    
    ' Defaults
    lbl114.Tag = "selected"
    lblPi314.Tag = "selected"
    lblRed638.Tag = "selected"
    ' Estado inicial do frame reducao � 1,70mm oculto e desabilitado
    lblRed9.Visible = False:   lblRed9.Enabled = False
    lblRed95.Visible = False:  lblRed95.Enabled = False
    lblRed10.Visible = False:  lblRed10.Enabled = False

    chkRelatorio.Value = True
    mCameronFilePath = ""
    lblCameronArquivo.Visible = False
    AtualizarLabelCameron

    AtualizarRadioVisual
    RecalcularTudo
End Sub

' ============================================================
' TEMA � FRAMES (padrao frmFlexo)
' ============================================================
Private Sub AplicarTemaFrames()
    Dim frms As Variant
    frms = Array("frameEspessura", "frameDimensoes", "frameEspacamento", _
                 "frameReducao", "frameOpcoes", "frameResultados")
    
    Dim icons As Variant
    icons = Array(ChrW(9679), ChrW(9670), ChrW(9632), _
                  ChrW(8600), ChrW(9881), ChrW(9733))
    
    Dim titles As Variant
    titles = Array("ESPESSURA FOTOPOLIMERO", "DIMENSOES", "ESPACAMENTO", _
                   "REDUCAO", "OPCOES", "RESULTADOS")
    
    Dim i As Long
    For i = 0 To UBound(frms)
        Dim frm As MSForms.Frame
        Set frm = Me.Controls(frms(i))
        With frm
            .BackColor = H(26, 32, 48)
            .ForeColor = H(106, 125, 150)
            .BorderColor = H(35, 45, 63)
            .Font.Name = "Segoe UI"
            .Font.Size = 8
            .Font.Bold = True
            .Caption = " " & icons(i) & "  " & titles(i)
        End With
    Next i
End Sub

' ============================================================
' TEMA � INPUTS (padrao frmFlexo)
' ============================================================
Private Sub AplicarTemaInputs()
    Dim txts As Variant
    txts = Array("txtZ", "txtLarguraFaca", "txtAlturaFaca", _
                 "txtLarguraMaterial", "txtPistas", "txtRepeticoes", "txtGapPistas")
    
    Dim i As Long
    For i = 0 To UBound(txts)
        Dim txt As MSForms.TextBox
        Set txt = Me.Controls(txts(i))
        With txt
            .BackColor = H(17, 24, 34)
            .ForeColor = H(154, 176, 200)
            .Font.Name = "Segoe UI"
            .Font.Size = 8
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
        End With
    Next i
    
    ' Gap Pistas desabilitado por padrao
    txtGapPistas.Enabled = False
    txtGapPistas.BackColor = H(24, 31, 44)
    txtGapPistas.ForeColor = H(58, 78, 98)
End Sub

' ============================================================
' TEMA � RADIOS (Labels simulando radio buttons)
' ============================================================
Private Sub AplicarTemaRadios()
    Dim radios As Variant
    radios = Array("lbl114", "lbl170", "lblPi314", "lblPi3175", _
                   "lblRed638", "lblRed622", "lblRed9", "lblRed95", "lblRed10")
    
    Dim i As Long
    For i = 0 To UBound(radios)
        Dim lbl As MSForms.Label
        Set lbl = Me.Controls(radios(i))
        With lbl
            .BackColor = H(30, 42, 58)
            .ForeColor = H(154, 176, 200)
            .Font.Name = "Segoe UI"
            .Font.Size = 8
            .Font.Bold = False
            .TextAlign = fmTextAlignCenter
            .BorderStyle = fmBorderStyleNone
            .Tag = ""
        End With
    Next i
End Sub

' ============================================================
' TEMA � RESULTADOS (labels somente leitura)
' ============================================================
Private Sub AplicarTemaResultados()
    Dim results As Variant
    results = Array("lblDesenvolvimento", "lblGapReps", "lblGapPistas", _
                    "lblReducao", "lblPasso")
    
    Dim i As Long
    For i = 0 To UBound(results)
        Dim lbl As MSForms.Label
        Set lbl = Me.Controls(results(i))
        With lbl
            .BackColor = H(26, 32, 48)
            .ForeColor = H(210, 180, 80)   ' Amarelo dourado � resultado
            .Font.Name = "Segoe UI"
            .Font.Size = 8
            .Font.Bold = True
            .TextAlign = fmTextAlignRight
            .Caption = ChrW(8212)   ' em dash
        End With
    Next i
End Sub

' ============================================================
' TEMA � BOTOES DE ACAO (Labels frmFlexo)
' ============================================================
Private Sub AplicarTemaBotoes()
    ' Montar � estilo acao (azul destaque)
    With Me.btnMontar
        .BackColor = H(26, 58, 94)
        .ForeColor = H(106, 172, 232)
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .BorderStyle = fmBorderStyleNone
        .Caption = vbCrLf & ChrW(9654) & "  MONTAR"
    End With
    
    ' Reset � estilo padrao
    With Me.btnReset
        .BackColor = H(30, 42, 58)
        .ForeColor = H(154, 176, 200)
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .BorderStyle = fmBorderStyleNone
        .Caption = vbCrLf & ChrW(8635) & "  RESET"
    End With
End Sub

' ============================================================
' TOOLTIPS
' ============================================================
Private Sub AplicarTooltips()
    Me.txtZ.ControlTipText = "Numero de dentes da engrenagem"
    Me.txtLarguraFaca.ControlTipText = "Largura da faca em mm"
    Me.txtAlturaFaca.ControlTipText = "Altura da faca em mm"
    Me.txtLarguraMaterial.ControlTipText = "Largura do material (opcional)"
    Me.txtPistas.ControlTipText = "Numero de pistas"
    Me.txtRepeticoes.ControlTipText = "Numero de repeticoes no cilindro"
    Me.txtGapPistas.ControlTipText = "Gap entre pistas em mm (ativo se pistas > 1)"
    Me.btnMontar.ControlTipText = "Executar montagem no documento ativo"
    Me.btnReset.ControlTipText = "Limpar todos os campos"
    Me.lblCameronArquivo.ControlTipText = "Clique para selecionar o arquivo CDR do Cameron"
    Me.chkCameron.ControlTipText = "Inserir Cameron externo (.cdr) na montagem"
    Me.chkCameronCenter.ControlTipText = "Centralizar Cameron entre as pistas"
    Me.chkRelatorio.ControlTipText = "Gerar relatorio apos a montagem"
End Sub

' ============================================================
' CAMERON — SELETOR DE ARQUIVO (FileDialog via PowerShell)
' Mesmo padrao do Mod07_InserirMicropontos
' ============================================================
Private Function EscolherArquivoCDRCameron() As String
    Dim oShell      As Object
    Dim sTmpFile    As String
    Dim sDirInicial As String
    Dim sScript     As String
    Dim sCmd        As String
    Dim iFile       As Integer
    Dim sResultado  As String

    sTmpFile = Environ("TEMP") & "\mod03_cameron_path.txt"

    If mCameronFilePath <> "" Then
        sDirInicial = Left(mCameronFilePath, InStrRev(mCameronFilePath, "\"))
    Else
        sDirInicial = "C:\"
    End If

    sScript = "Add-Type -AssemblyName System.Windows.Forms;" & _
              "[System.Windows.Forms.Application]::EnableVisualStyles();" & _
              "$d = New-Object System.Windows.Forms.OpenFileDialog;" & _
              "$d.Title = 'Step & Repeat - Selecione o CDR do Cameron';" & _
              "$d.Filter = 'CorelDRAW (*.cdr)|*.cdr|Todos os arquivos (*.*)|*.*';" & _
              "$d.FilterIndex = 1;" & _
              "$d.InitialDirectory = '" & sDirInicial & "';" & _
              "$d.CheckFileExists = $true;" & _
              "if ($d.ShowDialog() -eq 'OK') {" & _
              "  [System.IO.File]::WriteAllText('" & sTmpFile & "', $d.FileName)" & _
              "} else {" & _
              "  [System.IO.File]::WriteAllText('" & sTmpFile & "', '')" & _
              "}"

    sCmd = "powershell.exe -NoProfile -WindowStyle Hidden -Command """ & sScript & """"

    Set oShell = CreateObject("WScript.Shell")
    oShell.Run sCmd, 0, True
    Set oShell = Nothing

    sResultado = ""
    If Dir(sTmpFile) <> "" Then
        iFile = FreeFile
        Open sTmpFile For Input As #iFile
        If Not EOF(iFile) Then
            Line Input #iFile, sResultado
        End If
        Close #iFile
        Kill sTmpFile
    End If

    EscolherArquivoCDRCameron = Trim(sResultado)
End Function

' ============================================================
' CAMERON — ATUALIZAR LABEL DE ARQUIVO
' ============================================================
Private Sub AtualizarLabelCameron()
    With Me.lblCameronArquivo
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .BackColor = H(17, 24, 34)
        .BorderStyle = fmBorderStyleNone
        If mCameronFilePath = "" Then
            .Caption = "Selecionar..."
            .ForeColor = H(106, 172, 232)   ' Azul — convida ao clique
        Else
            Dim parts() As String
            parts = Split(mCameronFilePath, "\")
            .Caption = ChrW(9654) & " " & parts(UBound(parts))
            .ForeColor = H(154, 176, 200)   ' Cinza claro — preenchido
        End If
    End With
End Sub

' ============================================================
' HOVER / PRESS � padrao frmFlexo
' ============================================================
Private Sub AplicarHover(lbl As MSForms.Label)
    lbl.BackColor = H(36, 50, 68)
    lbl.ForeColor = H(192, 212, 232)
End Sub

Private Sub RemoverHover(lbl As MSForms.Label)
    If lbl.Name = "btnMontar" Then
        lbl.BackColor = H(26, 58, 94)
        lbl.ForeColor = H(106, 172, 232)
    Else
        lbl.BackColor = H(30, 42, 58)
        lbl.ForeColor = H(154, 176, 200)
    End If
End Sub

Private Sub AplicarPress(lbl As MSForms.Label)
    lbl.BackColor = H(21, 28, 43)
    lbl.ForeColor = H(192, 212, 232)
End Sub


' ============================================================
' EVENTOS � btnMontar
' ============================================================
Private Sub btnMontar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarHover Me.btnMontar
End Sub
Private Sub btnMontar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarPress Me.btnMontar
End Sub
Private Sub btnMontar_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RemoverHover Me.btnMontar
    ExecutarMontagemDoForm
End Sub

' ============================================================
' EVENTOS � btnReset
' ============================================================
Private Sub btnReset_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarHover Me.btnReset
End Sub
Private Sub btnReset_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AplicarPress Me.btnReset
End Sub
Private Sub btnReset_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RemoverHover Me.btnReset
    ResetarCampos
End Sub

' ============================================================
' EVENTOS � RADIOS ESPESSURA
' ============================================================
Private Sub lbl114_Click()
    lbl114.Tag = "selected"
    lbl170.Tag = ""
    AtualizarRadioVisual
    AtualizarFrameReducao
    RecalcularTudo
End Sub
Private Sub lbl170_Click()
    lbl170.Tag = "selected"
    lbl114.Tag = ""
    AtualizarRadioVisual
    AtualizarFrameReducao
    RecalcularTudo
End Sub

' ============================================================
' EVENTOS � RADIOS PI
' ============================================================
Private Sub lblPi314_Click()
    lblPi314.Tag = "selected"
    lblPi3175.Tag = ""
    AtualizarRadioVisual
    RecalcularTudo
End Sub
Private Sub lblPi3175_Click()
    lblPi3175.Tag = "selected"
    lblPi314.Tag = ""
    AtualizarRadioVisual
    RecalcularTudo
End Sub

' ============================================================
' EVENTOS � RADIOS REDUCAO 1,14
' ============================================================
Private Sub lblRed638_Click()
    lblRed638.Tag = "selected"
    lblRed622.Tag = ""
    AtualizarRadioVisual
    RecalcularTudo
End Sub
Private Sub lblRed622_Click()
    lblRed622.Tag = "selected"
    lblRed638.Tag = ""
    AtualizarRadioVisual
    RecalcularTudo
End Sub

' ============================================================
' EVENTOS � RADIOS REDUCAO 1,70
' ============================================================
Private Sub lblRed9_Click()
    lblRed9.Tag = "selected": lblRed95.Tag = "": lblRed10.Tag = ""
    AtualizarRadioVisual
    RecalcularTudo
End Sub
Private Sub lblRed95_Click()
    lblRed95.Tag = "selected": lblRed9.Tag = "": lblRed10.Tag = ""
    AtualizarRadioVisual
    RecalcularTudo
End Sub
Private Sub lblRed10_Click()
    lblRed10.Tag = "selected": lblRed9.Tag = "": lblRed95.Tag = ""
    AtualizarRadioVisual
    RecalcularTudo
End Sub

' ============================================================
' VISUAL � RADIO SELECTED/UNSELECTED
' ============================================================
Private Sub AtualizarRadioVisual()
    Dim radios As Variant
    radios = Array("lbl114", "lbl170", "lblPi314", "lblPi3175", _
                   "lblRed638", "lblRed622", "lblRed9", "lblRed95", "lblRed10")
    
    Dim i As Long
    For i = 0 To UBound(radios)
        Dim lbl As MSForms.Label
        Set lbl = Me.Controls(radios(i))
        If lbl.Tag = "selected" Then
            lbl.BackColor = H(26, 58, 94)     ' Azul acao
            lbl.ForeColor = H(106, 172, 232)   ' Azul claro
            lbl.Font.Bold = True
        Else
            lbl.BackColor = H(30, 42, 58)      ' Fundo padrao
            lbl.ForeColor = H(154, 176, 200)   ' Texto padrao
            lbl.Font.Bold = False
        End If
    Next i
End Sub

' ============================================================
' FRAME REDUCAO � alterna opcoes conforme espessura
' ============================================================
Private Sub AtualizarFrameReducao()
    Dim is114 As Boolean
    is114 = (lbl114.Tag = "selected")
    
    ' Mostrar/ocultar labels de reducao
    lblRed638.Visible = is114:      lblRed638.Enabled = is114
    lblRed622.Visible = is114:      lblRed622.Enabled = is114
    lblRed9.Visible = Not is114:    lblRed9.Enabled = Not is114
    lblRed95.Visible = Not is114:   lblRed95.Enabled = Not is114
    lblRed10.Visible = Not is114:   lblRed10.Enabled = Not is114

    
    ' Reset selecao de reducao
    If is114 Then
        lblRed638.Tag = "selected"
        lblRed622.Tag = ""
    Else
        lblRed9.Tag = "selected"
        lblRed95.Tag = ""
        lblRed10.Tag = ""
    End If
    AtualizarRadioVisual
End Sub

' ============================================================
' EVENTOS � INPUTS CHANGE
' ============================================================
Private Sub txtZ_Change()
    RecalcularTudo
End Sub
Private Sub txtAlturaFaca_Change()
    RecalcularTudo
End Sub
Private Sub txtLarguraFaca_Change()
    RecalcularTudo
End Sub
Private Sub txtPistas_Change()
    Dim p As Long
    p = val(txtPistas.Text)
    
    ' Habilitar Gap Pistas se > 1
    If p > 1 Then
        txtGapPistas.Enabled = True
        txtGapPistas.BackColor = H(17, 24, 34)
        txtGapPistas.ForeColor = H(154, 176, 200)
    Else
        txtGapPistas.Enabled = False
        txtGapPistas.BackColor = H(24, 31, 44)
        txtGapPistas.ForeColor = H(58, 78, 98)
        txtGapPistas.Text = ""
    End If
    
    ' Cameron centralizado so aparece com >= 2 pistas
    If p >= 2 Then
        chkCameronCenter.Visible = chkCameron.Value
    Else
        chkCameronCenter.Visible = False
        chkCameronCenter.Value = False
    End If
    
    RecalcularTudo
End Sub
Private Sub txtRepeticoes_Change()
    RecalcularTudo
End Sub
Private Sub txtGapPistas_Change()
    RecalcularTudo
End Sub

Private Sub chkCameron_Click()
    Dim p As Long
    p = val(txtPistas.Text)

    ' Cameron Central so com >= 2 pistas
    If p >= 2 Then
        chkCameronCenter.Visible = chkCameron.Value
    Else
        chkCameronCenter.Visible = False
        chkCameronCenter.Value = False
    End If

    ' Mostrar/ocultar label de arquivo
    lblCameronArquivo.Visible = chkCameron.Value

    ' Se marcou e ainda nao tem arquivo: abrir dialog automaticamente
    If chkCameron.Value And mCameronFilePath = "" Then
        Dim sPath As String
        sPath = EscolherArquivoCDRCameron()
        If sPath <> "" Then
            mCameronFilePath = sPath
        Else
            ' Usuário cancelou: desmarca o checkbox
            chkCameron.Value = False
            lblCameronArquivo.Visible = False
        End If
    End If

    AtualizarLabelCameron
End Sub

' ============================================================
' CAMERON — LABEL CLICAVEL (trocar arquivo)
' ============================================================
Private Sub lblCameronArquivo_Click()
    Dim sPath As String
    sPath = EscolherArquivoCDRCameron()
    If sPath <> "" Then
        mCameronFilePath = sPath
        AtualizarLabelCameron
    End If
End Sub

Private Sub lblCameronArquivo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblCameronArquivo.ForeColor = H(192, 212, 232)   ' Hover — azul mais claro
End Sub

' ============================================================
' RECALCULAR TUDO
' ============================================================
Private Sub RecalcularTudo()
    Dim piVal As Double
    If lblPi314.Tag = "selected" Then piVal = PI_PADRAO Else piVal = PI_ALT
    
    Dim zVal As Double
    zVal = val(txtZ.Text)
    
    Dim Desenvolvimento As Double
    Desenvolvimento = piVal * zVal
    
    Dim altFaca As Double
    altFaca = val(txtAlturaFaca.Text)
    
    Dim reps As Long
    reps = val(txtRepeticoes.Text)
    
    ' Reducao
    Dim Reducao As Double
    If lbl114.Tag = "selected" Then
        If lblRed638.Tag = "selected" Then Reducao = RED_114_638 Else Reducao = RED_114_622
    Else
        If lblRed9.Tag = "selected" Then
            Reducao = RED_170_9
        ElseIf lblRed95.Tag = "selected" Then
            Reducao = RED_170_95
        Else
            Reducao = RED_170_10
        End If
    End If
    
    ' Gap Reps
    Dim GapReps As Double
    If reps > 0 And Desenvolvimento > 0 Then
        GapReps = (Desenvolvimento / reps) - altFaca
    End If
    
    ' Passo (Distorcao)
    Dim Passo As Double
    Passo = Desenvolvimento - Reducao
    
    ' Gap Pistas
    Dim gapPistasVal As Double
    gapPistasVal = val(txtGapPistas.Text)
    
    ' Atualizar labels de resultado
    Dim hasData As Boolean
    hasData = (Desenvolvimento > 0 And reps > 0 And altFaca > 0)
    
    If hasData Then
        lblDesenvolvimento.Caption = Format(TruncarDecimal(Desenvolvimento, 2), "0.00") & " mm"
        lblGapReps.Caption = Format(TruncarDecimal(GapReps, 2), "0.00") & " mm"
        lblReducao.Caption = Format(Reducao, "0.00") & " mm"
        lblPasso.Caption = Format(TruncarDecimal(Passo, 2), "0.00") & " mm"
        
        If val(txtPistas.Text) > 1 And gapPistasVal > 0 Then
            lblGapPistas.Caption = Format(TruncarDecimal(gapPistasVal, 2), "0.00") & " mm"
        Else
            lblGapPistas.Caption = ChrW(8212)
        End If
        
        ' Cor do gap negativo
        If GapReps < 0 Then
            lblGapReps.ForeColor = H(220, 80, 80)   ' Vermelho alerta
        Else
            lblGapReps.ForeColor = H(210, 180, 80)   ' Amarelo dourado
        End If
    Else
        lblDesenvolvimento.Caption = ChrW(8212)
        lblGapReps.Caption = ChrW(8212)
        lblGapPistas.Caption = ChrW(8212)
        lblReducao.Caption = ChrW(8212)
        lblPasso.Caption = ChrW(8212)
        lblGapReps.ForeColor = H(210, 180, 80)
    End If
End Sub

' ============================================================
' EXECUTAR MONTAGEM
' ============================================================
Private Sub ExecutarMontagemDoForm()
    ' Validacao de campos obrigatorios
    If val(txtZ.Text) <= 0 Then
        MsgBox "Informe o numero de dentes (Z).", vbExclamation, "Step & Repeat"
        txtZ.SetFocus: Exit Sub
    End If
    If val(txtAlturaFaca.Text) <= 0 Then
        MsgBox "Informe a altura da faca.", vbExclamation, "Step & Repeat"
        txtAlturaFaca.SetFocus: Exit Sub
    End If
    If val(txtLarguraFaca.Text) <= 0 Then
        MsgBox "Informe a largura da faca.", vbExclamation, "Step & Repeat"
        txtLarguraFaca.SetFocus: Exit Sub
    End If
    If val(txtRepeticoes.Text) < 1 Then
        MsgBox "Informe o numero de repeticoes.", vbExclamation, "Step & Repeat"
        txtRepeticoes.SetFocus: Exit Sub
    End If
    If val(txtPistas.Text) > 1 And val(txtGapPistas.Text) <= 0 Then
        MsgBox "Informe o gap entre pistas.", vbExclamation, "Step & Repeat"
        txtGapPistas.SetFocus: Exit Sub
    End If
    If chkCameron.Value And mCameronFilePath = "" Then
        MsgBox "Selecione o arquivo CDR do Cameron antes de montar.", _
               vbExclamation, "Step & Repeat"
        Exit Sub
    End If

    Dim cfg As TStepRepeatConfig

    cfg.BandaEstreita = True
    cfg.Z = val(txtZ.Text)
    
    If lblPi314.Tag = "selected" Then
        cfg.PiValue = PI_PADRAO
    Else
        cfg.PiValue = PI_ALT
    End If
    
    cfg.Desenvolvimento = cfg.PiValue * cfg.Z
    cfg.LarguraFaca = val(txtLarguraFaca.Text)
    cfg.AlturaFaca = val(txtAlturaFaca.Text)
    cfg.LarguraMaterial = val(txtLarguraMaterial.Text)
    cfg.Pistas = val(txtPistas.Text)
    If cfg.Pistas < 1 Then cfg.Pistas = 1
    cfg.Repeticoes = val(txtRepeticoes.Text)
    cfg.GapPistas = val(txtGapPistas.Text)
    
    If lbl114.Tag = "selected" Then
        cfg.Foto114 = True
        If lblRed638.Tag = "selected" Then cfg.Reducao = RED_114_638 Else cfg.Reducao = RED_114_622
    Else
        cfg.Foto114 = False
        If lblRed9.Tag = "selected" Then
            cfg.Reducao = RED_170_9
        ElseIf lblRed95.Tag = "selected" Then
            cfg.Reducao = RED_170_95
        Else
            cfg.Reducao = RED_170_10
        End If
    End If
    
    cfg.GapReps = 0
    If cfg.Repeticoes > 0 And cfg.Desenvolvimento > 0 Then
        cfg.GapReps = (cfg.Desenvolvimento / cfg.Repeticoes) - cfg.AlturaFaca
    End If
    
    cfg.Passo = cfg.Desenvolvimento - cfg.Reducao
    cfg.IncluirCameron = chkCameron.Value
    cfg.CameronCentral = chkCameronCenter.Value
    cfg.CameronFilePath = mCameronFilePath

    cfg.GerarRelatorio = chkRelatorio.Value
    
    ' Executar
    Mod02_Montagem.ExecutarMontagem cfg
End Sub

' ============================================================
' RESET
' ============================================================
Private Sub ResetarCampos()
    txtZ.Text = ""
    txtLarguraFaca.Text = ""
    txtAlturaFaca.Text = ""
    txtLarguraMaterial.Text = ""
    txtPistas.Text = ""
    txtRepeticoes.Text = ""
    txtGapPistas.Text = ""
    
    chkCameron.Value = False
    chkCameronCenter.Value = False
    chkCameronCenter.Visible = False
    mCameronFilePath = ""
    lblCameronArquivo.Visible = False
    AtualizarLabelCameron
    chkRelatorio.Value = True
    
    ' Reset radios
    lbl114.Tag = "selected": lbl170.Tag = ""
    lblPi314.Tag = "selected": lblPi3175.Tag = ""
    lblRed638.Tag = "selected": lblRed622.Tag = ""
    AtualizarRadioVisual
    AtualizarFrameReducao
    RecalcularTudo
End Sub

' ============================================================
' LEAVE � remover hover quando mouse sai (padrao frmFlexo)
' ============================================================
Private Sub frameEspessura_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RemoverHover Me.btnMontar: RemoverHover Me.btnReset
End Sub
Private Sub frameDimensoes_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RemoverHover Me.btnMontar: RemoverHover Me.btnReset
End Sub
Private Sub frameEspacamento_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RemoverHover Me.btnMontar: RemoverHover Me.btnReset
End Sub
Private Sub frameReducao_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RemoverHover Me.btnMontar: RemoverHover Me.btnReset
End Sub
Private Sub frameOpcoes_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RemoverHover Me.btnMontar: RemoverHover Me.btnReset
End Sub
Private Sub frameResultados_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RemoverHover Me.btnMontar: RemoverHover Me.btnReset
End Sub

' ============================================================
' RADIO HOVER (labels de radio com hover)
' ============================================================
Private Sub lbl114_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If lbl114.Tag <> "selected" Then AplicarHover lbl114
End Sub
Private Sub lbl170_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If lbl170.Tag <> "selected" Then AplicarHover lbl170
End Sub
Private Sub lblPi314_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If lblPi314.Tag <> "selected" Then AplicarHover lblPi314
End Sub
Private Sub lblPi3175_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If lblPi3175.Tag <> "selected" Then AplicarHover lblPi3175
End Sub
Private Sub lblRed638_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If lblRed638.Tag <> "selected" Then AplicarHover lblRed638
End Sub
Private Sub lblRed622_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If lblRed622.Tag <> "selected" Then AplicarHover lblRed622
End Sub
Private Sub lblRed9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If lblRed9.Tag <> "selected" Then AplicarHover lblRed9
End Sub
Private Sub lblRed95_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If lblRed95.Tag <> "selected" Then AplicarHover lblRed95
End Sub
Private Sub lblRed10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If lblRed10.Tag <> "selected" Then AplicarHover lblRed10
End Sub


