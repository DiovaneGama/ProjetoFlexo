'==============================================================================
' MÓDULO  : Mod07_InserirMicropontos
' VERSÃO  : 2.7  — Removidos Application.Optimization e EventsEnabled (ausentes na API v27)
'                  Substituído Shapes.CreateRange por New ShapeRange
'                  Corrigido: ImportEx pertence a Layer, não a Page (API v27)
' APP     : CorelDRAW 2026 — API v27
' AUTOR   : [Seu Nome / Estúdio]
' DATA    : 2026
'
' DESCRIÇÃO:
'   Na primeira execução (ou quando o usuário pedir), abre um FileDialog
'   para localizar o arquivo .CDR do microponto (já em cor de Registro).
'   O caminho escolhido é guardado em uma variável de módulo (cache) e
'   reutilizado nas execuções seguintes da mesma sessão do CorelDRAW,
'   eliminando a necessidade de navegar toda vez.
'
'   Posiciona 4 cópias do microponto ao redor do objeto selecionado com
'   offset exato de 2 mm, agrupa os 4 pontos e move-os para a camada
'   "micropontos" (criada automaticamente se não existir).
'
' PRÉ-REQUISITOS:
'   1. Um único objeto deve estar selecionado na página ativa.
'   2. Um arquivo .CDR contendo APENAS o microponto em cor de Registro.
'
' COMO INSTALAR:
'   Ferramentas > Macros > Gerenciar Macros > abra/crie um projeto GMS
'   e cole este módulo. Para executar: chame InserirMicropontos().
'   Para redefinir o caminho sem reiniciar o Corel: chame RedefinirCaminho().
'==============================================================================

Option Explicit

' ─── CACHE DE SESSÃO ──────────────────────────────────────────────────────────
' Armazena o caminho escolhido pelo usuário durante a sessão atual.
' Resetado ao fechar o CorelDRAW ou ao chamar RedefinirCaminho().
Private m_CaminhoCDR As String

' Offset em milímetros para fora do bounding box do objeto alvo
Private Const OFFSET_MM As Double = 1.5

'==============================================================================
' FUNÇÃO AUXILIAR — Abre o seletor de arquivos via PowerShell + .NET
'
' Estratégia: invoca PowerShell de forma síncrona usando WScript.Shell.
' O script PS instancia System.Windows.Forms.OpenFileDialog (diálogo 100%
' nativo do Windows), exibe ao usuário e grava o caminho escolhido em um
' arquivo temporário .txt que o VBA então lê de volta.
'
' Vantagens:
'   - Funciona no GMS 32-bit e 64-bit sem qualquer Declare/API
'   - Suporta caminhos UNC (\\server\share\...), espaços e acentos
'   - Sem dependência de MSForms, FileDialog do Corel ou COMDLG32
' Retorna o caminho completo, ou "" se o usuário cancelar.
'==============================================================================
Private Function EscolherArquivoCDR(ByVal sCaminhoAnterior As String) As String

    Dim oShell    As Object
    Dim sTmpFile  As String
    Dim sDirInicial As String
    Dim sScript   As String
    Dim sCmd      As String
    Dim iFile     As Integer
    Dim sResultado As String

    ' ── Arquivo temporário onde o PS vai gravar o caminho escolhido ──────────
    sTmpFile = Environ("TEMP") & "\mod07_path.txt"

    ' ── Diretório inicial do diálogo ─────────────────────────────────────────
    If sCaminhoAnterior <> "" Then
        sDirInicial = Left(sCaminhoAnterior, InStrRev(sCaminhoAnterior, "\"))
    Else
        sDirInicial = "C:\"
    End If

    ' ── Script PowerShell (uma linha; executado inline via -Command) ──────────
    ' Carrega WinForms, cria o OpenFileDialog, exibe e grava resultado no .txt
    sScript = "Add-Type -AssemblyName System.Windows.Forms;" & _
              "[System.Windows.Forms.Application]::EnableVisualStyles();" & _
              "$d = New-Object System.Windows.Forms.OpenFileDialog;" & _
              "$d.Title = 'Mod07 - Selecione o arquivo CDR do Microponto';" & _
              "$d.Filter = 'CorelDRAW (*.cdr)|*.cdr|Todos os arquivos (*.*)|*.*';" & _
              "$d.FilterIndex = 1;" & _
              "$d.InitialDirectory = '" & sDirInicial & "';" & _
              "$d.CheckFileExists = $true;" & _
              "if ($d.ShowDialog() -eq 'OK') {" & _
              "  [System.IO.File]::WriteAllText('" & sTmpFile & "', $d.FileName)" & _
              "} else {" & _
              "  [System.IO.File]::WriteAllText('" & sTmpFile & "', '')" & _
              "}"

    ' ── Monta o comando: powershell.exe -NoProfile -WindowStyle Hidden ────────
    sCmd = "powershell.exe -NoProfile -WindowStyle Hidden -Command """ & sScript & """"

    ' ── Executa de forma SÍNCRONA (bWaitOnReturn = True) ─────────────────────
    Set oShell = CreateObject("WScript.Shell")
    oShell.Run sCmd, 0, True          ' 0 = janela oculta, True = aguarda fim
    Set oShell = Nothing

    ' ── Lê o resultado do arquivo temporário ─────────────────────────────────
    sResultado = ""
    If Dir(sTmpFile) <> "" Then
        iFile = FreeFile
        Open sTmpFile For Input As #iFile
        Line Input #iFile, sResultado
        Close #iFile
        Kill sTmpFile                 ' Limpa o arquivo temp
    End If

    EscolherArquivoCDR = Trim(sResultado)

End Function

'==============================================================================
' SUB PÚBLICA — Permite redefinir o arquivo CDR sem reiniciar o CorelDRAW.
' Útil quando o arquivo de microponto for substituído por uma versão nova.
'==============================================================================
Public Sub RedefinirCaminho()
    m_CaminhoCDR = ""
    MsgBox "Caminho do microponto redefinido." & vbCrLf & _
           "Na próxima execução, o seletor de arquivos será exibido novamente.", _
           vbInformation, "Mod07 — Caminho Redefinido"
End Sub

'==============================================================================
' PROCEDIMENTO PRINCIPAL
'==============================================================================
Public Sub InserirMicropontos()

    ' ── Referências de trabalho ──────────────────────────────────────────────
    Dim oDoc      As Document
    Dim oPage     As Page
    Dim oAlvo     As Shape       ' Objeto selecionado pelo usuário
    Dim origPonto As Shape       ' Microponto importado (original temporário)
    Dim ptTopo    As Shape       ' Clone — posição superior
    Dim ptBase    As Shape       ' Clone — posição inferior
    Dim ptEsq     As Shape       ' Clone — posição esquerda
    Dim ptDir     As Shape       ' Clone — posição direita
    Dim oGrupo    As Shape       ' Grupo final dos 4 micropontos
    Dim oLayer    As Layer       ' Camada "micropontos"
    Dim oLayers   As Layers
    Dim oRange    As ShapeRange
    Dim oImport   As ImportFilter

    ' ── Bounding box do alvo ─────────────────────────────────────────────────
    Dim dLeft    As Double
    Dim dRight   As Double
    Dim dTop     As Double
    Dim dBottom  As Double
    Dim dCentroX As Double
    Dim dCentroY As Double

    ' ── Dimensões do microponto ───────────────────────────────────────────────
    Dim dPontoW As Double
    Dim dPontoH As Double

    ' ── Auxiliares ───────────────────────────────────────────────────────────
    Dim i               As Integer
    Dim bCamadaExiste   As Boolean
    Dim sCaminhoEscolhido As String

    '==========================================================================
    ' ETAPA 1 — CONFIGURAÇÃO DE AMBIENTE
    '==========================================================================
    ' On Error GoTo ErroHandler  ' << DESABILITADO PARA DEPURAÇÃO — reativar em produção

    Set oDoc  = ActiveDocument
    Set oPage = ActivePage

    oDoc.Unit = cdrMillimeter
    oDoc.BeginCommandGroup "Inserir Micropontos"

    '==========================================================================
    ' ETAPA 2 — VALIDAÇÃO DO OBJETO ALVO
    '==========================================================================
    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Nenhum objeto selecionado." & vbCrLf & _
               "Selecione o objeto alvo antes de executar esta macro.", _
               vbExclamation, "Mod07 — Micropontos"
        GoTo Finalizar
    End If

    Set oAlvo = ActiveSelection.Shapes(1)

    '==========================================================================
    ' ETAPA 3 — RESOLUÇÃO DO CAMINHO  (cache de sessão + FileDialog)
    '
    '  Três cenários possíveis:
    '  A) Cache válido e arquivo ainda existe  → usa direto, sem diálogo.
    '  B) Cache preenchido, mas arquivo sumiu  → avisa e reabre o diálogo.
    '  C) Cache vazio (primeira execução)      → abre o diálogo.
    '==========================================================================

    ' Cenário B — arquivo em cache não existe mais no disco
    If m_CaminhoCDR <> "" And Dir(m_CaminhoCDR) = "" Then
        MsgBox "O arquivo CDR armazenado não foi encontrado:" & vbCrLf & _
               m_CaminhoCDR & vbCrLf & vbCrLf & _
               "Por favor, selecione o arquivo novamente.", _
               vbExclamation, "Mod07 — Arquivo não encontrado"
        m_CaminhoCDR = ""   ' Limpa cache para forçar nova seleção
    End If

    ' Cenários B e C — cache vazio: abre o seletor de arquivos
    If m_CaminhoCDR = "" Then

        ' Abre o seletor de arquivos nativo do Windows (COMDLG32)
        ' Não precisa restaurar EventsEnabled — API independente do Corel
        sCaminhoEscolhido = EscolherArquivoCDR(m_CaminhoCDR)

        ' Usuário pressionou Cancelar
        If sCaminhoEscolhido = "" Then
            MsgBox "Nenhum arquivo selecionado. Operação cancelada.", _
                   vbInformation, "Mod07 — Cancelado"
            GoTo Finalizar
        End If

        ' Validação extra: arquivo realmente existe
        If Dir(sCaminhoEscolhido) = "" Then
            MsgBox "O arquivo informado não foi encontrado:" & vbCrLf & _
                   sCaminhoEscolhido, _
                   vbCritical, "Mod07 — Arquivo inválido"
            GoTo Finalizar
        End If

        ' Persiste no cache de sessão
        m_CaminhoCDR = sCaminhoEscolhido

    End If
    ' Cenário A (e B/C após confirmação): m_CaminhoCDR é válido aqui.

    '==========================================================================
    ' ETAPA 4 — IMPORTAÇÃO ÚNICA DO MICROPONTO
    ' API v27: ImportEx pertence a Layer, não a Page.
    '==========================================================================
    Set oImport = oPage.ActiveLayer.ImportEx(m_CaminhoCDR, cdrCDR)
    oImport.Finish

    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "A importação não retornou nenhum objeto." & vbCrLf & _
               "Verifique se o arquivo CDR contém exatamente 1 objeto.", _
               vbCritical, "Mod07 — Erro de Importação"
        GoTo Finalizar
    End If

    Set origPonto = ActiveSelection.Shapes(1)
    dPontoW = origPonto.SizeWidth
    dPontoH = origPonto.SizeHeight

    '==========================================================================
    ' ETAPA 5 — BOUNDING BOX DO OBJETO ALVO
    '==========================================================================
    dLeft    = oAlvo.LeftX
    dRight   = oAlvo.RightX
    dTop     = oAlvo.TopY
    dBottom  = oAlvo.BottomY
    dCentroX = (dLeft  + dRight)  / 2#
    dCentroY = (dTop   + dBottom) / 2#

    '==========================================================================
    ' ETAPA 6 — DUPLICAÇÃO E POSICIONAMENTO
    '
    '  SetPosition(x, y) define o canto INFERIOR-ESQUERDO do shape.
    '  Metade das dimensões do microponto é descontada para centralizá-lo
    '  com precisão sobre cada ponto de destino.
    '==========================================================================

    ' ── Ponto SUPERIOR ── centro X | Y = top + 2 mm
    Set ptTopo = origPonto.Duplicate
    ptTopo.SetPosition _
        dCentroX - (dPontoW / 2#), _
        dTop + OFFSET_MM + dPontoH

    ' ── Ponto INFERIOR ── centro X | Y = bottom - 2 mm
    Set ptBase = origPonto.Duplicate
    ptBase.SetPosition _
        dCentroX - (dPontoW / 2#), _
        dBottom - OFFSET_MM

    ' ── Ponto ESQUERDO ── X = left - 2 mm | centro Y
    Set ptEsq = origPonto.Duplicate
    ptEsq.SetPosition _
        dLeft - OFFSET_MM - dPontoW, _
        dCentroY - (dPontoH / 2#)

    ' ── Ponto DIREITO ── X = right + 2 mm | centro Y
    Set ptDir = origPonto.Duplicate
    ptDir.SetPosition _
        dRight + OFFSET_MM, _
        dCentroY - (dPontoH / 2#)

    ' Deleta o original temporário
    origPonto.Delete
    Set origPonto = Nothing

    '==========================================================================
    ' ETAPA 7 — AGRUPAMENTO DOS 4 MICROPONTOS
    '==========================================================================
    Set oRange = New ShapeRange
    oRange.Add ptTopo
    oRange.Add ptBase
    oRange.Add ptEsq
    oRange.Add ptDir

    Set oGrupo = oRange.Group
    Set oRange = Nothing

    '==========================================================================
    ' ETAPA 8 — GESTÃO DA CAMADA "micropontos"
    '==========================================================================
    Set oLayers   = oPage.Layers
    bCamadaExiste = False

    For i = 1 To oLayers.Count
        If LCase(Trim(oLayers(i).Name)) = "micropontos" Then
            Set oLayer    = oLayers(i)
            bCamadaExiste = True
            Exit For
        End If
    Next i

    If Not bCamadaExiste Then
        Set oLayer = oPage.CreateLayer("micropontos")
    End If

    oLayer.Visible  = True
    oLayer.Editable = True
    oGrupo.Layer    = oLayer

    '==========================================================================
    ' FINALIZAÇÃO BEM-SUCEDIDA
    '==========================================================================
    MsgBox "Micropontos inseridos com sucesso!" & vbCrLf & vbCrLf & _
           "Arquivo utilizado:" & vbCrLf & m_CaminhoCDR, _
           vbInformation, "Mod07 — Concluído"
    GoTo Finalizar

    '==========================================================================
    ' TRATAMENTO DE ERROS  (desabilitado para depuração)
    '==========================================================================
'ErroHandler:
'    MsgBox "Erro inesperado durante a execução:" & vbCrLf & _
'           "Número : " & Err.Number      & vbCrLf & _
'           "Origem : " & Err.Source      & vbCrLf & _
'           "Detalhe: " & Err.Description, _
'           vbCritical, "Mod07 — Erro"

Finalizar:
    ' ── Restaura ambiente ────────────────────────────────────────────────────
    On Error Resume Next

    oDoc.EndCommandGroup

    Set oImport   = Nothing
    Set oRange    = Nothing
    Set oGrupo    = Nothing
    Set oLayer    = Nothing
    Set oLayers   = Nothing
    Set ptTopo    = Nothing
    Set ptBase    = Nothing
    Set ptEsq     = Nothing
    Set ptDir     = Nothing
    Set origPonto = Nothing
    Set oAlvo     = Nothing
    Set oPage     = Nothing
    Set oDoc      = Nothing

End Sub
'==============================================================================
' FIM DO MÓDULO Mod07_InserirMicropontos  (v2.7)
'==============================================================================
