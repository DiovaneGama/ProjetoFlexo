Attribute VB_Name = "Mod00_Versao"
' ============================================================
'  Mod00_Versao — Código completo
' ============================================================

Private Const VERSAO_ATUAL    As String = "2.1"
Private Const CAMINHO_VERSION As String = "X:\version.txt"
Private Const CAMINHO_GMS     As String = "X:\ConsoleFlexo.gms"
Private bAtualizando          As Boolean  ' <- flag de controle

Public Sub VerificarVersao()
    On Error GoTo Falhou
    
    ' Se já está no processo de atualização, ignora
    If bAtualizando Then Exit Sub
    
    Dim fileNum  As Integer
    Dim linha    As String
    Dim versao   As String
    Dim notas    As String
    Dim primeira As Boolean
    
    primeira = True
    fileNum = FreeFile
    
    Open CAMINHO_VERSION For Input As #fileNum
        Do While Not EOF(fileNum)
            Line Input #fileNum, linha
            If primeira Then
                versao = Trim(linha)
                primeira = False
            Else
                notas = notas & linha & vbCrLf
            End If
        Loop
    Close #fileNum
    
    If versao = VERSAO_ATUAL Then Exit Sub
    
    MsgBox "Nova versão disponível!" & vbCrLf & vbCrLf & _
           "  Versão instalada:    " & VERSAO_ATUAL & vbCrLf & _
           "  Versão disponível:  " & versao & vbCrLf & vbCrLf & _
           "O que há de novo:" & vbCrLf & Trim(notas) & vbCrLf & vbCrLf & _
           "O sistema será atualizado agora.", _
           vbOKOnly + vbInformation, "ConsoleFlexo — Atualização"
    
    bAtualizando = True  ' <- bloqueia nova execução durante o Load
    GlobalMacroStorage.Load CAMINHO_GMS
    bAtualizando = False
    
    MsgBox "Atualização concluída!" & vbCrLf & _
           "Por favor, execute o comando novamente.", _
           vbInformation, "ConsoleFlexo — Atualizado"
    Exit Sub
Falhou:
    If Not bAtualizando Then
        MsgBox "Não foi possível verificar atualizações." & vbCrLf & _
               "Verifique a conexão com o servidor.", _
               vbExclamation, "ConsoleFlexo — Aviso"
    End If
End Sub

