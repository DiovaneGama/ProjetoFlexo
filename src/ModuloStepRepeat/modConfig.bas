Attribute VB_Name = "modConfig"
' ==========================================================
' modConfig.bas - Step & Repeat - Constantes e Utilitarios
' ==========================================================
Option Explicit

' ==========================================================
' CONSTANTES DE PI
' ==========================================================
Public Const PI_PADRAO      As Double = 3.14159
Public Const PI_ALT         As Double = 3.175

' ==========================================================
' REDUCOES - FOTOPOLIMERO 1,14 mm
' ==========================================================
Public Const RED_114_638     As Double = 6.38
Public Const RED_114_622     As Double = 6.22

' ==========================================================
' REDUCOES - FOTOPOLIMERO 1,70 mm
' ==========================================================
Public Const RED_170_9       As Double = 9#
Public Const RED_170_95      As Double = 9.5
Public Const RED_170_10      As Double = 10#

' ==========================================================
' CAMERON
' ==========================================================
Public Const CAMERON_ESPESSURA   As Double = 1#

' ==========================================================
' TIPO CONFIGURACAO
' ==========================================================
Public Type TStepRepeatConfig
    BandaEstreita   As Boolean
    Z               As Double
    Cilindro        As Double
    PiValue         As Double
    LarguraFaca     As Double
    AlturaFaca      As Double
    LarguraMaterial As Double
    Pistas          As Long
    Repeticoes      As Long
    GapPistas       As Double
    Foto114         As Boolean
    Reducao         As Double
    IncluirCameron  As Boolean
    CameronCentral  As Boolean
    CameronFilePath As String
    GerarRelatorio  As Boolean
    Desenvolvimento As Double
    GapReps         As Double
    Passo           As Double
End Type

' ==========================================================
' SELECIONAR ARQUIVO CAMERON - File Dialog via PowerShell
' ==========================================================
Public Function SelecionarArquivoCDR() As String
    Dim sTempPS1    As String
    Dim sTempResult As String
    Dim oShell      As Object
    Dim ff          As Integer
    Dim sLine       As String

    sTempPS1    = Environ("TEMP") & "\sr_dialog.ps1"
    sTempResult = Environ("TEMP") & "\sr_result.txt"

    If Dir(sTempResult) <> "" Then Kill sTempResult

    ff = FreeFile
    Open sTempPS1 For Output As #ff
    Print #ff, "Add-Type -AssemblyName System.Windows.Forms"
    Print #ff, "$d = New-Object System.Windows.Forms.OpenFileDialog"
    Print #ff, "$d.Filter = 'CorelDRAW (*.cdr)|*.cdr'"
    Print #ff, "$d.Title = 'Selecionar arquivo Cameron'"
    Print #ff, "if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {"
    Print #ff, "    [System.IO.File]::WriteAllText('" & sTempResult & "', $d.FileName)"
    Print #ff, "}"
    Close #ff

    Set oShell = CreateObject("WScript.Shell")
    oShell.Run "powershell -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & sTempPS1 & """", 0, True

    If Dir(sTempResult) <> "" Then
        ff = FreeFile
        Open sTempResult For Input As #ff
        Line Input #ff, sLine
        Close #ff
        Kill sTempResult
        SelecionarArquivoCDR = Trim(sLine)
    End If

    If Dir(sTempPS1) <> "" Then Kill sTempPS1
End Function

' ==========================================================
' TRUNCAR SEM ARREDONDAR
' ==========================================================
Public Function TruncarDecimal(dVal As Double, iCasas As Integer) As Double
    Dim f As Double
    f = 10 ^ iCasas
    TruncarDecimal = Int(dVal * f) / f
End Function

' ==========================================================
' HELPER RGB
' ==========================================================
Public Function H(R As Long, G As Long, B As Long) As Long
    H = RGB(R, G, B)
End Function
