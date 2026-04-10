Attribute VB_Name = "modConfig"
' ============================================================
' modConfig.bas — Step & Repeat — Constantes e Utilitarios
' Padrao de design: frmFlexo (Console Flexo v2.0)
' ============================================================
Option Explicit

' ============================================================
' CONSTANTES DE PI
' ============================================================
Public Const PI_PADRAO      As Double = 3.14159
Public Const PI_ALT         As Double = 3.175

' ============================================================
' REDUCOES — FOTOPOLIMERO 1,14 mm
' ============================================================
Public Const RED_114_638     As Double = 6.38
Public Const RED_114_622     As Double = 6.22

' ============================================================
' REDUCOES — FOTOPOLIMERO 1,70 mm
' ============================================================
Public Const RED_170_9       As Double = 9#
Public Const RED_170_95      As Double = 9.5
Public Const RED_170_10      As Double = 10#

' ============================================================
' CAMERON
' ============================================================
Public Const CAMERON_ESPESSURA   As Double = 1#     ' mm

' ============================================================
' CORES DO TEMA — Paleta frmFlexo (valores RGB para H())
' ============================================================
' Fundo Form     : RGB( 26,  32,  48) = #1A2030
' Fundo Frame    : RGB( 35,  45,  63) = #232D3F
' Fundo Btn      : RGB( 30,  42,  58) = #1E2A3A
' Fundo Hover    : RGB( 36,  50,  68) = #243244
' Fundo Press    : RGB( 17,  24,  34) = #111822
' Fundo Done     : RGB( 24,  31,  44) = #181F2C
' Fundo Input    : RGB( 17,  24,  34) = #111822
' Fundo Acao     : RGB( 26,  58,  94) = #1A3A5E
' Texto Btn      : RGB(154, 176, 200) = #9AB0C8
' Texto Done     : RGB( 58,  78,  98) = #3A4E62
' Texto Label    : RGB( 74,  88, 112) = #4A5870
' Texto Sec      : RGB(106, 125, 150) = #6A7D96
' Azul Accent    : RGB(106, 172, 232) = #6AACE8
' Borda Sec      : RGB( 35,  45,  63) = #232D3F
' Texto Hover    : RGB(192, 212, 232) = #C0D4E8
' Texto Press    : RGB(230, 240, 252) = #E6F0FC
' Resultado      : RGB(210, 180,  80) = amarelo dourado

' ============================================================
' TIPO CONFIGURACAO
' ============================================================
Public Type TStepRepeatConfig
    ' Entrada
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
    Foto114         As Boolean      ' True = 1,14mm; False = 1,70mm
    Reducao         As Double
    IncluirCameron  As Boolean
    CameronCentral  As Boolean
    GerarRelatorio  As Boolean
    
    ' Calculado
    Desenvolvimento As Double
    GapReps         As Double
    Passo           As Double       ' Desenvolvimento - Reducao
End Type

' ============================================================
' TRUNCAR SEM ARREDONDAR
' ============================================================
Public Function TruncarDecimal(dVal As Double, iCasas As Integer) As Double
    Dim f As Double
    f = 10 ^ iCasas
    TruncarDecimal = Int(dVal * f) / f
End Function

' ============================================================
' HELPER RGB (compativel com frmFlexo)
' ============================================================
Public Function H(R As Long, G As Long, B As Long) As Long
    H = RGB(R, G, B)
End Function
