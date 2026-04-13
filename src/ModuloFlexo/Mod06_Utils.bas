Attribute VB_Name = "Mod06_Utils"
' ============================================================
' MODULO: Mod06_Utils (UTILITARIOS COMPARTILHADOS)
' DESCRICAO: Funcoes auxiliares usadas por mais de um modulo.
'            Centraliza logica que antes estava duplicada.
' ============================================================
Option Explicit

' ============================================================
' FUNCAO: TemBordaDura
' Retorna True se o gradiente do shape possui borda dura
' (hard edge) -- ou seja, algum canal vai a zero enquanto
' outro no tem valor significativo, ou combinacao Spot + branco CMYK.
'
' Usada por:
'   Mod02_Scanner_Engine.AnalisarGradiente  -- contagem no relatorio
'   (Mod02_Cores.CorrigirBordaDuraGradientes mantem Fase 1 propria
'    pois Fase 2 precisa das variaveis intermediarias maxC/temSpot/etc.)
' ============================================================
Public Function TemBordaDura(shp As Shape) As Boolean
    TemBordaDura = False
    On Error Resume Next

    Dim totalCores As Integer
    totalCores = 2 + shp.Fill.Fountain.Colors.Count
    If Err.Number <> 0 Then On Error GoTo 0: Exit Function

    Dim cores() As Color
    ReDim cores(1 To totalCores)
    Set cores(1) = shp.Fill.Fountain.StartColor
    Set cores(2) = shp.Fill.Fountain.EndColor
    Dim K As Integer
    For K = 0 To shp.Fill.Fountain.Colors.Count - 1
        Set cores(3 + K) = shp.Fill.Fountain.Colors(K).Color
    Next K

    ' Fase 1: le o DNA do gradiente
    Dim maxC As Long, maxM As Long, maxY As Long, maxK As Long, maxTint As Long
    Dim temSpot As Boolean, temBrancoCMYK As Boolean
    maxC = 0: maxM = 0: maxY = 0: maxK = 0: maxTint = 0
    temSpot = False: temBrancoCMYK = False

    For K = 1 To totalCores
        If cores(K).Type = cdrColorCMYK Then
            If cores(K).CMYKCyan    > maxC Then maxC = cores(K).CMYKCyan
            If cores(K).CMYKMagenta > maxM Then maxM = cores(K).CMYKMagenta
            If cores(K).CMYKYellow  > maxY Then maxY = cores(K).CMYKYellow
            If cores(K).CMYKBlack   > maxK Then maxK = cores(K).CMYKBlack
            If (cores(K).CMYKCyan + cores(K).CMYKMagenta + cores(K).CMYKYellow + cores(K).CMYKBlack) = 0 Then
                temBrancoCMYK = True
            End If
        ElseIf cores(K).Type = cdrColorSpot Then
            If cores(K).Tint > maxTint Then maxTint = cores(K).Tint
            If cores(K).Tint > 0 Then temSpot = True
        End If
    Next K

    ' Fase 2: detecta borda dura
    Dim resultado As Boolean: resultado = False
    For K = 1 To totalCores
        If cores(K).Type = cdrColorCMYK Then
            If maxC > 0 And cores(K).CMYKCyan    = 0 Then resultado = True: Exit For
            If maxM > 0 And cores(K).CMYKMagenta = 0 Then resultado = True: Exit For
            If maxY > 0 And cores(K).CMYKYellow  = 0 Then resultado = True: Exit For
            If maxK > 0 And cores(K).CMYKBlack   = 0 Then resultado = True: Exit For
        ElseIf cores(K).Type = cdrColorSpot Then
            If maxTint > 0 And cores(K).Tint = 0 Then resultado = True: Exit For
        End If
    Next K

    ' Combinacao Spot + branco CMYK = tambem e borda dura
    If Not resultado Then
        If temSpot And temBrancoCMYK Then resultado = True
    End If

    On Error GoTo 0
    TemBordaDura = resultado
End Function
