Attribute VB_Name = "Mod04_Relatorio"
' ============================================================
' Mod04_Relatorio.bas — Relatorio Tecnico Step & Repeat
' ============================================================
Option Explicit

' ============================================================
' GERAR RELATORIO
' ============================================================
Public Sub GerarRelatorio(cfg As TStepRepeatConfig)
    Dim sep As String
    sep = String(45, "=")
    
    Dim rpt As String
    rpt = sep & vbCrLf
    rpt = rpt & "  RELATORIO — STEP & REPEAT v1.0" & vbCrLf
    rpt = rpt & "  " & Format(Now, "dd/mm/yyyy hh:mm:ss") & vbCrLf
    rpt = rpt & sep & vbCrLf & vbCrLf
    
    ' Configuracao
    rpt = rpt & "CONFIGURACAO" & vbCrLf
    rpt = rpt & String(45, "-") & vbCrLf
    If cfg.BandaEstreita Then
        rpt = rpt & "  Tipo Montagem : Banda Estreita" & vbCrLf
        rpt = rpt & "  Z (dentes)    : " & cfg.Z & vbCrLf
        rpt = rpt & "  Pi            : " & cfg.PiValue & vbCrLf
    Else
        rpt = rpt & "  Tipo Montagem : Banda Larga" & vbCrLf
        rpt = rpt & "  Cilindro      : " & Format(cfg.Cilindro, "0.00") & " mm" & vbCrLf
    End If
    
    If cfg.Foto114 Then
        rpt = rpt & "  Fotopolimero  : 1,14 mm" & vbCrLf
    Else
        rpt = rpt & "  Fotopolimero  : 1,70 mm" & vbCrLf
    End If
    rpt = rpt & vbCrLf
    
    ' Dimensoes
    rpt = rpt & "DIMENSOES DA FACA" & vbCrLf
    rpt = rpt & String(45, "-") & vbCrLf
    rpt = rpt & "  Largura       : " & Format(cfg.LarguraFaca, "0.00") & " mm" & vbCrLf
    rpt = rpt & "  Altura        : " & Format(cfg.AlturaFaca, "0.00") & " mm" & vbCrLf
    If cfg.LarguraMaterial > 0 Then
        rpt = rpt & "  Larg. Material: " & Format(cfg.LarguraMaterial, "0.00") & " mm" & vbCrLf
    End If
    rpt = rpt & vbCrLf
    
    ' Layout
    rpt = rpt & "LAYOUT" & vbCrLf
    rpt = rpt & String(45, "-") & vbCrLf
    rpt = rpt & "  Pistas        : " & cfg.Pistas & vbCrLf
    rpt = rpt & "  Repeticoes    : " & cfg.Repeticoes & vbCrLf
    rpt = rpt & "  Total         : " & (cfg.Pistas * cfg.Repeticoes) & " unidades" & vbCrLf
    rpt = rpt & "  Cameron       : " & IIf(cfg.IncluirCameron, "Sim", "Nao") & vbCrLf
    If cfg.IncluirCameron And cfg.CameronCentral Then
        rpt = rpt & "  Cameron Pos.  : Centralizado entre pistas" & vbCrLf
    ElseIf cfg.IncluirCameron Then
        rpt = rpt & "  Cameron Pos.  : Laterais (colado, sem offset)" & vbCrLf
    End If
    rpt = rpt & vbCrLf
    
    ' Calculos
    rpt = rpt & "CALCULOS" & vbCrLf
    rpt = rpt & String(45, "-") & vbCrLf
    rpt = rpt & "  Desenvolvimento : " & Format(TruncarDecimal(cfg.Desenvolvimento, 2), "0.00") & " mm" & vbCrLf
    rpt = rpt & "  Reducao         : " & Format(cfg.Reducao, "0.00") & " mm" & vbCrLf
    rpt = rpt & "  Passo(Distorcao): " & Format(TruncarDecimal(cfg.Passo, 2), "0.00") & " mm" & vbCrLf
    rpt = rpt & "  Gap entre Reps  : " & Format(TruncarDecimal(cfg.GapReps, 2), "0.00") & " mm" & vbCrLf
    If cfg.Pistas > 1 Then
        rpt = rpt & "  Gap entre Pistas: " & Format(TruncarDecimal(cfg.GapPistas, 2), "0.00") & " mm" & vbCrLf
    End If
    rpt = rpt & vbCrLf
    
    ' Validacao
    rpt = rpt & "VALIDACAO" & vbCrLf
    rpt = rpt & String(45, "-") & vbCrLf
    If cfg.GapReps >= 0 Then
        rpt = rpt & "  Status: OK — Gap positivo, sem sobreposicao" & vbCrLf
    Else
        rpt = rpt & "  Status: ERRO — Gap negativo, ha sobreposicao!" & vbCrLf
    End If
    
    If cfg.LarguraMaterial > 0 Then
        Dim largTotal As Double
        largTotal = cfg.Pistas * cfg.LarguraFaca + (cfg.Pistas - 1) * cfg.GapPistas
        Dim aproveitamento As Double
        aproveitamento = (largTotal / cfg.LarguraMaterial) * 100
        rpt = rpt & "  Aproveitamento  : " & Format(aproveitamento, "0.0") & "%" & vbCrLf
    End If
    
    rpt = rpt & vbCrLf & sep
    
    ' Exibir
    MsgBox rpt, vbInformation, "Relatorio Step & Repeat"
End Sub
