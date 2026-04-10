Attribute VB_Name = "modStepRepeat"
' ============================================================
' modStepRepeat.bas — Logica de Posicionamento Step & Repeat
' Banda Estreita — CorelDRAW 2026 v27
' ============================================================
Option Explicit

' ============================================================
' EXECUTAR MONTAGEM
' ============================================================
Public Sub ExecutarMontagem(cfg As TStepRepeatConfig)
    ' Validar selecao
    If ActiveDocument Is Nothing Then
        MsgBox "Nenhum documento aberto.", vbExclamation, "Step & Repeat"
        Exit Sub
    End If
    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Selecione um objeto (faca/arte) antes de montar.", vbExclamation, "Step & Repeat"
        Exit Sub
    End If
    
    ' Validar gap negativo
    If cfg.GapReps < 0 Then
        MsgBox "Gap entre repeticoes negativo! Ha sobreposicao de facas." & vbCrLf & _
               "Gap = " & Format(cfg.GapReps, "0.00") & " mm", vbCritical, "Step & Repeat"
        Exit Sub
    End If
    
    ' Salvar estado
    Dim oldUnit As Long
    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter
    
    ' Performance
    Application.Optimization = True
    ActiveDocument.BeginCommandGroup "Step & Repeat - Montagem"
    
    On Error GoTo ErrHandler
    
    Dim shpOriginal As Shape
    Set shpOriginal = ActiveSelection.Shapes(1)
    
    Dim origemX As Double, origemY As Double
    origemX = shpOriginal.LeftX
    origemY = shpOriginal.TopY
    
    Dim allShapes As New ShapeRange
    allShapes.Add shpOriginal
    
    Dim p As Long, r As Long
    Dim posX As Double, posY As Double
    Dim shpCopy As Shape
    
    For p = 0 To cfg.Pistas - 1
        For r = 0 To cfg.Repeticoes - 1
            ' Pular o original (posicao 0,0)
            If p = 0 And r = 0 Then
                ' Posicionar o original na origem
                shpOriginal.SetPosition origemX, origemY
            Else
                ' Duplicar e posicionar
                Set shpCopy = shpOriginal.Duplicate
                posX = origemX + p * (cfg.LarguraFaca + cfg.GapPistas)
                posY = origemY - r * (cfg.AlturaFaca + cfg.GapReps)
                shpCopy.SetPosition posX, posY
                allShapes.Add shpCopy
            End If
        Next r
    Next p
    
    ' Agrupar tudo
    Dim grpFinal As Shape
    If allShapes.Count > 1 Then
        Set grpFinal = allShapes.Group
    Else
        Set grpFinal = shpOriginal
    End If
    
    ' Cameron
    If cfg.IncluirCameron Then
        modCameron.InserirCameron cfg, grpFinal
    End If
    
    ' Relatorio
    If cfg.GerarRelatorio Then
        modRelatorio.GerarRelatorio cfg
    End If
    
    ' Finalizar
    ActiveDocument.EndCommandGroup
    Application.Optimization = False
    ActiveDocument.Unit = oldUnit
    
    ' Feedback
    MsgBox "Montagem concluida!" & vbCrLf & _
           cfg.Pistas & " x " & cfg.Repeticoes & " = " & (cfg.Pistas * cfg.Repeticoes) & " repeticoes" & vbCrLf & _
           "Desenvolvimento: " & Format(TruncarDecimal(cfg.Desenvolvimento, 2), "0.00") & " mm" & vbCrLf & _
           "Gap Reps: " & Format(TruncarDecimal(cfg.GapReps, 2), "0.00") & " mm" & vbCrLf & _
           "Passo: " & Format(TruncarDecimal(cfg.Passo, 2), "0.00") & " mm", _
           vbInformation, "Step & Repeat"
    Exit Sub

ErrHandler:
    ActiveDocument.EndCommandGroup
    Application.Optimization = False
    ActiveDocument.Unit = oldUnit
    MsgBox "Erro durante a montagem: " & Err.Description, vbCritical, "Step & Repeat"
End Sub
