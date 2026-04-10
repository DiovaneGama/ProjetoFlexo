Attribute VB_Name = "Mod02_Montagem"
' ============================================================
' Mod02_Montagem.bas — Logica de Posicionamento Step & Repeat
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
    If ActiveDocument.Selection.Shapes.Count = 0 Then
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
    Dim oldUnit As cdrUnit
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
    origemY = shpOriginal.BottomY
    
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

    ' Aplicar reducao: a altura final do grupo deve ser exatamente cfg.Passo
    ' (= Desenvolvimento - Reducao), que e a circunferencia distorcida real.
    If cfg.Reducao > 0 And cfg.Passo > 0 Then
        grpFinal.SizeHeight = cfg.Passo
    End If

    ' Desabilitar otimizacao antes do Cameron para garantir leitura correta
    ' das coordenadas do grupo apos o SizeHeight (evita valores stale)
    Application.Optimization = False

    ' Cameron + agrupar tudo + centralizar
    Dim tudo As New ShapeRange
    tudo.Add grpFinal

    If cfg.IncluirCameron Then
        Dim camShapes As ShapeRange
        Set camShapes = Mod03_Cameron.InserirCameron(cfg, grpFinal)
        Dim i As Long
        For i = 1 To camShapes.Count
            tudo.Add camShapes(i)
        Next i
    End If

    ' Agrupar montagem + cameron num unico grupo
    Dim grpTotal As Shape
    If tudo.Count > 1 Then
        Set grpTotal = tudo.Group
    Else
        Set grpTotal = grpFinal
    End If

    ' Centralizar na pagina ativa
    Dim pgW As Double, pgH As Double
    pgW = ActivePage.SizeWidth
    pgH = ActivePage.SizeHeight
    grpTotal.SetPosition (pgW - grpTotal.SizeWidth) / 2, (pgH - grpTotal.SizeHeight) / 2

    ' Relatorio
    If cfg.GerarRelatorio Then
        Mod04_Relatorio.GerarRelatorio cfg
    End If

    ' Finalizar
    ActiveDocument.EndCommandGroup
    Application.Optimization = False
    ActiveDocument.Unit = oldUnit
    ActiveWindow.Refresh
    
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
    ActiveWindow.Refresh
    MsgBox "Erro durante a montagem: " & Err.Description, vbCritical, "Step & Repeat"
End Sub

' ============================================================
' PONTO DE ENTRADA — aparece no dialogo "Executar Macro"
' Ferramentas > Macros > Executar Macro > StepRepeat.MostrarStepRepeat
' ============================================================
Public Sub MostrarStepRepeat()
    frmStepRepeat.Show 0   ' 0 = vbModeless (nao bloqueia o CorelDRAW)
End Sub
