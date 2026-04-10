Attribute VB_Name = "Mod03_Cameron"
' ============================================================
' Mod03_Cameron.bas - Marcas de Registro Cameron
' Step & Repeat - Banda Estreita
' ============================================================
Option Explicit

' ============================================================
' INSERIR CAMERON
' Importa o CDR do Cameron e posiciona copias ao redor do grupo.
' Retorna ShapeRange com os shapes criados.
' Interface publica inalterada — Mod02_Montagem nao precisa mudar.
' ============================================================
Public Function InserirCameron(cfg As TStepRepeatConfig, grp As Shape) As ShapeRange

    Dim resultado As New ShapeRange

    ' Fallback seguro: se nao ha caminho configurado, retorna vazio
    ' O loop em Mod02_Montagem (Count=0) simplesmente nao itera
    If cfg.CameronFilePath = "" Then
        Set InserirCameron = resultado
        Exit Function
    End If

    Dim lyr As Layer
    Set lyr = ActivePage.ActiveLayer

    ' ── Bounding box do grupo ja com reducao aplicada ────────────────────────
    Dim leftX   As Double
    Dim rightX  As Double
    Dim topY    As Double
    Dim bottomY As Double
    Dim centroX As Double
    Dim centroY As Double

    leftX   = grp.LeftX
    rightX  = grp.RightX
    topY    = grp.TopY
    bottomY = grp.BottomY
    centroX = (leftX + rightX)  / 2#
    centroY = (topY  + bottomY) / 2#

    ' ── Importar o CDR original (unico shape esperado no arquivo) ────────────
    Dim oImport  As ImportFilter
    Dim origCam  As Shape
    Dim camLarg  As Double
    Dim camAltura As Double

    Set oImport = lyr.ImportEx(cfg.CameronFilePath, cdrCDR)
    oImport.Finish

    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "A importacao do Cameron nao retornou nenhum objeto." & vbCrLf & _
               "Verifique o arquivo CDR: " & cfg.CameronFilePath, _
               vbCritical, "Step & Repeat — Cameron"
        Set InserirCameron = resultado
        Exit Function
    End If

    Set origCam  = ActiveSelection.Shapes(1)
    camLarg      = origCam.SizeWidth
    camAltura    = origCam.SizeHeight

    ' ── Posicionar copias e construir resultado ───────────────────────────────
    ' O CDR do Cameron e mantido nas proporcoes originais (sem distorcao).
    ' Centralizado verticalmente em relacao ao grupo.
    Dim shpCam As Shape

    If cfg.CameronCentral And cfg.Pistas >= 2 Then
        ' ── CAMERON CENTRALIZADO — entre pistas ──────────────────────────────
        Set shpCam = origCam.Duplicate
        shpCam.SetPosition centroX - (camLarg / 2#), centroY - (camAltura / 2#)
        shpCam.Name = "Cameron_Centro"
        resultado.Add shpCam
    Else
        ' ── CAMERON LATERAL — colado na montagem ─────────────────────────────
        ' Esquerda: borda direita do Cameron encosta em leftX do grupo
        Set shpCam = origCam.Duplicate
        shpCam.SetPosition leftX - camLarg, centroY - (camAltura / 2#)
        shpCam.Name = "Cameron_Esq"
        resultado.Add shpCam

        ' Direita: borda esquerda do Cameron encosta em rightX do grupo
        Set shpCam = origCam.Duplicate
        shpCam.SetPosition rightX, centroY - (camAltura / 2#)
        shpCam.Name = "Cameron_Dir"
        resultado.Add shpCam
    End If

    ' Deleta o original importado (era apenas referencia para duplicar)
    origCam.Delete
    Set origCam = Nothing

    Set InserirCameron = resultado

End Function
