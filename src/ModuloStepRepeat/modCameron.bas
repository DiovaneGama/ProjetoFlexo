Attribute VB_Name = "modCameron"
' ============================================================
' modCameron.bas — Marcas de Registro Cameron
' Step & Repeat — Banda Estreita
' Usa arquivo .cdr externo como modelo, escalado na altura do grupo
' ============================================================
Option Explicit

' ============================================================
' INSERIR CAMERON
' ============================================================
Public Sub InserirCameron(cfg As TStepRepeatConfig, grp As Shape)
    ' Validar arquivo
    If Len(Trim(cfg.CameronFilePath)) = 0 Then
        MsgBox "Nenhum arquivo Cameron selecionado.", vbExclamation, "Cameron"
        Exit Sub
    End If
    If Dir(cfg.CameronFilePath) = "" Then
        MsgBox "Arquivo Cameron nao encontrado:" & vbCrLf & cfg.CameronFilePath, vbExclamation, "Cameron"
        Exit Sub
    End If

    Dim alturaGrupo As Double
    alturaGrupo = grp.TopY - grp.BottomY   ' altura ja com reducao aplicada

    If cfg.CameronCentral And cfg.Pistas >= 2 Then
        ' ============================================
        ' CAMERON CENTRALIZADO — entre pistas
        ' ============================================
        Dim shpCentro As Shape
        Set shpCentro = ImportarEscalarCameron(cfg.CameronFilePath, alturaGrupo)

        Dim centroX As Double
        centroX = (grp.LeftX + grp.RightX) / 2 - (shpCentro.SizeWidth / 2)
        shpCentro.SetPosition centroX, grp.BottomY
        shpCentro.Name = "Cameron_Centro"
    Else
        ' ============================================
        ' CAMERON LATERAL — esquerda + direita
        ' ============================================
        Dim shpEsq As Shape
        Set shpEsq = ImportarEscalarCameron(cfg.CameronFilePath, alturaGrupo)
        shpEsq.SetPosition grp.LeftX - shpEsq.SizeWidth, grp.BottomY
        shpEsq.Name = "Cameron_Esq"

        Dim shpDir As Shape
        Set shpDir = ImportarEscalarCameron(cfg.CameronFilePath, alturaGrupo)
        shpDir.SetPosition grp.RightX, grp.BottomY
        shpDir.Name = "Cameron_Dir"
    End If
End Sub

' ============================================================
' IMPORTAR E ESCALAR CAMERON
' ============================================================
Private Function ImportarEscalarCameron(filePath As String, alturaAlvo As Double) As Shape
    ' Importar arquivo .cdr
    ActivePage.Import filePath

    ' Pegar o objeto importado (fica selecionado apos import)
    Dim shp As Shape
    If ActiveDocument.Selection.Shapes.Count = 0 Then
        MsgBox "Falha ao importar Cameron. Verifique o arquivo.", vbCritical, "Cameron"
        Exit Function
    End If

    ' Se importou multiplos objetos, agrupar
    If ActiveDocument.Selection.Shapes.Count > 1 Then
        Set shp = ActiveDocument.Selection.Group
    Else
        Set shp = ActiveDocument.Selection.Shapes(1)
    End If

    ' Escalar altura mantendo proporcao da largura
    If shp.SizeHeight > 0 Then
        Dim escala As Double
        escala = alturaAlvo / shp.SizeHeight
        shp.SizeWidth  = shp.SizeWidth * escala
        shp.SizeHeight = alturaAlvo
    End If

    Set ImportarEscalarCameron = shp
End Function
