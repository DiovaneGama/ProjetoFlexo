Attribute VB_Name = "modCameron"
' ============================================================
' modCameron.bas - Marcas de Registro Cameron
' Step & Repeat - Banda Estreita
' ============================================================
Option Explicit

' ============================================================
' INSERIR CAMERON
' Retorna ShapeRange com os shapes de Cameron criados
' ============================================================
Public Function InserirCameron(cfg As TStepRepeatConfig, grp As Shape) As ShapeRange
    Dim lyr As Layer
    Set lyr = ActivePage.ActiveLayer

    Dim leftX   As Double, rightX  As Double
    Dim bottomY As Double
    leftX   = grp.LeftX
    rightX  = grp.RightX
    bottomY = grp.BottomY

    ' Altura igual ao grupo ja com reducao aplicada
    Dim camAltura As Double
    camAltura = grp.SizeHeight

    Dim camLarg As Double
    camLarg = CAMERON_ESPESSURA   ' 1 mm

    Dim resultado As New ShapeRange
    Dim shpCam As Shape

    If cfg.CameronCentral And cfg.Pistas >= 2 Then
        ' ============================================
        ' CAMERON CENTRALIZADO - entre pistas
        ' ============================================
        Dim centroX As Double
        centroX = (leftX + rightX) / 2 - (camLarg / 2)

        Set shpCam = lyr.CreateRectangle2(centroX, bottomY, camLarg, camAltura)
        shpCam.Fill.UniformColor.RGBAssign 0, 180, 0
        shpCam.Outline.SetNoOutline
        shpCam.Name = "Cameron_Centro"
        resultado.Add shpCam
    Else
        ' ============================================
        ' CAMERON LATERAL - colado na montagem (sem offset)
        ' ============================================
        ' Esquerda
        Set shpCam = lyr.CreateRectangle2(leftX - camLarg, bottomY, camLarg, camAltura)
        shpCam.Fill.UniformColor.RGBAssign 255, 0, 0
        shpCam.Outline.SetNoOutline
        shpCam.Name = "Cameron_Esq"
        resultado.Add shpCam

        ' Direita
        Set shpCam = lyr.CreateRectangle2(rightX, bottomY, camLarg, camAltura)
        shpCam.Fill.UniformColor.RGBAssign 255, 0, 0
        shpCam.Outline.SetNoOutline
        shpCam.Name = "Cameron_Dir"
        resultado.Add shpCam
    End If

    Set InserirCameron = resultado
End Function
