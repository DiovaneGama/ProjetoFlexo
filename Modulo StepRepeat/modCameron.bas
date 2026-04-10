Attribute VB_Name = "modCameron"
' ============================================================
' modCameron.bas — Marcas de Registro Cameron
' Step & Repeat — Banda Estreita
' ============================================================
Option Explicit

' ============================================================
' INSERIR CAMERON
' ============================================================
Public Sub InserirCameron(cfg As TStepRepeatConfig, grp As Shape)
    Dim lyr As Layer
    Set lyr = ActivePage.ActiveLayer
    
    Dim leftX As Double, rightX As Double
    Dim topY As Double, bottomY As Double
    leftX = grp.LeftX
    rightX = grp.RightX
    topY = grp.TopY
    bottomY = grp.BottomY
    
    Dim camAltura As Double
    camAltura = cfg.Desenvolvimento
    
    Dim camLarg As Double
    camLarg = CAMERON_ESPESSURA   ' 1 mm
    
    Dim shpCam As Shape
    
    If cfg.CameronCentral And cfg.Pistas >= 2 Then
        ' ============================================
        ' CAMERON CENTRALIZADO — entre pistas
        ' ============================================
        Dim centroX As Double
        centroX = (leftX + rightX) / 2 - (camLarg / 2)
        
        Set shpCam = lyr.CreateRectangle2(centroX, bottomY, camLarg, camAltura)
        shpCam.Fill.UniformColor.RGBAssign 0, 180, 0  ' Verde
        shpCam.Outline.SetNoOutline
        shpCam.Name = "Cameron_Centro"
    Else
        ' ============================================
        ' CAMERON LATERAL — esquerda + direita
        ' ============================================
        Dim camLeftX As Double, camRightX As Double
        camLeftX = leftX - CAMERON_OFFSET - camLarg
        camRightX = rightX + CAMERON_OFFSET
        
        ' Esquerda
        Set shpCam = lyr.CreateRectangle2(camLeftX, bottomY, camLarg, camAltura)
        shpCam.Fill.UniformColor.RGBAssign 255, 0, 0  ' Vermelho
        shpCam.Outline.SetNoOutline
        shpCam.Name = "Cameron_Esq"
        
        ' Direita
        Set shpCam = lyr.CreateRectangle2(camRightX, bottomY, camLarg, camAltura)
        shpCam.Fill.UniformColor.RGBAssign 255, 0, 0  ' Vermelho
        shpCam.Outline.SetNoOutline
        shpCam.Name = "Cameron_Dir"
    End If
End Sub
