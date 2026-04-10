Attribute VB_Name = "Mod01_Setup"
Sub AbrirPainelFlexo()
    ' vbModeless permite que você continue editando o arquivo
    ' enquanto o painel de ferramentas fica aberto ao lado.
    VerificarVersao
    frmFlexo.Show vbModeless
End Sub
