Attribute VB_Name = "Mod01_AbrirScanner"
Public Sub IniciarScanner()
    ' Agora o Mod_Scanner_Engine existe e será encontrado!
    VerificarVersao
    Mod02_Scanner_Engine.ExecutarScanner
End Sub
