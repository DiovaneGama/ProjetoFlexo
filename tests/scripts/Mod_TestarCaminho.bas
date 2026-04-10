Attribute VB_Name = "Mod_TestarCaminho"
Sub TestarCaminho()
    MsgBox Dir("X:\version.txt")
End Sub
Sub IdentificarDrives()
    Dim msg As String
    Dim drives(2) As String
    Dim i As Integer
    
    drives(0) = "G:\"
    drives(1) = "X:\"
    drives(2) = "Y:\"
    
    For i = 0 To 2
        msg = msg & drives(i) & " ? " & Dir(drives(i), vbDirectory) & vbCrLf
    Next i
    
    MsgBox msg
End Sub

Sub TestarDrives()
    MsgBox "G: " & Dir("G:\Trabalhos Geral Manos\Programas\ConsoleFlexo2.0\version.txt") & vbCrLf & _
           "X: " & Dir("X:\Trabalhos Geral Manos\Programas\ConsoleFlexo2.0\version.txt") & vbCrLf & _
           "Y: " & Dir("Y:\Trabalhos Geral Manos\Programas\ConsoleFlexo2.0\version.txt")
End Sub

Sub ListarPastasRaiz()
    Dim pasta As String
    Dim msg As String
    
    ' Listar pastas raiz de cada drive
    msg = "G:\ ? " & vbCrLf
    pasta = Dir("G:\", vbDirectory)
    Do While pasta <> ""
        If pasta <> "." And pasta <> ".." Then msg = msg & "  " & pasta & vbCrLf
        pasta = Dir
    Loop
    
    msg = msg & vbCrLf & "X:\ ? " & vbCrLf
    pasta = Dir("X:\", vbDirectory)
    Do While pasta <> ""
        If pasta <> "." And pasta <> ".." Then msg = msg & "  " & pasta & vbCrLf
        pasta = Dir
    Loop
    
    msg = msg & vbCrLf & "Y:\ ? " & vbCrLf
    pasta = Dir("Y:\", vbDirectory)
    Do While pasta <> ""
        If pasta <> "." And pasta <> ".." Then msg = msg & "  " & pasta & vbCrLf
        pasta = Dir
    Loop
    
    MsgBox msg
End Sub
