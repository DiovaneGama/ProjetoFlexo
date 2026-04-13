' ============================================================
'  MACRO: PadronizarLayers
'  Aplicativo: CorelDRAW 2026 (v27)
'  Descrição:  Garante que a PÁGINA ATIVA possui as layers
'              abaixo, criando as que faltarem e reordenando
'              todas na sequência correta.
'
'  Ordem final no painel (topo → base):
'    1. Trimbox
'    2. Informacoes
'    3. Micropontos
'    4. Branco
'    5. Arte
'    6. Material
'
'  API oficial (Programming Guide v27):
'    Page.CreateLayer(name)   → cria layer no topo
'    Layer.MoveAbove(layer)   → move acima de outra layer
'    Layer.MoveBelow(layer)   → move abaixo de outra layer
'    Page.Layers(name)        → acessa layer pelo nome
' ============================================================

Option Explicit

Private Const NUM_LAYERS As Integer = 6

Private Function LayerOrder() As String()
    Dim arr(1 To NUM_LAYERS) As String
    arr(1) = "Trimbox"
    arr(2) = "Informacoes"
    arr(3) = "Micropontos"
    arr(4) = "Branco"
    arr(5) = "Arte"
    arr(6) = "Material"
    LayerOrder = arr
End Function

' ============================================================
'  PONTO DE ENTRADA PRINCIPAL
' ============================================================
Public Sub PadronizarLayers()

    If ActiveDocument Is Nothing Then
        MsgBox "Nenhum documento aberto.", vbExclamation, "PadronizarLayers"
        Exit Sub
    End If

    Dim pg As Page
    Set pg = ActiveDocument.ActivePage

    Dim names() As String
    names = LayerOrder()

    ' --------------------------------------------------------
    ' 1. Criar layers que ainda não existem
    '    Page.CreateLayer insere sempre no topo
    ' --------------------------------------------------------
    Dim i As Integer
    For i = 1 To NUM_LAYERS
        If Not LayerExists(pg, names(i)) Then
            pg.CreateLayer names(i)
        End If
    Next i

    ' --------------------------------------------------------
    ' 2. Reordenar usando MoveAbove / MoveBelow
    '    Estratégia: percorre do fundo (Material) para o topo
    '    (Saida), posicionando cada layer abaixo da anterior.
    '
    '    Resultado: Saida(1) > Trimbox(2) > ... > Material(7)
    ' --------------------------------------------------------

    ' Coloca "Trimbox" no topo via loop até que seja a primeira layer
    Dim lTrimbox As Layer
    Set lTrimbox = pg.Layers(names(1))
    Do While pg.Layers.Item(1).Name <> names(1)
        lTrimbox.MoveAbove pg.Layers.Item(1)
    Loop

    ' Encadeia cada layer abaixo da anterior
    ' names(1)=Trimbox já está no topo; posiciona 2..6 em cascata
    Dim lRef  As Layer
    Dim lMove As Layer
    Set lRef = pg.Layers(names(1))   ' referência inicial = Trimbox

    For i = 2 To NUM_LAYERS
        Set lMove = pg.Layers(names(i))
        lMove.MoveBelow lRef          ' coloca lMove logo abaixo de lRef
        Set lRef = lMove              ' próxima referência é a que acabou de mover
    Next i

    MsgBox "Layers padronizadas com sucesso na página ativa!" & vbCrLf & vbCrLf & _
           "  1. Trimbox" & vbCrLf & _
           "  2. Informacoes" & vbCrLf & _
           "  3. Micropontos" & vbCrLf & _
           "  4. Branco" & vbCrLf & _
           "  5. Arte" & vbCrLf & _
           "  6. Material", _
           vbInformation, "PadronizarLayers"

End Sub

' ============================================================
'  FUNÇÕES AUXILIARES
' ============================================================

' ------------------------------------------------------------
' Verifica se uma layer com o nome dado existe na página.
' Usa acesso por nome via pg.Layers(name) com tratamento de
' erro, pois a coleção lança erro se o nome não existir.
' ------------------------------------------------------------
Private Function LayerExists(pg As Page, layerName As String) As Boolean
    Dim l As Layer
    LayerExists = False
    For Each l In pg.Layers
        If StrComp(l.Name, layerName, vbTextCompare) = 0 Then
            LayerExists = True
            Exit Function
        End If
    Next l
End Function
