' ============================================================
'  MACRO: PadronizarLayers
'  Aplicativo: CorelDRAW 2026 (v27)
'  Descrição:  Garante que a PÁGINA ATIVA possui as layers
'              abaixo, criando as que faltarem e reordenando
'              todas na sequência correta.
'
'  Ordem final no painel (topo → base):
'    1. Saida
'    2. Trimbox
'    3. Informacoes
'    4. Micropontos
'    5. Branco
'    6. Arte
'    7. Material
'
'  API oficial (Programming Guide v27):
'    Page.CreateLayer(name)   → cria layer no topo
'    Layer.MoveAbove(layer)   → move acima de outra layer
'    Layer.MoveBelow(layer)   → move abaixo de outra layer
'    Page.Layers(name)        → acessa layer pelo nome
' ============================================================

Option Explicit

Private Const NUM_LAYERS As Integer = 7

Private Function LayerOrder() As String()
    Dim arr(1 To NUM_LAYERS) As String
    arr(1) = "Saida"
    arr(2) = "Trimbox"
    arr(3) = "Informacoes"
    arr(4) = "Micropontos"
    arr(5) = "Branco"
    arr(6) = "Arte"
    arr(7) = "Material"
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

    ' Coloca "Saida" no topo: move acima de todas as outras
    ' fazendo MoveAbove da primeira layer da coleção
    Dim lSaida As Layer
    Set lSaida = pg.Layers(names(1))

    ' Move Saida para o topo (acima da layer que está em pos 1)
    ' Repetimos até que Saida seja a primeira
    Dim lTop As Layer
    Set lTop = pg.Layers.Item(1)
    If lTop.Name <> names(1) Then
        lSaida.MoveAbove lTop
    End If

    ' Agora encadeia cada layer abaixo da anterior
    ' names(1)=Saida já está no topo; posiciona 2..7 em cascata
    Dim lRef  As Layer
    Dim lMove As Layer
    Set lRef = pg.Layers(names(1))   ' referência inicial = Saida

    For i = 2 To NUM_LAYERS
        Set lMove = pg.Layers(names(i))
        lMove.MoveBelow lRef          ' coloca lMove logo abaixo de lRef
        Set lRef = lMove              ' próxima referência é a que acabou de mover
    Next i

    MsgBox "Layers padronizadas com sucesso na página ativa!" & vbCrLf & vbCrLf & _
           "  1. Saida" & vbCrLf & _
           "  2. Trimbox" & vbCrLf & _
           "  3. Informacoes" & vbCrLf & _
           "  4. Micropontos" & vbCrLf & _
           "  5. Branco" & vbCrLf & _
           "  6. Arte" & vbCrLf & _
           "  7. Material", _
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
