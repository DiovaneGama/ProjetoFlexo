Attribute VB_Name = "Mod03_Vetores"
' ============================================================
' M�DULO: Mod03_Vetores (TRATAMENTO ESTRUTURAL E FIRST)
' DESCRI��O: Convers�o de Fontes, Otimiza��o de N�s e Espessura
' ============================================================

Option Explicit

' ============================================================
' FERRAMENTA 1: CONVERTER TODOS OS TEXTOS EM CURVAS
' ============================================================
Public Sub ConverterTextosEmCurvas(Optional silencioso As Boolean = False)
    Dim srTextos As ShapeRange
    Set srTextos = CreateShapeRange
    Dim s As Shape

    For Each s In ActivePage.shapes
        CrawlerBuscaTexto s, srTextos
    Next s

    If srTextos.Count = 0 Then
        If Not silencioso Then MsgBox "Nenhum texto vivo encontrado.", vbInformation, "Console Flexo"
        Exit Sub
    End If

    ActiveDocument.BeginCommandGroup "Console Flexo - Textos em Curvas"
    On Error GoTo FimErro

    Dim i As Integer
    Dim convertidos As Integer: convertidos = 0
    For i = srTextos.Count To 1 Step -1
        srTextos(i).ConvertToCurves
        convertidos = convertidos + 1
    Next i

    ActiveDocument.EndCommandGroup
    If Not silencioso Then MsgBox convertidos & " textos convertidos em curvas!", vbInformation, "Console Flexo"
    Exit Sub

FimErro:
    ActiveDocument.EndCommandGroup
    If Not silencioso Then MsgBox "Erro: " & Err.Description, vbCritical, "Console Flexo"
End Sub

Private Sub CrawlerBuscaTexto(s As Shape, ByRef sacola As ShapeRange)
    Dim subS As Shape
    On Error Resume Next
    If s.Type = cdrTextShape Then sacola.Add s
    On Error GoTo 0
    
    If s.Type = cdrGroupShape Then
        For Each subS In s.shapes: CrawlerBuscaTexto subS, sacola: Next subS
    End If
    If Not s.PowerClip Is Nothing Then
        For Each subS In s.PowerClip.shapes: CrawlerBuscaTexto subS, sacola: Next subS
    End If
End Sub

' ============================================================
' FERRAMENTA 2: INSPETOR DE N�S (Atua apenas na Sele��o)
' ============================================================
Public Sub InspecionarNos()
    ' Trava: Verifica se o operador selecionou algo para inspecionar
    If ActiveSelection.shapes.Count = 0 Then
        MsgBox "Selecione os objetos ou o grupo que voc� deseja inspecionar primeiro!", vbExclamation, "Console Flexo"
        Exit Sub
    End If

    Dim limiteNos As String
    limiteNos = InputBox("Digite a quantidade m�xima de n�s permitida por objeto:" & vbCrLf & "(Acima de 1500 costuma ser lixo de Rastreio Autom�tico)", "Inspetor de N�s", "1500")
    If limiteNos = "" Or Not IsNumeric(limiteNos) Then Exit Sub
    
    Dim maxNos As Long: maxNos = CLng(limiteNos)
    Dim srNos As ShapeRange: Set srNos = CreateShapeRange
    Dim s As Shape
    
    ' Muda a varredura: Agora olha S� para o que o operador selecionou!
    For Each s In ActiveSelection.shapes
        CrawlerBuscaNos s, maxNos, srNos
    Next s
    
    ' Limpa a sele��o inicial do operador para n�o confundir
    ActiveDocument.ClearSelection
    
    If srNos.Count > 0 Then
        ' Seleciona APENAS os objetos defeituosos dentro do grupo que ele havia selecionado
        srNos.CreateSelection
        MsgBox "Aten��o! " & srNos.Count & " objetos DENTRO DA SUA SELE��O possuem mais de " & maxNos & " n�s e foram isolados." & vbCrLf & vbCrLf & _
               "Analise se � poss�vel utilizar o bot�o de Reduzir N�s sem deformar a arte.", vbExclamation, "Console Flexo"
    Else
        MsgBox "Sele��o limpa! Nenhum objeto inspecionado possui excesso de n�s.", vbInformation, "Console Flexo"
    End If
End Sub
' ============================================================
' FERRAMENTA 4: REDUTOR DE N�S SEGURO (AutoReduce)
' ============================================================
Public Sub ReduzirNosSeguro()
    ' Trava de seguran�a
    If ActiveSelection.shapes.Count = 0 Then
        MsgBox "Selecione as curvas que deseja suavizar primeiro!", vbExclamation, "Console Flexo"
        Exit Sub
    End If
    
    Dim s As Shape
    Dim totalNosAntes As Long: totalNosAntes = 0
    Dim totalNosDepois As Long: totalNosDepois = 0
    Dim curvasAfetadas As Integer: curvasAfetadas = 0
    
    ' Fator de Suaviza��o (0.005 � considerado seguro no Corel para n�o deformar logotipos)
    Dim fatorSuavizacao As Double: fatorSuavizacao = 0.005
    
    ' Inicia a varredura blindada APENAS na sele��o
    For Each s In ActiveSelection.shapes
        CrawlerReduzirNos s, fatorSuavizacao, totalNosAntes, totalNosDepois, curvasAfetadas
    Next s
    
    Dim nosRemovidos As Long
    nosRemovidos = totalNosAntes - totalNosDepois
    
    If nosRemovidos > 0 Then
        MsgBox "Limpeza conclu�da! Foram removidos " & nosRemovidos & " n�s in�teis de " & curvasAfetadas & " curva(s)." & vbCrLf & vbCrLf & _
               "DICA: D� um zoom e verifique visualmente se a arte original n�o sofreu deforma��es.", vbInformation, "Console Flexo"
    Else
        MsgBox "Nenhum n� p�de ser removido. Os objetos selecionados j� est�o otimizados ou n�o s�o curvas.", vbInformation, "Console Flexo"
    End If
End Sub

' ------------------------------------------------------------
' O CRAWLER: Varredura Segura para Remo��o de N�s
' ------------------------------------------------------------
Private Sub CrawlerReduzirNos(s As Shape, fator As Double, ByRef nosAntes As Long, ByRef nosDepois As Long, ByRef afetadas As Integer)
    Dim subS As Shape
    On Error Resume Next
    
    ' 1. Modifica o objeto se for Curva
    If s.Type = cdrCurveShape Then
        Dim antes As Long: antes = s.Curve.Nodes.Count
        
        ' Aplica a redu��o nativa do Corel
        s.Curve.AutoReduceNodes fator
        
        Dim depois As Long: depois = s.Curve.Nodes.Count
        
        ' S� contabiliza se realmente conseguiu remover algum n�
        If antes > depois Then
            nosAntes = nosAntes + antes
            nosDepois = nosDepois + depois
            afetadas = afetadas + 1
        End If
    End If
    On Error GoTo 0
    
    ' 2. Mergulha nos Grupos
    If s.Type = cdrGroupShape Then
        For Each subS In s.shapes: CrawlerReduzirNos subS, fator, nosAntes, nosDepois, afetadas: Next subS
    End If
    
    ' 3. Mergulha nos PowerClips
    If Not s.PowerClip Is Nothing Then
        For Each subS In s.PowerClip.shapes: CrawlerReduzirNos subS, fator, nosAntes, nosDepois, afetadas: Next subS
    End If
End Sub

' O CrawlerBuscaNos continua exatamente o mesmo que voc� j� tem a�!
' Ele vai mergulhar nos grupos e PowerClips da sele��o normalmente.

Private Sub CrawlerBuscaNos(s As Shape, limite As Long, ByRef sacola As ShapeRange)
    Dim subS As Shape
    On Error Resume Next
    If s.Type = cdrCurveShape Then
        If s.Curve.Nodes.Count > limite Then sacola.Add s
    End If
    On Error GoTo 0
    
    If s.Type = cdrGroupShape Then
        For Each subS In s.shapes: CrawlerBuscaNos subS, limite, sacola: Next subS
    End If
    If Not s.PowerClip Is Nothing Then
        For Each subS In s.PowerClip.shapes: CrawlerBuscaNos subS, limite, sacola: Next subS
    End If
End Sub

' ============================================================
' FERRAMENTA 3: RADAR DE ESPESSURA M�NIMA (Estilo ArtPro)
' ============================================================
Public Sub InspecionarEspessuraMinima()
    ' Padroniza a unidade matem�tica do documento para Mil�metros temporariamente
    Dim unidOriginal As cdrUnit
    unidOriginal = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter
    
    Dim limiteMinimo As Double: limiteMinimo = 0.1 ' 0.1 mm (Padr�o FIRST)
    Dim srFinos As ShapeRange: Set srFinos = CreateShapeRange
    Dim s As Shape
    
    ' Inicia a varredura
    For Each s In ActivePage.shapes
        CrawlerEspessura s, limiteMinimo, srFinos
    Next s
    
    ' Devolve a unidade original para o Corel do usu�rio
    ActiveDocument.Unit = unidOriginal
    
    ' Resultado
    If srFinos.Count > 0 Then
        srFinos.CreateSelection
        MsgBox "Alerta Cr�tico FIRST! " & srFinos.Count & " objetos ou contornos possuem espessura f�sica menor que 0,1 mm." & vbCrLf & vbCrLf & _
               "Eles foram selecionados para que voc� possa engross�-los, sob risco de quebra na chapa.", vbCritical, "Console Flexo"
    Else
        MsgBox "Aprovado! Nenhuma linha ou objeto fino o suficiente para quebrar na chapa foi encontrado.", vbInformation, "Console Flexo"
    End If
End Sub

Private Sub CrawlerEspessura(s As Shape, limit As Double, ByRef sacola As ShapeRange)
    Dim subS As Shape
    On Error Resume Next
    
    If s.Type <> cdrGroupShape And s.Type <> cdrGuidelineShape Then
        Dim W As Double: W = s.SizeWidth
        Dim H As Double: H = s.SizeHeight
        Dim outW As Double: outW = 0
        Dim sinalizar As Boolean: sinalizar = False
        
        ' 1. AVALIA��O DE CONTORNO VIVO
        If s.Outline.Type = cdrOutline Then
            outW = s.Outline.Width
            ' Se tem contorno e � menor que o limite, j� falha na hora!
            If outW > 0 And outW < limit Then sinalizar = True
        End If
        
        ' 2. AVALIA��O DE OBJETO CONVERTIDO (Ctrl+Shift+Q)
        ' Objeto convertido nao tem contorno ativo -- dimensao fisica = espessura original
        ' [T20] Usa GetBoundingBox que e mais confiavel que SizeWidth para objetos convertidos
        If s.Type <> cdrBitmapShape And s.Type <> cdrTextShape Then
            Dim semContorno As Boolean: semContorno = False
            If s.Outline.Type <> cdrOutline Then
                semContorno = True
            ElseIf Round(outW, 3) <= 0 Then
                semContorno = True
            End If
            If semContorno Then
                ' Usa GetBoundingBox para obter dimensoes reais em mm
                Dim bbX As Double, bbY As Double, bbW As Double, bbH As Double
                s.GetBoundingBox bbX, bbY, bbW, bbH
                If Round(bbW, 3) <= limit Or Round(bbH, 3) <= limit Then
                    sinalizar = True
                End If
            End If
        End If
        
        If sinalizar Then sacola.Add s
    End If
    On Error GoTo 0
    
    ' Mergulhos
    If s.Type = cdrGroupShape Then
        For Each subS In s.shapes: CrawlerEspessura subS, limit, sacola: Next subS
    End If
    If Not s.PowerClip Is Nothing Then
        For Each subS In s.PowerClip.shapes: CrawlerEspessura subS, limit, sacola: Next subS
    End If
End Sub
' ============================================================
' FERRAMENTA: PADRONIZADOR DE CONTORNOS (0,2mm + Sele��o Final)
' ============================================================
Public Sub PadronizarContornosFinos(Optional silencioso As Boolean = False)
    Dim srProblemas As ShapeRange: Set srProblemas = CreateShapeRange
    Dim srCorrigidos As ShapeRange: Set srCorrigidos = CreateShapeRange
    Dim s As Shape

    Dim unidadeOriginal As cdrUnit: unidadeOriginal = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter

    For Each s In ActivePage.shapes
        CrawlerBuscaContornos s, srProblemas
    Next s

    If srProblemas.Count = 0 Then
        ActiveDocument.Unit = unidadeOriginal
        If Not silencioso Then MsgBox "Nenhum contorno abaixo de 0,1mm encontrado.", vbInformation, "Console Flexo"
        Exit Sub
    End If

    ActiveDocument.BeginCommandGroup "Console Flexo - Corrigir Contornos"
    Application.Optimization = True
    On Error GoTo FimErro

    Dim obj As Shape
    For Each obj In srProblemas
        obj.Outline.Width = 0.2
        srCorrigidos.Add obj
    Next obj

    ActiveDocument.EndCommandGroup
    Application.Optimization = False
    Application.Refresh
    ActiveDocument.Unit = unidadeOriginal

    If srCorrigidos.Count > 0 And Not silencioso Then
        srCorrigidos.CreateSelection
        MsgBox srCorrigidos.Count & " contorno(s) padronizados para 0,2mm.", vbInformation, "Console Flexo"
    End If
    Exit Sub

FimErro:
    ActiveDocument.EndCommandGroup
    Application.Optimization = False
    ActiveDocument.Unit = unidadeOriginal
    Application.Refresh
    If Not silencioso Then MsgBox "Erro: " & Err.Description, vbCritical, "Console Flexo"
End Sub

' ------------------------------------------------------------
' CRAWLER: CA�ADOR DE CONTORNOS (Ignora preenchimentos e tamanhos)
' ------------------------------------------------------------
Private Sub CrawlerBuscaContornos(s As Shape, ByRef sacola As ShapeRange)
    Dim subS As Shape
    On Error Resume Next
    
    ' Ignora Bitmaps e Grupos na an�lise individual
    If s.Type <> cdrBitmapShape And s.Type <> cdrGroupShape Then
        ' Verifica se o objeto TEM um contorno aplicado
        If s.Outline.Type = cdrOutline Then
            ' Se a espessura for maior que zero e menor ou igual a 0.101mm (Margem de erro do VBA)
            If s.Outline.Width > 0 And s.Outline.Width <= 0.101 Then
                sacola.Add s
            End If
        End If
    End If
    
    ' Mergulha em Grupos
    If s.Type = cdrGroupShape Then
        For Each subS In s.shapes: CrawlerBuscaContornos subS, sacola: Next subS
    End If
    
    ' Mergulha em PowerClips
    If Not s.PowerClip Is Nothing Then
        For Each subS In s.PowerClip.shapes: CrawlerBuscaContornos subS, sacola: Next subS
    End If
    On Error GoTo 0
End Sub

' ============================================================
' FERRAMENTA: DESBLOQUEAR OBJETOS DA PAGINA ATIVA
' ============================================================
Public Sub DesbloquearObjetos()
    Dim s As Shape
    Dim desbloqueados As Integer: desbloqueados = 0

    On Error Resume Next
    For Each s In ActivePage.shapes
        CrawlerDesbloquear s, desbloqueados
    Next s
    On Error GoTo 0

    If desbloqueados > 0 Then
        MsgBox desbloqueados & " objeto(s) desbloqueado(s) com sucesso!", vbInformation, "Console Flexo"
    Else
        MsgBox "Nenhum objeto bloqueado encontrado na p" & ChrW(225) & "gina ativa.", vbInformation, "Console Flexo"
    End If
End Sub

Private Sub CrawlerDesbloquear(s As Shape, ByRef contador As Integer)
    Dim subS As Shape
    On Error Resume Next
    If s.Locked Then
        s.Locked = False
        contador = contador + 1
    End If
    If s.Type = cdrGroupShape Then
        For Each subS In s.shapes: CrawlerDesbloquear subS, contador: Next subS
    End If
    If Not s.PowerClip Is Nothing Then
        For Each subS In s.PowerClip.shapes: CrawlerDesbloquear subS, contador: Next subS
    End If
    On Error GoTo 0
End Sub
