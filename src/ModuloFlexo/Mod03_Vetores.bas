Attribute VB_Name = "Mod03_Vetores"
' ============================================================
' MÓDULO: Mod03_Vetores (TRATAMENTO ESTRUTURAL E FIRST)
' DESCRIÇÃO: Conversão de Fontes, Otimização de Nós e Espessura
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

    If Not silencioso Then ActiveDocument.BeginCommandGroup "Console Flexo - Textos em Curvas"
    On Error GoTo FimErro

    Dim i As Integer
    Dim convertidos As Integer: convertidos = 0
    For i = srTextos.Count To 1 Step -1
        srTextos(i).ConvertToCurves
        convertidos = convertidos + 1
    Next i

    If Not silencioso Then ActiveDocument.EndCommandGroup
    If Not silencioso Then MsgBox convertidos & " textos convertidos em curvas!", vbInformation, "Console Flexo"
    Exit Sub

FimErro:
    If Not silencioso Then ActiveDocument.EndCommandGroup
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
' FERRAMENTA 2: INSPETOR DE NÓS (Atua apenas na Seleção)
' ============================================================
Public Sub InspecionarNos()
    ' Trava: Verifica se o operador selecionou algo para inspecionar
    If ActiveSelection.shapes.Count = 0 Then
        MsgBox "Selecione os objetos ou o grupo que você deseja inspecionar primeiro!", vbExclamation, "Console Flexo"
        Exit Sub
    End If

    Dim limiteNos As String
    limiteNos = InputBox("Digite a quantidade máxima de nós permitida por objeto:" & vbCrLf & "(Acima de 1500 costuma ser lixo de Rastreio Automático)", "Inspetor de Nós", "1500")
    If limiteNos = "" Or Not IsNumeric(limiteNos) Then Exit Sub
    
    Dim maxNos As Long: maxNos = CLng(limiteNos)
    Dim srNos As ShapeRange: Set srNos = CreateShapeRange
    Dim s As Shape
    
    ' Muda a varredura: Agora olha SÓ para o que o operador selecionou!
    For Each s In ActiveSelection.shapes
        CrawlerBuscaNos s, maxNos, srNos
    Next s
    
    ' Limpa a seleção inicial do operador para não confundir
    ActiveDocument.ClearSelection
    
    If srNos.Count > 0 Then
        ' Seleciona APENAS os objetos defeituosos dentro do grupo que ele havia selecionado
        srNos.CreateSelection
        MsgBox "Atenção! " & srNos.Count & " objetos DENTRO DA SUA SELEÇÃO possuem mais de " & maxNos & " nós e foram isolados." & vbCrLf & vbCrLf & _
               "Analise se é possível utilizar o botão de Reduzir Nós sem deformar a arte.", vbExclamation, "Console Flexo"
    Else
        MsgBox "Seleção limpa! Nenhum objeto inspecionado possui excesso de nós.", vbInformation, "Console Flexo"
    End If
End Sub
' ============================================================
' FERRAMENTA 4: REDUTOR DE NÓS SEGURO (AutoReduce)
' ============================================================
Public Sub ReduzirNosSeguro()
    ' Trava de segurança
    If ActiveSelection.shapes.Count = 0 Then
        MsgBox "Selecione as curvas que deseja suavizar primeiro!", vbExclamation, "Console Flexo"
        Exit Sub
    End If
    
    Dim s As Shape
    Dim totalNosAntes As Long: totalNosAntes = 0
    Dim totalNosDepois As Long: totalNosDepois = 0
    Dim curvasAfetadas As Integer: curvasAfetadas = 0
    
    ' Fator de Suavização (0.005 é considerado seguro no Corel para não deformar logotipos)
    Dim fatorSuavizacao As Double: fatorSuavizacao = 0.005
    
    ' Inicia a varredura blindada APENAS na seleção
    For Each s In ActiveSelection.shapes
        CrawlerReduzirNos s, fatorSuavizacao, totalNosAntes, totalNosDepois, curvasAfetadas
    Next s
    
    Dim nosRemovidos As Long
    nosRemovidos = totalNosAntes - totalNosDepois
    
    If nosRemovidos > 0 Then
        MsgBox "Limpeza concluída! Foram removidos " & nosRemovidos & " nós inúteis de " & curvasAfetadas & " curva(s)." & vbCrLf & vbCrLf & _
               "DICA: Dê um zoom e verifique visualmente se a arte original não sofreu deformações.", vbInformation, "Console Flexo"
    Else
        MsgBox "Nenhum nó pôde ser removido. Os objetos selecionados já estão otimizados ou não são curvas.", vbInformation, "Console Flexo"
    End If
End Sub

' ------------------------------------------------------------
' O CRAWLER: Varredura Segura para Remoção de Nós
' ------------------------------------------------------------
Private Sub CrawlerReduzirNos(s As Shape, fator As Double, ByRef nosAntes As Long, ByRef nosDepois As Long, ByRef afetadas As Integer)
    Dim subS As Shape
    On Error Resume Next
    
    ' 1. Modifica o objeto se for Curva
    If s.Type = cdrCurveShape Then
        Dim antes As Long: antes = s.Curve.Nodes.Count
        
        ' Aplica a redução nativa do Corel
        s.Curve.AutoReduceNodes fator
        
        Dim depois As Long: depois = s.Curve.Nodes.Count
        
        ' Só contabiliza se realmente conseguiu remover algum nó
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

' O CrawlerBuscaNos continua exatamente o mesmo que você já tem aí!
' Ele vai mergulhar nos grupos e PowerClips da seleção normalmente.

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
' FERRAMENTA 3: RADAR DE ESPESSURA MÍNIMA (Estilo ArtPro)
' ============================================================
Public Sub InspecionarEspessuraMinima()
    ' Padroniza a unidade matemática do documento para Milímetros temporariamente
    Dim unidOriginal As cdrUnit
    unidOriginal = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter
    
    Dim limiteMinimo As Double: limiteMinimo = 0.1 ' 0.1 mm (Padrão FIRST)
    Dim srFinos As ShapeRange: Set srFinos = CreateShapeRange
    Dim s As Shape
    
    ' Inicia a varredura
    For Each s In ActivePage.shapes
        CrawlerEspessura s, limiteMinimo, srFinos
    Next s
    
    ' Devolve a unidade original para o Corel do usuário
    ActiveDocument.Unit = unidOriginal
    
    ' Resultado
    If srFinos.Count > 0 Then
        srFinos.CreateSelection
        MsgBox "Alerta Crítico FIRST! " & srFinos.Count & " objetos ou contornos possuem espessura física menor que 0,1 mm." & vbCrLf & vbCrLf & _
               "Eles foram selecionados para que você possa engrossá-los, sob risco de quebra na chapa.", vbCritical, "Console Flexo"
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
        
        ' 1. AVALIAÇÃO DE CONTORNO VIVO
        If s.Outline.Type = cdrOutline Then
            outW = s.Outline.Width
            ' Se tem contorno e é menor que o limite, já falha na hora!
            If outW > 0 And outW < limit Then sinalizar = True
        End If
        
        ' 2. AVALIAÇÃO DE OBJETO CONVERTIDO (Ctrl+Shift+Q)
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
' FERRAMENTA: PADRONIZADOR DE CONTORNOS (0,2mm + Seleção Final)
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

    If Not silencioso Then ActiveDocument.BeginCommandGroup "Console Flexo - Corrigir Contornos"
    Application.Optimization = True
    On Error GoTo FimErro

    Dim obj As Shape
    For Each obj In srProblemas
        obj.Outline.Width = 0.2
        srCorrigidos.Add obj
    Next obj

    If Not silencioso Then ActiveDocument.EndCommandGroup
    Application.Optimization = False
    If Not silencioso Then Application.Refresh
    ActiveDocument.Unit = unidadeOriginal

    If srCorrigidos.Count > 0 And Not silencioso Then
        srCorrigidos.CreateSelection
        MsgBox srCorrigidos.Count & " contorno(s) padronizados para 0,2mm.", vbInformation, "Console Flexo"
    End If
    Exit Sub

FimErro:
    If Not silencioso Then ActiveDocument.EndCommandGroup
    Application.Optimization = False
    ActiveDocument.Unit = unidadeOriginal
    If Not silencioso Then Application.Refresh
    If Not silencioso Then MsgBox "Erro: " & Err.Description, vbCritical, "Console Flexo"
End Sub

' ------------------------------------------------------------
' CRAWLER: CAÇADOR DE CONTORNOS (Ignora preenchimentos e tamanhos)
' ------------------------------------------------------------
Private Sub CrawlerBuscaContornos(s As Shape, ByRef sacola As ShapeRange)
    Dim subS As Shape
    On Error Resume Next
    
    ' Ignora Bitmaps e Grupos na análise individual
    If s.Type <> cdrBitmapShape And s.Type <> cdrGroupShape Then
        ' Verifica se o objeto TEM um contorno aplicado
        If s.Outline.Type = cdrOutline Then
            Dim espW As Double: espW = s.Outline.Width
            ' Se a espessura for maior que zero e menor ou igual a 0.101mm (Margem de erro do VBA)
            If espW > 0 And espW <= 0.101 Then
                ' Regra de negócio: contornos brancos são intencionais (não sinalizar)
                Dim ehIntencional As Boolean: ehIntencional = False
                If EhContornoBranco(s.Outline.Color) Then ehIntencional = True
                ' Contornos <= 0,05mm com preenchimento também são intencionais
                If espW <= 0.05 Then
                    If s.Fill.Type = cdrUniformFill Or _
                       s.Fill.Type = cdrFountainFill Or _
                       s.Fill.Type = cdrTextureFill Or _
                       s.Fill.Type = cdrPatternFill Then
                        ehIntencional = True
                    End If
                End If
                If Not ehIntencional Then sacola.Add s
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

' ------------------------------------------------------------
' Função auxiliar: identifica contorno de cor branca
' ------------------------------------------------------------
Private Function EhContornoBranco(C As Color) As Boolean
    EhContornoBranco = False
    On Error Resume Next
    If C.Type = cdrColorCMYK Then
        If (C.CMYKCyan + C.CMYKMagenta + C.CMYKYellow + C.CMYKBlack) = 0 Then EhContornoBranco = True
    ElseIf C.Type = cdrColorRGB Then
        If C.RGBRed = 255 And C.RGBGreen = 255 And C.RGBBlue = 255 Then EhContornoBranco = True
    ElseIf C.IsSpot Then
        If C.Tint = 0 Then EhContornoBranco = True
    End If
    On Error GoTo 0
End Function
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
