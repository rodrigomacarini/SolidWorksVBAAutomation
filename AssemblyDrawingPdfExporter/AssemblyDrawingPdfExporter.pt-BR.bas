Option Explicit

'===============================================================================
' SOLIDWORKS 2025
' Marcador: 20261105-04-PDF-MONTAGEM-SAVEAS-DIRETO
'
' Exporta para PDF:
'   - o desenho da montagem ativa, se existir;
'   - os desenhos das pecas e submontagens fisicas da montagem.
'
' Cada desenho precisa estar na mesma pasta do respectivo modelo e ter o mesmo
' nome-base. Componentes repetidos geram somente um PDF. Os PDFs sao gravados
' junto aos desenhos. A macro nao forca rebuild para manter o fluxo rapido.
' Cada desenho e ativado antes do SaveAs3 e fechado depois da exportacao,
' descartando alteracoes em memoria sem salvar arquivos SOLIDWORKS.
'===============================================================================

Private Const MACRO_VERSION As String = "20261105-04-PDF-MONTAGEM-SAVEAS-DIRETO"
Private Const SW_DOC_ASSEMBLY As Long = 2
Private Const SW_DOC_DRAWING As Long = 3
Private Const SW_OPEN_DOC_SILENT As Long = 1
Private Const SW_SAVE_AS_CURRENT_VERSION As Long = 0
Private Const SW_SAVE_AS_OPTIONS_SILENT As Long = 1
Private Const SW_DONT_REBUILD_ACTIVE_DOC As Long = 1

Private swApp As SldWorks.SldWorks
Private fso As Object
Private caminhosModelo As Object
Private desenhosParaExportar As Object

Private modelosVerificados As Long
Private desenhosEncontrados As Long
Private desenhosAusentes As Long
Private componentesSuprimidos As Long
Private componentesVirtuais As Long
Private componentesSemCaminho As Long
Private desenhosDuplicados As Long
Private pdfsExistentes As Long
Private pdfsCriados As Long
Private falhas As Long

Sub main()
    Dim montagem As SldWorks.ModelDoc2
    Dim montagemAssembly As SldWorks.AssemblyDoc
    Dim componentes As Variant
    Dim limiteInferior As Long
    Dim limiteSuperior As Long
    Dim i As Long
    Dim componente As SldWorks.Component2
    Dim caminhoComponente As String
    Dim tipoComponente As Long
    Dim componenteVirtual As Boolean
    Dim falhasAntesSupressao As Long
    Dim resposta As VbMsgBoxResult
    Dim caminhoMontagem As String
    Dim errosAtivacao As Long

    On Error GoTo ErroFatal

    ReiniciarContadores
    Debug.Print "[VERSION] " & MACRO_VERSION
    Debug.Print "[INICIO] Exportacao PDF da montagem ativa"

    Set swApp = Application.SldWorks
    If swApp Is Nothing Then
        MsgBox "O SOLIDWORKS nao esta disponivel.", vbExclamation, "PDF da montagem"
        Exit Sub
    End If

    Set montagem = swApp.ActiveDoc
    If montagem Is Nothing Then
        MsgBox "Abra a montagem que deseja processar antes de executar a macro.", vbExclamation, "PDF da montagem"
        Exit Sub
    End If
    If montagem.GetType <> SW_DOC_ASSEMBLY Then
        MsgBox "O documento ativo precisa ser uma montagem (.SLDASM).", vbExclamation, "PDF da montagem"
        Exit Sub
    End If

    caminhoMontagem = montagem.GetPathName
    If Len(caminhoMontagem) = 0 Then
        MsgBox "Salve a montagem antes de executar a macro.", vbExclamation, "PDF da montagem"
        Exit Sub
    End If

    Set montagemAssembly = montagem
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set caminhosModelo = CreateObject("Scripting.Dictionary")
    Set desenhosParaExportar = CreateObject("Scripting.Dictionary")
    caminhosModelo.CompareMode = vbTextCompare
    desenhosParaExportar.CompareMode = vbTextCompare
    Debug.Print "[MONTAGEM] " & caminhoMontagem
    ' A montagem principal entra na lista, assim como todos os modelos fisicos.
    ConsiderarModelo caminhoMontagem, True

    componentes = montagemAssembly.GetComponents(False)
    If IsArray(componentes) Then
        On Error Resume Next
        Err.Clear
        limiteInferior = LBound(componentes)
        limiteSuperior = UBound(componentes)
        If Err.Number <> 0 Then
            Debug.Print "[AVISO] Array de componentes vazio ou sem limites validos."
            Err.Clear
            On Error GoTo ErroFatal
            GoTo ConfirmarLista
        End If
        On Error GoTo ErroFatal

        For i = limiteInferior To limiteSuperior
            Set componente = Nothing
            On Error Resume Next
            Err.Clear
            Set componente = componentes(i)
            If Err.Number <> 0 Or componente Is Nothing Then
                Debug.Print "[ERRO COMPONENTE] Indice=" & i & " | Erro=" & Err.Number & " | " & Err.Description
                falhas = falhas + 1
                Err.Clear
                On Error GoTo ErroFatal
                GoTo ProximoComponente
            End If
            On Error GoTo ErroFatal

            falhasAntesSupressao = falhas
            If ComponenteOuAncestralSuprimido(componente) Then
                If falhas > falhasAntesSupressao Then
                    Debug.Print "[IGNORADO CONSULTA SUPRESSAO] " & componente.Name2
                Else
                    componentesSuprimidos = componentesSuprimidos + 1
                    Debug.Print "[IGNORADO SUPRIMIDO] " & componente.Name2
                End If
                GoTo ProximoComponente
            End If

            On Error Resume Next
            Err.Clear
            componenteVirtual = componente.IsVirtual
            If Err.Number <> 0 Then
                Debug.Print "[ERRO CONSULTA VIRTUAL] " & componente.Name2 & _
                            " | Erro=" & Err.Number & " | " & Err.Description
                falhas = falhas + 1
                Err.Clear
                On Error GoTo ErroFatal
                GoTo ProximoComponente
            End If
            On Error GoTo ErroFatal

            If componenteVirtual Then
                componentesVirtuais = componentesVirtuais + 1
                Debug.Print "[IGNORADO VIRTUAL] " & componente.Name2
                GoTo ProximoComponente
            End If

            On Error Resume Next
            Err.Clear
            caminhoComponente = componente.GetPathName
            If Err.Number <> 0 Then
                Debug.Print "[ERRO CONSULTA CAMINHO] " & componente.Name2 & _
                            " | Erro=" & Err.Number & " | " & Err.Description
                falhas = falhas + 1
                Err.Clear
                On Error GoTo ErroFatal
                GoTo ProximoComponente
            End If
            On Error GoTo ErroFatal

            If Len(caminhoComponente) = 0 Then
                componentesSemCaminho = componentesSemCaminho + 1
                Debug.Print "[IGNORADO SEM CAMINHO] " & componente.Name2
                GoTo ProximoComponente
            End If

            tipoComponente = TipoPeloCaminho(caminhoComponente)
            If tipoComponente = 0 Then
                Debug.Print "[IGNORADO EXTENSAO] " & caminhoComponente
                GoTo ProximoComponente
            End If

            ConsiderarModelo caminhoComponente, False

ProximoComponente:
        Next i
    Else
        Debug.Print "[AVISO] GetComponents(False) nao retornou array."
    End If

ConfirmarLista:
    If desenhosParaExportar.Count = 0 Then
        MsgBox "Nenhum desenho correspondente foi encontrado." & vbCrLf & vbCrLf & _
               "A macro procurou o desenho da montagem principal e de cada peca/submontagem fisica.", _
               vbInformation, "PDF da montagem"
        Debug.Print "[RESUMO] Nenhum desenho para exportar."
        Exit Sub
    End If

    pdfsExistentes = ContarPDFsExistentes
    resposta = MsgBox( _
        "Desenhos encontrados: " & desenhosParaExportar.Count & vbCrLf & _
        "Desenhos correspondentes ausentes: " & desenhosAusentes & vbCrLf & _
        "PDFs que ja existem e serao substituidos: " & pdfsExistentes & vbCrLf & vbCrLf & _
        "Os PDFs serao gravados ao lado dos desenhos." & vbCrLf & _
        "A macro nao fara rebuild forcado; usara o estado atual/salvo de cada desenho." & vbCrLf & _
        "Deseja continuar?", _
        vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar exportacao PDF")
    If resposta <> vbYes Then
        Debug.Print "[CANCELADO] Confirmacao recusada"
        Exit Sub
    End If

    ExportarDesenhos

    ' Restaurar a montagem principal como documento ativo, caso a abertura de
    ' algum desenho tenha alterado o foco da janela.
    On Error Resume Next
    swApp.ActivateDoc3 montagem.GetTitle, False, SW_DONT_REBUILD_ACTIVE_DOC, errosAtivacao
    Debug.Print "[RESTAURACAO] Montagem ativa | ErroActivateDoc3=" & errosAtivacao
    On Error GoTo 0

    Debug.Print "[RESUMO] Desenhos encontrados=" & desenhosEncontrados & _
                " | PDFs criados=" & pdfsCriados & _
                " | Ausentes=" & desenhosAusentes & _
                " | Modelos unicos verificados=" & modelosVerificados & _
                " | Suprimidos=" & componentesSuprimidos & _
                " | Virtuais=" & componentesVirtuais & _
                " | Duplicados=" & desenhosDuplicados & _
                " | Falhas=" & falhas

    MsgBox "Exportacao concluida." & vbCrLf & vbCrLf & _
        "Desenhos encontrados: " & desenhosEncontrados & vbCrLf & _
        "PDFs criados: " & pdfsCriados & vbCrLf & _
        "Desenhos sem arquivo correspondente: " & desenhosAusentes & vbCrLf & _
        "Modelos unicos verificados: " & modelosVerificados & vbCrLf & _
        "Componentes suprimidos ignorados: " & componentesSuprimidos & vbCrLf & _
        "Componentes virtuais ignorados: " & componentesVirtuais & vbCrLf & _
        "Componentes sem caminho: " & componentesSemCaminho & vbCrLf & _
        "Desenhos duplicados consolidados: " & desenhosDuplicados & vbCrLf & _
        "Falhas: " & falhas & vbCrLf & _
        "Veja o Immediate Window para caminhos, retornos e erros.", _
        IIf(falhas > 0, vbExclamation, vbInformation), "Resultado da exportacao PDF"
    Exit Sub

ErroFatal:
    Debug.Print "[ERRO FATAL] Numero=" & Err.Number & " | Descricao=" & Err.Description & " | Origem=" & Err.Source
    MsgBox "Ocorreu um erro inesperado." & vbCrLf & _
           "Erro " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Consulte o Immediate Window do editor VBA.", vbCritical, "PDF da montagem"
End Sub

Private Sub ReiniciarContadores()
    modelosVerificados = 0
    desenhosEncontrados = 0
    desenhosAusentes = 0
    componentesSuprimidos = 0
    componentesVirtuais = 0
    componentesSemCaminho = 0
    desenhosDuplicados = 0
    pdfsExistentes = 0
    pdfsCriados = 0
    falhas = 0
End Sub

Private Function ComponenteOuAncestralSuprimido(ByVal componente As SldWorks.Component2) As Boolean
    Dim atual As SldWorks.Component2
    Dim estado As Long

    ComponenteOuAncestralSuprimido = True
    Set atual = componente
    On Error GoTo TratarErro

    Do While Not atual Is Nothing
        estado = atual.GetSuppression2
        If estado = swComponentSuppressed Then Exit Function
        Set atual = atual.GetParent
    Loop

    ComponenteOuAncestralSuprimido = False
    Exit Function

TratarErro:
    Debug.Print "[ERRO CONSULTA SUPRESSAO] Erro=" & Err.Number & " | " & Err.Description
    falhas = falhas + 1
End Function

Private Sub ConsiderarModelo(ByVal caminhoModelo As String, ByVal eMontagemPrincipal As Boolean)
    Dim chaveModelo As String
    Dim extensao As String
    Dim nomeBase As String
    Dim caminhoDesenho As String

    On Error GoTo TratarErro
    If Len(caminhoModelo) = 0 Then Exit Sub

    chaveModelo = LCase$(caminhoModelo)
    If caminhosModelo.Exists(chaveModelo) Then Exit Sub
    caminhosModelo.Add chaveModelo, True
    modelosVerificados = modelosVerificados + 1

    extensao = UCase$(fso.GetExtensionName(caminhoModelo))
    If extensao <> "SLDPRT" And extensao <> "SLDASM" Then
        Debug.Print "[IGNORADO EXTENSAO MODELO] " & caminhoModelo
        Exit Sub
    End If

    nomeBase = fso.GetBaseName(caminhoModelo)
    caminhoDesenho = fso.BuildPath(fso.GetParentFolderName(caminhoModelo), nomeBase & ".SLDDRW")
    If Not fso.FileExists(caminhoDesenho) Then
        desenhosAusentes = desenhosAusentes + 1
        Debug.Print "[DESENHO AUSENTE] Modelo=" & caminhoModelo & " | Procurado=" & caminhoDesenho
        Exit Sub
    End If

    If desenhosParaExportar.Exists(LCase$(caminhoDesenho)) Then
        desenhosDuplicados = desenhosDuplicados + 1
        Debug.Print "[DUPLICADO CONSOLIDADO] " & caminhoDesenho
        Exit Sub
    End If

    desenhosParaExportar.Add LCase$(caminhoDesenho), caminhoDesenho
    desenhosEncontrados = desenhosParaExportar.Count
    Debug.Print "[DESENHO ENCONTRADO] " & caminhoDesenho & _
                IIf(eMontagemPrincipal, " | Origem=Montagem principal", " | Origem=Componente")
    Exit Sub

TratarErro:
    Debug.Print "[ERRO CATALOGAR MODELO] Caminho=" & caminhoModelo & _
                " | Erro=" & Err.Number & " | " & Err.Description
    falhas = falhas + 1
End Sub

Private Function ContarPDFsExistentes() As Long
    Dim chave As Variant
    Dim caminhoDesenho As String
    Dim caminhoPDF As String

    ContarPDFsExistentes = 0
    For Each chave In desenhosParaExportar.Keys
        caminhoDesenho = CStr(desenhosParaExportar(chave))
        caminhoPDF = ObterCaminhoPDF(caminhoDesenho)
        If fso.FileExists(caminhoPDF) Then ContarPDFsExistentes = ContarPDFsExistentes + 1
    Next chave
End Function

Private Sub ExportarDesenhos()
    Dim chave As Variant
    Dim caminhoDesenho As String
    Dim caminhoPDF As String

    For Each chave In desenhosParaExportar.Keys
        caminhoDesenho = CStr(desenhosParaExportar(chave))
        caminhoPDF = ObterCaminhoPDF(caminhoDesenho)
        Debug.Print "[PROCESSAR] Desenho=" & caminhoDesenho & " | PDF=" & caminhoPDF
        If ExportarUmDesenho(caminhoDesenho, caminhoPDF) Then
            pdfsCriados = pdfsCriados + 1
        Else
            falhas = falhas + 1
        End If
    Next chave
End Sub

Private Function ExportarUmDesenho(ByVal caminhoDesenho As String, ByVal caminhoPDF As String) As Boolean
    Dim documento As SldWorks.ModelDoc2
    Dim documentoAtivo As SldWorks.ModelDoc2
    Dim documentoVerificacaoFechamento As SldWorks.ModelDoc2
    Dim extensaoModelo As SldWorks.ModelDocExtension
    Dim errosAbertura As Long
    Dim avisosAbertura As Long
    Dim errosAtivacaoDrawing As Long
    Dim errosSaveAs As Long
    Dim avisosSaveAs As Long
    Dim retornoSaveAs As Boolean
    Dim tamanhoPDF As Double

    ExportarUmDesenho = False
    On Error GoTo TratarErro

    Set documento = swApp.OpenDoc6( _
        caminhoDesenho, SW_DOC_DRAWING, SW_OPEN_DOC_SILENT, "", errosAbertura, avisosAbertura)
    Debug.Print "[OPEN DOC6] Documento=" & caminhoDesenho & _
                " | ObjetoNothing=" & CStr(documento Is Nothing) & _
                " | Erros=" & errosAbertura & " | Avisos=" & avisosAbertura
    If documento Is Nothing Then Exit Function

    If documento.GetType <> SW_DOC_DRAWING Then
        Debug.Print "[FALHA] O arquivo carregado nao e um desenho: " & caminhoDesenho
        GoTo Finalizar
    End If

    ' SaveAs3 exporta o documento ativo. Ativar sem rebuild preserva o
    ' comportamento rapido e usa o estado carregado/salvo do desenho.
    Set documentoAtivo = swApp.ActivateDoc3( _
        documento.GetTitle, False, SW_DONT_REBUILD_ACTIVE_DOC, errosAtivacaoDrawing)
    Debug.Print "[ACTIVATE DOC3] Titulo=" & documento.GetTitle & _
                " | ObjetoNothing=" & CStr(documentoAtivo Is Nothing) & _
                " | Erros/avisos=" & errosAtivacaoDrawing
    If documentoAtivo Is Nothing Then GoTo Finalizar
    If StrComp(documentoAtivo.GetPathName, caminhoDesenho, vbTextCompare) <> 0 Then
        Debug.Print "[FALHA] Documento ativo nao corresponde ao caminho esperado. Ativo=" & _
                    documentoAtivo.GetPathName & " | Esperado=" & caminhoDesenho
        GoTo Finalizar
    End If
    Set documento = documentoAtivo
    Debug.Print "[REBUILD] Nao executado; exportando o estado carregado do drawing."

    Set extensaoModelo = documento.Extension
    If extensaoModelo Is Nothing Then
        Debug.Print "[FALHA] IModelDocExtension indisponivel: " & caminhoDesenho
        GoTo Finalizar
    End If

    Debug.Print "[CHAMADA PDF] SaveAs3 direto | ExportData=Nothing (todas as folhas) | " & caminhoPDF
    retornoSaveAs = extensaoModelo.SaveAs3( _
        caminhoPDF, _
        SW_SAVE_AS_CURRENT_VERSION, _
        SW_SAVE_AS_OPTIONS_SILENT, _
        Nothing, _
        Nothing, _
        errosSaveAs, _
        avisosSaveAs)

    Debug.Print "[API] IModelDocExtension.SaveAs3 retorno=" & CStr(retornoSaveAs) & _
                " | Erros=" & errosSaveAs & " | Avisos=" & avisosSaveAs

    If retornoSaveAs Then
        If fso.FileExists(caminhoPDF) Then
            tamanhoPDF = CDbl(fso.GetFile(caminhoPDF).Size)
            If tamanhoPDF > 0 Then
                ExportarUmDesenho = True
                Debug.Print "[PDF OK] Caminho=" & caminhoPDF & " | Bytes=" & tamanhoPDF
            Else
                Debug.Print "[FALHA] PDF vazio: " & caminhoPDF
            End If
        Else
            Debug.Print "[FALHA] SaveAs3 retornou sucesso, mas o arquivo nao existe: " & caminhoPDF
        End If
    End If

Finalizar:
    If Not documento Is Nothing Then
        On Error Resume Next
        Err.Clear
        swApp.QuitDoc documento.GetTitle
        If Err.Number <> 0 Then
            Debug.Print "[ERRO AO FECHAR] Caminho=" & caminhoDesenho & _
                        " | Erro=" & Err.Number & " | " & Err.Description
            falhas = falhas + 1
            Err.Clear
        Else
            Set documentoVerificacaoFechamento = swApp.GetOpenDocumentByName(caminhoDesenho)
            If documentoVerificacaoFechamento Is Nothing Then
                Debug.Print "[FECHADO SEM SALVAR] Drawing fechado apos a exportacao: " & caminhoDesenho
            Else
                Debug.Print "[ERRO AO FECHAR] Documento ainda consta aberto: " & caminhoDesenho
                falhas = falhas + 1
            End If
        End If
        On Error GoTo 0
    End If

    Set extensaoModelo = Nothing
    Set documentoVerificacaoFechamento = Nothing
    Set documentoAtivo = Nothing
    Set documento = Nothing
    Exit Function

TratarErro:
    Debug.Print "[ERRO PDF] Caminho=" & caminhoDesenho & _
                " | Erro=" & Err.Number & " | " & Err.Description & " | Origem=" & Err.Source
    Resume Finalizar
End Function

Private Function TipoPeloCaminho(ByVal caminho As String) As Long
    Dim extensao As String

    extensao = UCase$(fso.GetExtensionName(caminho))
    Select Case extensao
        Case "SLDPRT": TipoPeloCaminho = 1
        Case "SLDASM": TipoPeloCaminho = SW_DOC_ASSEMBLY
        Case Else: TipoPeloCaminho = 0
    End Select
End Function

Private Function ObterCaminhoPDF(ByVal caminhoDesenho As String) As String
    ObterCaminhoPDF = fso.BuildPath( _
        fso.GetParentFolderName(caminhoDesenho), _
        fso.GetBaseName(caminhoDesenho) & ".pdf")
End Function
