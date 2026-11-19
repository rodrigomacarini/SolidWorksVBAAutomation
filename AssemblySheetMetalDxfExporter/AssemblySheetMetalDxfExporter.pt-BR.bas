Option Explicit

' Prompt de depuracao: envie [VERSION], [RESUMO] e linhas relevantes de [ERRO DXF]/[RESULTADO DA PECA], mais a versao do SOLIDWORKS e um teste anonimizado.
' Nao inclua nomes de clientes, caminhos de rede ou dados proprietarios de desenhos.

' ============================================================
' MACRO SOLIDWORKS 2025
'
' EXPORTA A CUT LIST DE CADA PECA FISICA UNICA DA MONTAGEM:
'
' CHAPAS METALICAS:
' -> Um arquivo DXF por item da Cut List.
' -> Somente o FlatPattern associado ao Body2 correto.
' -> Geometria planificada e linhas de dobra.
'
' ITENS QUE NAO SAO CHAPA METALICA:
' -> Sao ignorados.
'
' ============================================================

Dim swApp           As SldWorks.SldWorks
Dim swModel         As SldWorks.ModelDoc2
Dim swPart          As SldWorks.PartDoc
Dim swFeat          As SldWorks.Feature
Dim swBodyFolder    As SldWorks.BodyFolder
Dim swBody          As SldWorks.Body2

Dim filePath        As String
Dim dxfOutputFolder As String
Dim fileName        As String
Dim drawingNumber   As String

Const MACRO_VERSION As String = "20261106-02-DXF-MONTAGEM-PASTA-PRINCIPAL"

Private Const SW_DOC_PART As Long = 1
Private Const SW_DOC_ASSEMBLY As Long = 2
Private Const SW_OPEN_DOC_SILENT As Long = 1
Private Const SW_DONT_REBUILD_ACTIVE_DOC As Long = 1

Private totalPartsVisited As Long
Private totalPartsActivated As Long
Private totalPartsClosed As Long
Private totalProcessed As Long
Private totalDxfSuccess As Long
Private totalDxfFailures As Long
Private totalIgnoredItems As Long
Private totalSkippedSuppressed As Long
Private totalSkippedVirtual As Long
Private totalDuplicateParts As Long
Private totalOpenFailures As Long


' ============================================================
' PROPRIEDADES DO DOCUMENTO
' ============================================================

Function GetFileProperty(propName As String) As String
    ' Le a propriedade primeiro na configuracao ativa e depois nas propriedades gerais. A busca alternativa pelo simbolo de grau e preservada.

    On Error Resume Next

    Dim mgr As SldWorks.CustomPropertyManager
    Dim configName As String
    Dim valOut As String
    Dim resolved As String

    GetFileProperty = ""

    configName = _
        swModel.ConfigurationManager.ActiveConfiguration.Name

    Set mgr = _
        swModel.Extension.CustomPropertyManager(configName)

    valOut = ""
    resolved = ""

    mgr.Get4 propName, False, valOut, resolved

    If Trim(resolved) = "" _
        And InStr(propName, "N" & ChrW(176)) > 0 Then

        valOut = ""
        resolved = ""

        mgr.Get4 _
            Replace(propName, "N" & ChrW(176), "N" & ChrW(186)), _
            False, _
            valOut, _
            resolved

    End If

    If Trim(resolved) = "" Then

        Set mgr = _
            swModel.Extension.CustomPropertyManager("")

        valOut = ""
        resolved = ""

        mgr.Get4 propName, False, valOut, resolved

        If Trim(resolved) = "" _
            And InStr(propName, "N" & ChrW(176)) > 0 Then

            valOut = ""
            resolved = ""

            mgr.Get4 _
                Replace(propName, "N" & ChrW(176), "N" & ChrW(186)), _
                False, _
                valOut, _
                resolved

        End If

    End If

    GetFileProperty = Trim(resolved)

    On Error GoTo 0

End Function


' ============================================================
' PROPRIEDADES DO ITEM DA CUT LIST
' ============================================================

Function GetProp( _
    propMgr As SldWorks.CustomPropertyManager, _
    propName As String) As String
    ' Le o valor resolvido de uma propriedade personalizada da pasta atual da Cut List.

    On Error Resume Next

    Dim valOut As String
    Dim resolved As String

    GetProp = ""

    valOut = ""
    resolved = ""

    propMgr.Get4 _
        propName, _
        False, _
        valOut, _
        resolved

    GetProp = Trim(resolved)

    On Error GoTo 0

End Function


' ============================================================
' LIMPEZA DO NOME DO ARQUIVO DE SAIDA
' ============================================================

Function CleanFileName(txt As String) As String
    ' Substitui caracteres invalidos para nome de arquivo no Windows e reduz espacos repetidos antes de montar o caminho.

    Dim chars As Variant

    chars = Array( _
        "\", _
        "/", _
        ":", _
        "*", _
        "?", _
        """", _
        "<", _
        ">", _
        "|" _
    )

    Dim i As Long

    For i = LBound(chars) To UBound(chars)

        txt = Replace(txt, chars(i), "_")

    Next i

    Do While InStr(txt, "  ") > 0

        txt = Replace(txt, "  ", " ")

    Loop

    CleanFileName = Trim(txt)

End Function


' ============================================================
' IDENTIFICA CHAPA METALICA
' ============================================================

Function CheckIfSheetMetal( _
    swFeat As SldWorks.Feature) As Boolean
    ' Classifica o item pelos nomes de propriedades da Cut List. Os termos de busca sao criterios funcionais e permanecem inalterados.

    On Error Resume Next

    Dim swCustPropMgr As _
        SldWorks.CustomPropertyManager

    Dim propNames As Variant

    CheckIfSheetMetal = False

    Set swCustPropMgr = _
        swFeat.CustomPropertyManager

    propNames = swCustPropMgr.GetNames

    If Not IsEmpty(propNames) Then

        Dim j As Long
        Dim pName As String

        For j = 0 To UBound(propNames)

            pName = _
                LCase(CStr(propNames(j)))

            If InStr( _
                    1, _
                    pName, _
                    "sheet metal", _
                    vbTextCompare) > 0 _
                Or InStr( _
                    1, _
                    pName, _
                    "chapa", _
                    vbTextCompare) > 0 _
                Or InStr( _
                    1, _
                    pName, _
                    "espessura", _
                    vbTextCompare) > 0 _
                Or InStr( _
                    1, _
                    pName, _
                    "thickness", _
                    vbTextCompare) > 0 Then

                CheckIfSheetMetal = True
                Exit Function

            End If

        Next j

    End If

    On Error GoTo 0

End Function


' ============================================================
' COMPARA DOIS OBJETOS BODY2
' ============================================================

Function SameBody( _
    bodyA As SldWorks.Body2, _
    bodyB As SldWorks.Body2) As Boolean
    ' Compara identificadores de selecao para associar o FlatPattern ao corpo esperado.

    On Error GoTo fail

    SameBody = False

    If bodyA Is Nothing Then Exit Function
    If bodyB Is Nothing Then Exit Function

    Dim idA As String
    Dim idB As String

    idA = bodyA.GetSelectionId
    idB = bodyB.GetSelectionId

    If Trim(idA) = "" Then Exit Function
    If Trim(idB) = "" Then Exit Function

    SameBody = (idA = idB)

    Exit Function

fail:

    SameBody = False

End Function


' ============================================================
' OBTEM O BODY ASSOCIADO AO FLAT-PATTERN
' ============================================================

Function GetBodyFromFlatPattern( _
    flatFeat As SldWorks.Feature) As SldWorks.Body2
    ' Segue FlatPattern -> GetDefinition -> FixedFace2 -> GetBody. O fallback para aresta e mantido.

    On Error GoTo fail

    Set GetBodyFromFlatPattern = Nothing

    If flatFeat Is Nothing Then Exit Function

    Dim flatData As _
        SldWorks.FlatPatternFeatureData

    Set flatData = _
        flatFeat.GetDefinition

    If flatData Is Nothing Then Exit Function

    Dim fixedObj As Object

    Set fixedObj = flatData.FixedFace2

    If fixedObj Is Nothing Then Exit Function

    Dim swFace As SldWorks.Face2
    Dim swEdge As SldWorks.Edge

    On Error Resume Next

    Set swFace = fixedObj

    On Error GoTo fail

    If Not swFace Is Nothing Then

        Set GetBodyFromFlatPattern = _
            swFace.GetBody

        If Not GetBodyFromFlatPattern Is Nothing Then
            Exit Function
        End If

    End If

    On Error Resume Next

    Set swEdge = fixedObj

    On Error GoTo fail

    If Not swEdge Is Nothing Then

        Set GetBodyFromFlatPattern = _
            swEdge.GetBody

        If Not GetBodyFromFlatPattern Is Nothing Then
            Exit Function
        End If

    End If

    Exit Function

fail:

    Set GetBodyFromFlatPattern = Nothing

End Function


' ============================================================
' LOCALIZA O FLAT-PATTERN DO BODY
' ============================================================

Function FindFlatPatternForBody( _
    targetBody As SldWorks.Body2) As SldWorks.Feature
    ' Percorre a arvore de recursos e retorna o FlatPattern associado ao corpo solicitado.

    On Error GoTo fail

    Set FindFlatPatternForBody = Nothing

    If targetBody Is Nothing Then Exit Function

    Dim feat As SldWorks.Feature
    Dim flatBody As SldWorks.Body2

    Set feat = swModel.FirstFeature

    Do While Not feat Is Nothing

        If feat.GetTypeName2 = "FlatPattern" Then

            Set flatBody = _
                GetBodyFromFlatPattern(feat)

            If Not flatBody Is Nothing Then

                Debug.Print _
                    "      FlatPattern encontrado: " & _
                    feat.Name

                Debug.Print _
                    "      Corpo associado: " & _
                    flatBody.Name

                If SameBody( _
                        targetBody, _
                        flatBody) Then

                    Set FindFlatPatternForBody = feat
                    Exit Function

                End If

            End If

        End If

        Set feat = feat.GetNextFeature

    Loop

    Exit Function

fail:

    Debug.Print _
        "      [ERRO] FindFlatPatternForBody: " & _
        Err.Number & " - " & Err.Description

    Set FindFlatPatternForBody = Nothing

End Function


' ============================================================
' NOMES DE PROPRIEDADES E SUFIXO LEGADO
' ============================================================

Function PcsSuffix() As String
    ' Retorna o sufixo legado em portugues usado nos nomes DXF; altera-lo pode afetar a programacao posterior do laser.

    PcsSuffix = "p" & ChrW(231) & "s"

End Function


Function PropDescricao() As String

    PropDescricao = _
        "Descri" & ChrW(231) & ChrW(227) & "o"

End Function


Function PropEspessuraChapa() As String

    PropEspessuraChapa = _
        "Espessura da Chapa met" & ChrW(225) & "lica"

End Function


Function PropDimensoes() As String

    PropDimensoes = _
        "Dimens" & ChrW(245) & "es"

End Function


Function PropNumeroDesenho() As String

    PropNumeroDesenho = _
        "N" & ChrW(250) & "mero do Desenho"

End Function


Function FormatThicknessForFileName( _
    thicknessValue As String) As String
    ' Preserva a convencao de separador decimal nos nomes DXF gerados.

    FormatThicknessForFileName = _
        Replace(Trim(thicknessValue), ".", ",")

End Function


' ============================================================
' DIAGNOSTICO DE BOOLEANOS
' ============================================================

Sub DebugBooleanResult( _
    label As String, _
    value As Boolean)
    ' Registra o valor Boolean, o tipo VBA e o resultado do ramo para diagnosticar retornos contraditorios da API.

    Debug.Print _
        "      " & label & " = " & CStr(value)

    Debug.Print _
        "      TypeName(" & label & ") = " & TypeName(value)

    If value Then
        Debug.Print "      RESULTADO BOOLEAN = TRUE"
    Else
        Debug.Print "      RESULTADO BOOLEAN = FALSE"
    End If

End Sub


' ============================================================
' EXPORTACAO DOS ITENS DE CUT LIST DAS PECAS DA MONTAGEM:
'
' A logica de exportacao foi mantida; comentarios e mensagens foram preparados para este pacote.
' ============================================================

Function ExportFlatPatternToDXF( _
    targetBody As SldWorks.Body2, _
    flatFeat As SldWorks.Feature, _
    outPath As String) As Boolean
    ' Seleciona e valida o FlatPattern, chama ExportToDWG2 e exige retorno positivo da API e arquivo nao vazio. A selecao e limpa ao sair.

    On Error GoTo fail

    ExportFlatPatternToDXF = False

    If targetBody Is Nothing Then

        Debug.Print _
            "      [ERRO DXF] Body = Nothing"

        Exit Function

    End If

    If flatFeat Is Nothing Then

        Debug.Print _
            "      [ERRO DXF] FlatPattern = Nothing"

        Exit Function

    End If

    Debug.Print _
        "      Corpo alvo: " & _
        targetBody.Name

    Debug.Print _
        "      FlatPattern alvo: " & _
        flatFeat.Name

    Debug.Print _
        "      Tipo: " & _
        flatFeat.GetTypeName2

    ' --------------------------------------------------------
    ' LIMPA A SELECAO
    ' --------------------------------------------------------

    swModel.ClearSelection2 True

    ' --------------------------------------------------------
    ' SELECIONA O FLAT-PATTERN ALVO
    ' --------------------------------------------------------

    Dim selected As Boolean

    selected = _
        flatFeat.Select2(False, 0)

    Debug.Print _
        "      FlatPattern.Select2 = " & _
        CStr(selected)

    DebugBooleanResult _
        "FlatPattern.Select2", _
        selected

    If selected = False Then

        Debug.Print _
            "      [ERRO DXF] FlatPattern.Select2 retornou False."

        Exit Function

    End If

    ' --------------------------------------------------------
    ' VALIDA A SELECAO PELO SELECTION MANAGER
    ' --------------------------------------------------------

    Dim swSelMgr As SldWorks.SelectionMgr

    Set swSelMgr = _
        swModel.SelectionManager

    Dim selectedCount As Long

    selectedCount = _
        swSelMgr.GetSelectedObjectCount2(-1)

    Debug.Print _
        "      Objetos selecionados: " & _
        selectedCount

    If selectedCount <> 1 Then

        Debug.Print _
            "      [ERRO DXF] Nao existe exatamente um objeto selecionado."

        Exit Function

    End If

    ' --------------------------------------------------------
    ' VALIDA O OBJETO SELECIONADO
    ' --------------------------------------------------------

    Dim selectedObject As Object

    Set selectedObject = _
        swSelMgr.GetSelectedObject6(1, -1)

    If selectedObject Is Nothing Then

        Debug.Print _
            "      [ERRO DXF] GetSelectedObject6 retornou Nothing."

        Exit Function

    End If

    Debug.Print _
        "      Objeto selecionado: " & _
        TypeName(selectedObject)

    ' --------------------------------------------------------
    ' CONVERTE O OBJETO SELECIONADO PARA FEATURE
    ' --------------------------------------------------------

    Dim selectedFeature As SldWorks.Feature

    On Error Resume Next

    Set selectedFeature = selectedObject

    On Error GoTo fail

    If selectedFeature Is Nothing Then

        Debug.Print _
            "      [ERRO DXF] Objeto selecionado nao e Feature."

        Exit Function

    End If

    Debug.Print _
        "      Feature selecionada: " & _
        selectedFeature.Name

    Debug.Print _
        "      Tipo selecionado: " & _
        selectedFeature.GetTypeName2

    ' --------------------------------------------------------
    ' CONFIRMA QUE A FEATURE E O FLAT-PATTERN ESPERADO
    ' --------------------------------------------------------

    If selectedFeature.Name <> flatFeat.Name Then

        Debug.Print _
            "      [ERRO DXF] FlatPattern selecionado e diferente do alvo."

        Debug.Print _
            "      Alvo: " & flatFeat.Name

        Debug.Print _
            "      Selecionado: " & selectedFeature.Name

        Exit Function

    End If

    Debug.Print _
        "      FlatPattern correto confirmado."

    ' --------------------------------------------------------
    ' OPCOES DE EXPORTACAO DXF
    ' --------------------------------------------------------

    Dim sheetMetalOptions As Long

    sheetMetalOptions = 5

    Debug.Print _
        "      SheetMetalOptions = " & _
        sheetMetalOptions

    Debug.Print _
        "      Geometria = SIM"

    Debug.Print _
        "      Linhas de dobra = SIM"

    ' --------------------------------------------------------
    ' ALINHAMENTO DA EXPORTACAO
    ' --------------------------------------------------------

    Dim alignment(11) As Double
    Dim i As Long

    For i = 0 To 11

        alignment(i) = 0#

    Next i

    ' --------------------------------------------------------
    ' CHAMADA DE EXPORTACAO
    ' --------------------------------------------------------

    Dim dxfOK As Boolean

    Debug.Print _
        "      Chamando ExportToDWG2..."

    Debug.Print _
        "      Caminho: " & outPath

    dxfOK = _
        swPart.ExportToDWG2( _
            outPath, _
            swModel.GetPathName, _
            swExportToDWG_e.swExportToDWG_ExportSheetMetal, _
            True, _
            alignment, _
            False, _
            False, _
            sheetMetalOptions, _
            Empty)

    Debug.Print _
        "      ExportToDWG2 = " & _
        CStr(dxfOK)

    Debug.Print _
        "      TypeName(ExportToDWG2) = " & TypeName(dxfOK) & _
        " | VarType=" & CStr(VarType(dxfOK))

    Debug.Print "      [CONDICAO DXF] If dxfOK Then | valor=" & CStr(dxfOK)

    If dxfOK Then

        Debug.Print "      [RAMO DXF] dxfOK=True; verificando arquivo exportado."

    Else

        Debug.Print "      [RAMO DXF] dxfOK=False; exportacao classificada como falha."
        Debug.Print _
            "      [ERRO DXF] ExportToDWG2 retornou False."

        swModel.ClearSelection2 True
        Exit Function

    End If
    If Dir$(outPath) = "" Then
        Debug.Print "      [ERRO DXF] ExportToDWG2 retornou True, mas o arquivo nao existe: " & outPath
        swModel.ClearSelection2 True
        Exit Function
    End If

    If FileLen(outPath) <= 0 Then
        Debug.Print "      [ERRO DXF] Arquivo exportado vazio: " & outPath
        swModel.ClearSelection2 True
        Exit Function
    End If

    Debug.Print "      DXF criado com sucesso: " & outPath
    ExportFlatPatternToDXF = True

    swModel.ClearSelection2 True

    Exit Function

    ' --------------------------------------------------------
    ' VALIDA O ARQUIVO GERADO
    ' --------------------------------------------------------

    If Dir(outPath) = "" Then

        Debug.Print _
            "      [ERRO DXF] Arquivo nao encontrado."

        Exit Function

    End If

    If FileLen(outPath) = 0 Then

        Debug.Print _
            "      [ERRO DXF] Arquivo vazio."

        Exit Function

    End If

    Debug.Print _
        "      DXF criado com sucesso!"

    Debug.Print _
        "      Arquivo: " & outPath

    ExportFlatPatternToDXF = True

    swModel.ClearSelection2 True

    Exit Function

fail:

    Debug.Print _
        "      [ERRO DXF] " & _
        Err.Number & _
        " - " & _
        Err.Description

    On Error Resume Next

    swModel.ClearSelection2 True

    ExportFlatPatternToDXF = False

End Function


' ============================================================
' ROTINA PRINCIPAL
' ============================================================

' Constantes e contadores do modulo ficam antes das rotinas, conforme a sintaxe do VBA.

Private Function IsSuppressedInHierarchy( _
    ByVal component As SldWorks.Component2, _
    ByRef queryFailed As Boolean) As Boolean
    ' Verifica o componente e todos os ancestrais. Falhas de consulta sao tratadas conservadoramente como supressao.

    Dim current As SldWorks.Component2
    Dim parentComponent As SldWorks.Component2
    Dim suppressionState As Long

    IsSuppressedInHierarchy = True
    queryFailed = False
    Set current = component

    On Error GoTo QueryError
    Do While Not current Is Nothing
        suppressionState = current.GetSuppression2
        If suppressionState = swComponentSuppressed Then Exit Function
        Set parentComponent = current.GetParent
        Set current = parentComponent
    Loop

    IsSuppressedInHierarchy = False
    Exit Function

QueryError:
    queryFailed = True
    IsSuppressedInHierarchy = True
    Debug.Print "[ERRO HIERARQUIA] Componente=" & component.Name2 & _
                " | Erro=" & Err.Number & " | " & Err.Description
End Function

Private Sub ProcessarPeca( _
    ByVal partModel As SldWorks.ModelDoc2, _
    ByRef processedOut As Long, _
    ByRef successOut As Long, _
    ByRef failureOut As Long, _
    ByRef ignoredOut As Long)
    ' Processa uma peca ativa: le quantidades da configuracao e da Cut List, inspeciona os itens e exporta corpos elegiveis.

    Dim processedCount As Long
    Dim successCount As Long
    Dim failCount As Long
    Dim ignoredCount As Long

    processedOut = 0
    successOut = 0
    failureOut = 0
    ignoredOut = 0

    On Error GoTo PartError
    Set swModel = partModel
    Debug.Print "[PECA] " & swModel.GetPathName & _
                " | SaveFlag inicio=" & CStr(swModel.GetSaveFlag)
    filePath = swModel.GetPathName

    If filePath = "" Then

        MsgBox _
            "Salve o arquivo da peca antes de executar a macro.", _
            vbExclamation, _
            "Macro"

        Exit Sub

    End If

    fileName = _
        Mid( _
            filePath, _
            InStrRev(filePath, "\") + 1, _
            InStrRev(filePath, ".") - _
            InStrRev(filePath, "\") - 1)

    ' --------------------------------------------------------
    ' OBTEM O NUMERO DE DESENHO
    ' --------------------------------------------------------

    drawingNumber = _
        GetFileProperty(PropNumeroDesenho())

    If drawingNumber = "" Then

        drawingNumber = _
            GetFileProperty("N" & ChrW(176) & " Desenho")

    End If

    If drawingNumber = "" Then

        drawingNumber = _
            GetFileProperty("Numero")

    End If

    If drawingNumber = "" Then

        drawingNumber = fileName

    End If

    If InStr(drawingNumber, " ") > 0 _
        And drawingNumber = fileName Then

        drawingNumber = _
            Split(drawingNumber, " ")(0)

    End If

    Debug.Print _
        "Numero do desenho: " & drawingNumber

    ' --------------------------------------------------------
    ' OBTEM A QUANTIDADE DA CONFIGURACAO
    ' --------------------------------------------------------

    Dim configQtyStr As String
    Dim configQty As Long

    configQtyStr = _
        GetFileProperty("Quantidade")

    If IsNumeric(configQtyStr) _
        And configQtyStr <> "" Then

        configQty = _
            CLng(configQtyStr)

    Else

        configQty = 1

    End If

    If configQty <= 0 Then

        configQty = 1

    End If

    Debug.Print _
        "Quantidade da configuracao: " & configQty

    ' --------------------------------------------------------
    ' DOCUMENTO DE PECA
    ' --------------------------------------------------------

    Set swPart = swModel

    ' --------------------------------------------------------
    ' OBTEM OS CORPOS SOLIDOS
    ' --------------------------------------------------------

    Dim allBodies As Variant

    allBodies = _
        swPart.GetBodies2( _
            swSolidBody, _
            False)

    If IsEmpty(allBodies) Then
        Debug.Print "[SEM CORPOS SOLIDOS] " & filePath
        ignoredCount = ignoredCount + 1
        ignoredOut = ignoredCount
        Exit Sub
    End If

    ' --------------------------------------------------------
    ' PERCORRE A CUT LIST EXISTENTE
    ' --------------------------------------------------------

    Set swFeat = _
        swModel.FirstFeature

    Dim itemIndex As Integer

    itemIndex = 0
    processedCount = 0
    successCount = 0
    failCount = 0
    ignoredCount = 0

    Do While Not swFeat Is Nothing

        If swFeat.GetTypeName2 = "CutListFolder" Then

            Set swBodyFolder = _
                swFeat.GetSpecificFeature2

            ' Nao chame UpdateCutList durante a exportacao. Essa chamada
            ' pode alterar a Cut List do documento nativo e marcar a peca
            ' e os documentos que a referenciam como alterados.

            Dim vBodies As Variant

            vBodies = _
                swBodyFolder.GetBodies

            If Not IsEmpty(vBodies) Then

                If UBound(vBodies) >= 0 Then

                    Set swBody = _
                        vBodies(0)

                    itemIndex = _
                        itemIndex + 1

                    processedCount = _
                        processedCount + 1

                    Debug.Print ""
                    Debug.Print "------------------------------------------------------"
                    Debug.Print _
                        "ITEM CUT LIST #" & _
                        itemIndex
                    Debug.Print "------------------------------------------------------"

                    ' ------------------------------------------------
                    ' LE PROPRIEDADES DO ITEM
                    ' ------------------------------------------------

                    Dim propMgr As _
                        SldWorks.CustomPropertyManager

                    Set propMgr = _
                        swFeat.CustomPropertyManager

                    Dim desc As String
                    Dim material As String
                    Dim thickness As String
                    Dim cutListQtyStr As String
                    Dim dimensions As String
                    Dim lengthValue As String

                    desc = _
                        GetProp( _
                            propMgr, _
                            PropDescricao())

                    If desc = "" Then

                        desc = _
                            GetProp( _
                                propMgr, _
                                "Descricao")

                    End If

                    material = _
                        GetProp( _
                            propMgr, _
                            "MATERIAL")

                    If material = "" Then

                        material = _
                            GetProp( _
                                propMgr, _
                                "Material")

                    End If

                    thickness = _
                        GetProp( _
                            propMgr, _
                            PropEspessuraChapa())

                    If thickness = "" Then

                        thickness = _
                            GetProp( _
                                propMgr, _
                                "Espessura")

                    End If

                    cutListQtyStr = _
                        GetProp( _
                            propMgr, _
                            "QUANTITY")

                    If cutListQtyStr = "" Then

                        cutListQtyStr = _
                            GetProp( _
                                propMgr, _
                                "Quantidade")

                    End If

                    ' ------------------------------------------------
                    ' CALCULA A QUANTIDADE DO ITEM DA CUT LIST
                    ' ------------------------------------------------

                    Dim cutListQty As Long

                    If IsNumeric(cutListQtyStr) _
                        And cutListQtyStr <> "" Then

                        cutListQty = _
                            CLng(cutListQtyStr)

                    Else

                        cutListQty = _
                            swBodyFolder.GetBodyCount

                    End If

                    If cutListQty <= 0 Then

                        cutListQty = 1

                    End If

                    ' ------------------------------------------------
                    ' CALCULA A QUANTIDADE TOTAL
                    ' ------------------------------------------------

                    Dim totalQty As Long

                    totalQty = _
                        cutListQty * configQty

                    dimensions = _
                        GetProp( _
                            propMgr, _
                            PropDimensoes())

                    If dimensions = "" Then

                        dimensions = _
                            GetProp( _
                                propMgr, _
                                "Dimensoes")

                    End If

                    lengthValue = _
                        GetProp( _
                            propMgr, _
                            "COMPRIMENTO")

                    If lengthValue = "" Then

                        lengthValue = _
                            GetProp( _
                                propMgr, _
                                "Comprimento")

                    End If

                    ' ------------------------------------------------
                    ' IDENTIFICA ITENS DE CHAPA
                    ' ------------------------------------------------

                    Dim isSheetMetal As Boolean

                    isSheetMetal = _
                        CheckIfSheetMetal(swFeat) _
                        Or (thickness <> "")

                    Debug.Print _
                        "Descricao: " & desc

                    Debug.Print _
                        "Material: " & material

                    Debug.Print _
                        "Espessura: " & thickness

                    Debug.Print _
                        "Dimensoes: " & dimensions

                    Debug.Print _
                        "Comprimento: " & lengthValue

                    Debug.Print _
                        "Qtd Cut List: " & cutListQty

                    Debug.Print _
                        "Qtd configuracao: " & configQty

                    Debug.Print _
                        "Qtd final: " & totalQty

                    Debug.Print _
                        "Corpo: " & swBody.Name


                    ' =================================================
                    ' EXPORTA SOMENTE ITENS DE CHAPA METALICA
                    ' =================================================

                    If isSheetMetal Then

                        Debug.Print ""
                        Debug.Print ">> TIPO: CHAPA"

                        ' ------------------------------------------------
                        ' MONTA O NOME LEGADO PARA A PROGRAMACAO DO LASER
                        ' ------------------------------------------------

                        Dim finalFileName As String
                        Dim outPathComplete As String

                        finalFileName = material

                        If thickness <> "" Then

                            finalFileName = _
                                finalFileName & _
                                " #" & _
                                FormatThicknessForFileName(thickness)

                        End If

                        finalFileName = _
                            finalFileName & _
                            " " & drawingNumber & _
                            "-" & itemIndex

                        finalFileName = _
                            finalFileName & _
                            " " & totalQty & PcsSuffix()

                        finalFileName = _
                            CleanFileName(finalFileName)

                        outPathComplete = _
                            dxfOutputFolder & _
                            finalFileName & _
                            ".dxf"

                        Debug.Print _
                            "Nome DXF: " & _
                            finalFileName

                        ' ------------------------------------------------
                        ' LOCALIZA O FLAT-PATTERN ASSOCIADO
                        ' ------------------------------------------------

                        Debug.Print _
                            "   Procurando FlatPattern do corpo..."

                        Dim targetFlatPattern As _
                            SldWorks.Feature

                        Set targetFlatPattern = _
                            FindFlatPatternForBody(swBody)

                        If targetFlatPattern Is Nothing Then

                            Debug.Print _
                                "   [ERRO DXF] FlatPattern nao encontrado."

                            failCount = _
                                failCount + 1

                        Else

                            Debug.Print _
                                "   FlatPattern alvo: " & _
                                targetFlatPattern.Name

                            If ExportFlatPatternToDXF( _
                                    swBody, _
                                    targetFlatPattern, _
                                    outPathComplete) Then

                                successCount = _
                                    successCount + 1

                            Else

                                failCount = _
                                    failCount + 1

                            End If

                        End If


                    ' =================================================
                    ' NAO E CHAPA METALICA -> IGNORA
                    ' =================================================

                    Else

                        ignoredCount = _
                            ignoredCount + 1

                        Debug.Print ""
                        Debug.Print _
                            ">> IGNORANDO: nao e chapa metalica."

                    End If

                End If

            End If

        End If

        Set swFeat = _
            swFeat.GetNextFeature

    Loop

    ' --------------------------------------------------------
    ' LIMPA A SELECAO
    swModel.ClearSelection2 True
    Debug.Print "[PECA RESULTADO] Caminho=" & filePath & _
                " | CutListItens=" & processedCount & _
                " | DXF=" & successCount & _
                " | Falhas=" & failCount & _
                " | Ignorados=" & ignoredCount & _
                " | SaveFlag fim=" & CStr(swModel.GetSaveFlag)
    processedOut = processedCount
    successOut = successCount
    failureOut = failCount
    ignoredOut = ignoredCount
    Exit Sub

PartError:
    Debug.Print "[ERRO PROCESSAR PECA] Caminho=" & partModel.GetPathName & _
                " | Numero=" & Err.Number & " | Descricao=" & Err.Description
    processedOut = processedCount
    successOut = successCount
    failureOut = failCount + 1
    ignoredOut = ignoredCount
    On Error Resume Next
    If Not swModel Is Nothing Then swModel.ClearSelection2 True
    On Error GoTo 0
End Sub

Sub main()
    ' Valida a montagem, enumera componentes uma vez, elimina duplicatas pelo caminho, processa a primeira configuracao, fecha cada peca sem salvar, restaura a montagem e informa totais.
    Dim assemblyModel As SldWorks.ModelDoc2
    Dim assemblyDoc As SldWorks.AssemblyDoc
    Dim componentArray As Variant
    Dim component As SldWorks.Component2
    Dim partModel As SldWorks.ModelDoc2
    Dim activePart As SldWorks.ModelDoc2
    Dim processedPaths As Object
    Dim i As Long
    Dim lowerBound As Long
    Dim upperBound As Long
    Dim partPath As String
    Dim targetConfiguration As String
    Dim originalConfiguration As String
    Dim isVirtual As Boolean
    Dim suppressedQueryFailed As Boolean
    Dim openedByMacro As Boolean
    Dim configurationOK As Boolean
    Dim partTitle As String
    Dim remainingPart As SldWorks.ModelDoc2
    Dim partSaveFlag As Boolean
    Dim openErrors As Long
    Dim openWarnings As Long
    Dim activateErrors As Long
    Dim processedDelta As Long
    Dim successDelta As Long
    Dim failureDelta As Long
    Dim ignoredDelta As Long
    Dim restoredAssembly As SldWorks.ModelDoc2

    On Error GoTo FatalError
    totalPartsVisited = 0
    totalPartsActivated = 0
    totalPartsClosed = 0
    totalProcessed = 0
    totalDxfSuccess = 0
    totalDxfFailures = 0
    totalIgnoredItems = 0
    totalSkippedSuppressed = 0
    totalSkippedVirtual = 0
    totalDuplicateParts = 0
    totalOpenFailures = 0

    Set swApp = Application.SldWorks
    Debug.Print "[VERSION] _2.chapas_metalicas_dxf_montagem " & MACRO_VERSION
    Debug.Print "[INICIO] Exportacao DXF de chapas metalicas da montagem ativa"

    If swApp Is Nothing Then
        MsgBox "O SOLIDWORKS nao esta disponivel.", vbCritical, "DXF da montagem"
        Exit Sub
    End If

    Set assemblyModel = swApp.ActiveDoc
    If assemblyModel Is Nothing Then
        MsgBox "Abra a montagem que deseja processar.", vbExclamation, "DXF da montagem"
        Exit Sub
    End If
    If assemblyModel.GetType <> SW_DOC_ASSEMBLY Then
        MsgBox "O documento ativo precisa ser uma montagem (.SLDASM).", vbExclamation, "DXF da montagem"
        Exit Sub
    End If
    If Len(assemblyModel.GetPathName) = 0 Then
        MsgBox "Salve a montagem antes de executar a macro.", vbExclamation, "DXF da montagem"
        Exit Sub
    End If

    dxfOutputFolder = Left$(assemblyModel.GetPathName, InStrRev(assemblyModel.GetPathName, "\"))
    If Len(dxfOutputFolder) = 0 Then
        MsgBox "Nao foi possivel determinar a pasta da montagem principal.", vbCritical, "DXF da montagem"
        Exit Sub
    End If

    Set assemblyDoc = assemblyModel
    Set processedPaths = CreateObject("Scripting.Dictionary")
    processedPaths.CompareMode = vbTextCompare
    Debug.Print "[MONTAGEM] " & assemblyModel.GetPathName
    Debug.Print "[PASTA DESTINO DXF] " & dxfOutputFolder
    Debug.Print "[CUT LIST] UpdateCutList nao sera chamado; arquivos SOLIDWORKS nao serao salvos pela macro."

    componentArray = assemblyDoc.GetComponents(False)
    If Not IsArray(componentArray) Then
        MsgBox "GetComponents(False) nao retornou a lista de componentes.", vbExclamation, "DXF da montagem"
        GoTo RestoreAndReport
    End If

    On Error Resume Next
    Err.Clear
    lowerBound = LBound(componentArray)
    upperBound = UBound(componentArray)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo FatalError
        GoTo RestoreAndReport
    End If
    On Error GoTo FatalError

    For i = lowerBound To upperBound
        Set component = Nothing
        On Error Resume Next
        Err.Clear
        Set component = componentArray(i)
        If Err.Number <> 0 Or component Is Nothing Then
            Debug.Print "[ERRO COMPONENTE] Indice=" & i & " | " & Err.Number & " | " & Err.Description
            totalOpenFailures = totalOpenFailures + 1
            Err.Clear
            On Error GoTo FatalError
            GoTo NextComponent
        End If
        On Error GoTo FatalError

        If IsSuppressedInHierarchy(component, suppressedQueryFailed) Then
            If suppressedQueryFailed Then
                totalOpenFailures = totalOpenFailures + 1
            Else
                totalSkippedSuppressed = totalSkippedSuppressed + 1
            End If
            Debug.Print "[IGNORADO SUPRIMIDO/HIERARQUIA] " & component.Name2 & _
                        " | ConsultaComErro=" & CStr(suppressedQueryFailed)
            GoTo NextComponent
        End If

        On Error Resume Next
        Err.Clear
        isVirtual = component.IsVirtual
        partPath = component.GetPathName
        If Err.Number <> 0 Then
            Debug.Print "[ERRO COMPONENTE PATH/VIRTUAL] " & component.Name2 & _
                        " | Erro=" & Err.Number & " | " & Err.Description
            totalOpenFailures = totalOpenFailures + 1
            Err.Clear
            On Error GoTo FatalError
            GoTo NextComponent
        End If
        On Error GoTo FatalError

        If isVirtual Then
            totalSkippedVirtual = totalSkippedVirtual + 1
            Debug.Print "[IGNORADO VIRTUAL] " & component.Name2
            GoTo NextComponent
        End If
        If Len(partPath) = 0 Then
            Debug.Print "[IGNORADO SEM CAMINHO] " & component.Name2
            GoTo NextComponent
        End If
        If UCase$(Right$(partPath, 7)) <> ".SLDPRT" Then GoTo NextComponent

        On Error Resume Next
        Err.Clear
        targetConfiguration = component.ReferencedConfiguration
        If Err.Number <> 0 Then
            Debug.Print "[ERRO CONFIG REFERENCIADA] " & component.Name2 & _
                        " | Erro=" & Err.Number & " | " & Err.Description
            totalOpenFailures = totalOpenFailures + 1
            Err.Clear
            On Error GoTo FatalError
            GoTo NextComponent
        End If
        On Error GoTo FatalError

        totalPartsVisited = totalPartsVisited + 1
        If processedPaths.Exists(partPath) Then
            totalDuplicateParts = totalDuplicateParts + 1
            If StrComp(CStr(processedPaths(partPath)), targetConfiguration, vbTextCompare) <> 0 Then
                Debug.Print "[DUPLICADO OUTRA CONFIGURACAO] Arquivo=" & partPath & _
                            " | Exportada primeiro=" & CStr(processedPaths(partPath)) & _
                            " | Ignorada depois=" & targetConfiguration
            Else
                Debug.Print "[DUPLICADO CONSOLIDADO] " & partPath & _
                            " | Configuracao=" & targetConfiguration
            End If
            GoTo NextComponent
        End If
        processedPaths.Add partPath, targetConfiguration

        Set partModel = swApp.GetOpenDocumentByName(partPath)
        openedByMacro = (partModel Is Nothing)
        originalConfiguration = ""
        If openedByMacro Then
            openErrors = 0
            openWarnings = 0
            Set partModel = swApp.OpenDoc6( _
                partPath, SW_DOC_PART, SW_OPEN_DOC_SILENT, targetConfiguration, openErrors, openWarnings)
            Debug.Print "[OPEN DOC6] Caminho=" & partPath & _
                        " | Nothing=" & CStr(partModel Is Nothing) & _
                        " | Erros=" & openErrors & " | Avisos=" & openWarnings
        Else
            Debug.Print "[JA ABERTA] Reutilizando documento existente: " & partPath
        End If
        If partModel Is Nothing Then
            totalOpenFailures = totalOpenFailures + 1
            GoTo NextComponent
        End If

        If partModel.GetType <> SW_DOC_PART Then
            Debug.Print "[ERRO TIPO DOCUMENTO] Caminho=" & partPath & _
                        " | Tipo=" & partModel.GetType & " | Esperado=SLDPRT"
            totalOpenFailures = totalOpenFailures + 1
            GoTo CloseMacroOpenedPart
        End If

        On Error Resume Next
        Err.Clear
        originalConfiguration = partModel.ConfigurationManager.ActiveConfiguration.Name
        If Err.Number <> 0 Then
            Debug.Print "[ERRO CONFIG ATIVA] " & partPath & _
                        " | Erro=" & Err.Number & " | " & Err.Description
            totalOpenFailures = totalOpenFailures + 1
            Err.Clear
            On Error GoTo FatalError
            GoTo CloseMacroOpenedPart
        End If
        On Error GoTo FatalError

        activateErrors = 0
        Set activePart = swApp.ActivateDoc3( _
            partModel.GetTitle, False, SW_DONT_REBUILD_ACTIVE_DOC, activateErrors)
        Debug.Print "[ACTIVATE DOC3] Documento=" & partPath & _
                    " | Nothing=" & CStr(activePart Is Nothing) & _
                    " | Erros/Avisos=" & activateErrors
        If activePart Is Nothing Then
            totalOpenFailures = totalOpenFailures + 1
            GoTo CloseMacroOpenedPart
        End If
        If StrComp(activePart.GetPathName, partPath, vbTextCompare) <> 0 Then
            Debug.Print "[FALHA ATIVACAO] Caminho ativo diferente. Ativo=" & _
                        activePart.GetPathName & " | Esperado=" & partPath
            totalOpenFailures = totalOpenFailures + 1
            GoTo CloseMacroOpenedPart
        End If

        If Len(targetConfiguration) > 0 And _
           StrComp(originalConfiguration, targetConfiguration, vbTextCompare) <> 0 Then
            configurationOK = activePart.ShowConfiguration2(targetConfiguration)
            Debug.Print "[SHOW CONFIGURATION2] Config=" & targetConfiguration & _
                        " | Retorno=" & CStr(configurationOK)
            If Not configurationOK Then
                totalOpenFailures = totalOpenFailures + 1
                GoTo CloseMacroOpenedPart
            End If
        End If

        Set partModel = activePart
        totalPartsActivated = totalPartsActivated + 1
        ProcessarPeca partModel, processedDelta, successDelta, failureDelta, ignoredDelta
        totalProcessed = totalProcessed + processedDelta
        totalDxfSuccess = totalDxfSuccess + successDelta
        totalDxfFailures = totalDxfFailures + failureDelta
        totalIgnoredItems = totalIgnoredItems + ignoredDelta

CloseMacroOpenedPart:
        On Error Resume Next
        partTitle = ""
        partSaveFlag = False
        If Not partModel Is Nothing Then
            partTitle = partModel.GetTitle
            partSaveFlag = partModel.GetSaveFlag
            Debug.Print "[FECHAMENTO PREPARADO] Caminho=" & partPath & _
                        " | Ja estava aberta=" & CStr(Not openedByMacro) & _
                        " | SaveFlag=" & CStr(partSaveFlag) & _
                        " | Fechamento sem salvar=True"
        End If
        Set swFeat = Nothing
        Set swBodyFolder = Nothing
        Set swBody = Nothing
        Set swPart = Nothing
        Set swModel = Nothing
        Set activePart = Nothing
        Set partModel = Nothing
        If Len(partTitle) > 0 Then
            Err.Clear
            swApp.QuitDoc partTitle
            If Err.Number <> 0 Then
                Debug.Print "[ERRO FECHAR PECA SEM SALVAR] " & partPath & _
                            " | Erro=" & Err.Number & " | " & Err.Description
                totalOpenFailures = totalOpenFailures + 1
                Err.Clear
            Else
                Err.Clear
                Set remainingPart = swApp.GetOpenDocumentByName(partPath)
                If Err.Number <> 0 Then
                    Debug.Print "[ERRO VERIFICAR FECHAMENTO] " & partPath & _
                                " | Erro=" & Err.Number & " | " & Err.Description
                    totalOpenFailures = totalOpenFailures + 1
                    Err.Clear
                ElseIf Not remainingPart Is Nothing Then
                    Debug.Print "[ERRO PECA AINDA ABERTA] " & partPath
                    totalOpenFailures = totalOpenFailures + 1
                ElseIf openedByMacro Then
                    Debug.Print "[PECA FECHADA SEM SALVAR] Aberta pela macro: " & partPath
                    totalPartsClosed = totalPartsClosed + 1
                Else
                    Debug.Print "[PECA FECHADA SEM SALVAR] Ja estava aberta: " & partPath
                    totalPartsClosed = totalPartsClosed + 1
                End If
            End If
        Else
            Debug.Print "[ERRO FECHAR PECA] Nao foi possivel obter o titulo do documento: " & partPath
            totalOpenFailures = totalOpenFailures + 1
        End If
        Set remainingPart = Nothing
        On Error GoTo FatalError

NextComponent:
    Next i

RestoreAndReport:
    On Error Resume Next
    Set restoredAssembly = swApp.ActivateDoc3( _
        assemblyModel.GetTitle, False, SW_DONT_REBUILD_ACTIVE_DOC, activateErrors)
    Debug.Print "[RESTAURACAO MONTAGEM] Nothing=" & CStr(restoredAssembly Is Nothing) & _
                " | Erros/Avisos=" & activateErrors
    On Error GoTo 0

    Debug.Print "[RESUMO] Pecas fisicas unicas ativadas=" & totalPartsActivated & _
                " | Pecas fechadas=" & totalPartsClosed & _
                " | Modelos candidatos=" & totalPartsVisited & _
                " | Duplicados=" & totalDuplicateParts & _
                " | Suprimidos=" & totalSkippedSuppressed & _
                " | Virtuais=" & totalSkippedVirtual & _
                " | CutList itens=" & totalProcessed & _
                " | DXFs criados=" & totalDxfSuccess & _
                " | Falhas DXF=" & totalDxfFailures & _
                " | Falhas abertura/consulta=" & totalOpenFailures & _
                " | Itens ignorados=" & totalIgnoredItems

    MsgBox "Exportacao DXF da montagem concluida." & vbCrLf & vbCrLf & _
           "Pecas fisicas unicas ativadas: " & totalPartsActivated & vbCrLf & _
           "Pecas fechadas: " & totalPartsClosed & vbCrLf & _
           "DXFs criados: " & totalDxfSuccess & vbCrLf & _
           "Falhas DXF: " & totalDxfFailures & vbCrLf & _
           "Falhas de abertura/consulta: " & totalOpenFailures & vbCrLf & _
           "Pecas repetidas consolidadas: " & totalDuplicateParts & vbCrLf & _
           "Consulte o Immediate Window para detalhes.", _
           IIf(totalDxfFailures + totalOpenFailures > 0, vbExclamation, vbInformation), _
           "DXF da montagem"
    Exit Sub

FatalError:
    Debug.Print "[ERRO FATAL] Numero=" & Err.Number & _
                " | Descricao=" & Err.Description & " | Origem=" & Err.Source
    On Error Resume Next
    If Not partModel Is Nothing Then
        partTitle = partModel.GetTitle
        Set swFeat = Nothing
        Set swBodyFolder = Nothing
        Set swBody = Nothing
        Set swPart = Nothing
        Set swModel = Nothing
        Set activePart = Nothing
        Set partModel = Nothing
        If Len(partTitle) > 0 Then
            swApp.QuitDoc partTitle
            Debug.Print "[FECHAMENTO ERRO FATAL] " & partTitle
        End If
    End If
    If Not assemblyModel Is Nothing Then
        swApp.ActivateDoc3 assemblyModel.GetTitle, False, SW_DONT_REBUILD_ACTIVE_DOC, activateErrors
    End If
    MsgBox "A macro encontrou um erro inesperado." & vbCrLf & _
           "Erro " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Consulte o Immediate Window.", vbCritical, "DXF da montagem"
End Sub


