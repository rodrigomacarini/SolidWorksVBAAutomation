Option Explicit

' Debugging prompt: Share [VERSION], [SUMMARY], and relevant [DXF ERROR]/[PART RESULT] lines, plus SOLIDWORKS version and an anonymized test setup.
' Do not include customer names, network paths, or proprietary drawing data.

' ============================================================
' MACRO SOLIDWORKS 2025
'
' EXPORTS THE CUT LIST FROM EACH UNIQUE PHYSICAL PART IN THE ASSEMBLY:
'
' SHEET METAL PARTS:
' -> One DXF file per Cut List item.
' -> Only the FlatPattern associated with the matching Body2.
' -> Flat geometry and bend lines.
'
' NON-SHEET-METAL ITEMS:
' -> Skipped.
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
' DOCUMENT PROPERTIES
' ============================================================

Function GetFileProperty(propName As String) As String
    ' Reads a property from the active configuration first, then falls back to document-level properties. The existing alternate degree-symbol lookup is preserved.

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
' CUT LIST ITEM PROPERTIES
' ============================================================

Function GetProp( _
    propMgr As SldWorks.CustomPropertyManager, _
    propName As String) As String
    ' Reads a resolved custom property from the current Cut List folder.

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
' OUTPUT FILENAME SANITIZATION
' ============================================================

Function CleanFileName(txt As String) As String
    ' Replaces Windows-invalid filename characters and collapses repeated spaces before building an output path.

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
' IDENTIFIES SHEET METAL
' ============================================================

Function CheckIfSheetMetal( _
    swFeat As SldWorks.Feature) As Boolean
    ' Classifies an item by its Cut List property names. Portuguese and English search terms are functional compatibility tokens and remain unchanged.

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
' COMPARES TWO BODY2 OBJECTS
' ============================================================

Function SameBody( _
    bodyA As SldWorks.Body2, _
    bodyB As SldWorks.Body2) As Boolean
    ' Compares body selection identifiers so the located FlatPattern is tied to the intended body.

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
' GETS THE BODY ASSOCIATED WITH A FLAT-PATTERN
' ============================================================

Function GetBodyFromFlatPattern( _
    flatFeat As SldWorks.Feature) As SldWorks.Body2
    ' Follows FlatPattern -> GetDefinition -> FixedFace2 -> GetBody. The edge fallback is retained.

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
' FINDS THE FLAT-PATTERN FOR A BODY
' ============================================================

Function FindFlatPatternForBody( _
    targetBody As SldWorks.Body2) As SldWorks.Feature
    ' Walks the feature tree and returns the FlatPattern associated with the requested body.

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
                    "      FlatPattern found: " & _
                    feat.Name

                Debug.Print _
                    "      Associated body: " & _
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
' PROPERTY NAMES AND LEGACY SUFFIX
' ============================================================

Function PcsSuffix() As String
    ' Returns the legacy Portuguese quantity suffix used in DXF filenames; changing it could affect downstream laser programming.

    PcsSuffix = "p" & ChrW(231) & "s"

End Function


Function DescriptionPropertyName() As String
    ' Builds the exact SOLIDWORKS property name with ChrW to avoid source-encoding corruption.

    DescriptionPropertyName = _
        "Descri" & ChrW(231) & ChrW(227) & "o"

End Function


Function SheetThicknessPropertyName() As String
    ' Builds the exact SOLIDWORKS property name; this functional key is not translated.

    SheetThicknessPropertyName = _
        "Espessura da Chapa met" & ChrW(225) & "lica"

End Function


Function DimensionsPropertyName() As String
    ' Builds the exact SOLIDWORKS property name with ChrW.

    DimensionsPropertyName = _
        "Dimens" & ChrW(245) & "es"

End Function


Function DrawingNumberPropertyName() As String
    ' Builds the exact SOLIDWORKS drawing-number property name with ChrW.

    DrawingNumberPropertyName = _
        "N" & ChrW(250) & "mero do Desenho"

End Function


Function FormatThicknessForFileName( _
    thicknessValue As String) As String
    ' Preserves the existing decimal-separator convention in generated DXF filenames.

    FormatThicknessForFileName = _
        Replace(Trim(thicknessValue), ".", ",")

End Function


' ============================================================
' BOOLEAN DIAGNOSTICS
' ============================================================

Sub DebugBooleanResult( _
    label As String, _
    value As Boolean)
    ' Logs the Boolean value, VBA type, and branch result to diagnose contradictory API results.

    Debug.Print _
        "      " & label & " = " & CStr(value)

    Debug.Print _
        "      TypeName(" & label & ") = " & TypeName(value)

    If value Then
        Debug.Print "      BOOLEAN RESULT = TRUE"
    Else
        Debug.Print "      BOOLEAN RESULT = FALSE"
    End If

End Sub


' ============================================================
' EXPORTS CUT LIST ITEMS FROM PHYSICAL PARTS IN THE ASSEMBLY:
'
' Export logic is preserved; comments and messages are prepared for this portfolio package.
' ============================================================

Function ExportFlatPatternToDXF( _
    targetBody As SldWorks.Body2, _
    flatFeat As SldWorks.Feature, _
    outPath As String) As Boolean
    ' Selects and validates the target FlatPattern, calls ExportToDWG2, and requires a successful API result plus a non-empty output file. Selection is cleared on exit.

    On Error GoTo fail

    ExportFlatPatternToDXF = False

    If targetBody Is Nothing Then

        Debug.Print _
            "      [DXF ERROR] Body = Nothing"

        Exit Function

    End If

    If flatFeat Is Nothing Then

        Debug.Print _
            "      [DXF ERROR] FlatPattern = Nothing"

        Exit Function

    End If

    Debug.Print _
        "      Target body: " & _
        targetBody.Name

    Debug.Print _
        "      Target FlatPattern: " & _
        flatFeat.Name

    Debug.Print _
        "      Tipo: " & _
        flatFeat.GetTypeName2

    ' --------------------------------------------------------
    ' CLEARS THE SELECTION
    ' --------------------------------------------------------

    swModel.ClearSelection2 True

    ' --------------------------------------------------------
    ' SELECTS THE TARGET FLAT-PATTERN
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
            "      [DXF ERROR] FlatPattern.Select2 returned False."

        Exit Function

    End If

    ' --------------------------------------------------------
    ' VALIDATES THE SELECTION THROUGH THE SELECTION MANAGER
    ' --------------------------------------------------------

    Dim swSelMgr As SldWorks.SelectionMgr

    Set swSelMgr = _
        swModel.SelectionManager

    Dim selectedCount As Long

    selectedCount = _
        swSelMgr.GetSelectedObjectCount2(-1)

    Debug.Print _
        "      Selected objects: " & _
        selectedCount

    If selectedCount <> 1 Then

        Debug.Print _
            "      [DXF ERROR] Selection must contain exactly one object."

        Exit Function

    End If

    ' --------------------------------------------------------
    ' VALIDATES THE SELECTED OBJECT
    ' --------------------------------------------------------

    Dim selectedObject As Object

    Set selectedObject = _
        swSelMgr.GetSelectedObject6(1, -1)

    If selectedObject Is Nothing Then

        Debug.Print _
            "      [DXF ERROR] GetSelectedObject6 returned Nothing."

        Exit Function

    End If

    Debug.Print _
        "      Selected object: " & _
        TypeName(selectedObject)

    ' --------------------------------------------------------
    ' CONVERTS THE SELECTED OBJECT TO A FEATURE
    ' --------------------------------------------------------

    Dim selectedFeature As SldWorks.Feature

    On Error Resume Next

    Set selectedFeature = selectedObject

    On Error GoTo fail

    If selectedFeature Is Nothing Then

        Debug.Print _
            "      [DXF ERROR] Selected object is not a Feature."

        Exit Function

    End If

    Debug.Print _
        "      Selected Feature: " & _
        selectedFeature.Name

    Debug.Print _
        "      Selected type: " & _
        selectedFeature.GetTypeName2

    ' --------------------------------------------------------
    ' CONFIRMS THAT THE FEATURE IS THE EXPECTED FLAT-PATTERN
    ' --------------------------------------------------------

    If selectedFeature.Name <> flatFeat.Name Then

        Debug.Print _
            "      [DXF ERROR] Selected FlatPattern does not match the target."

        Debug.Print _
            "      Alvo: " & flatFeat.Name

        Debug.Print _
            "      Selecionado: " & selectedFeature.Name

        Exit Function

    End If

    Debug.Print _
        "      Correct FlatPattern confirmed."

    ' --------------------------------------------------------
    ' DXF EXPORT OPTIONS
    ' --------------------------------------------------------

    Dim sheetMetalOptions As Long

    sheetMetalOptions = 5

    Debug.Print _
        "      SheetMetalOptions = " & _
        sheetMetalOptions

    Debug.Print _
        "      Geometry = YES"

    Debug.Print _
        "      Bend lines = YES"

    ' --------------------------------------------------------
    ' EXPORT ALIGNMENT
    ' --------------------------------------------------------

    Dim alignment(11) As Double
    Dim i As Long

    For i = 0 To 11

        alignment(i) = 0#

    Next i

    ' --------------------------------------------------------
    ' EXPORT CALL
    ' --------------------------------------------------------

    Dim dxfOK As Boolean

    Debug.Print _
        "      Calling ExportToDWG2..."

    Debug.Print _
        "      Path: " & outPath

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

    Debug.Print "      [DXF CONDITION] If dxfOK Then | value=" & CStr(dxfOK)

    If dxfOK Then

        Debug.Print "      [DXF BRANCH] dxfOK=True; checking exported file."

    Else

        Debug.Print "      [DXF BRANCH] dxfOK=False; export classified as a failure."
        Debug.Print _
            "      [DXF ERROR] ExportToDWG2 returned False."

        swModel.ClearSelection2 True
        Exit Function

    End If
    If Dir$(outPath) = "" Then
        Debug.Print "      [DXF ERROR] ExportToDWG2 returned True, but the file does not exist: " & outPath
        swModel.ClearSelection2 True
        Exit Function
    End If

    If FileLen(outPath) <= 0 Then
        Debug.Print "      [DXF ERROR] Exported file is empty: " & outPath
        swModel.ClearSelection2 True
        Exit Function
    End If

    Debug.Print "      DXF created successfully: " & outPath
    ExportFlatPatternToDXF = True

    swModel.ClearSelection2 True

    Exit Function

    ' --------------------------------------------------------
    ' VALIDATES THE OUTPUT FILE
    ' --------------------------------------------------------

    If Dir(outPath) = "" Then

        Debug.Print _
            "      [DXF ERROR] File not found."

        Exit Function

    End If

    If FileLen(outPath) = 0 Then

        Debug.Print _
            "      [DXF ERROR] File is empty."

        Exit Function

    End If

    Debug.Print _
        "      DXF created successfully!"

    Debug.Print _
        "      File: " & outPath

    ExportFlatPatternToDXF = True

    swModel.ClearSelection2 True

    Exit Function

fail:

    Debug.Print _
        "      [DXF ERROR] " & _
        Err.Number & _
        " - " & _
        Err.Description

    On Error Resume Next

    swModel.ClearSelection2 True

    ExportFlatPatternToDXF = False

End Function


' ============================================================
' MAIN PROCEDURE
' ============================================================

' Module constants and counters are declared before procedures, as required by VBA syntax.

Private Function IsSuppressedInHierarchy( _
    ByVal component As SldWorks.Component2, _
    ByRef queryFailed As Boolean) As Boolean
    ' Checks the component and every parent. An API query failure is treated conservatively as suppressed.

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
    Debug.Print "[HIERARCHY ERROR] Component=" & component.Name2 & _
                " | Error=" & Err.Number & " | " & Err.Description
End Function

Private Sub ProcessPart( _
    ByVal partModel As SldWorks.ModelDoc2, _
    ByRef processedOut As Long, _
    ByRef successOut As Long, _
    ByRef failureOut As Long, _
    ByRef ignoredOut As Long)
    ' Processes one active part: reads configuration and Cut List quantities, inspects each item, and exports eligible sheet-metal bodies.

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
    Debug.Print "[PART] " & swModel.GetPathName & _
                " | Initial SaveFlag=" & CStr(swModel.GetSaveFlag)
    filePath = swModel.GetPathName

    If filePath = "" Then

        MsgBox _
            "Save the part file before running the macro.", _
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
    ' READS THE DRAWING NUMBER
    ' --------------------------------------------------------

    drawingNumber = _
        GetFileProperty(DrawingNumberPropertyName())

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
        "Drawing number: " & drawingNumber

    ' --------------------------------------------------------
    ' READS THE CONFIGURATION QUANTITY
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
        "Configuration quantity: " & configQty

    ' --------------------------------------------------------
    ' PART DOCUMENT
    ' --------------------------------------------------------

    Set swPart = swModel

    ' --------------------------------------------------------
    ' GETS THE SOLID BODIES
    ' --------------------------------------------------------

    Dim allBodies As Variant

    allBodies = _
        swPart.GetBodies2( _
            swSolidBody, _
            False)

    If IsEmpty(allBodies) Then
        Debug.Print "[NO SOLID BODIES] " & filePath
        ignoredCount = ignoredCount + 1
        ignoredOut = ignoredCount
        Exit Sub
    End If

    ' --------------------------------------------------------
    ' ITERATES THE EXISTING CUT LIST
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

            ' Do not call UpdateCutList during export. That call
            ' can modify the native document Cut List and mark the part
            ' and referencing documents as modified.

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
                        "CUT LIST ITEM #" & _
                        itemIndex
                    Debug.Print "------------------------------------------------------"

                    ' ------------------------------------------------
                    ' READS THE CUT LIST ITEM PROPERTIES
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
                            DescriptionPropertyName())

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
                            SheetThicknessPropertyName())

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
                    ' CALCULATES THE CUT LIST ITEM QUANTITY
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
                    ' CALCULATES THE TOTAL QUANTITY
                    ' ------------------------------------------------

                    Dim totalQty As Long

                    totalQty = _
                        cutListQty * configQty

                    dimensions = _
                        GetProp( _
                            propMgr, _
                            DimensionsPropertyName())

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
                    ' IDENTIFIES SHEET METAL ITEMS
                    ' ------------------------------------------------

                    Dim isSheetMetal As Boolean

                    isSheetMetal = _
                        CheckIfSheetMetal(swFeat) _
                        Or (thickness <> "")

                    Debug.Print _
                        "Description: " & desc

                    Debug.Print _
                        "Material: " & material

                    Debug.Print _
                        "Thickness: " & thickness

                    Debug.Print _
                        "Dimensions: " & dimensions

                    Debug.Print _
                        "Length: " & lengthValue

                    Debug.Print _
                        "Cut List qty: " & cutListQty

                    Debug.Print _
                        "Configuration qty: " & configQty

                    Debug.Print _
                        "Total qty: " & totalQty

                    Debug.Print _
                        "Body: " & swBody.Name


                    ' =================================================
                    ' EXPORTS SHEET METAL ITEMS ONLY
                    ' =================================================

                    If isSheetMetal Then

                        Debug.Print ""
                        Debug.Print ">> TYPE: SHEET METAL"

                        ' ------------------------------------------------
                        ' BUILDS THE LEGACY LASER-PROGRAMMING FILENAME
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
                            "DXF filename: " & _
                            finalFileName

                        ' ------------------------------------------------
                        ' FINDS THE ASSOCIATED FLAT-PATTERN
                        ' ------------------------------------------------

                        Debug.Print _
                            "   Looking for the FlatPattern associated with this body..."

                        Dim targetFlatPattern As _
                            SldWorks.Feature

                        Set targetFlatPattern = _
                            FindFlatPatternForBody(swBody)

                        If targetFlatPattern Is Nothing Then

                            Debug.Print _
                                "   [DXF ERROR] FlatPattern not found."

                            failCount = _
                                failCount + 1

                        Else

                            Debug.Print _
                                "   Target FlatPattern: " & _
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
                    ' NOT SHEET METAL -> SKIP
                    ' =================================================

                    Else

                        ignoredCount = _
                            ignoredCount + 1

                        Debug.Print ""
                        Debug.Print _
                            ">> SKIPPED: not a sheet metal item."

                    End If

                End If

            End If

        End If

        Set swFeat = _
            swFeat.GetNextFeature

    Loop

    ' --------------------------------------------------------
    ' CLEARS THE SELECTION
    swModel.ClearSelection2 True
    Debug.Print "[PART RESULT] Path=" & filePath & _
                " | Cut List items=" & processedCount & _
                " | DXF=" & successCount & _
                " | Failures=" & failCount & _
                " | Skipped=" & ignoredCount & _
                " | Final SaveFlag=" & CStr(swModel.GetSaveFlag)
    processedOut = processedCount
    successOut = successCount
    failureOut = failCount
    ignoredOut = ignoredCount
    Exit Sub

PartError:
    Debug.Print "[PART PROCESSING ERROR] Path=" & partModel.GetPathName & _
                " | Number=" & Err.Number & " | Description=" & Err.Description
    processedOut = processedCount
    successOut = successCount
    failureOut = failCount + 1
    ignoredOut = ignoredCount
    On Error Resume Next
    If Not swModel Is Nothing Then swModel.ClearSelection2 True
    On Error GoTo 0
End Sub

Sub main()
    ' Validates the active assembly, enumerates components once, deduplicates by full path, processes the first referenced configuration, closes each part without saving, restores the assembly, and reports totals.
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
    Debug.Print "[VERSION] AssemblySheetMetalDxfExporter " & MACRO_VERSION
    Debug.Print "[START] Exportacao DXF de chapas metalicas da montagem ativa"

    If swApp Is Nothing Then
        MsgBox "SOLIDWORKS is not available.", vbCritical, "DXF da montagem"
        Exit Sub
    End If

    Set assemblyModel = swApp.ActiveDoc
    If assemblyModel Is Nothing Then
        MsgBox "Open the assembly you want to process.", vbExclamation, "DXF da montagem"
        Exit Sub
    End If
    If assemblyModel.GetType <> SW_DOC_ASSEMBLY Then
        MsgBox "The active document must be an assembly (.SLDASM).", vbExclamation, "DXF da montagem"
        Exit Sub
    End If
    If Len(assemblyModel.GetPathName) = 0 Then
        MsgBox "Save the assembly before running the macro.", vbExclamation, "DXF da montagem"
        Exit Sub
    End If

    dxfOutputFolder = Left$(assemblyModel.GetPathName, InStrRev(assemblyModel.GetPathName, "\"))
    If Len(dxfOutputFolder) = 0 Then
        MsgBox "Could not determine the main assembly folder.", vbCritical, "DXF da montagem"
        Exit Sub
    End If

    Set assemblyDoc = assemblyModel
    Set processedPaths = CreateObject("Scripting.Dictionary")
    processedPaths.CompareMode = vbTextCompare
    Debug.Print "[ASSEMBLY] " & assemblyModel.GetPathName
    Debug.Print "[DXF OUTPUT FOLDER] " & dxfOutputFolder
    Debug.Print "[CUT LIST] UpdateCutList will not be called; the macro will not save SOLIDWORKS documents."

    componentArray = assemblyDoc.GetComponents(False)
    If Not IsArray(componentArray) Then
        MsgBox "GetComponents(False) did not return a component list.", vbExclamation, "DXF da montagem"
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
            Debug.Print "[COMPONENT ERROR] Index=" & i & " | " & Err.Number & " | " & Err.Description
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
            Debug.Print "[SKIPPED: SUPPRESSED/HIERARCHY] " & component.Name2 & _
                        " | QueryFailed=" & CStr(suppressedQueryFailed)
            GoTo NextComponent
        End If

        On Error Resume Next
        Err.Clear
        isVirtual = component.IsVirtual
        partPath = component.GetPathName
        If Err.Number <> 0 Then
            Debug.Print "[COMPONENT PATH/VIRTUAL ERROR] " & component.Name2 & _
                        " | Error=" & Err.Number & " | " & Err.Description
            totalOpenFailures = totalOpenFailures + 1
            Err.Clear
            On Error GoTo FatalError
            GoTo NextComponent
        End If
        On Error GoTo FatalError

        If isVirtual Then
            totalSkippedVirtual = totalSkippedVirtual + 1
            Debug.Print "[SKIPPED VIRTUAL] " & component.Name2
            GoTo NextComponent
        End If
        If Len(partPath) = 0 Then
            Debug.Print "[SKIPPED: NO PATH] " & component.Name2
            GoTo NextComponent
        End If
        If UCase$(Right$(partPath, 7)) <> ".SLDPRT" Then GoTo NextComponent

        On Error Resume Next
        Err.Clear
        targetConfiguration = component.ReferencedConfiguration
        If Err.Number <> 0 Then
            Debug.Print "[REFERENCED CONFIGURATION ERROR] " & component.Name2 & _
                        " | Error=" & Err.Number & " | " & Err.Description
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
                Debug.Print "[DUPLICATE: DIFFERENT CONFIGURATION] File=" & partPath & _
                            " | Exportada primeiro=" & CStr(processedPaths(partPath)) & _
                            " | Ignorada depois=" & targetConfiguration
            Else
                Debug.Print "[DUPLICATE CONSOLIDATED] " & partPath & _
                            " | Configuration=" & targetConfiguration
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
            Debug.Print "[OPEN DOC6] Path=" & partPath & _
                        " | Nothing=" & CStr(partModel Is Nothing) & _
                        " | Errors=" & openErrors & " | Warnings=" & openWarnings
        Else
            Debug.Print "[ALREADY OPEN] Reusing existing document: " & partPath
        End If
        If partModel Is Nothing Then
            totalOpenFailures = totalOpenFailures + 1
            GoTo NextComponent
        End If

        If partModel.GetType <> SW_DOC_PART Then
            Debug.Print "[DOCUMENT TYPE ERROR] Path=" & partPath & _
                        " | Tipo=" & partModel.GetType & " | Esperado=SLDPRT"
            totalOpenFailures = totalOpenFailures + 1
            GoTo CloseMacroOpenedPart
        End If

        On Error Resume Next
        Err.Clear
        originalConfiguration = partModel.ConfigurationManager.ActiveConfiguration.Name
        If Err.Number <> 0 Then
            Debug.Print "[ACTIVE CONFIGURATION ERROR] " & partPath & _
                        " | Error=" & Err.Number & " | " & Err.Description
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
                    " | Errors/Warnings=" & activateErrors
        If activePart Is Nothing Then
            totalOpenFailures = totalOpenFailures + 1
            GoTo CloseMacroOpenedPart
        End If
        If StrComp(activePart.GetPathName, partPath, vbTextCompare) <> 0 Then
            Debug.Print "[ACTIVATION FAILED] Active path mismatch. Active=" & _
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
        ProcessPart partModel, processedDelta, successDelta, failureDelta, ignoredDelta
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
            Debug.Print "[CLOSE PREPARED] Path=" & partPath & _
                        " | Was already open=" & CStr(Not openedByMacro) & _
                        " | SaveFlag=" & CStr(partSaveFlag) & _
                        " | Closed without saving=True"
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
                Debug.Print "[ERROR CLOSING PART WITHOUT SAVING] " & partPath & _
                            " | Error=" & Err.Number & " | " & Err.Description
                totalOpenFailures = totalOpenFailures + 1
                Err.Clear
            Else
                Err.Clear
                Set remainingPart = swApp.GetOpenDocumentByName(partPath)
                If Err.Number <> 0 Then
                    Debug.Print "[CLOSE VERIFICATION ERROR] " & partPath & _
                                " | Error=" & Err.Number & " | " & Err.Description
                    totalOpenFailures = totalOpenFailures + 1
                    Err.Clear
                ElseIf Not remainingPart Is Nothing Then
                    Debug.Print "[ERROR: PART STILL OPEN] " & partPath
                    totalOpenFailures = totalOpenFailures + 1
                ElseIf openedByMacro Then
                    Debug.Print "[PART CLOSED WITHOUT SAVING] Opened by macro: " & partPath
                    totalPartsClosed = totalPartsClosed + 1
                Else
                    Debug.Print "[PART CLOSED WITHOUT SAVING] Was already open: " & partPath
                    totalPartsClosed = totalPartsClosed + 1
                End If
            End If
        Else
            Debug.Print "[ERROR CLOSING PART] Could not get the document title: " & partPath
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
    Debug.Print "[ASSEMBLY RESTORATION] Nothing=" & CStr(restoredAssembly Is Nothing) & _
                " | Errors/Warnings=" & activateErrors
    On Error GoTo 0

    Debug.Print "[SUMMARY] Unique physical parts activated=" & totalPartsActivated & _
                " | Parts closed=" & totalPartsClosed & _
                " | Candidate models=" & totalPartsVisited & _
                " | Duplicates=" & totalDuplicateParts & _
                " | Suppressed=" & totalSkippedSuppressed & _
                " | Virtual=" & totalSkippedVirtual & _
                " | Cut List items=" & totalProcessed & _
                " | DXFs created=" & totalDxfSuccess & _
                " | DXF failures=" & totalDxfFailures & _
                " | Open/query failures=" & totalOpenFailures & _
                " | Items skipped=" & totalIgnoredItems

    MsgBox "Assembly DXF export completed." & vbCrLf & vbCrLf & _
           "Unique physical parts activated: " & totalPartsActivated & vbCrLf & _
           "Parts closed: " & totalPartsClosed & vbCrLf & _
           "DXFs created: " & totalDxfSuccess & vbCrLf & _
           "DXF failures: " & totalDxfFailures & vbCrLf & _
           "Open/query failures: " & totalOpenFailures & vbCrLf & _
           "Duplicate parts consolidated: " & totalDuplicateParts & vbCrLf & _
           "See the Immediate Window for details.", _
           IIf(totalDxfFailures + totalOpenFailures > 0, vbExclamation, vbInformation), _
           "DXF da montagem"
    Exit Sub

FatalError:
    Debug.Print "[FATAL ERROR] Number=" & Err.Number & _
                " | Description=" & Err.Description & " | Origem=" & Err.Source
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
            Debug.Print "[FATAL ERROR CLEANUP] " & partTitle
        End If
    End If
    If Not assemblyModel Is Nothing Then
        swApp.ActivateDoc3 assemblyModel.GetTitle, False, SW_DONT_REBUILD_ACTIVE_DOC, activateErrors
    End If
    MsgBox "The macro encountered an unexpected error." & vbCrLf & _
           "Erro " & Err.Number & ": " & Err.Description & vbCrLf & _
           "See the Immediate Window.", vbCritical, "DXF da montagem"
End Sub


