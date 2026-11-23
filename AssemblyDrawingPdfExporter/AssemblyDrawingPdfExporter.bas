Option Explicit

' ===============================================================================
' SOLIDWORKS 2025
' Marker: 20261105-04-PDF-MONTAGEM-SAVEAS-DIRETO
'
' Exports to PDF:
' - the active assembly drawing, if it exists;
' - drawings for physical parts and subassemblies in the assembly.
'
' Each drawing must be in the same folder as its model and have the same
' base name. Repeated components produce only one PDF. PDFs are saved
' next to their drawings. The macro does not force a rebuild.
' Each drawing is activated before SaveAs3 and closed after export,
' discarding in-memory changes without saving SOLIDWORKS documents.
' ===============================================================================

Private Const MACRO_VERSION As String = "20261105-04-PDF-MONTAGEM-SAVEAS-DIRETO"
Private Const SW_DOC_ASSEMBLY As Long = 2
Private Const SW_DOC_DRAWING As Long = 3
Private Const SW_OPEN_DOC_SILENT As Long = 1
Private Const SW_SAVE_AS_CURRENT_VERSION As Long = 0
Private Const SW_SAVE_AS_OPTIONS_SILENT As Long = 1
Private Const SW_DONT_REBUILD_ACTIVE_DOC As Long = 1

Private swApp As SldWorks.SldWorks
Private fso As Object
Private modelPaths As Object
Private drawingsToExport As Object

Private modelsChecked As Long
Private drawingsFound As Long
Private drawingsMissing As Long
Private suppressedComponents As Long
Private virtualComponents As Long
Private componentsWithoutPath As Long
Private duplicateDrawings As Long
Private existingPdfs As Long
Private createdPdfs As Long
Private failures As Long

Sub main()
    Dim assemblyModel As SldWorks.ModelDoc2
    Dim assemblyDoc As SldWorks.AssemblyDoc
    Dim components As Variant
    Dim lowerBound As Long
    Dim upperBound As Long
    Dim i As Long
    Dim component As SldWorks.Component2
    Dim componentPath As String
    Dim componentType As Long
    Dim isVirtual As Boolean
    Dim failuresBeforeSuppression As Long
    Dim confirmationResponse As VbMsgBoxResult
    Dim assemblyPath As String
    Dim activationErrors As Long

    On Error GoTo FatalErrorHandler

    ResetCounters
    Debug.Print "[VERSION] " & MACRO_VERSION
    Debug.Print "[START] PDF export for the active assembly"

    Set swApp = Application.SldWorks
    If swApp Is Nothing Then
        MsgBox "SOLIDWORKS is not available.", vbExclamation, "Assembly PDF"
        Exit Sub
    End If

    Set assemblyModel = swApp.ActiveDoc
    If assemblyModel Is Nothing Then
        MsgBox "Open the assembly you want to process before running the macro.", vbExclamation, "Assembly PDF"
        Exit Sub
    End If
    If assemblyModel.GetType <> SW_DOC_ASSEMBLY Then
        MsgBox "The active document must be an assembly (.SLDASM).", vbExclamation, "Assembly PDF"
        Exit Sub
    End If

    assemblyPath = assemblyModel.GetPathName
    If Len(assemblyPath) = 0 Then
        MsgBox "Save the assembly before running the macro.", vbExclamation, "Assembly PDF"
        Exit Sub
    End If

    Set assemblyDoc = assemblyModel
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set modelPaths = CreateObject("Scripting.Dictionary")
    Set drawingsToExport = CreateObject("Scripting.Dictionary")
    modelPaths.CompareMode = vbTextCompare
    drawingsToExport.CompareMode = vbTextCompare
    Debug.Print "[ASSEMBLY] " & assemblyPath
    ' The main assembly and all physical models are added to the processing list.
    ConsiderModel assemblyPath, True

    components = assemblyDoc.GetComponents(False)
    If IsArray(components) Then
        On Error Resume Next
        Err.Clear
        lowerBound = LBound(components)
        upperBound = UBound(components)
        If Err.Number <> 0 Then
            Debug.Print "[WARNING] The component array is empty or has invalid bounds."
            Err.Clear
            On Error GoTo FatalErrorHandler
            GoTo ConfirmDrawingList
        End If
        On Error GoTo FatalErrorHandler

        For i = lowerBound To upperBound
            Set component = Nothing
            On Error Resume Next
            Err.Clear
            Set component = components(i)
            If Err.Number <> 0 Or component Is Nothing Then
                Debug.Print "[COMPONENT ERROR] Index=" & i & " | Error=" & Err.Number & " | " & Err.Description
                failures = failures + 1
                Err.Clear
                On Error GoTo FatalErrorHandler
                GoTo NextComponent
            End If
            On Error GoTo FatalErrorHandler

            failuresBeforeSuppression = failures
            If IsComponentOrAncestorSuppressed(component) Then
                If failures > failuresBeforeSuppression Then
                    Debug.Print "[SKIPPED: SUPPRESSION QUERY ERROR] " & component.Name2
                Else
                    suppressedComponents = suppressedComponents + 1
                    Debug.Print "[SKIPPED: SUPPRESSED] " & component.Name2
                End If
                GoTo NextComponent
            End If

            On Error Resume Next
            Err.Clear
            isVirtual = component.IsVirtual
            If Err.Number <> 0 Then
                Debug.Print "[VIRTUAL STATUS QUERY ERROR] " & component.Name2 & _
                            " | Error=" & Err.Number & " | " & Err.Description
                failures = failures + 1
                Err.Clear
                On Error GoTo FatalErrorHandler
                GoTo NextComponent
            End If
            On Error GoTo FatalErrorHandler

            If isVirtual Then
                virtualComponents = virtualComponents + 1
                Debug.Print "[SKIPPED: VIRTUAL] " & component.Name2
                GoTo NextComponent
            End If

            On Error Resume Next
            Err.Clear
            componentPath = component.GetPathName
            If Err.Number <> 0 Then
                Debug.Print "[PATH QUERY ERROR] " & component.Name2 & _
                            " | Error=" & Err.Number & " | " & Err.Description
                failures = failures + 1
                Err.Clear
                On Error GoTo FatalErrorHandler
                GoTo NextComponent
            End If
            On Error GoTo FatalErrorHandler

            If Len(componentPath) = 0 Then
                componentsWithoutPath = componentsWithoutPath + 1
                Debug.Print "[SKIPPED: NO PATH] " & component.Name2
                GoTo NextComponent
            End If

            componentType = GetDocumentTypeFromPath(componentPath)
            If componentType = 0 Then
                Debug.Print "[SKIPPED: EXTENSION] " & componentPath
                GoTo NextComponent
            End If

            ConsiderModel componentPath, False

NextComponent:
        Next i
    Else
        Debug.Print "[WARNING] GetComponents(False) did not return an array."
    End If

ConfirmDrawingList:
    If drawingsToExport.Count = 0 Then
        MsgBox "No matching drawings were found." & vbCrLf & vbCrLf & _
               "The macro searched for the main assembly drawing and each physical part/subassembly drawing.", _
               vbInformation, "Assembly PDF"
        Debug.Print "[SUMMARY] No drawings to export."
        Exit Sub
    End If

    existingPdfs = CountExistingPDFs
    confirmationResponse = MsgBox( _
        "Drawings found: " & drawingsToExport.Count & vbCrLf & _
        "Matching drawings missing: " & drawingsMissing & vbCrLf & _
        "Existing PDFs that will be overwritten: " & existingPdfs & vbCrLf & vbCrLf & _
        "PDFs will be saved next to their drawings." & vbCrLf & _
        "The macro will not force a rebuild; it will use each drawing's current/saved state." & vbCrLf & _
        "Continue?", _
        vbQuestion + vbYesNo + vbDefaultButton2, "Confirm PDF export")
    If confirmationResponse <> vbYes Then
        Debug.Print "[CANCELLED] Confirmation declined"
        Exit Sub
    End If

    ExportDrawings

    ' Restore the main assembly as the active document if opening
    ' a drawing changed the active window.
    On Error Resume Next
    swApp.ActivateDoc3 assemblyModel.GetTitle, False, SW_DONT_REBUILD_ACTIVE_DOC, activationErrors
    Debug.Print "[RESTORE] Assembly active | ActivateDoc3Error=" & activationErrors
    On Error GoTo 0

    Debug.Print "[SUMMARY] Drawings found=" & drawingsFound & _
                " | PDFs created=" & createdPdfs & _
                " | Missing=" & drawingsMissing & _
                " | Unique models checked=" & modelsChecked & _
                " | Suppressed=" & suppressedComponents & _
                " | Virtual=" & virtualComponents & _
                " | Duplicates=" & duplicateDrawings & _
                " | Failures=" & failures

    MsgBox "Export completed." & vbCrLf & vbCrLf & _
        "Drawings found: " & drawingsFound & vbCrLf & _
        "PDFs created: " & createdPdfs & vbCrLf & _
        "Drawings without a matching file: " & drawingsMissing & vbCrLf & _
        "Unique models checked: " & modelsChecked & vbCrLf & _
        "Suppressed components skipped: " & suppressedComponents & vbCrLf & _
        "Virtual components skipped: " & virtualComponents & vbCrLf & _
        "Components without a path: " & componentsWithoutPath & vbCrLf & _
        "Duplicate drawings consolidated: " & duplicateDrawings & vbCrLf & _
        "Failures: " & failures & vbCrLf & _
        "See the Immediate Window for paths, API results, and errors.", _
        IIf(failures > 0, vbExclamation, vbInformation), "PDF export result"
    Exit Sub

FatalErrorHandler:
    Debug.Print "[FATAL ERROR] Number=" & Err.Number & " | Description=" & Err.Description & " | Source=" & Err.Source
    MsgBox "An unexpected error occurred." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "See the VBA editor's Immediate Window.", vbCritical, "Assembly PDF"
End Sub

Private Sub ResetCounters()
    modelsChecked = 0
    drawingsFound = 0
    drawingsMissing = 0
    suppressedComponents = 0
    virtualComponents = 0
    componentsWithoutPath = 0
    duplicateDrawings = 0
    existingPdfs = 0
    createdPdfs = 0
    failures = 0
End Sub

Private Function IsComponentOrAncestorSuppressed(ByVal component As SldWorks.Component2) As Boolean
    Dim currentComponent As SldWorks.Component2
    Dim suppressionState As Long

    IsComponentOrAncestorSuppressed = True
    Set currentComponent = component
    On Error GoTo HandleError

    Do While Not currentComponent Is Nothing
        suppressionState = currentComponent.GetSuppression2
        If suppressionState = swComponentSuppressed Then Exit Function
        Set currentComponent = currentComponent.GetParent
    Loop

    IsComponentOrAncestorSuppressed = False
    Exit Function

HandleError:
    Debug.Print "[SUPPRESSION QUERY ERROR] Error=" & Err.Number & " | " & Err.Description
    failures = failures + 1
End Function

Private Sub ConsiderModel(ByVal modelPath As String, ByVal isMainAssembly As Boolean)
    Dim modelKey As String
    Dim extension As String
    Dim baseName As String
    Dim drawingPath As String

    On Error GoTo HandleError
    If Len(modelPath) = 0 Then Exit Sub

    modelKey = LCase$(modelPath)
    If modelPaths.Exists(modelKey) Then Exit Sub
    modelPaths.Add modelKey, True
    modelsChecked = modelsChecked + 1

    extension = UCase$(fso.GetExtensionName(modelPath))
    If extension <> "SLDPRT" And extension <> "SLDASM" Then
        Debug.Print "[SKIPPED: MODEL EXTENSION] " & modelPath
        Exit Sub
    End If

    baseName = fso.GetBaseName(modelPath)
    drawingPath = fso.BuildPath(fso.GetParentFolderName(modelPath), baseName & ".SLDDRW")
    If Not fso.FileExists(drawingPath) Then
        drawingsMissing = drawingsMissing + 1
        Debug.Print "[DRAWING MISSING] Model=" & modelPath & " | Searched=" & drawingPath
        Exit Sub
    End If

    If drawingsToExport.Exists(LCase$(drawingPath)) Then
        duplicateDrawings = duplicateDrawings + 1
        Debug.Print "[DUPLICATE CONSOLIDATED] " & drawingPath
        Exit Sub
    End If

    drawingsToExport.Add LCase$(drawingPath), drawingPath
    drawingsFound = drawingsToExport.Count
    Debug.Print "[DRAWING FOUND] " & drawingPath & _
                IIf(isMainAssembly, " | Source=Main assembly", " | Source=Component")
    Exit Sub

HandleError:
    Debug.Print "[MODEL CATALOG ERROR] Path=" & modelPath & _
                " | Error=" & Err.Number & " | " & Err.Description
    failures = failures + 1
End Sub

Private Function CountExistingPDFs() As Long
    Dim key As Variant
    Dim drawingPath As String
    Dim pdfPath As String

    CountExistingPDFs = 0
    For Each key In drawingsToExport.Keys
        drawingPath = CStr(drawingsToExport(key))
        pdfPath = GetPDFPath(drawingPath)
        If fso.FileExists(pdfPath) Then CountExistingPDFs = CountExistingPDFs + 1
    Next key
End Function

Private Sub ExportDrawings()
    Dim key As Variant
    Dim drawingPath As String
    Dim pdfPath As String

    For Each key In drawingsToExport.Keys
        drawingPath = CStr(drawingsToExport(key))
        pdfPath = GetPDFPath(drawingPath)
        Debug.Print "[PROCESSING] Drawing=" & drawingPath & " | PDF=" & pdfPath
        If ExportOneDrawing(drawingPath, pdfPath) Then
            createdPdfs = createdPdfs + 1
        Else
            failures = failures + 1
        End If
    Next key
End Sub

Private Function ExportOneDrawing(ByVal drawingPath As String, ByVal pdfPath As String) As Boolean
    Dim document As SldWorks.ModelDoc2
    Dim activeDocument As SldWorks.ModelDoc2
    Dim closingCheckDocument As SldWorks.ModelDoc2
    Dim modelExtension As SldWorks.ModelDocExtension
    Dim openErrors As Long
    Dim openWarnings As Long
    Dim drawingActivationErrors As Long
    Dim saveErrors As Long
    Dim saveWarnings As Long
    Dim saveResult As Boolean
    Dim pdfSize As Double

    ExportOneDrawing = False
    On Error GoTo HandleError

    Set document = swApp.OpenDoc6( _
        drawingPath, SW_DOC_DRAWING, SW_OPEN_DOC_SILENT, "", openErrors, openWarnings)
    Debug.Print "[OPEN DOC6] Document=" & drawingPath & _
                " | ObjectIsNothing=" & CStr(document Is Nothing) & _
                " | Errors=" & openErrors & " | Warnings=" & openWarnings
    If document Is Nothing Then Exit Function

    If document.GetType <> SW_DOC_DRAWING Then
        Debug.Print "[FAILURE] The opened file is not a drawing: " & drawingPath
        GoTo Cleanup
    End If

    ' SaveAs3 exports the active document. Activating without a rebuild preserves
    ' the faster workflow and uses the drawing's loaded/saved state.
    Set activeDocument = swApp.ActivateDoc3( _
        document.GetTitle, False, SW_DONT_REBUILD_ACTIVE_DOC, drawingActivationErrors)
    Debug.Print "[ACTIVATE DOC3] Title=" & document.GetTitle & _
                " | ObjectIsNothing=" & CStr(activeDocument Is Nothing) & _
                " | Errors/warnings=" & drawingActivationErrors
    If activeDocument Is Nothing Then GoTo Cleanup
    If StrComp(activeDocument.GetPathName, drawingPath, vbTextCompare) <> 0 Then
        Debug.Print "[FAILURE] Active document path does not match the expected path. Active=" & _
                    activeDocument.GetPathName & " | Expected=" & drawingPath
        GoTo Cleanup
    End If
    Set document = activeDocument
    Debug.Print "[REBUILD] Not performed; exporting the drawing's loaded state."

    Set modelExtension = document.Extension
    If modelExtension Is Nothing Then
        Debug.Print "[FAILURE] IModelDocExtension unavailable: " & drawingPath
        GoTo Cleanup
    End If

    Debug.Print "[PDF EXPORT CALL] Direct SaveAs3 | ExportData=Nothing (all sheets) | " & pdfPath
    saveResult = modelExtension.SaveAs3( _
        pdfPath, _
        SW_SAVE_AS_CURRENT_VERSION, _
        SW_SAVE_AS_OPTIONS_SILENT, _
        Nothing, _
        Nothing, _
        saveErrors, _
        saveWarnings)

    Debug.Print "[API] IModelDocExtension.SaveAs3 result=" & CStr(saveResult) & _
                " | Errors=" & saveErrors & " | Warnings=" & saveWarnings

    If saveResult Then
        If fso.FileExists(pdfPath) Then
            pdfSize = CDbl(fso.GetFile(pdfPath).Size)
            If pdfSize > 0 Then
                ExportOneDrawing = True
                Debug.Print "[PDF OK] Path=" & pdfPath & " | Bytes=" & pdfSize
            Else
                Debug.Print "[FAILURE] PDF is empty: " & pdfPath
            End If
        Else
            Debug.Print "[FAILURE] SaveAs3 returned success, but the file does not exist: " & pdfPath
        End If
    End If

Cleanup:
    If Not document Is Nothing Then
        On Error Resume Next
        Err.Clear
        swApp.QuitDoc document.GetTitle
        If Err.Number <> 0 Then
            Debug.Print "[CLOSE ERROR] Path=" & drawingPath & _
                        " | Error=" & Err.Number & " | " & Err.Description
            failures = failures + 1
            Err.Clear
        Else
            Set closingCheckDocument = swApp.GetOpenDocumentByName(drawingPath)
            If closingCheckDocument Is Nothing Then
                Debug.Print "[CLOSED WITHOUT SAVING] Drawing closed after export: " & drawingPath
            Else
                Debug.Print "[CLOSE ERROR] Document still appears to be open: " & drawingPath
                failures = failures + 1
            End If
        End If
        On Error GoTo 0
    End If

    Set modelExtension = Nothing
    Set closingCheckDocument = Nothing
    Set activeDocument = Nothing
    Set document = Nothing
    Exit Function

HandleError:
    Debug.Print "[PDF ERROR] Path=" & drawingPath & _
                " | Error=" & Err.Number & " | " & Err.Description & " | Source=" & Err.Source
    Resume Cleanup
End Function

Private Function GetDocumentTypeFromPath(ByVal path As String) As Long
    Dim extension As String

    extension = UCase$(fso.GetExtensionName(path))
    Select Case extension
        Case "SLDPRT": GetDocumentTypeFromPath = 1
        Case "SLDASM": GetDocumentTypeFromPath = SW_DOC_ASSEMBLY
        Case Else: GetDocumentTypeFromPath = 0
    End Select
End Function

Private Function GetPDFPath(ByVal drawingPath As String) As String
    GetPDFPath = fso.BuildPath( _
        fso.GetParentFolderName(drawingPath), _
        fso.GetBaseName(drawingPath) & ".pdf")
End Function
