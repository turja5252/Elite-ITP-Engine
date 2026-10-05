Attribute VB_Name = "Module1"
Private Const TEMPLATE_SHEET As String = "Generated ITP Template"
Private Const TEMPLATE_SHEET_TYPE_2 As String = "Generated ITP Template Type 2"
Private Const OUTPUT_SHEET As String = "Generated ITP"
Private Const OUTPUT_SHEET_MARKER As String = "GeneratedITP"
Private Const OUTPUT_SHEET_NAME_REF As String = "GeneratedITPSheetName"
Private Const WORKING_SHEET_SUFFIX As String = "-WIP"
Private Const LEGACY_WORKING_SHEET_SUFFIX As String = " - WIP"
Private Const FIRST_DATA_ROW As Long = 10
Private Const TEMPLATE_SECTION_ROW As Long = 10
Private Const TEMPLATE_ACTIVITY_ROW As Long = 11

Private Const LIB_SECTION_TITLE_COL As String = "C"
Private Const LIB_ACTIVITY_DESC_COL As String = "D"
Private Const LIB_CONTROL_DOCS_COL As String = "E"
Private Const LIB_ACCEPTANCE_COL As String = "F"
Private Const LIB_VERIFICATION_COL As String = "G"
Private Const LIB_ELITE_POINT_COL As String = "H"
Private Const LIB_APPROVER1_POINT_COL As String = "I"
Private Const LIB_APPROVER2_POINT_COL As String = "J"
Private Const LIB_COMMENTS_COL As String = "K"
Private Const LIB_HEADER_ROW As Long = 1
Private Const LIB_LOCATION_HEADER As String = "Shop / Field"
Private Const LIB_ROW_TYPE_HEADER As String = "Row Type"
Private Const LIB_LOCATION_INSERT_COL As Long = 12
Private Const LIB_LOCATION_VALIDATION_LIST As String = "Shop,Field,Shop and Field"
Private Const BUTTON_MARGIN_X As Double = 2.5
Private Const BUTTON_MARGIN_Y As Double = 2.5
Private Const REFRESH_BUTTON_MARGIN_X As Double = 5.5
Private Const REFRESH_BUTTON_MARGIN_Y As Double = 3
Private Const REVISION_BUTTON_MARGIN_X As Double = 5.5
Private Const REVISION_BUTTON_MARGIN_Y As Double = 3
Private Const FLAGGED_COMMENT_FILL_COLOR As Long = 65535
Private Const SECTION_DIFF_HIGHLIGHT_COLOR As Long = 255 ' Red highlighter for section title differentiators
Private Const ELITE_HEADER_LOGO_FILE As String = "Elite_Header_Logo.png"
Private Const CLIENT_HEADER_LOGO_FILE As String = "Client_Header_Logo.png"
Private Const COMBINED_HEADER_LOGO_FILE As String = "Elite_Client_Header_Logo.png"
Private Const POSITIONED_CLIENT_HEADER_LOGO_FILE As String = "Client_Header_Positioned.png"
Private Const POSITIONED_CLIENT_HEADER_LOGO_FILE_TYPE_2 As String = "Client_Header_Positioned_Type2.png"
Private Const HEADER_LOGO_LOCKED_HEIGHT As Double = 72
Private Const HEADER_LOGO_EXPORT_SCALE As Double = 4
Private Const CLIENT_LOGO_ENABLED_REF As String = "ClientLogoEnabled"
Private Const ELITE_LOGO_ENABLED_REF As String = "EliteLogoEnabled"
' Client logo UI/feature disabled for now. Set True to restore dual-logo behavior.
Private Const CLIENT_LOGO_FEATURE_ENABLED As Boolean = False
Private Const LOGO_ASSET_SHEET As String = "Logo Assets"
Private Const LOGO_ASSET_ELITE_SHAPE As String = "EliteLogoAsset"

Private Const INPUT_CHAPTER_FIRST_ROW As Long = 6
Private Const INPUT_CHAPTER_LAST_ROW As Long = 40
Private Const INPUT_CHAPTER_COL As Long = 7
Private Const INPUT_SCOPE_COL As Long = 8
Private Const INPUT_WO_COL As Long = 9
Private Const INPUT_DWG_COL As Long = 10
Private Const INPUT_CUSTOM_TITLE_COL As Long = 11
Private Const INPUT_CUSTOM_ROWS_COL As Long = 12
Private Const INPUT_INSPECTION_TYPE_CELL As String = "E10"
Private Const INPUT_INSPECTION_TYPE_SHOP As String = "SHOP"
Private Const INPUT_INSPECTION_TYPE_FIELD As String = "FIELD"

Private Const WPS_FIRST_ROW As Long = 6
Private Const WPS_LAST_ROW As Long = 55
Private Const INPUT_WO_LIST_FIRST_ROW As Long = 26
Private Const INPUT_WO_LIST_LAST_ROW As Long = 35
Private Const INPUT_DWG_LIST_FIRST_ROW As Long = 26
Private Const INPUT_DWG_LIST_LAST_ROW As Long = 64
Private Const INPUT_CUSTOM_FORMAT_FIRST_ROW As Long = 6
Private Const INPUT_CUSTOM_FORMAT_LAST_ROW As Long = 79
Private Const INPUT_CUSTOM_DISABLED_COLOR As Long = 14277081
Private Const LISTS_SCOPE_ALL_COL As String = "G"
Private Const LISTS_SCOPE_SHOP_COL As String = "H"
Private Const LISTS_SCOPE_FIELD_COL As String = "I"
Private Const SCOPE_PREVIEW_MAX_LENGTH As Long = 1000
Private Const WORKBOOK_VERSION_LABEL As String = "Version 1.2.0"
Private Const WORKBOOK_RELEASE_DATE_TEXT As String = "05-Oct-2026"

Public ChapterPickerConfirmed As Boolean
Public ChapterPickerScopeName As String
Public ChapterPickerWOText As String
Public ChapterPickerDWGText As String
Public ChapterPickerCustomTitle As String
Public ChapterPickerCustomRows As Long
Public ChapterPickerAfterChapterText As String

Private flaggedCommentRows As Collection

Public Sub GenerateITP()

    Dim wsIn As Worksheet, wsLib As Worksheet
    Dim wsT As Worksheet, wsOut As Worksheet
    Dim existingGeneratedSheet As Worksheet
    Dim lastLib As Long, libRow As Long, chRow As Long
    Dim outputSheetIndex As Long
    Dim scopeName As String
    Dim outRow As Long, notesRow As Long
    Dim sectionNo As Long, itemNo As Long
    Dim isSectionRow As Boolean
    Dim sectionTitle As String
    Dim thisWO As String, thisDWG As String
    Dim customTitle As String
    Dim customRows As Long
    Dim i As Long
    Dim wpsList As String
    Dim blockFound As Boolean
    Dim desiredSheetName As String
    Dim preservePriorRevision As VbMsgBoxResult
    Dim keepExistingRevisionSheet As Boolean
    Dim generationCompleted As Boolean

    Set wsIn = ThisWorkbook.Worksheets("Inputs")
    Set wsLib = ThisWorkbook.Worksheets("Activity Library")
    Set wsT = GetTemplateWorksheet(wsIn)
    EnsureActivityLibraryLocationColumn wsLib
    ResetFlaggedCommentRows

    If Not ValidateGenerateITPInputs(wsIn) Then Exit Sub

    RefreshScopeList

    wpsList = GetSelectedWPSList(wsIn)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set existingGeneratedSheet = GetGeneratedITPWorksheet()
    desiredSheetName = BuildGeneratedITPSheetName(wsIn)
    outputSheetIndex = GetGeneratedITPSheetIndex()

    If ShouldPromptToPreserveRevision(existingGeneratedSheet, desiredSheetName) Then
        preservePriorRevision = MsgBox("A generated ITP already exists for a different revision:" & vbCrLf & vbCrLf & _
                                       existingGeneratedSheet.Name & vbCrLf & vbCrLf & _
                                       "Yes = keep that revision and create a new generated ITP for " & desiredSheetName & "." & vbCrLf & _
                                       "No = replace the existing generated ITP." & vbCrLf & _
                                       "Cancel = stop.", _
                                       vbQuestion + vbYesNoCancel + vbDefaultButton1, "Elite ITP Builder")

        If preservePriorRevision = vbCancel Then GoTo SafeExit

        keepExistingRevisionSheet = (preservePriorRevision = vbYes)
    End If

    If Not keepExistingRevisionSheet Then DeleteGeneratedITPSheetIfExists

    wsT.Visible = xlSheetVisible

    If outputSheetIndex <= 0 Then
        wsT.Copy After:=wsT
    ElseIf outputSheetIndex = 1 Then
        wsT.Copy Before:=ThisWorkbook.Worksheets(1)
    ElseIf outputSheetIndex > ThisWorkbook.Worksheets.Count Then
        wsT.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Else
        wsT.Copy Before:=ThisWorkbook.Worksheets(outputSheetIndex)
    End If

    Set wsOut = ActiveSheet
    wsOut.Name = OUTPUT_SHEET

    wsT.Visible = xlSheetHidden

    Application.DisplayAlerts = True

    ApplyHeaderFromInputs wsOut, wsIn
    MarkGeneratedITPSheet wsOut
    RenameGeneratedITPSheet wsOut, wsIn
    ApplyTemplateInputValues wsOut, wsIn

    notesRow = FindRowContains(wsOut, "Notes / Comments", 1, 1, 250)

    If notesRow = 0 Then
        MsgBox "Could not find the Notes / Comments section in the template.", vbExclamation, "Elite ITP Builder"
        GoTo SafeExit
    End If

    If notesRow > FIRST_DATA_ROW Then
        wsOut.Rows(FIRST_DATA_ROW & ":" & notesRow - 1).Delete
    End If

    outRow = FIRST_DATA_ROW
    sectionNo = 0
    itemNo = 0

    lastLib = GetActivityLibraryLastDataRow(wsLib)

    For chRow = INPUT_CHAPTER_FIRST_ROW To GetInputChapterLastRow(wsIn)

        scopeName = Trim(CStr(wsIn.Cells(chRow, INPUT_SCOPE_COL).Value))

        If Len(scopeName) > 0 Then

            thisWO = FormatPrefixedValue(Trim(CStr(wsIn.Cells(chRow, INPUT_WO_COL).Value)), "WO-")
            thisDWG = FormatPrefixedValue(Trim(CStr(wsIn.Cells(chRow, INPUT_DWG_COL).Value)), "DWG-")

            If UCase$(scopeName) = "CUSTOM" Then

                customTitle = Trim(CStr(wsIn.Cells(chRow, INPUT_CUSTOM_TITLE_COL).Value))
                If Len(customTitle) = 0 Then customTitle = "Custom"

                customRows = Val(wsIn.Cells(chRow, INPUT_CUSTOM_ROWS_COL).Value)
                If customRows <= 0 Then customRows = 5

                sectionNo = sectionNo + 1
                itemNo = 0

                InsertTemplateRow wsT, wsOut, TEMPLATE_SECTION_ROW, outRow
                WriteSectionRow wsOut, outRow, CStr(sectionNo) & ".0   " & _
                    FormatSectionHeading(ReplaceTokens(customTitle, wsIn, thisWO, thisDWG, wpsList, False), thisWO, thisDWG)
                outRow = outRow + 1

                For i = 1 To customRows
                    itemNo = itemNo + 1
                    InsertTemplateRow wsT, wsOut, TEMPLATE_ACTIVITY_ROW, outRow
                    WriteActivityRow wsOut, outRow, CStr(sectionNo) & "." & CStr(itemNo), "", "", "", "", "", "", "", ""
                    outRow = outRow + 1
                Next i

            Else

                blockFound = False

                For libRow = 2 To lastLib

                    If Trim(CStr(wsLib.Cells(libRow, "A").Value)) = scopeName Then

                        blockFound = True

                                If Len(Trim(CStr(wsLib.Cells(libRow, LIB_SECTION_TITLE_COL).Value))) > 0 Or _
                                    Len(Trim(CStr(wsLib.Cells(libRow, LIB_ACTIVITY_DESC_COL).Value))) > 0 Then

                                     isSectionRow = (Len(Trim(CStr(wsLib.Cells(libRow, LIB_ACTIVITY_DESC_COL).Value))) = 0)

                            If isSectionRow Then

                                sectionNo = sectionNo + 1
                                itemNo = 0

                                sectionTitle = StripLeadingSectionNumber( _
                                    ReplaceTokens(CStr(wsLib.Cells(libRow, LIB_SECTION_TITLE_COL).Value), wsIn, thisWO, thisDWG, wpsList, False))

                                InsertTemplateRow wsT, wsOut, TEMPLATE_SECTION_ROW, outRow
                                WriteSectionRow wsOut, outRow, _
                                    CStr(sectionNo) & ".0   " & FormatSectionHeading(sectionTitle, thisWO, thisDWG), _
                                    wsLib.Cells(libRow, LIB_SECTION_TITLE_COL), _
                                    sectionTitle

                            Else

                                If sectionNo = 0 Then
                                    sectionNo = sectionNo + 1
                                    itemNo = 0
                                    InsertTemplateRow wsT, wsOut, TEMPLATE_SECTION_ROW, outRow
                                    WriteSectionRow wsOut, outRow, _
                                        CStr(sectionNo) & ".0   " & FormatSectionHeading(scopeName, thisWO, thisDWG), _
                                        wsLib.Cells(libRow, "A"), _
                                        scopeName
                                    outRow = outRow + 1
                                End If

                                itemNo = itemNo + 1

                                InsertTemplateRow wsT, wsOut, TEMPLATE_ACTIVITY_ROW, outRow

                                WriteActivityRow wsOut, outRow, _
                                    CStr(sectionNo) & "." & CStr(itemNo), _
                                    ReplaceTokens(CStr(wsLib.Cells(libRow, LIB_ACTIVITY_DESC_COL).Value), wsIn, thisWO, thisDWG, wpsList), _
                                    ReplaceTokens(CStr(wsLib.Cells(libRow, LIB_CONTROL_DOCS_COL).Value), wsIn, thisWO, thisDWG, wpsList), _
                                    ReplaceTokens(CStr(wsLib.Cells(libRow, LIB_ACCEPTANCE_COL).Value), wsIn, thisWO, thisDWG, wpsList), _
                                    ReplaceTokens(CStr(wsLib.Cells(libRow, LIB_VERIFICATION_COL).Value), wsIn, thisWO, thisDWG, wpsList), _
                                    ReplaceTokens(CStr(wsLib.Cells(libRow, LIB_ELITE_POINT_COL).Value), wsIn, thisWO, thisDWG, wpsList), _
                                    ReplaceTokens(CStr(wsLib.Cells(libRow, LIB_APPROVER1_POINT_COL).Value), wsIn, thisWO, thisDWG, wpsList), _
                                    ReplaceTokens(CStr(wsLib.Cells(libRow, LIB_APPROVER2_POINT_COL).Value), wsIn, thisWO, thisDWG, wpsList), _
                                    ReplaceTokens(CStr(wsLib.Cells(libRow, LIB_COMMENTS_COL).Value), wsIn, thisWO, thisDWG, wpsList), _
                                    IsActivityLibraryCommentFlagged(wsLib.Cells(libRow, LIB_COMMENTS_COL)) Or _
                                        IsActivityLibraryCellFormatted(wsLib.Cells(libRow, LIB_ACTIVITY_DESC_COL)), _
                                    wsLib.Cells(libRow, LIB_COMMENTS_COL), _
                                    wsLib.Cells(libRow, LIB_ACTIVITY_DESC_COL)

                            End If

                            outRow = outRow + 1

                        End If

                    End If

                Next libRow

                If blockFound = False Then
                    sectionNo = sectionNo + 1
                    itemNo = 0
                    InsertTemplateRow wsT, wsOut, TEMPLATE_SECTION_ROW, outRow
                    WriteSectionRow wsOut, outRow, CStr(sectionNo) & ".0   " & _
                        FormatSectionHeading(scopeName & " - BLOCK NOT FOUND IN ACTIVITY LIBRARY", thisWO, thisDWG)
                    outRow = outRow + 1
                End If

            End If

        End If

NextInputChapter:

    Next chRow

    RenumberGeneratedITPInSheet wsOut, False
    FormatFinalOutput wsOut
    RefreshGeneratedITPRowButtons wsOut
    ApplyAllFlaggedCommentStyles wsOut

    wsOut.Activate
    wsOut.DisplayPageBreaks = True
    ActiveWindow.View = xlPageBreakPreview
    ActiveWindow.Zoom = 100
    ActiveWindow.DisplayGridlines = False
    Application.GoTo wsOut.Range("A1"), True
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1

    If keepExistingRevisionSheet Then
        FreezeGeneratedITPWorksheet existingGeneratedSheet
    End If

    generationCompleted = True
    If keepExistingRevisionSheet And Not existingGeneratedSheet Is Nothing Then
        MsgBox "ITP generated successfully using the chapter builder." & vbCrLf & vbCrLf & _
               "Archived revision: " & existingGeneratedSheet.Name, vbInformation, "Elite ITP Builder"
    Else
        MsgBox "ITP generated successfully using the chapter builder.", vbInformation, "Elite ITP Builder"
    End If

SafeExit:
    On Error Resume Next
    If Not generationCompleted Then
        If Not wsOut Is Nothing Then
            Application.DisplayAlerts = False
            wsOut.Delete
        End If
        If keepExistingRevisionSheet Then
            If Not existingGeneratedSheet Is Nothing Then SetStoredGeneratedITPSheetName existingGeneratedSheet.Name
        End If
    End If
    wsT.Visible = xlSheetHidden
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

Private Function ShouldPromptToPreserveRevision(ByVal existingGeneratedSheet As Worksheet, ByVal desiredSheetName As String) As Boolean

    If existingGeneratedSheet Is Nothing Then Exit Function
    If Len(desiredSheetName) = 0 Then Exit Function

    ShouldPromptToPreserveRevision = (StrComp(existingGeneratedSheet.Name, desiredSheetName, vbTextCompare) <> 0)

End Function

Public Sub ApplyWorkbookVersionHeader(Optional ByVal wsIn As Worksheet = Nothing)

    If wsIn Is Nothing Then Set wsIn = ThisWorkbook.Worksheets("Inputs")

    If wsIn.Range("A1:B1").MergeCells = False Then wsIn.Range("A1:B1").Merge
    wsIn.Range("D1").NumberFormat = "dd-mmm-yyyy"

    wsIn.Range("A1:B1").HorizontalAlignment = xlCenter
    wsIn.Range("A1:B1").VerticalAlignment = xlCenter
    wsIn.Range("A1:B1").Font.Bold = True
    wsIn.Range("D1").Font.Bold = True

End Sub

Public Sub RefreshInputsSheet()

    Dim wsIn As Worksheet
    Dim lastChapterRow As Long
    Dim lastWORow As Long
    Dim lastDWGRow As Long
    Dim rowNum As Long
    Dim response As VbMsgBoxResult

    Set wsIn = ThisWorkbook.Worksheets("Inputs")

    response = MsgBox("Are you sure you want to clear the Inputs sheet values, scope builder, WO list, drawing list, and reset weld approvals to No?", _
                      vbQuestion + vbYesNo + vbDefaultButton2, "Elite ITP Builder")
    If response <> vbYes Then Exit Sub

    Application.ScreenUpdating = False

    wsIn.Range("B5:B12").ClearContents
    wsIn.Range("E5:E12").ClearContents
    wsIn.Range("P6").ClearContents

    lastChapterRow = GetInputChapterLastRow(wsIn)
    wsIn.Range(wsIn.Cells(INPUT_CHAPTER_FIRST_ROW, INPUT_SCOPE_COL), wsIn.Cells(lastChapterRow, INPUT_CUSTOM_ROWS_COL)).ClearContents
    wsIn.Range("K" & INPUT_CUSTOM_FORMAT_FIRST_ROW & ":L" & INPUT_CUSTOM_FORMAT_LAST_ROW).Interior.Pattern = xlSolid
    wsIn.Range("K" & INPUT_CUSTOM_FORMAT_FIRST_ROW & ":L" & INPUT_CUSTOM_FORMAT_LAST_ROW).Interior.Color = INPUT_CUSTOM_DISABLED_COLOR
    UpdateInputChapterSerials wsIn

    lastWORow = GetDynamicInputListRange(wsIn, "A", INPUT_WO_LIST_FIRST_ROW).Cells(GetDynamicInputListRange(wsIn, "A", INPUT_WO_LIST_FIRST_ROW).Cells.Count).Row
    lastDWGRow = GetDynamicInputListRange(wsIn, "D", INPUT_DWG_LIST_FIRST_ROW).Cells(GetDynamicInputListRange(wsIn, "D", INPUT_DWG_LIST_FIRST_ROW).Cells.Count).Row

    GetDynamicInputListClearRange(wsIn, "A", INPUT_WO_LIST_FIRST_ROW, lastWORow).ClearContents
    GetDynamicInputListClearRange(wsIn, "D", INPUT_DWG_LIST_FIRST_ROW, lastDWGRow).ClearContents

    For rowNum = WPS_FIRST_ROW To WPS_LAST_ROW
        wsIn.Cells(rowNum, "N").Value = "No"
    Next rowNum

    RefreshScopeList

    Application.ScreenUpdating = True

    wsIn.Activate
    wsIn.Range("B6").Select

    MsgBox "Inputs refreshed.", vbInformation, "Elite ITP Builder"

End Sub

Public Sub RenumberGeneratedITP()

    Dim wsOut As Worksheet

    On Error Resume Next
    Set wsOut = GetGeneratedITPWorksheet()
    On Error GoTo 0

    If wsOut Is Nothing Then
        MsgBox "Could not find the Generated ITP sheet.", vbExclamation, "Elite ITP Builder"
        Exit Sub
    End If

    RenumberGeneratedITPInSheet wsOut, True
    RefreshGeneratedITPRowButtons wsOut

End Sub

Public Sub RefreshGeneratedITPButtons()

    Dim wsOut As Worksheet

    On Error Resume Next
    Set wsOut = GetGeneratedITPWorksheet()
    On Error GoTo 0

    If wsOut Is Nothing Then Exit Sub

    RenumberGeneratedITPInSheet wsOut, False
    RefreshGeneratedITPRowButtons wsOut

End Sub

Public Sub DeleteGeneratedITPSelectedRow()

    Dim wsOut As Worksheet
    Dim wsIn As Worksheet
    Dim targetRow As Long
    Dim notesRow As Long
    Dim deleteLastRow As Long
    Dim cellText As String
    Dim sectionNo As Long
    Dim isSectionDelete As Boolean
    Dim confirmSectionDelete As VbMsgBoxResult

    On Error Resume Next
    Set wsOut = GetGeneratedITPWorksheet()
    Set wsIn = ThisWorkbook.Worksheets("Inputs")
    On Error GoTo 0

    If wsOut Is Nothing Then
        MsgBox "Could not find the Generated ITP sheet.", vbExclamation, "Elite ITP Builder"
        Exit Sub
    End If

    targetRow = GetGeneratedITPButtonRow(wsOut)
    If targetRow = 0 Then Exit Sub

    notesRow = FindRowContains(wsOut, "Notes / Comments", 1, 1, 250)
    If targetRow < FIRST_DATA_ROW Then Exit Sub
    If notesRow > 0 And targetRow >= notesRow Then Exit Sub

    cellText = Trim(CStr(wsOut.Range("A" & targetRow).Value))
    If IsGeneratedSectionRow(cellText) Then
        isSectionDelete = True
        sectionNo = GetGeneratedSectionNumber(cellText)
        deleteLastRow = GetGeneratedSectionEndRow(wsOut, targetRow, notesRow)
        confirmSectionDelete = MsgBox("Delete this entire chapter and its activities?" & vbCrLf & vbCrLf & cellText, _
                                      vbQuestion + vbYesNo + vbDefaultButton2, "Elite ITP Builder")
        If confirmSectionDelete <> vbYes Then Exit Sub
    Else
        deleteLastRow = targetRow
    End If

    Application.ScreenUpdating = False

    If isSectionDelete And Not wsIn Is Nothing Then
        DeleteInputChapterBySectionNo wsIn, sectionNo
    End If

    wsOut.Rows(targetRow & ":" & deleteLastRow).Delete
    RenumberGeneratedITPInSheet wsOut, False
    RefreshGeneratedITPRowButtons wsOut

    Application.ScreenUpdating = True

End Sub

Public Sub AddGeneratedITPBlankRow()

    Dim wsOut As Worksheet
    Dim targetRow As Long
    Dim insertRow As Long
    Dim notesRow As Long
    Dim cellText As String
    Dim clearLastColumn As String

    On Error Resume Next
    Set wsOut = GetGeneratedITPWorksheet()
    On Error GoTo 0

    If wsOut Is Nothing Then
        MsgBox "Could not find the Generated ITP sheet.", vbExclamation, "Elite ITP Builder"
        Exit Sub
    End If

    targetRow = GetGeneratedITPButtonRow(wsOut)
    If targetRow = 0 Then Exit Sub

    notesRow = FindRowContains(wsOut, "Notes / Comments", 1, 1, 250)
    If targetRow < FIRST_DATA_ROW Then Exit Sub
    If notesRow > 0 And targetRow >= notesRow Then Exit Sub

    cellText = Trim(CStr(wsOut.Range("A" & targetRow).Value))
    If Not IsGeneratedActivityRow(cellText) Then Exit Sub

    insertRow = targetRow + 1

    Application.ScreenUpdating = False

    wsOut.Rows(insertRow).Insert Shift:=xlDown
    wsOut.Rows(targetRow).Copy
    wsOut.Rows(insertRow).PasteSpecial Paste:=xlPasteFormats
    wsOut.Rows(insertRow).RowHeight = wsOut.Rows(targetRow).RowHeight
    Application.CutCopyMode = False

    clearLastColumn = GetPrintLastColumn(wsOut)
    wsOut.Range("A" & insertRow).Value = wsOut.Range("A" & targetRow).Value
    wsOut.Range("B" & insertRow & ":" & clearLastColumn & insertRow).ClearContents

    RenumberGeneratedITPInSheet wsOut, False
    RefreshGeneratedITPRowButtons wsOut

    Application.ScreenUpdating = True

End Sub

Public Sub MoveGeneratedITPActivityUp()

    MoveGeneratedITPActivity -1

End Sub

Public Sub MoveGeneratedITPActivityDown()

    MoveGeneratedITPActivity 1

End Sub

Public Sub MoveGeneratedITPChapterUp()

    MoveGeneratedITPChapter -1

End Sub

Public Sub MoveGeneratedITPChapterDown()

    MoveGeneratedITPChapter 1

End Sub

Private Sub MoveGeneratedITPActivity(ByVal direction As Long)

    Dim wsOut As Worksheet
    Dim targetRow As Long
    Dim swapRow As Long
    Dim notesRow As Long
    Dim targetText As String
    Dim swapText As String

    If direction <> -1 And direction <> 1 Then Exit Sub

    On Error Resume Next
    Set wsOut = GetGeneratedITPWorksheet()
    On Error GoTo 0

    If wsOut Is Nothing Then
        MsgBox "Could not find the Generated ITP sheet.", vbExclamation, "Elite ITP Builder"
        Exit Sub
    End If

    targetRow = GetGeneratedITPButtonRow(wsOut)
    If targetRow = 0 Then Exit Sub

    notesRow = FindRowContains(wsOut, "Notes / Comments", 1, 1, 250)
    If targetRow < FIRST_DATA_ROW Then Exit Sub
    If notesRow > 0 And targetRow >= notesRow Then Exit Sub

    targetText = Trim(CStr(wsOut.Range("A" & targetRow).Value))
    If Not IsGeneratedActivityRow(targetText) Then Exit Sub

    swapRow = targetRow + direction
    If swapRow < FIRST_DATA_ROW Then Exit Sub
    If notesRow > 0 And swapRow >= notesRow Then Exit Sub

    swapText = Trim(CStr(wsOut.Range("A" & swapRow).Value))
    If Not IsGeneratedActivityRow(swapText) Then Exit Sub

    Application.ScreenUpdating = False

    If direction = -1 Then
        wsOut.Rows(targetRow).Cut
        wsOut.Rows(swapRow).Insert Shift:=xlDown
    Else
        wsOut.Rows(targetRow).Cut
        wsOut.Rows(targetRow + 2).Insert Shift:=xlDown
    End If

    Application.CutCopyMode = False
    RenumberGeneratedITPInSheet wsOut, False
    RefreshGeneratedITPRowButtons wsOut

    Application.ScreenUpdating = True

End Sub

Private Sub MoveGeneratedITPChapter(ByVal direction As Long)

    Dim wsOut As Worksheet
    Dim wsIn As Worksheet
    Dim sectionRow As Long
    Dim sectionEndRow As Long
    Dim notesRow As Long
    Dim sectionNo As Long
    Dim destinationRow As Long
    Dim nextSectionEndRow As Long
    Dim cellText As String

    If direction <> -1 And direction <> 1 Then Exit Sub

    On Error Resume Next
    Set wsOut = GetGeneratedITPWorksheet()
    Set wsIn = ThisWorkbook.Worksheets("Inputs")
    On Error GoTo 0

    If wsOut Is Nothing Then
        MsgBox "Could not find the Generated ITP sheet.", vbExclamation, "Elite ITP Builder"
        Exit Sub
    End If

    sectionRow = GetGeneratedITPButtonRow(wsOut)
    If sectionRow = 0 Then Exit Sub

    notesRow = FindRowContains(wsOut, "Notes / Comments", 1, 1, 250)
    If sectionRow < FIRST_DATA_ROW Then Exit Sub
    If notesRow > 0 And sectionRow >= notesRow Then Exit Sub

    cellText = Trim(CStr(wsOut.Range("A" & sectionRow).Value))
    If Not IsGeneratedSectionRow(cellText) Then Exit Sub

    sectionNo = GetGeneratedSectionNumber(cellText)
    sectionEndRow = GetGeneratedSectionEndRow(wsOut, sectionRow, notesRow)

    If direction = -1 Then
        destinationRow = GetPreviousGeneratedSectionRow(wsOut, sectionRow)
        If destinationRow = 0 Then Exit Sub
    Else
        destinationRow = sectionEndRow + 1
        If notesRow > 0 And destinationRow >= notesRow Then Exit Sub
        If Not IsGeneratedSectionRow(Trim(CStr(wsOut.Range("A" & destinationRow).Value))) Then Exit Sub
        nextSectionEndRow = GetGeneratedSectionEndRow(wsOut, destinationRow, notesRow)
    End If

    Application.ScreenUpdating = False

    If Not wsIn Is Nothing Then MoveInputChapter wsIn, sectionNo, direction

    If direction = -1 Then
        wsOut.Rows(sectionRow & ":" & sectionEndRow).Cut
        wsOut.Rows(destinationRow).Insert Shift:=xlDown
    Else
        wsOut.Rows(sectionRow & ":" & sectionEndRow).Cut
        wsOut.Rows(nextSectionEndRow + 1).Insert Shift:=xlDown
    End If

    Application.CutCopyMode = False
    RenumberGeneratedITPInSheet wsOut, False
    RefreshGeneratedITPRowButtons wsOut

    Application.ScreenUpdating = True

End Sub

Public Sub AddGeneratedITPChapterBlock()

    Dim wsOut As Worksheet
    Dim wsIn As Worksheet
    Dim wsLib As Worksheet
    Dim wsT As Worksheet
    Dim targetRow As Long
    Dim insertRow As Long
    Dim notesRow As Long
    Dim cellText As String
    Dim selectedSectionNo As Long
    Dim scopeName As String
    Dim woText As String
    Dim dwgText As String
    Dim rawWOText As String
    Dim rawDWGText As String
    Dim customTitle As String
    Dim customRows As Long
    Dim wpsList As String

    On Error Resume Next
    Set wsOut = GetGeneratedITPWorksheet()
    Set wsIn = ThisWorkbook.Worksheets("Inputs")
    Set wsLib = ThisWorkbook.Worksheets("Activity Library")
    If Not wsOut Is Nothing And IsTwoApproverWorksheet(wsOut) Then
        Set wsT = ThisWorkbook.Worksheets(TEMPLATE_SHEET_TYPE_2)
    Else
        Set wsT = ThisWorkbook.Worksheets(TEMPLATE_SHEET)
    End If
    On Error GoTo 0

    If wsOut Is Nothing Or wsIn Is Nothing Or wsLib Is Nothing Or wsT Is Nothing Then
        MsgBox "Could not find one or more required sheets for adding a chapter block.", vbExclamation, "Elite ITP Builder"
        Exit Sub
    End If

    targetRow = GetGeneratedITPButtonRow(wsOut)
    If targetRow = 0 Then Exit Sub

    notesRow = FindRowContains(wsOut, "Notes / Comments", 1, 1, 250)
    If targetRow < FIRST_DATA_ROW Then Exit Sub
    If notesRow > 0 And targetRow >= notesRow Then Exit Sub

    cellText = Trim(CStr(wsOut.Range("A" & targetRow).Value))
    If Not IsGeneratedSectionRow(cellText) Then Exit Sub
    selectedSectionNo = GetGeneratedSectionNumber(cellText)

    If Not ShowGeneratedITPChapterPicker(cellText, scopeName, woText, dwgText, customTitle, customRows) Then Exit Sub

    If Len(scopeName) = 0 Then
        MsgBox "Choose a scope block before adding a chapter.", vbExclamation, "Elite ITP Builder"
        Exit Sub
    End If

    rawWOText = RemovePrefixIfPresent(woText, "WO-")
    rawDWGText = RemovePrefixIfPresent(dwgText, "DWG-")

    wpsList = GetSelectedWPSList(wsIn)

    insertRow = GetGeneratedSectionEndRow(wsOut, targetRow, notesRow) + 1

    Application.ScreenUpdating = False

    SyncGeneratedChapterSelectionToInputs wsIn, selectedSectionNo, scopeName, rawWOText, rawDWGText, customTitle, customRows

    InsertGeneratedScopeBlock wsIn, wsLib, wsT, wsOut, scopeName, woText, dwgText, wpsList, insertRow, customTitle, customRows

    RenumberGeneratedITPInSheet wsOut, False
    RefreshGeneratedITPRowButtons wsOut

    Application.ScreenUpdating = True

    wsOut.Activate
    wsOut.Range("A" & insertRow).Select

End Sub

Private Sub SyncGeneratedChapterSelectionToInputs(ByVal wsIn As Worksheet, ByVal afterSectionNo As Long, ByVal scopeName As String, _
                                                  ByVal rawWOText As String, ByVal rawDWGText As String, _
                                                  ByVal customTitle As String, ByVal customRows As Long)

    Dim targetRow As Long

    targetRow = GetOrderedInputChapterRow(wsIn, afterSectionNo)
    If targetRow = 0 Then Exit Sub

    wsIn.Cells(targetRow, INPUT_SCOPE_COL).Value = scopeName
    wsIn.Cells(targetRow, INPUT_WO_COL).Value = rawWOText
    wsIn.Cells(targetRow, INPUT_DWG_COL).Value = rawDWGText

    If UCase$(scopeName) = "CUSTOM" Then
        wsIn.Cells(targetRow, INPUT_CUSTOM_TITLE_COL).Value = customTitle
        wsIn.Cells(targetRow, INPUT_CUSTOM_ROWS_COL).Value = customRows
    Else
        wsIn.Cells(targetRow, INPUT_CUSTOM_TITLE_COL).ClearContents
        wsIn.Cells(targetRow, INPUT_CUSTOM_ROWS_COL).ClearContents
    End If

    EnsureValueInInputList wsIn, "A", INPUT_WO_LIST_FIRST_ROW, rawWOText
    EnsureValueInInputList wsIn, "D", INPUT_DWG_LIST_FIRST_ROW, rawDWGText
    UpdateInputChapterSerials wsIn

End Sub

Private Sub DeleteInputChapterBySectionNo(ByVal wsIn As Worksheet, ByVal sectionNo As Long)

    Dim targetRow As Long
    Dim lastRow As Long

    targetRow = GetInputChapterRowBySectionNo(wsIn, sectionNo)
    If targetRow = 0 Then Exit Sub

    lastRow = GetInputChapterLastRow(wsIn)

    If targetRow < lastRow Then
        wsIn.Range(wsIn.Cells(targetRow + 1, INPUT_CHAPTER_COL), wsIn.Cells(lastRow, INPUT_CUSTOM_ROWS_COL)).Copy
        wsIn.Range(wsIn.Cells(targetRow, INPUT_CHAPTER_COL), wsIn.Cells(lastRow - 1, INPUT_CUSTOM_ROWS_COL)).PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
    End If

    wsIn.Range(wsIn.Cells(lastRow, INPUT_CHAPTER_COL), wsIn.Cells(lastRow, INPUT_CUSTOM_ROWS_COL)).ClearContents
    UpdateInputChapterSerials wsIn

End Sub

Private Sub MoveInputChapter(ByVal wsIn As Worksheet, ByVal sectionNo As Long, ByVal direction As Long)

    Dim targetRow As Long
    Dim swapRow As Long
    Dim tempValues As Variant

    targetRow = GetInputChapterRowBySectionNo(wsIn, sectionNo)
    If targetRow = 0 Then Exit Sub

    swapRow = targetRow + direction
    If swapRow < INPUT_CHAPTER_FIRST_ROW Or swapRow > GetInputChapterLastRow(wsIn) Then Exit Sub
    If Len(Trim(CStr(wsIn.Cells(swapRow, INPUT_SCOPE_COL).Value))) = 0 Then Exit Sub

    tempValues = wsIn.Range(wsIn.Cells(targetRow, INPUT_SCOPE_COL), wsIn.Cells(targetRow, INPUT_CUSTOM_ROWS_COL)).Value
    wsIn.Range(wsIn.Cells(targetRow, INPUT_SCOPE_COL), wsIn.Cells(targetRow, INPUT_CUSTOM_ROWS_COL)).Value = _
        wsIn.Range(wsIn.Cells(swapRow, INPUT_SCOPE_COL), wsIn.Cells(swapRow, INPUT_CUSTOM_ROWS_COL)).Value
    wsIn.Range(wsIn.Cells(swapRow, INPUT_SCOPE_COL), wsIn.Cells(swapRow, INPUT_CUSTOM_ROWS_COL)).Value = tempValues

    UpdateInputChapterSerials wsIn

End Sub

Private Function GetInputChapterRowBySectionNo(ByVal wsIn As Worksheet, ByVal sectionNo As Long) As Long

    Dim rowNum As Long
    Dim currentSectionNo As Long
    Dim scopeName As String

    If sectionNo <= 0 Then Exit Function

    For rowNum = INPUT_CHAPTER_FIRST_ROW To GetInputChapterLastRow(wsIn)
        scopeName = Trim(CStr(wsIn.Cells(rowNum, INPUT_SCOPE_COL).Value))
        If Len(scopeName) > 0 Then
            currentSectionNo = currentSectionNo + 1
            If currentSectionNo = sectionNo Then
                GetInputChapterRowBySectionNo = rowNum
                Exit Function
            End If
        End If
    Next rowNum

End Function

Private Function ShowGeneratedITPChapterPicker(ByVal afterChapterText As String, ByRef scopeName As String, ByRef woText As String, _
                                               ByRef dwgText As String, ByRef customTitle As String, ByRef customRows As Long) As Boolean

    ResetGeneratedITPChapterPickerState
    ChapterPickerAfterChapterText = Trim(afterChapterText)

    On Error GoTo ShowFormError
    frmITPChapterPicker.Show vbModal
    On Error GoTo 0

    If Not ChapterPickerConfirmed Then Exit Function

    scopeName = ChapterPickerScopeName
    woText = FormatPrefixedValue(ChapterPickerWOText, "WO-")
    dwgText = FormatPrefixedValue(ChapterPickerDWGText, "DWG-")
    customTitle = ChapterPickerCustomTitle
    customRows = ChapterPickerCustomRows
    ShowGeneratedITPChapterPicker = True
    Exit Function

ShowFormError:
    MsgBox "Could not open the chapter picker form. " & Err.Description, vbExclamation, "Elite ITP Builder"

End Function

Public Sub ResetGeneratedITPChapterPickerState()

    ChapterPickerConfirmed = False
    ChapterPickerScopeName = ""
    ChapterPickerWOText = ""
    ChapterPickerDWGText = ""
    ChapterPickerCustomTitle = ""
    ChapterPickerCustomRows = 0
    ChapterPickerAfterChapterText = ""

End Sub

Private Sub InsertGeneratedScopeBlock(ByVal wsIn As Worksheet, ByVal wsLib As Worksheet, ByVal wsT As Worksheet, _
                                     ByVal wsOut As Worksheet, ByVal scopeName As String, ByVal thisWO As String, _
                                     ByVal thisDWG As String, ByVal wpsList As String, ByVal insertRow As Long, _
                                     Optional ByVal customTitle As String = "", Optional ByVal customRows As Long = 0)

    Dim lastLib As Long
    Dim libRow As Long
    Dim outRow As Long
    Dim itemNo As Long
    Dim blockFound As Boolean
    Dim sectionInserted As Boolean
    Dim isSectionRow As Boolean
    Dim sectionTitle As String

    outRow = insertRow
    EnsureActivityLibraryLocationColumn wsLib

    If UCase$(scopeName) = "CUSTOM" Then
        If Len(customTitle) = 0 Then customTitle = "Custom"
        If customRows <= 0 Then customRows = 5

        InsertTemplateRow wsT, wsOut, TEMPLATE_SECTION_ROW, outRow
        WriteSectionRow wsOut, outRow, "0.0   " & FormatSectionHeading(ReplaceTokens(customTitle, wsIn, thisWO, thisDWG, wpsList, False), thisWO, thisDWG)
        outRow = outRow + 1

        For itemNo = 1 To customRows
            InsertTemplateRow wsT, wsOut, TEMPLATE_ACTIVITY_ROW, outRow
            WriteActivityRow wsOut, outRow, "0." & CStr(itemNo), "", "", "", "", "", "", "", ""
            outRow = outRow + 1
        Next itemNo

        Exit Sub
    End If

    lastLib = GetActivityLibraryLastDataRow(wsLib)
    blockFound = False
    sectionInserted = False
    itemNo = 0

    For libRow = 2 To lastLib
        If Trim(CStr(wsLib.Cells(libRow, "A").Value)) = scopeName Then
            blockFound = True

            If Len(Trim(CStr(wsLib.Cells(libRow, LIB_SECTION_TITLE_COL).Value))) > 0 Or _
               Len(Trim(CStr(wsLib.Cells(libRow, LIB_ACTIVITY_DESC_COL).Value))) > 0 Then

                isSectionRow = (Len(Trim(CStr(wsLib.Cells(libRow, LIB_ACTIVITY_DESC_COL).Value))) = 0)

                If isSectionRow Then
                    sectionTitle = StripLeadingSectionNumber( _
                        ReplaceTokens(CStr(wsLib.Cells(libRow, LIB_SECTION_TITLE_COL).Value), wsIn, thisWO, thisDWG, wpsList, False))

                    InsertTemplateRow wsT, wsOut, TEMPLATE_SECTION_ROW, outRow
                    WriteSectionRow wsOut, outRow, _
                        "0.0   " & FormatSectionHeading(sectionTitle, thisWO, thisDWG), _
                        wsLib.Cells(libRow, LIB_SECTION_TITLE_COL), _
                        sectionTitle
                    outRow = outRow + 1
                    itemNo = 0
                    sectionInserted = True
                Else
                    If Not sectionInserted Then
                        InsertTemplateRow wsT, wsOut, TEMPLATE_SECTION_ROW, outRow
                        WriteSectionRow wsOut, outRow, _
                            "0.0   " & FormatSectionHeading(scopeName, thisWO, thisDWG), _
                            wsLib.Cells(libRow, "A"), _
                            scopeName
                        outRow = outRow + 1
                        sectionInserted = True
                    End If

                    itemNo = itemNo + 1
                    InsertTemplateRow wsT, wsOut, TEMPLATE_ACTIVITY_ROW, outRow
                    WriteActivityRow wsOut, outRow, _
                        "0." & CStr(itemNo), _
                        ReplaceTokens(CStr(wsLib.Cells(libRow, LIB_ACTIVITY_DESC_COL).Value), wsIn, thisWO, thisDWG, wpsList), _
                        ReplaceTokens(CStr(wsLib.Cells(libRow, LIB_CONTROL_DOCS_COL).Value), wsIn, thisWO, thisDWG, wpsList), _
                        ReplaceTokens(CStr(wsLib.Cells(libRow, LIB_ACCEPTANCE_COL).Value), wsIn, thisWO, thisDWG, wpsList), _
                        ReplaceTokens(CStr(wsLib.Cells(libRow, LIB_VERIFICATION_COL).Value), wsIn, thisWO, thisDWG, wpsList), _
                        ReplaceTokens(CStr(wsLib.Cells(libRow, LIB_ELITE_POINT_COL).Value), wsIn, thisWO, thisDWG, wpsList), _
                        ReplaceTokens(CStr(wsLib.Cells(libRow, LIB_APPROVER1_POINT_COL).Value), wsIn, thisWO, thisDWG, wpsList), _
                        ReplaceTokens(CStr(wsLib.Cells(libRow, LIB_APPROVER2_POINT_COL).Value), wsIn, thisWO, thisDWG, wpsList), _
                        ReplaceTokens(CStr(wsLib.Cells(libRow, LIB_COMMENTS_COL).Value), wsIn, thisWO, thisDWG, wpsList), _
                        IsActivityLibraryCommentFlagged(wsLib.Cells(libRow, LIB_COMMENTS_COL)) Or _
                            IsActivityLibraryCellFormatted(wsLib.Cells(libRow, LIB_ACTIVITY_DESC_COL)), _
                        wsLib.Cells(libRow, LIB_COMMENTS_COL), _
                        wsLib.Cells(libRow, LIB_ACTIVITY_DESC_COL)
                    outRow = outRow + 1
                End If
            End If
        End If
    Next libRow

    If Not blockFound Then
        InsertTemplateRow wsT, wsOut, TEMPLATE_SECTION_ROW, outRow
        WriteSectionRow wsOut, outRow, "0.0   " & FormatSectionHeading(scopeName & " - BLOCK NOT FOUND IN ACTIVITY LIBRARY", thisWO, thisDWG)
    End If

End Sub

Private Function PromptForRangeChoice(ByVal sourceRange As Range, ByVal promptText As String, ByVal titleText As String, _
                                      ByVal allowBlank As Boolean, ByVal allowCustomEntry As Boolean) As String

    Dim promptBody As String
    Dim userInput As String
    Dim selectedValue As String

    promptBody = promptText & vbCrLf & vbCrLf & BuildChoicePrompt(sourceRange)
    userInput = InputBox(promptBody, titleText)
    userInput = Trim(userInput)

    If Len(userInput) = 0 Then
        If allowBlank Then PromptForRangeChoice = ""
        Exit Function
    End If

    selectedValue = ResolveRangeChoice(sourceRange, userInput)
    If Len(selectedValue) > 0 Then
        PromptForRangeChoice = selectedValue
    ElseIf allowCustomEntry Then
        PromptForRangeChoice = userInput
    Else
        MsgBox "Choose a value from the listed scope blocks.", vbExclamation, "Elite ITP Builder"
    End If

End Function

Private Function BuildChoicePrompt(ByVal sourceRange As Range) As String

    Dim c As Range
    Dim optionNo As Long
    Dim promptBody As String

    For Each c In sourceRange.Cells
        If Len(Trim(CStr(c.Value))) > 0 Then
            optionNo = optionNo + 1
            promptBody = promptBody & CStr(optionNo) & ". " & Trim(CStr(c.Value)) & vbCrLf
        End If
    Next c

    BuildChoicePrompt = promptBody

End Function

Private Function ResolveRangeChoice(ByVal sourceRange As Range, ByVal userInput As String) As String

    Dim c As Range
    Dim optionNo As Long

    If IsNumeric(userInput) Then
        For Each c In sourceRange.Cells
            If Len(Trim(CStr(c.Value))) > 0 Then
                optionNo = optionNo + 1
                If optionNo = CLng(userInput) Then
                    ResolveRangeChoice = Trim(CStr(c.Value))
                    Exit Function
                End If
            End If
        Next c
    End If

    For Each c In sourceRange.Cells
        If StrComp(Trim(CStr(c.Value)), userInput, vbTextCompare) = 0 Then
            ResolveRangeChoice = Trim(CStr(c.Value))
            Exit Function
        End If
    Next c

End Function

Private Sub RenumberGeneratedITPInSheet(ByVal wsOut As Worksheet, ByVal showMessage As Boolean)

    Dim notesRow As Long
    Dim lastRow As Long
    Dim rowNum As Long
    Dim sectionNo As Long
    Dim itemNo As Long
    Dim cellText As String
    Dim headingText As String

    notesRow = FindRowContains(wsOut, "Notes / Comments", 1, 1, 250)
    If notesRow = 0 Then
        lastRow = wsOut.Cells(wsOut.Rows.Count, "A").End(xlUp).Row
    Else
        lastRow = notesRow - 1
    End If

    If lastRow < FIRST_DATA_ROW Then
        If showMessage Then
            MsgBox "No generated ITP activities were found to renumber.", vbInformation, "Elite ITP Builder"
        End If
        Exit Sub
    End If

    sectionNo = 0
    itemNo = 0

    For rowNum = FIRST_DATA_ROW To lastRow
        cellText = Trim(CStr(wsOut.Range("A" & rowNum).Value))

        If IsGeneratedSectionRow(cellText) Then
            sectionNo = sectionNo + 1
            itemNo = 0
            headingText = GetSectionHeadingText(cellText)
            RewriteSectionHeadingPreservingFormat wsOut.Range("A" & rowNum), _
                                                  CStr(sectionNo) & ".0   " & headingText
        ElseIf IsGeneratedActivityRow(cellText) Then
            If sectionNo > 0 Then
                itemNo = itemNo + 1
                wsOut.Range("A" & rowNum).NumberFormat = "@"
                wsOut.Range("A" & rowNum).Value = CStr(sectionNo) & "." & CStr(itemNo)
            End If
        End If
    Next rowNum

    If showMessage Then
        MsgBox "Generated ITP renumbered successfully.", vbInformation, "Elite ITP Builder"
    End If

End Sub

Private Sub RewriteSectionHeadingPreservingFormat(ByVal targetCell As Range, ByVal newText As String)

    Dim destCell As Range
    Dim oldText As String
    Dim oldTitle As String
    Dim newTitle As String
    Dim oldStart As Long
    Dim newStart As Long
    Dim titleLen As Long
    Dim i As Long
    Dim boldArr() As Variant
    Dim italicArr() As Variant
    Dim underlineArr() As Variant
    Dim colorArr() As Variant

    If targetCell Is Nothing Then Exit Sub

    If targetCell.MergeCells Then
        Set destCell = targetCell.MergeArea.Cells(1, 1)
    Else
        Set destCell = targetCell
    End If

    oldText = CStr(destCell.Value)
    oldTitle = GetSectionHeadingText(oldText)
    newTitle = GetSectionHeadingText(newText)
    titleLen = Len(oldTitle)

    If titleLen > 0 And StrComp(oldTitle, newTitle, vbBinaryCompare) = 0 Then
        oldStart = InStr(1, oldText, oldTitle, vbBinaryCompare)
        If oldStart = 0 Then oldStart = InStr(1, oldText, oldTitle, vbTextCompare)
        If oldStart > 0 Then
            ReDim boldArr(1 To titleLen)
            ReDim italicArr(1 To titleLen)
            ReDim underlineArr(1 To titleLen)
            ReDim colorArr(1 To titleLen)

            On Error Resume Next
            For i = 1 To titleLen
                With destCell.Characters(oldStart + i - 1, 1).Font
                    boldArr(i) = .Bold
                    italicArr(i) = .Italic
                    underlineArr(i) = .Underline
                    colorArr(i) = .Color
                End With
            Next i
            On Error GoTo 0
        Else
            titleLen = 0
        End If
    Else
        titleLen = 0
    End If

    destCell.Value = newText

    If titleLen = 0 Then Exit Sub

    newStart = InStr(1, newText, newTitle, vbBinaryCompare)
    If newStart = 0 Then newStart = InStr(1, newText, newTitle, vbTextCompare)
    If newStart = 0 Then Exit Sub

    On Error Resume Next
    For i = 1 To titleLen
        With destCell.Characters(newStart + i - 1, 1).Font
            If Not IsNull(boldArr(i)) Then .Bold = boldArr(i)
            If Not IsNull(italicArr(i)) Then .Italic = italicArr(i)
            If Not IsNull(underlineArr(i)) Then .Underline = underlineArr(i)
            If Not IsNull(colorArr(i)) Then .Color = colorArr(i)
        End With
    Next i
    On Error GoTo 0

End Sub

Private Sub RefreshGeneratedITPRowButtons(ByVal wsOut As Worksheet)

    Dim shapeIndex As Long
    Dim shapeName As String
    Dim notesRow As Long
    Dim lastRow As Long
    Dim rowNum As Long
    Dim cellText As String

    For shapeIndex = wsOut.Shapes.Count To 1 Step -1
        shapeName = wsOut.Shapes(shapeIndex).Name
          If shapeName Like "ITP_Delete_Row_*" Or shapeName Like "ITP_Add_Row_*" Or _
              shapeName Like "ITP_Move_Up_Row_*" Or shapeName Like "ITP_Move_Down_Row_*" Or _
              shapeName = "ITP_Refresh_Rows_Button" Or shapeName = "ITP_Update_Revision_Button" Or _
              shapeName = "ITP_Finalize_Button" Or shapeName = "ITP_Add_Client_Logo_Button" Or _
              shapeName = "ITP_Clear_Client_Logo_Button" Or shapeName = "ITP_Add_Elite_Logo_Button" Or _
              shapeName = "ITP_Clear_Elite_Logo_Button" Then
            wsOut.Shapes(shapeIndex).Delete
        End If
    Next shapeIndex

    notesRow = FindRowContains(wsOut, "Notes / Comments", 1, 1, 250)
    If notesRow = 0 Then
        lastRow = wsOut.Cells(wsOut.Rows.Count, "A").End(xlUp).Row
    Else
        lastRow = notesRow - 1
    End If

    If lastRow < FIRST_DATA_ROW Then Exit Sub

    wsOut.Columns(GetPrintSafetyColumn(wsOut)).ColumnWidth = wsOut.Columns(GetPrintLastColumn(wsOut)).ColumnWidth
    wsOut.Columns(GetDeleteButtonColumn(wsOut)).ColumnWidth = 8.5
    wsOut.Columns(GetAddButtonColumn(wsOut)).ColumnWidth = 8.5
    wsOut.Columns(GetMoveUpButtonColumn(wsOut)).ColumnWidth = 8.5
    wsOut.Columns(GetMoveDownButtonColumn(wsOut)).ColumnWidth = 8.5

    AddGeneratedITPRefreshButton wsOut
    AddGeneratedITPRevisionButton wsOut
    AddGeneratedITPFinalizeButton wsOut
    AddGeneratedITPClientLogoButtons wsOut

    For rowNum = FIRST_DATA_ROW To lastRow
        cellText = Trim(CStr(wsOut.Range("A" & rowNum).Value))
        If Len(cellText) > 0 Then
            AddGeneratedITPRowDeleteButton wsOut, rowNum
            If IsGeneratedSectionRow(cellText) Then
                AddGeneratedITPChapterAddButton wsOut, rowNum
                AddGeneratedITPChapterMoveButton wsOut, rowNum, True
                AddGeneratedITPChapterMoveButton wsOut, rowNum, False
            ElseIf IsGeneratedActivityRow(cellText) Then
                AddGeneratedITPRowAddButton wsOut, rowNum
                AddGeneratedITPActivityMoveButton wsOut, rowNum, True
                AddGeneratedITPActivityMoveButton wsOut, rowNum, False
            End If
        End If
    Next rowNum

End Sub

Private Sub FreezeGeneratedITPWorksheet(ByVal wsOut As Worksheet)

    If wsOut Is Nothing Then Exit Sub

    On Error Resume Next
    wsOut.Unprotect
    On Error GoTo 0

    RemoveGeneratedITPInteractiveShapes wsOut
    RenameWorksheetForArchivedState wsOut

    wsOut.Cells.Locked = True
    wsOut.Columns(GetDeleteButtonColumn(wsOut) & ":" & GetMoveDownButtonColumn(wsOut)).Hidden = True
    wsOut.Tab.Color = RGB(166, 166, 166)
    ClearGeneratedITPSheetMarker wsOut

    wsOut.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    wsOut.EnableSelection = xlNoSelection

End Sub

Private Sub RemoveGeneratedITPInteractiveShapes(ByVal wsOut As Worksheet)

    Dim shapeIndex As Long
    Dim shapeName As String
    Dim actionName As String

    For shapeIndex = wsOut.Shapes.Count To 1 Step -1
        shapeName = wsOut.Shapes(shapeIndex).Name
        On Error Resume Next
        actionName = CStr(wsOut.Shapes(shapeIndex).OnAction)
        On Error GoTo 0

          If shapeName Like "ITP_Delete_Row_*" Or shapeName Like "ITP_Add_Row_*" Or _
              shapeName Like "ITP_Move_Up_Row_*" Or shapeName Like "ITP_Move_Down_Row_*" Or _
              shapeName = "ITP_Refresh_Rows_Button" Or _
              actionName = "DeleteGeneratedITPSelectedRow" Or actionName = "AddGeneratedITPBlankRow" Or _
              actionName = "MoveGeneratedITPActivityUp" Or actionName = "MoveGeneratedITPActivityDown" Or _
              actionName = "AddGeneratedITPChapterBlock" Or actionName = "RenumberGeneratedITP" Or _
              shapeName = "ITP_Update_Revision_Button" Or actionName = "UpdateGeneratedITPRevision" Or _
              shapeName = "ITP_Finalize_Button" Or actionName = "FinalizeGeneratedITP" Or _
              shapeName = "ITP_Add_Client_Logo_Button" Or actionName = "AddClientHeaderLogo" Or _
              shapeName = "ITP_Clear_Client_Logo_Button" Or actionName = "ClearClientHeaderLogo" Or _
              shapeName = "ITP_Add_Elite_Logo_Button" Or actionName = "AddEliteHeaderLogo" Or _
              shapeName = "ITP_Clear_Elite_Logo_Button" Or actionName = "ClearEliteHeaderLogo" Then
            wsOut.Shapes(shapeIndex).Delete
        End If
    Next shapeIndex

End Sub

Private Sub AddGeneratedITPRefreshButton(ByVal ws As Worksheet)

    Dim targetRange As Range
    Dim shp As Shape
    Dim buttonWidth As Double
    Dim buttonHeight As Double

    Set targetRange = ws.Range(GetDeleteButtonColumn(ws) & "9:" & GetAddButtonColumn(ws) & "9")
    buttonWidth = targetRange.Width - (REFRESH_BUTTON_MARGIN_X * 2)
    buttonHeight = targetRange.Height - (REFRESH_BUTTON_MARGIN_Y * 2)

    If buttonWidth <= 0 Or buttonHeight <= 0 Then Exit Sub

    On Error Resume Next
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, targetRange.Left + REFRESH_BUTTON_MARGIN_X, targetRange.Top + REFRESH_BUTTON_MARGIN_Y, _
                                 buttonWidth, buttonHeight)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub

    shp.Name = "ITP_Refresh_Rows_Button"
    shp.TextFrame.Characters.Text = "REFRESH ORDER"
    shp.Fill.ForeColor.RGB = RGB(112, 48, 160)
    shp.Line.ForeColor.RGB = RGB(80, 35, 120)
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.Characters.Font.Bold = True
    shp.TextFrame.Characters.Font.Size = 9
    shp.OnAction = "RenumberGeneratedITP"
    shp.Placement = xlMoveAndSize
    DisableShapePrinting shp

End Sub

Private Sub AddGeneratedITPRevisionButton(ByVal ws As Worksheet)

    Dim targetRange As Range
    Dim shp As Shape
    Dim buttonWidth As Double
    Dim buttonHeight As Double

    Set targetRange = ws.Range(GetDeleteButtonColumn(ws) & "2:" & GetAddButtonColumn(ws) & "2")
    buttonWidth = targetRange.Width - (REVISION_BUTTON_MARGIN_X * 2)
    buttonHeight = targetRange.Height - (REVISION_BUTTON_MARGIN_Y * 2)

    If buttonWidth <= 0 Or buttonHeight <= 0 Then Exit Sub

    On Error Resume Next
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, targetRange.Left + REVISION_BUTTON_MARGIN_X, targetRange.Top + REVISION_BUTTON_MARGIN_Y, _
                                 buttonWidth, buttonHeight)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub

    shp.Name = "ITP_Update_Revision_Button"
    shp.TextFrame.Characters.Text = "UPDATE REV"
    shp.Fill.ForeColor.RGB = RGB(191, 144, 0)
    shp.Line.ForeColor.RGB = RGB(127, 96, 0)
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.Characters.Font.Bold = True
    shp.TextFrame.Characters.Font.Size = 9
    shp.OnAction = "UpdateGeneratedITPRevision"
    shp.Placement = xlMoveAndSize
    DisableShapePrinting shp

End Sub

Private Sub AddGeneratedITPFinalizeButton(ByVal ws As Worksheet)

    Dim targetRange As Range
    Dim shp As Shape
    Dim buttonWidth As Double
    Dim buttonHeight As Double

    Set targetRange = ws.Range(GetDeleteButtonColumn(ws) & "1:" & GetAddButtonColumn(ws) & "1")
    buttonWidth = targetRange.Width - (REVISION_BUTTON_MARGIN_X * 2)
    buttonHeight = targetRange.Height - (REVISION_BUTTON_MARGIN_Y * 2)

    If buttonWidth <= 0 Or buttonHeight <= 0 Then Exit Sub

    On Error Resume Next
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, targetRange.Left + REVISION_BUTTON_MARGIN_X, targetRange.Top + REVISION_BUTTON_MARGIN_Y, _
                                 buttonWidth, buttonHeight)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub

    shp.Name = "ITP_Finalize_Button"
    shp.TextFrame.Characters.Text = "FINALIZE"
    shp.Fill.ForeColor.RGB = RGB(0, 176, 80)
    shp.Line.ForeColor.RGB = RGB(0, 97, 46)
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.Characters.Font.Bold = True
    shp.TextFrame.Characters.Font.Size = 9
    shp.OnAction = "FinalizeGeneratedITP"
    shp.Placement = xlMoveAndSize
    DisableShapePrinting shp

End Sub

Private Sub AddGeneratedITPClientLogoButtons(ByVal ws As Worksheet)

    If Not IsITPWorksheetForHeader(ws) Then Exit Sub

    ' Client logo UI disabled for now (code kept for possible restore):
    ' AddGeneratedITPClientLogoButton ws, "ITP_Add_Client_Logo_Button", "ADD CLIENT LOGO", 3, _
    '                                 RGB(237, 125, 49), RGB(197, 90, 17), "AddClientHeaderLogo"
    ' AddGeneratedITPClientLogoButton ws, "ITP_Clear_Client_Logo_Button", "CLEAR CLIENT LOGO", 4, _
    '                                 RGB(127, 127, 127), RGB(89, 89, 89), "ClearClientHeaderLogo"
    AddGeneratedITPClientLogoButton ws, "ITP_Add_Elite_Logo_Button", "ADD LOGO", 3, _
                                    RGB(237, 125, 49), RGB(197, 90, 17), "AddEliteHeaderLogo"
    AddGeneratedITPClientLogoButton ws, "ITP_Clear_Elite_Logo_Button", "CLEAR LOGO", 4, _
                                    RGB(127, 127, 127), RGB(89, 89, 89), "ClearEliteHeaderLogo"

End Sub

Private Sub AddGeneratedITPClientLogoButton(ByVal ws As Worksheet, ByVal buttonName As String, _
                                            ByVal buttonText As String, ByVal targetRow As Long, _
                                            ByVal fillColor As Long, ByVal lineColor As Long, _
                                            ByVal macroName As String)

    Dim targetRange As Range
    Dim shp As Shape
    Dim buttonWidth As Double
    Dim buttonHeight As Double

    Set targetRange = ws.Range(GetDeleteButtonColumn(ws) & targetRow & ":" & GetAddButtonColumn(ws) & targetRow)
    buttonWidth = targetRange.Width - (REVISION_BUTTON_MARGIN_X * 2)
    buttonHeight = targetRange.Height - (REVISION_BUTTON_MARGIN_Y * 2)

    If buttonWidth <= 0 Or buttonHeight <= 0 Then Exit Sub

    On Error Resume Next
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, targetRange.Left + REVISION_BUTTON_MARGIN_X, targetRange.Top + REVISION_BUTTON_MARGIN_Y, _
                                 buttonWidth, buttonHeight)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub

    shp.Name = buttonName
    shp.TextFrame.Characters.Text = buttonText
    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.ForeColor.RGB = lineColor
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.Characters.Font.Bold = True
    shp.TextFrame.Characters.Font.Size = 9
    shp.OnAction = macroName
    shp.Placement = xlMoveAndSize
    DisableShapePrinting shp

End Sub

Public Sub UpdateGeneratedITPRevision()

    Dim wsCurrent As Worksheet
    Dim wsNew As Worksheet
    Dim wsIn As Worksheet
    Dim newRevision As String
    Dim oldRevision As String
    Dim desiredSheetName As String
    Dim revisionPrompt As String
    Dim archivedName As String
    Dim alreadyFinalized As Boolean
    Dim updateCompleted As Boolean

    On Error Resume Next
    Set wsCurrent = GetGeneratedITPWorksheet()
    Set wsIn = ThisWorkbook.Worksheets("Inputs")
    On Error GoTo 0

    If wsCurrent Is Nothing Or wsIn Is Nothing Then
        MsgBox "Could not find the current generated ITP or Inputs sheet.", vbExclamation, "Elite ITP Builder"
        Exit Sub
    End If

    oldRevision = Trim$(CStr(wsIn.Range("E6").Value))
    revisionPrompt = "Enter the new revision number for this ITP copy."
    If Len(oldRevision) > 0 Then revisionPrompt = revisionPrompt & vbCrLf & vbCrLf & "Current revision: " & oldRevision

    newRevision = Trim$(InputBox(revisionPrompt, "Elite ITP Builder", oldRevision))
    If Len(newRevision) = 0 Then Exit Sub

    If StrComp(newRevision, oldRevision, vbTextCompare) = 0 Then
        MsgBox "Enter a revision number that is different from the current one.", vbExclamation, "Elite ITP Builder"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    wsCurrent.Copy After:=wsCurrent
    Set wsNew = ActiveSheet

    wsIn.Range("E6").Value = newRevision
    desiredSheetName = BuildGeneratedITPSheetName(wsIn)

    MarkGeneratedITPSheet wsNew
    RenameGeneratedITPSheet wsNew, wsIn
    ApplyHeaderFromInputs wsNew, wsIn
    ApplyTemplateInputValues wsNew, wsIn
    ApplyRevisionValueToGeneratedSheet wsNew, newRevision
    RefreshGeneratedITPRowButtons wsNew

    archivedName = BuildArchivedWorksheetName(wsCurrent.Name)
    alreadyFinalized = WorksheetNameExists(archivedName, wsCurrent)

    If alreadyFinalized Then
        ClearGeneratedITPSheetMarker wsCurrent
        wsCurrent.Delete
        Set wsCurrent = Nothing
    Else
        FreezeGeneratedITPWorksheet wsCurrent
    End If

    wsNew.Activate
    wsNew.DisplayPageBreaks = True
    ActiveWindow.View = xlPageBreakPreview
    ActiveWindow.Zoom = 100
    ActiveWindow.DisplayGridlines = False
    Application.GoTo wsNew.Range("A1"), True

    updateCompleted = True
    If alreadyFinalized Then
        MsgBox "Revision updated successfully." & vbCrLf & vbCrLf & _
               "Existing finalized revision kept: " & archivedName & vbCrLf & _
               "Active revision: " & wsNew.Name, vbInformation, "Elite ITP Builder"
    Else
        MsgBox "Revision updated successfully." & vbCrLf & vbCrLf & _
               "Archived revision: " & wsCurrent.Name & vbCrLf & _
               "Active revision: " & wsNew.Name, vbInformation, "Elite ITP Builder"
    End If

SafeExit:
    On Error Resume Next
    If Not updateCompleted Then
        wsIn.Range("E6").Value = oldRevision
        If Not wsNew Is Nothing Then wsNew.Delete
    End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

Public Sub FinalizeGeneratedITP()

    Dim wsCurrent As Worksheet
    Dim wsArchived As Worksheet
    Dim archivedBaseName As String
    Dim archivedName As String
    Dim duplicateArchiveResponse As VbMsgBoxResult
    Dim finalizeCompleted As Boolean

    On Error Resume Next
    Set wsCurrent = GetGeneratedITPWorksheet()
    On Error GoTo 0

    If wsCurrent Is Nothing Then
        MsgBox "Could not find the current generated ITP tab to finalize.", vbExclamation, "Elite ITP Builder"
        Exit Sub
    End If

    archivedBaseName = BuildArchivedWorksheetName(wsCurrent.Name)

    If WorksheetNameExists(archivedBaseName, Nothing) Then
        duplicateArchiveResponse = MsgBox("A finalized archive already exists with this name:" & vbCrLf & vbCrLf & _
                                          archivedBaseName & vbCrLf & vbCrLf & _
                                          "Yes = create another finalized copy with a numbered suffix." & vbCrLf & _
                                          "No = update the revision instead." & vbCrLf & _
                                          "Cancel = stop.", _
                                          vbQuestion + vbYesNoCancel + vbDefaultButton2, "Elite ITP Builder")

        If duplicateArchiveResponse = vbCancel Then Exit Sub
        If duplicateArchiveResponse = vbNo Then
            UpdateGeneratedITPRevision
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    wsCurrent.Copy After:=wsCurrent
    Set wsArchived = ActiveSheet

    archivedName = EnsureUniqueWorksheetName(archivedBaseName, wsArchived)
    On Error Resume Next
    wsArchived.Name = archivedName
    On Error GoTo 0

    FreezeGeneratedITPWorksheet wsArchived

    wsCurrent.Activate
    Application.GoTo wsCurrent.Range("A1"), True

    finalizeCompleted = True
    MsgBox "Finalized copy created successfully." & vbCrLf & vbCrLf & _
           "Working tab: " & wsCurrent.Name & vbCrLf & _
           "Finalized tab: " & wsArchived.Name, vbInformation, "Elite ITP Builder"

SafeExit:
    On Error Resume Next
    If Not finalizeCompleted Then
        If Not wsArchived Is Nothing Then wsArchived.Delete
    End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

Private Sub ApplyRevisionValueToGeneratedSheet(ByVal wsOut As Worksheet, ByVal revisionText As String)

    SetCellBesideLabel wsOut, "Revision", revisionText
    wsOut.PageSetup.LeftFooter = FormatRevisionFooterText(revisionText)

End Sub

Private Function GetGeneratedITPButtonRow(ByVal wsOut As Worksheet) As Long

    Dim shapeName As String
    Dim shp As Shape

    On Error Resume Next
    shapeName = CStr(Application.Caller)
    If Len(shapeName) > 0 Then Set shp = wsOut.Shapes(shapeName)
    On Error GoTo 0

    If shp Is Nothing Then Exit Function

    ' The top edge of a short section-row button falls into the row above. Use the center.
    GetGeneratedITPButtonRow = RowAtVerticalPosition(wsOut, shp.Top + (shp.Height / 2))

End Function

Private Function RowAtVerticalPosition(ByVal ws As Worksheet, ByVal verticalPosition As Double) As Long

    Dim rowNum As Long
    Dim rowTop As Double
    Dim rowHeight As Double
    Dim lastRow As Long

    If verticalPosition < 0 Then Exit Function

    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count + 5
    If lastRow < 1 Then lastRow = 1
    If lastRow > ws.Rows.Count Then lastRow = ws.Rows.Count

    For rowNum = 1 To lastRow
        rowTop = ws.Rows(rowNum).Top
        rowHeight = ws.Rows(rowNum).Height
        If rowHeight > 0 Then
            If verticalPosition >= rowTop And verticalPosition < (rowTop + rowHeight) Then
                RowAtVerticalPosition = rowNum
                Exit Function
            End If
        End If
        If rowTop > verticalPosition Then Exit Function
    Next rowNum

End Function

Private Sub SetRowButtonBounds(ByVal targetCell As Range, ByRef buttonLeft As Double, ByRef buttonTop As Double, _
                               ByRef buttonWidth As Double, ByRef buttonHeight As Double)

    Dim marginY As Double

    marginY = BUTTON_MARGIN_Y
    If targetCell.Height <= 18 Then marginY = 0.6

    buttonLeft = targetCell.Left + BUTTON_MARGIN_X
    buttonWidth = targetCell.Width - (BUTTON_MARGIN_X * 2)
    buttonTop = targetCell.Top + marginY
    buttonHeight = targetCell.Height - (marginY * 2)

End Sub

Private Sub AddGeneratedITPRowDeleteButton(ByVal ws As Worksheet, ByVal rowNum As Long)

    Dim targetCell As Range
    Dim shp As Shape
    Dim buttonLeft As Double
    Dim buttonTop As Double
    Dim buttonWidth As Double
    Dim buttonHeight As Double

    Set targetCell = ws.Range(GetDeleteButtonColumn(ws) & rowNum)
    If targetCell.MergeCells Then Exit Sub

    SetRowButtonBounds targetCell, buttonLeft, buttonTop, buttonWidth, buttonHeight

    If buttonWidth <= 0 Or buttonHeight <= 0 Then Exit Sub

    On Error Resume Next
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, buttonLeft, buttonTop, buttonWidth, buttonHeight)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub

    shp.Name = "ITP_Delete_Row_" & CStr(rowNum)
    shp.TextFrame.Characters.Text = "-"
    shp.Fill.ForeColor.RGB = RGB(192, 0, 0)
    shp.Line.ForeColor.RGB = RGB(128, 0, 0)
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.Characters.Font.Bold = True
    shp.TextFrame.Characters.Font.Size = 8
    shp.OnAction = "DeleteGeneratedITPSelectedRow"
    shp.Placement = xlMoveAndSize
    DisableShapePrinting shp

End Sub

Private Sub AddGeneratedITPRowAddButton(ByVal ws As Worksheet, ByVal rowNum As Long)

    Dim targetCell As Range
    Dim shp As Shape
    Dim buttonLeft As Double
    Dim buttonTop As Double
    Dim buttonWidth As Double
    Dim buttonHeight As Double

    Set targetCell = ws.Range(GetAddButtonColumn(ws) & rowNum)
    If targetCell.MergeCells Then Exit Sub

    SetRowButtonBounds targetCell, buttonLeft, buttonTop, buttonWidth, buttonHeight

    If buttonWidth <= 0 Or buttonHeight <= 0 Then Exit Sub

    On Error Resume Next
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, buttonLeft, buttonTop, buttonWidth, buttonHeight)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub

    shp.Name = "ITP_Add_Row_" & CStr(rowNum)
    shp.TextFrame.Characters.Text = "+"
    shp.Fill.ForeColor.RGB = RGB(0, 112, 192)
    shp.Line.ForeColor.RGB = RGB(0, 78, 134)
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.Characters.Font.Bold = True
    shp.TextFrame.Characters.Font.Size = 9
    shp.OnAction = "AddGeneratedITPBlankRow"
    shp.Placement = xlMoveAndSize
    DisableShapePrinting shp

End Sub

Private Sub AddGeneratedITPChapterAddButton(ByVal ws As Worksheet, ByVal rowNum As Long)

    Dim targetCell As Range
    Dim shp As Shape
    Dim buttonLeft As Double
    Dim buttonTop As Double
    Dim buttonWidth As Double
    Dim buttonHeight As Double

    Set targetCell = ws.Range(GetAddButtonColumn(ws) & rowNum)
    If targetCell.MergeCells Then Exit Sub

    SetRowButtonBounds targetCell, buttonLeft, buttonTop, buttonWidth, buttonHeight

    If buttonWidth <= 0 Or buttonHeight <= 0 Then Exit Sub

    On Error Resume Next
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, buttonLeft, buttonTop, buttonWidth, buttonHeight)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub

    shp.Name = "ITP_Add_Row_" & CStr(rowNum)
    shp.TextFrame.Characters.Text = "+"
    shp.Fill.ForeColor.RGB = RGB(0, 176, 80)
    shp.Line.ForeColor.RGB = RGB(0, 97, 46)
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.Characters.Font.Bold = True
    shp.TextFrame.Characters.Font.Size = 9
    shp.OnAction = "AddGeneratedITPChapterBlock"
    shp.Placement = xlMoveAndSize
    DisableShapePrinting shp

End Sub

Private Sub AddGeneratedITPChapterMoveButton(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal moveUp As Boolean)

    Dim targetCell As Range
    Dim shp As Shape
    Dim buttonLeft As Double
    Dim buttonTop As Double
    Dim buttonWidth As Double
    Dim buttonHeight As Double

    If moveUp Then
        Set targetCell = ws.Range(GetMoveUpButtonColumn(ws) & rowNum)
    Else
        Set targetCell = ws.Range(GetMoveDownButtonColumn(ws) & rowNum)
    End If
    If targetCell.MergeCells Then Exit Sub

    SetRowButtonBounds targetCell, buttonLeft, buttonTop, buttonWidth, buttonHeight

    If buttonWidth <= 0 Or buttonHeight <= 0 Then Exit Sub

    On Error Resume Next
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, buttonLeft, buttonTop, buttonWidth, buttonHeight)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub

    If moveUp Then
        shp.Name = "ITP_Move_Up_Row_" & CStr(rowNum)
        shp.TextFrame.Characters.Text = "UP"
        shp.OnAction = "MoveGeneratedITPChapterUp"
    Else
        shp.Name = "ITP_Move_Down_Row_" & CStr(rowNum)
        shp.TextFrame.Characters.Text = "DN"
        shp.OnAction = "MoveGeneratedITPChapterDown"
    End If
    shp.Fill.ForeColor.RGB = RGB(112, 48, 160)
    shp.Line.ForeColor.RGB = RGB(80, 35, 120)
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.Characters.Font.Bold = True
    shp.TextFrame.Characters.Font.Size = 8
    shp.Placement = xlMoveAndSize
    DisableShapePrinting shp

End Sub

Private Sub AddGeneratedITPActivityMoveButton(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal moveUp As Boolean)

    Dim targetCell As Range
    Dim shp As Shape
    Dim buttonLeft As Double
    Dim buttonTop As Double
    Dim buttonWidth As Double
    Dim buttonHeight As Double

    If moveUp Then
        Set targetCell = ws.Range(GetMoveUpButtonColumn(ws) & rowNum)
    Else
        Set targetCell = ws.Range(GetMoveDownButtonColumn(ws) & rowNum)
    End If
    If targetCell.MergeCells Then Exit Sub

    SetRowButtonBounds targetCell, buttonLeft, buttonTop, buttonWidth, buttonHeight

    If buttonWidth <= 0 Or buttonHeight <= 0 Then Exit Sub

    On Error Resume Next
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, buttonLeft, buttonTop, buttonWidth, buttonHeight)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub

    If moveUp Then
        shp.Name = "ITP_Move_Up_Row_" & CStr(rowNum)
        shp.TextFrame.Characters.Text = "UP"
        shp.OnAction = "MoveGeneratedITPActivityUp"
    Else
        shp.Name = "ITP_Move_Down_Row_" & CStr(rowNum)
        shp.TextFrame.Characters.Text = "DN"
        shp.OnAction = "MoveGeneratedITPActivityDown"
    End If
    shp.Fill.ForeColor.RGB = RGB(89, 89, 89)
    shp.Line.ForeColor.RGB = RGB(64, 64, 64)
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.Characters.Font.Bold = True
    shp.TextFrame.Characters.Font.Size = 8
    shp.Placement = xlMoveAndSize
    DisableShapePrinting shp

End Sub

Private Sub DisableShapePrinting(ByVal shp As Shape)

    On Error Resume Next
    shp.PrintObject = False
    On Error GoTo 0

End Sub

Private Function GetGeneratedSectionEndRow(ByVal wsOut As Worksheet, ByVal sectionRow As Long, ByVal notesRow As Long) As Long

    Dim lastRow As Long
    Dim rowNum As Long
    Dim cellText As String

    If notesRow = 0 Then
        lastRow = wsOut.Cells(wsOut.Rows.Count, "A").End(xlUp).Row
    Else
        lastRow = notesRow - 1
    End If

    GetGeneratedSectionEndRow = sectionRow

    For rowNum = sectionRow + 1 To lastRow
        cellText = Trim(CStr(wsOut.Range("A" & rowNum).Value))
        If IsGeneratedSectionRow(cellText) Then Exit Function
        GetGeneratedSectionEndRow = rowNum
    Next rowNum

End Function

Private Function GetPreviousGeneratedSectionRow(ByVal wsOut As Worksheet, ByVal sectionRow As Long) As Long

    Dim rowNum As Long
    Dim cellText As String

    For rowNum = sectionRow - 1 To FIRST_DATA_ROW Step -1
        cellText = Trim(CStr(wsOut.Range("A" & rowNum).Value))
        If IsGeneratedSectionRow(cellText) Then
            GetPreviousGeneratedSectionRow = rowNum
            Exit Function
        End If
    Next rowNum

End Function

Private Function GetInputChapterLastRow(ByVal wsIn As Worksheet) As Long

    Dim lastRow As Long
    Dim candidateRow As Long
    Dim colNum As Long

    lastRow = INPUT_CHAPTER_LAST_ROW

    For colNum = INPUT_SCOPE_COL To INPUT_CUSTOM_ROWS_COL
        candidateRow = wsIn.Cells(wsIn.Rows.Count, colNum).End(xlUp).Row
        If candidateRow > lastRow Then lastRow = candidateRow
    Next colNum

    GetInputChapterLastRow = lastRow

End Function

Private Function GetOrderedInputChapterRow(ByVal wsIn As Worksheet, ByVal afterSectionNo As Long) As Long

    Dim lastRow As Long
    Dim templateRow As Long
    Dim insertRow As Long

    lastRow = GetInputChapterLastRow(wsIn)
    insertRow = INPUT_CHAPTER_FIRST_ROW + afterSectionNo

    If insertRow < INPUT_CHAPTER_FIRST_ROW Then insertRow = INPUT_CHAPTER_FIRST_ROW

    If insertRow > lastRow Then
        GetOrderedInputChapterRow = GetNextInputChapterRow(wsIn)
        Exit Function
    End If

    If Len(Trim(CStr(wsIn.Cells(insertRow, INPUT_SCOPE_COL).Value))) = 0 Then
        GetOrderedInputChapterRow = insertRow
        Exit Function
    End If

    templateRow = insertRow - 1
    If templateRow < INPUT_CHAPTER_FIRST_ROW Then templateRow = insertRow

    wsIn.Range(wsIn.Cells(insertRow, INPUT_CHAPTER_COL), wsIn.Cells(insertRow, INPUT_CUSTOM_ROWS_COL)).Insert Shift:=xlDown
    wsIn.Range(wsIn.Cells(templateRow, INPUT_CHAPTER_COL), wsIn.Cells(templateRow, INPUT_CUSTOM_ROWS_COL)).Copy
    wsIn.Range(wsIn.Cells(insertRow, INPUT_CHAPTER_COL), wsIn.Cells(insertRow, INPUT_CUSTOM_ROWS_COL)).PasteSpecial Paste:=xlPasteFormats
    On Error Resume Next
    wsIn.Range(wsIn.Cells(insertRow, INPUT_CHAPTER_COL), wsIn.Cells(insertRow, INPUT_CUSTOM_ROWS_COL)).PasteSpecial Paste:=xlPasteValidation
    On Error GoTo 0
    wsIn.Rows(insertRow).RowHeight = wsIn.Rows(templateRow).RowHeight
    Application.CutCopyMode = False
    wsIn.Range(wsIn.Cells(insertRow, INPUT_SCOPE_COL), wsIn.Cells(insertRow, INPUT_CUSTOM_ROWS_COL)).ClearContents
    RefreshInputScopeValidation wsIn

    GetOrderedInputChapterRow = insertRow

End Function

Private Function GetNextInputChapterRow(ByVal wsIn As Worksheet) As Long

    Dim rowNum As Long
    Dim lastRow As Long
    Dim templateRow As Long

    lastRow = GetInputChapterLastRow(wsIn)

    For rowNum = INPUT_CHAPTER_FIRST_ROW To lastRow
        If Len(Trim(CStr(wsIn.Cells(rowNum, INPUT_SCOPE_COL).Value))) = 0 Then
            GetNextInputChapterRow = rowNum
            Exit Function
        End If
    Next rowNum

    GetNextInputChapterRow = lastRow + 1
    templateRow = WorksheetFunction.Max(INPUT_CHAPTER_LAST_ROW, lastRow)

    wsIn.Range(wsIn.Cells(templateRow, INPUT_CHAPTER_COL), wsIn.Cells(templateRow, INPUT_CUSTOM_ROWS_COL)).Copy
    wsIn.Range(wsIn.Cells(GetNextInputChapterRow, INPUT_CHAPTER_COL), wsIn.Cells(GetNextInputChapterRow, INPUT_CUSTOM_ROWS_COL)).PasteSpecial Paste:=xlPasteFormats
    On Error Resume Next
    wsIn.Range(wsIn.Cells(GetNextInputChapterRow, INPUT_CHAPTER_COL), wsIn.Cells(GetNextInputChapterRow, INPUT_CUSTOM_ROWS_COL)).PasteSpecial Paste:=xlPasteValidation
    On Error GoTo 0
    wsIn.Rows(GetNextInputChapterRow).RowHeight = wsIn.Rows(templateRow).RowHeight
    Application.CutCopyMode = False
    wsIn.Range(wsIn.Cells(GetNextInputChapterRow, INPUT_SCOPE_COL), wsIn.Cells(GetNextInputChapterRow, INPUT_CUSTOM_ROWS_COL)).ClearContents
    RefreshInputScopeValidation wsIn

End Function

Private Sub EnsureValueInInputList(ByVal wsIn As Worksheet, ByVal colLetter As String, ByVal firstRow As Long, ByVal valueText As String)

    Dim targetRange As Range
    Dim c As Range
    Dim blankCell As Range

    valueText = Trim(valueText)
    If Len(valueText) = 0 Then Exit Sub

    Set targetRange = GetDynamicInputListRange(wsIn, colLetter, firstRow)

    For Each c In targetRange.Cells
        If StrComp(Trim(CStr(c.Value)), valueText, vbTextCompare) = 0 Then Exit Sub
        If blankCell Is Nothing Then
            If Len(Trim(CStr(c.Value))) = 0 Then Set blankCell = c
        End If
    Next c

    If blankCell Is Nothing Then Exit Sub

    blankCell.Value = valueText

End Sub

Private Sub UpdateInputChapterSerials(ByVal wsIn As Worksheet)

    Dim rowNum As Long
    Dim lastRow As Long

    lastRow = GetInputChapterLastRow(wsIn)

    For rowNum = INPUT_CHAPTER_FIRST_ROW To lastRow
        wsIn.Cells(rowNum, INPUT_CHAPTER_COL).Value = rowNum - INPUT_CHAPTER_FIRST_ROW + 1
    Next rowNum

End Sub

Public Sub RefreshScopeList()

    Dim wsLib As Worksheet
    Dim wsIn As Worksheet
    Dim wsLists As Worksheet
    Dim dAll As Object
    Dim dShop As Object
    Dim dField As Object
    Dim scopeLocations As Object
    Dim lastLib As Long
    Dim r As Long
    Dim k As Variant
    Dim v As String
    Dim locationText As String
    Dim locationCol As Long
    Dim lastAllRow As Long
    Dim lastShopRow As Long
    Dim lastFieldRow As Long

    Set wsIn = ThisWorkbook.Worksheets("Inputs")
    Set wsLib = ThisWorkbook.Worksheets("Activity Library")
    Set wsLists = ThisWorkbook.Worksheets("Lists")
    Set dAll = CreateObject("Scripting.Dictionary")
    Set dShop = CreateObject("Scripting.Dictionary")
    Set dField = CreateObject("Scripting.Dictionary")
    Set scopeLocations = CreateObject("Scripting.Dictionary")
    dAll.CompareMode = vbTextCompare
    dShop.CompareMode = vbTextCompare
    dField.CompareMode = vbTextCompare
    scopeLocations.CompareMode = vbTextCompare
    EnsureActivityLibraryLocationColumn wsLib
    EnsureActivityLibraryScopeBlocks wsLib
    EnsureActivityLibraryTableCoversData wsLib

    lastLib = GetActivityLibraryLastDataRow(wsLib)
    locationCol = FindActivityLibraryHeaderColumn(wsLib, LIB_LOCATION_HEADER)

    For r = 2 To lastLib
        v = Trim(CStr(wsLib.Cells(r, "A").Value))
        If Len(v) > 0 Then
            If UCase$(v) <> "FINAL ACCEPTANCE" Then
                dAll(v) = True

                If locationCol > 0 Then
                    locationText = Trim$(CStr(wsLib.Cells(r, locationCol).Value))
                Else
                    locationText = ""
                End If

                If Not scopeLocations.Exists(v) Then scopeLocations(v) = ""
                scopeLocations(v) = MergeScopeLocationValues(CStr(scopeLocations(v)), locationText)
            End If
        End If
    Next r

    For Each k In dAll.Keys
        locationText = ""
        If scopeLocations.Exists(CStr(k)) Then locationText = CStr(scopeLocations(k))

        If ScopeMatchesInspectionType(locationText, INPUT_INSPECTION_TYPE_SHOP) Then dShop(CStr(k)) = True
        If ScopeMatchesInspectionType(locationText, INPUT_INSPECTION_TYPE_FIELD) Then dField(CStr(k)) = True
    Next k

    wsLists.Range(LISTS_SCOPE_ALL_COL & "2:" & LISTS_SCOPE_FIELD_COL & wsLists.Rows.Count).ClearContents

    lastAllRow = WriteScopeListColumn(wsLists, LISTS_SCOPE_ALL_COL, dAll)
    lastShopRow = WriteScopeListColumn(wsLists, LISTS_SCOPE_SHOP_COL, dShop)
    lastFieldRow = WriteScopeListColumn(wsLists, LISTS_SCOPE_FIELD_COL, dField)

    UpdateScopeBlockListName wsLists, lastAllRow
    UpdateNamedScopeList wsLists, "ScopeBlockListShop", LISTS_SCOPE_SHOP_COL, lastShopRow
    UpdateNamedScopeList wsLists, "ScopeBlockListField", LISTS_SCOPE_FIELD_COL, lastFieldRow
    RefreshInputScopeValidation wsIn

End Sub

Private Function WriteScopeListColumn(ByVal wsLists As Worksheet, ByVal colLetter As String, ByVal scopeDict As Object) As Long

    Dim k As Variant
    Dim outRow As Long

    outRow = 2

    For Each k In scopeDict.Keys
        If UCase$(CStr(k)) <> "CUSTOM" Then
            wsLists.Cells(outRow, colLetter).Value = CStr(k)
            outRow = outRow + 1
        End If
    Next k

    If outRow > 2 Then
        wsLists.Range(colLetter & "2:" & colLetter & outRow - 1).Sort Key1:=wsLists.Range(colLetter & "2"), Order1:=xlAscending, Header:=xlNo
    End If

    wsLists.Cells(outRow, colLetter).Value = "Custom"
    WriteScopeListColumn = outRow

End Function

Public Sub RefreshInputScopeValidation(Optional ByVal wsIn As Worksheet = Nothing)

    Dim lastRow As Long
    Dim targetRange As Range
    Dim listName As String

    If wsIn Is Nothing Then Set wsIn = ThisWorkbook.Worksheets("Inputs")

    lastRow = GetInputChapterLastRow(wsIn)
    If lastRow < INPUT_CHAPTER_LAST_ROW Then lastRow = INPUT_CHAPTER_LAST_ROW

    Set targetRange = wsIn.Range(wsIn.Cells(INPUT_CHAPTER_FIRST_ROW, INPUT_SCOPE_COL), _
                                 wsIn.Cells(lastRow, INPUT_SCOPE_COL))

    listName = GetScopeBlockListNameForInspectionType(GetInspectionTypeValue(wsIn))

    On Error Resume Next
    targetRange.Validation.Delete
    On Error GoTo 0

    With targetRange.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="=" & listName
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Scope Block"
        .InputMessage = "Choose a scope block that matches the selected inspection type."
        .ErrorTitle = "Invalid Scope Block"
        .ErrorMessage = "Select a scope block from the filtered list."
        .ShowInput = True
        .ShowError = True
    End With

End Sub

Public Function GetScopeBlockRangeForCurrentInspection(Optional ByVal wsIn As Worksheet = Nothing, _
                                                       Optional ByVal wsLists As Worksheet = Nothing) As Range

    Dim listName As String

    If wsIn Is Nothing Then Set wsIn = ThisWorkbook.Worksheets("Inputs")
    If wsLists Is Nothing Then Set wsLists = ThisWorkbook.Worksheets("Lists")

    listName = GetScopeBlockListNameForInspectionType(GetInspectionTypeValue(wsIn))

    On Error Resume Next
    Set GetScopeBlockRangeForCurrentInspection = ThisWorkbook.Names(listName).RefersToRange
    On Error GoTo 0

    If GetScopeBlockRangeForCurrentInspection Is Nothing Then
        Set GetScopeBlockRangeForCurrentInspection = ThisWorkbook.Names("ScopeBlockList").RefersToRange
    End If

End Function

Private Function GetInspectionTypeValue(ByVal wsIn As Worksheet) As String

    GetInspectionTypeValue = UCase$(Trim$(CStr(wsIn.Range(INPUT_INSPECTION_TYPE_CELL).Value)))

End Function

Private Function ScopeMatchesInspectionType(ByVal locationText As String, ByVal inspectionType As String) As Boolean

    locationText = NormalizeActivityLibraryLocation(locationText)
    inspectionType = UCase$(Trim$(inspectionType))

    If Len(locationText) = 0 Then
        ScopeMatchesInspectionType = True
        Exit Function
    End If

    Select Case inspectionType
        Case INPUT_INSPECTION_TYPE_SHOP
            ScopeMatchesInspectionType = (locationText = "SHOP" Or locationText = "SHOP AND FIELD")
        Case INPUT_INSPECTION_TYPE_FIELD
            ScopeMatchesInspectionType = (locationText = "FIELD" Or locationText = "SHOP AND FIELD")
        Case Else
            ScopeMatchesInspectionType = True
    End Select

End Function

Private Function MergeScopeLocationValues(ByVal existingLocation As String, ByVal newLocation As String) As String

    existingLocation = NormalizeActivityLibraryLocation(existingLocation)
    newLocation = NormalizeActivityLibraryLocation(newLocation)

    If Len(existingLocation) = 0 Then
        MergeScopeLocationValues = newLocation
        Exit Function
    End If

    If Len(newLocation) = 0 Or existingLocation = newLocation Then
        MergeScopeLocationValues = existingLocation
        Exit Function
    End If

    MergeScopeLocationValues = "SHOP AND FIELD"

End Function

Private Function NormalizeActivityLibraryLocation(ByVal locationText As String) As String

    locationText = UCase$(Trim$(locationText))
    locationText = Replace(locationText, "/", " AND ")
    locationText = Replace(locationText, "&", " AND ")

    Do While InStr(locationText, "  ") > 0
        locationText = Replace(locationText, "  ", " ")
    Loop

    Select Case locationText
        Case "SHOP"
            NormalizeActivityLibraryLocation = "SHOP"
        Case "FIELD"
            NormalizeActivityLibraryLocation = "FIELD"
        Case "SHOP AND FIELD"
            NormalizeActivityLibraryLocation = "SHOP AND FIELD"
        Case Else
            NormalizeActivityLibraryLocation = ""
    End Select

End Function

Private Function GetScopeBlockListNameForInspectionType(ByVal inspectionType As String) As String

    Select Case UCase$(Trim$(inspectionType))
        Case INPUT_INSPECTION_TYPE_SHOP
            GetScopeBlockListNameForInspectionType = "ScopeBlockListShop"
        Case INPUT_INSPECTION_TYPE_FIELD
            GetScopeBlockListNameForInspectionType = "ScopeBlockListField"
        Case Else
            GetScopeBlockListNameForInspectionType = "ScopeBlockList"
    End Select

End Function

Private Sub EnsureActivityLibraryLocationColumn(Optional ByVal wsLib As Worksheet = Nothing)

    Dim locationCol As Long
    Dim rowTypeCol As Long
    Dim lastLib As Long
    Dim validationLastRow As Long

    If wsLib Is Nothing Then Set wsLib = ThisWorkbook.Worksheets("Activity Library")

    locationCol = FindActivityLibraryHeaderColumn(wsLib, LIB_LOCATION_HEADER)

    If locationCol = 0 Then
        rowTypeCol = FindActivityLibraryHeaderColumn(wsLib, LIB_ROW_TYPE_HEADER)
        If rowTypeCol = 0 Then rowTypeCol = LIB_LOCATION_INSERT_COL

        wsLib.Columns(rowTypeCol).Insert Shift:=xlToRight

        If rowTypeCol < wsLib.Columns.Count Then
            wsLib.Columns(rowTypeCol + 1).Copy
            wsLib.Columns(rowTypeCol).PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False
            wsLib.Columns(rowTypeCol).ColumnWidth = wsLib.Columns(rowTypeCol + 1).ColumnWidth
        End If

        wsLib.Cells(LIB_HEADER_ROW, rowTypeCol).Value = LIB_LOCATION_HEADER
        locationCol = rowTypeCol
    End If

    lastLib = GetActivityLibraryLastDataRow(wsLib)
    validationLastRow = lastLib
    If validationLastRow < 500 Then validationLastRow = 500

    ApplyActivityLibraryLocationValidation wsLib.Range(wsLib.Cells(LIB_HEADER_ROW + 1, locationCol), _
                                                       wsLib.Cells(validationLastRow, locationCol))

End Sub

Private Function GetActivityLibraryLastDataRow(Optional ByVal wsLib As Worksheet = Nothing) As Long

    Dim lastA As Long
    Dim lastC As Long
    Dim lastD As Long
    Dim lastRow As Long

    If wsLib Is Nothing Then Set wsLib = ThisWorkbook.Worksheets("Activity Library")

    lastA = wsLib.Cells(wsLib.Rows.Count, "A").End(xlUp).Row
    lastC = wsLib.Cells(wsLib.Rows.Count, "C").End(xlUp).Row
    lastD = wsLib.Cells(wsLib.Rows.Count, "D").End(xlUp).Row

    lastRow = lastA
    If lastC > lastRow Then lastRow = lastC
    If lastD > lastRow Then lastRow = lastD
    If lastRow < LIB_HEADER_ROW + 1 Then lastRow = LIB_HEADER_ROW + 1

    GetActivityLibraryLastDataRow = lastRow

End Function

Private Sub EnsureActivityLibraryScopeBlocks(Optional ByVal wsLib As Worksheet = Nothing)

    Dim lastLib As Long
    Dim r As Long
    Dim scopeName As String
    Dim sectionTitle As String
    Dim activityDesc As String
    Dim currentScope As String

    If wsLib Is Nothing Then Set wsLib = ThisWorkbook.Worksheets("Activity Library")

    lastLib = GetActivityLibraryLastDataRow(wsLib)
    currentScope = vbNullString

    For r = LIB_HEADER_ROW + 1 To lastLib
        scopeName = Trim$(CStr(wsLib.Cells(r, "A").Value))
        sectionTitle = Trim$(CStr(wsLib.Cells(r, LIB_SECTION_TITLE_COL).Value))
        activityDesc = Trim$(CStr(wsLib.Cells(r, LIB_ACTIVITY_DESC_COL).Value))

        If Len(activityDesc) = 0 And (Len(sectionTitle) > 0 Or Len(scopeName) > 0) Then
            ' Section header row
            If Len(scopeName) = 0 And Len(sectionTitle) > 0 Then
                wsLib.Cells(r, "A").Value = sectionTitle
                scopeName = sectionTitle
            End If
            If Len(scopeName) > 0 Then currentScope = scopeName
        ElseIf Len(activityDesc) > 0 Then
            ' Activity row - keep Scope Block filled for dropdown/generate matching
            If Len(scopeName) = 0 And Len(currentScope) > 0 Then
                wsLib.Cells(r, "A").Value = currentScope
            ElseIf Len(scopeName) > 0 Then
                currentScope = scopeName
            End If
        End If
    Next r

End Sub

Private Sub EnsureActivityLibraryTableCoversData(Optional ByVal wsLib As Worksheet = Nothing)

    Dim lo As ListObject
    Dim lastLib As Long
    Dim lastCol As Long

    If wsLib Is Nothing Then Set wsLib = ThisWorkbook.Worksheets("Activity Library")

    On Error Resume Next
    Set lo = wsLib.ListObjects("tblActivityLibrary")
    If lo Is Nothing And wsLib.ListObjects.Count > 0 Then Set lo = wsLib.ListObjects(1)
    On Error GoTo 0

    If lo Is Nothing Then Exit Sub

    lastLib = GetActivityLibraryLastDataRow(wsLib)
    lastCol = lo.HeaderRowRange.Columns.Count
    If lastCol < 1 Then Exit Sub
    If lastLib <= lo.HeaderRowRange.Row Then Exit Sub

    On Error Resume Next
    lo.Resize wsLib.Range(wsLib.Cells(lo.HeaderRowRange.Row, lo.HeaderRowRange.Column), _
                          wsLib.Cells(lastLib, lo.HeaderRowRange.Column + lastCol - 1))
    On Error GoTo 0

End Sub

Private Function FindActivityLibraryHeaderColumn(ByVal wsLib As Worksheet, ByVal headerText As String) As Long

    Dim lastCol As Long
    Dim colNum As Long

    lastCol = wsLib.Cells(LIB_HEADER_ROW, wsLib.Columns.Count).End(xlToLeft).Column

    For colNum = 1 To lastCol
        If StrComp(Trim$(CStr(wsLib.Cells(LIB_HEADER_ROW, colNum).Value)), headerText, vbTextCompare) = 0 Then
            FindActivityLibraryHeaderColumn = colNum
            Exit Function
        End If
    Next colNum

End Function

Private Sub ApplyActivityLibraryLocationValidation(ByVal targetRange As Range)

    If targetRange Is Nothing Then Exit Sub

    On Error Resume Next
    targetRange.Validation.Delete
    On Error GoTo 0

    With targetRange.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:=LIB_LOCATION_VALIDATION_LIST
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = LIB_LOCATION_HEADER
        .InputMessage = "Choose Shop, Field, or Shop and Field."
        .ErrorTitle = "Invalid Selection"
        .ErrorMessage = "Select Shop, Field, or Shop and Field."
        .ShowInput = True
        .ShowError = True
    End With

End Sub

Public Sub ShowScopeBlockPreview(ByVal wsIn As Worksheet, ByVal scopeName As String, _
                                 Optional ByVal scopeWO As String = "", _
                                 Optional ByVal scopeDrawing As String = "")

    Dim previewText As String

    previewText = BuildScopeBlockPreview(wsIn, scopeName, scopeWO, scopeDrawing)
    If Len(previewText) = 0 Then Exit Sub

    MsgBox previewText, vbInformation, "Elite ITP Builder"

End Sub

Public Sub PromptForCustomChapterInputs(ByVal wsIn As Worksheet, ByVal targetRow As Long)

    Dim customTitle As String
    Dim currentRows As Long
    Dim userInput As String

    If wsIn Is Nothing Then Exit Sub
    If targetRow < INPUT_CHAPTER_FIRST_ROW Then Exit Sub
    If UCase$(Trim$(CStr(wsIn.Cells(targetRow, INPUT_SCOPE_COL).Value))) <> "CUSTOM" Then Exit Sub

    customTitle = Trim$(CStr(wsIn.Cells(targetRow, INPUT_CUSTOM_TITLE_COL).Value))
    If Len(customTitle) = 0 Then customTitle = "Custom"

    userInput = InputBox("Enter the custom chapter title for this row." & vbCrLf & _
                         "This will be saved in column K.", _
                         "Elite ITP Builder", customTitle)
    userInput = Trim$(userInput)
    If Len(userInput) > 0 Then
        customTitle = userInput
    End If

    currentRows = Val(wsIn.Cells(targetRow, INPUT_CUSTOM_ROWS_COL).Value)
    If currentRows <= 0 Then currentRows = 5

    Do
        userInput = InputBox("Enter the number of blank activities for this custom chapter." & vbCrLf & _
                             "This will be saved in column L.", _
                             "Elite ITP Builder", CStr(currentRows))

        If Len(Trim$(userInput)) = 0 Then Exit Do
        If IsNumeric(userInput) Then
            currentRows = CLng(Val(userInput))
            If currentRows > 0 And CDbl(currentRows) = CDbl(Val(userInput)) Then Exit Do
        End If

        MsgBox "Enter a whole number greater than zero for the number of blank activities.", vbExclamation, "Elite ITP Builder"
    Loop

    Application.EnableEvents = False
    wsIn.Cells(targetRow, INPUT_CUSTOM_TITLE_COL).Value = customTitle
    wsIn.Cells(targetRow, INPUT_CUSTOM_ROWS_COL).Value = currentRows
    Application.EnableEvents = True

    wsIn.Activate
    wsIn.Cells(targetRow, INPUT_CUSTOM_TITLE_COL).Select

End Sub

Private Function BuildScopeBlockPreview(ByVal wsIn As Worksheet, ByVal scopeName As String, _
                                        Optional ByVal scopeWO As String = "", _
                                        Optional ByVal scopeDrawing As String = "") As String

    Dim wsLib As Worksheet
    Dim lastLib As Long
    Dim libRow As Long
    Dim wpsList As String
    Dim previewText As String
    Dim activityText As String
    Dim blockFound As Boolean
    Dim sectionInserted As Boolean
    Dim isSectionRow As Boolean
    Dim itemNo As Long
    Dim previousLineBlank As Boolean

    scopeName = Trim(scopeName)
    If Len(scopeName) = 0 Then Exit Function

    If UCase$(scopeName) = "CUSTOM" Then
        BuildScopeBlockPreview = "Scope Block: Custom" & vbCrLf & vbCrLf & _
            "This block is user-defined. It will use the custom title in column K and blank activity count in column L for that row on the Inputs sheet."
        Exit Function
    End If

    Set wsLib = ThisWorkbook.Worksheets("Activity Library")
    EnsureActivityLibraryLocationColumn wsLib
    lastLib = GetActivityLibraryLastDataRow(wsLib)
    wpsList = GetSelectedWPSList(wsIn)

    previewText = "Operations in " & scopeName & ":"

    For libRow = 2 To lastLib
        If Trim(CStr(wsLib.Cells(libRow, "A").Value)) = scopeName Then
            blockFound = True

                If Len(Trim(CStr(wsLib.Cells(libRow, LIB_SECTION_TITLE_COL).Value))) > 0 Or _
                    Len(Trim(CStr(wsLib.Cells(libRow, LIB_ACTIVITY_DESC_COL).Value))) > 0 Then

                     isSectionRow = (Len(Trim(CStr(wsLib.Cells(libRow, LIB_ACTIVITY_DESC_COL).Value))) = 0)

                If isSectionRow Then
                    If Not previousLineBlank Then
                        previewText = AppendScopePreviewLine(previewText, "")
                        previousLineBlank = True
                    End If
                    itemNo = 0
                    sectionInserted = True
                Else
                    If Not sectionInserted Then
                        sectionInserted = True
                    End If

                    itemNo = itemNo + 1
                    activityText = ReplaceTokens(CStr(wsLib.Cells(libRow, LIB_ACTIVITY_DESC_COL).Value), wsIn, scopeWO, scopeDrawing, wpsList)

                    previewText = AppendScopePreviewLine(previewText, CStr(itemNo) & ". " & activityText)
                    previousLineBlank = False
                End If
            End If
        End If
    Next libRow

    If Not blockFound Then
        previewText = AppendScopePreviewLine(previewText, "Block not found in Activity Library.")
    End If

    BuildScopeBlockPreview = LimitScopePreviewText(previewText)

End Function

Private Function AppendScopePreviewLine(ByVal previewText As String, ByVal lineText As String) As String

    If Len(previewText) = 0 Then
        AppendScopePreviewLine = lineText
    Else
        AppendScopePreviewLine = previewText & vbCrLf & lineText
    End If

End Function

Private Function LimitScopePreviewText(ByVal previewText As String) As String

    Const truncationNotice As String = vbCrLf & vbCrLf & "(Preview truncated. The selected block contains additional details.)"

    If Len(previewText) <= SCOPE_PREVIEW_MAX_LENGTH Then
        LimitScopePreviewText = previewText
    Else
        LimitScopePreviewText = Left$(previewText, SCOPE_PREVIEW_MAX_LENGTH - Len(truncationNotice)) & truncationNotice
    End If

End Function

Private Sub UpdateScopeBlockListName(ByVal wsLists As Worksheet, ByVal lastScopeRow As Long)

    UpdateNamedScopeList wsLists, "ScopeBlockList", LISTS_SCOPE_ALL_COL, lastScopeRow

End Sub

Private Sub UpdateNamedScopeList(ByVal wsLists As Worksheet, ByVal rangeName As String, ByVal colLetter As String, _
                                 ByVal lastScopeRow As Long)

    Dim refersToText As String

    If lastScopeRow < 2 Then lastScopeRow = 2
    refersToText = "=" & wsLists.Name & "!$" & colLetter & "$2:$" & colLetter & "$" & CStr(lastScopeRow)

    On Error Resume Next
    ThisWorkbook.Names(rangeName).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:=rangeName, RefersTo:=refersToText

End Sub

Private Sub InsertTemplateRow(ByVal wsT As Worksheet, ByVal wsOut As Worksheet, _
                              ByVal templateRow As Long, ByVal targetRow As Long)

    wsOut.Rows(targetRow).Insert Shift:=xlDown
    wsT.Rows(templateRow).Copy
    wsOut.Rows(targetRow).PasteSpecial Paste:=xlPasteAll
    wsOut.Rows(targetRow).RowHeight = wsT.Rows(templateRow).RowHeight
    Application.CutCopyMode = False

End Sub

Private Sub WriteSectionRow(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal headingText As String, _
                            Optional ByVal sourceSectionCell As Range = Nothing, _
                            Optional ByVal mappedTitleText As String = "")

    Dim headingCell As Range

    ClearRowMergeAreaContents ws, rowNum
    Set headingCell = ws.Range("A" & rowNum)
    SetMergedRangeValue headingCell, headingText

    ' Headings default to non-bold. Only true library differentiator words become bold+red.
    If Not sourceSectionCell Is Nothing Then
        If Len(mappedTitleText) = 0 Then mappedTitleText = CStr(sourceSectionCell.Value)
        ApplyLibraryTextDifferentiatorFormatting sourceSectionCell, headingCell, mappedTitleText
    Else
        On Error Resume Next
        If headingCell.MergeCells Then
            headingCell.MergeArea.Cells(1, 1).Font.Bold = False
        Else
            headingCell.Font.Bold = False
        End If
        On Error GoTo 0
    End If

End Sub

Private Sub WriteActivityRow(ByVal ws As Worksheet, ByVal rowNum As Long, _
                             ByVal lineNo As String, _
                             ByVal activityDesc As String, _
                             ByVal controlDocs As String, _
                             ByVal acceptanceCriteria As String, _
                             ByVal verificationDocs As String, _
                             ByVal elitePoint As String, _
                             ByVal approver1Point As String, _
                             ByVal approver2Point As String, _
                             ByVal comments As String, _
                             Optional ByVal commentIsFlagged As Boolean = False, _
                             Optional ByVal sourceCommentCell As Range = Nothing, _
                             Optional ByVal sourceActivityCell As Range = Nothing)

    Dim activityCell As Range

    SetMergedRangeValue ws.Range("A" & rowNum), lineNo
    Set activityCell = ws.Range("B" & rowNum)
    SetMergedRangeValue activityCell, activityDesc
    ' Same rule as section titles: uniform bold/plain stays plain; outliers become bold+red.
    If Not sourceActivityCell Is Nothing Then
        ApplyLibraryTextDifferentiatorFormatting sourceActivityCell, activityCell, CStr(sourceActivityCell.Value)
    End If

    SetMergedRangeValue ws.Range("K" & rowNum), controlDocs
    SetMergedRangeValue ws.Range("N" & rowNum), acceptanceCriteria
    SetMergedRangeValue ws.Range("Q" & rowNum), verificationDocs
    SetMergedRangeValue ws.Range("T" & rowNum), elitePoint
    SetMergedRangeValue GetApprover1TargetCell(ws, rowNum), approver1Point
    If IsTwoApproverWorksheet(ws) Then SetMergedRangeValue GetApprover2TargetCell(ws, rowNum), approver2Point
    ws.Range("A" & rowNum).NumberFormat = "@"
    SetMergedRangeValue ws.Range("A" & rowNum), lineNo

    WriteCommentsValue ws, rowNum, comments, commentIsFlagged, sourceCommentCell

End Sub

Private Function IsActivityLibraryCommentFlagged(ByVal commentCell As Range) As Boolean

    Dim cellText As String

    If commentCell Is Nothing Then Exit Function

    cellText = CStr(commentCell.Value)

    ' Reviewer-note convention used throughout the Activity Library.
    If Len(Trim$(cellText)) > 0 Then
        If InStr(1, cellText, "DELETE NOTE", vbTextCompare) > 0 Then
            IsActivityLibraryCommentFlagged = True
            Exit Function
        End If
    End If

    IsActivityLibraryCommentFlagged = IsActivityLibraryCellFormatted(commentCell)

End Function

Private Function IsActivityLibraryCellFormatted(ByVal sourceCell As Range) As Boolean

    Dim cellText As String
    Dim i As Long

    If sourceCell Is Nothing Then Exit Function

    If IsActivityLibraryCommentFillFlagged(sourceCell) Then
        IsActivityLibraryCellFormatted = True
        Exit Function
    End If

    If HasNonDefaultCommentFont(sourceCell.Font) Then
        IsActivityLibraryCellFormatted = True
        Exit Function
    End If

    cellText = CStr(sourceCell.Value)
    For i = 1 To Len(cellText)
        If HasNonDefaultCommentFont(sourceCell.Characters(i, 1).Font) Then
            IsActivityLibraryCellFormatted = True
            Exit Function
        End If
    Next i

End Function

Private Function IsActivityLibraryCommentFillFlagged(ByVal commentCell As Range) As Boolean

    Dim fillColor As Variant
    Dim fillPattern As Variant
    Dim fillColorIndex As Variant
    Dim displayColor As Variant

    If commentCell Is Nothing Then Exit Function

    On Error GoTo CleanFail
    fillPattern = commentCell.Interior.Pattern
    fillColor = commentCell.Interior.Color
    fillColorIndex = commentCell.Interior.ColorIndex

    If Not IsNull(fillPattern) Then
        If CLng(fillPattern) <> xlNone Then
            If Not IsNull(fillColor) Then
                If CLng(fillColor) <> vbWhite And CLng(fillColor) <> 16777215 Then
                    IsActivityLibraryCommentFillFlagged = True
                    Exit Function
                End If
            End If
        End If
    End If

    ' Yellow is ColorIndex 6 in the standard palette.
    If Not IsNull(fillColorIndex) Then
        If CLng(fillColorIndex) = 6 Then
            IsActivityLibraryCommentFillFlagged = True
            Exit Function
        End If
    End If

    displayColor = commentCell.DisplayFormat.Interior.Color
    If Not IsNull(displayColor) Then
        If CLng(displayColor) = FLAGGED_COMMENT_FILL_COLOR Then
            IsActivityLibraryCommentFillFlagged = True
        End If
    End If
    Exit Function

CleanFail:
    On Error GoTo 0

End Function

Private Function HasNonDefaultCommentFont(ByVal sourceFont As Font) As Boolean

    If IsFontPropertyTrue(sourceFont.Bold) Then
        HasNonDefaultCommentFont = True
    ElseIf IsFontPropertyTrue(sourceFont.Italic) Then
        HasNonDefaultCommentFont = True
    ElseIf IsFontUnderlineApplied(sourceFont.Underline) Then
        HasNonDefaultCommentFont = True
    ElseIf IsFontColorNonBlack(sourceFont.Color) Then
        HasNonDefaultCommentFont = True
    End If

End Function

Private Function IsFontPropertyTrue(ByVal propertyValue As Variant) As Boolean

    If IsNull(propertyValue) Then
        IsFontPropertyTrue = False
    Else
        IsFontPropertyTrue = CBool(propertyValue)
    End If

End Function

Private Function IsFontUnderlineApplied(ByVal propertyValue As Variant) As Boolean

    If IsNull(propertyValue) Then
        IsFontUnderlineApplied = False
    Else
        IsFontUnderlineApplied = (CLng(propertyValue) <> xlUnderlineStyleNone)
    End If

End Function

Private Function IsFontColorNonBlack(ByVal propertyValue As Variant) As Boolean

    If IsNull(propertyValue) Then
        IsFontColorNonBlack = False
    Else
        IsFontColorNonBlack = (CLng(propertyValue) <> vbBlack)
    End If

End Function

Private Function IsGeneratedSectionRow(ByVal cellText As String) As Boolean

    Dim numberPart As String
    Dim dotPos As Long

    cellText = Trim(cellText)
    If Len(cellText) = 0 Then Exit Function

    dotPos = InStr(1, cellText, ".0", vbTextCompare)
    If dotPos <= 1 Then Exit Function

    numberPart = Left$(cellText, dotPos - 1)
    If Not IsNumeric(numberPart) Then Exit Function

    If Mid$(cellText, dotPos + 2, 1) <> " " Then Exit Function

    IsGeneratedSectionRow = True

End Function

Private Function IsGeneratedActivityRow(ByVal cellText As String) As Boolean

    Dim parts() As String

    cellText = Trim(cellText)
    If Len(cellText) = 0 Then Exit Function
    If InStr(cellText, " ") > 0 Then Exit Function

    parts = Split(cellText, ".")
    If UBound(parts) <> 1 Then Exit Function
    If Not IsNumeric(parts(0)) Then Exit Function
    If Not IsNumeric(parts(1)) Then Exit Function
    If CLng(parts(1)) = 0 Then Exit Function

    IsGeneratedActivityRow = True

End Function

Private Function GetSectionHeadingText(ByVal cellText As String) As String

    Dim zeroPos As Long

    zeroPos = InStr(1, cellText, ".0", vbTextCompare)
    If zeroPos = 0 Then
        GetSectionHeadingText = Trim(cellText)
    Else
        GetSectionHeadingText = Trim(Mid$(cellText, zeroPos + 2))
    End If

End Function

Private Function GetGeneratedSectionNumber(ByVal cellText As String) As Long

    Dim zeroPos As Long
    Dim numberText As String

    cellText = Trim(cellText)
    zeroPos = InStr(1, cellText, ".0", vbTextCompare)
    If zeroPos <= 1 Then Exit Function

    numberText = Trim(Left$(cellText, zeroPos - 1))
    If IsNumeric(numberText) Then GetGeneratedSectionNumber = CLng(numberText)

End Function

Private Sub ApplyHeaderFromInputs(ByVal wsOut As Worksheet, ByVal wsIn As Worksheet)

    Dim customerName As String
    Dim approver2Name As String
    Dim contractText As String
    Dim descriptionText As String
    Dim itpNumber As String
    Dim inspectionType As String

    customerName = Trim(CStr(GetInputValue(wsIn, "Customer")))
    approver2Name = GetApprover2Name(wsIn)
    contractText = Trim(CStr(wsIn.Range("E8").Value))
    descriptionText = Trim(CStr(wsIn.Range("B8").Value))
    inspectionType = UCase$(Trim(CStr(wsIn.Range("E10").Value)))
    If Len(inspectionType) = 0 Then inspectionType = "FIELD"

    itpNumber = NormalizeITPNumber(CStr(GetInputValue(wsIn, "ITP Number")))

    wsOut.Range("A1").Value = inspectionType & " INSPECTION TEST PLAN (ITP)" & vbLf & itpNumber

    wsOut.Range("D2").Value = customerName
    wsOut.Range("D3").Value = wsIn.Range("B6").Value
    wsOut.Range("D4").Value = contractText
    wsOut.Range("D5").Value = descriptionText

    wsOut.Range("T8").Value = "Elite"
    wsOut.Range("Y8").Value = customerName

    wsOut.Range("S2").Value = GetInputValue(wsIn, "Elite Engineering")
    wsOut.Range("S3").Value = wsIn.Range("B11").Value
    wsOut.Range("S4").Value = GetInputValue(wsIn, "Tag Number")

    If IsTwoApproverTemplate(wsIn) Then
        wsOut.Range("AD8").Value = approver2Name
        wsOut.Range("AI8").Value = "Comments"
        SetMergedRangeValue wsOut.Range("AG2:AN2"), GetInputValue(wsIn, "Elite Foreman")
        SetMergedRangeValue wsOut.Range("AG3:AN3"), wsIn.Range("E9").Value
        SetMergedRangeValue wsOut.Range("AG4:AN4"), wsIn.Range("B9").Value
    Else
        wsOut.Range("AD8").Value = "Comments"
        wsOut.Range("AE2").Value = GetInputValue(wsIn, "Elite Foreman")
        SetMergedRangeValue wsOut.Range("AD3:AI3"), wsIn.Range("E9").Value
        SetMergedRangeValue wsOut.Range("AD4:AI4"), wsIn.Range("B9").Value
    End If

End Sub

Private Sub ApplyTemplateInputValues(ByVal wsOut As Worksheet, ByVal wsIn As Worksheet)

    Dim clientSiteRep As String
    Dim customerName As String
    Dim approver2Name As String

    clientSiteRep = CStr(wsIn.Range("E12").Value)
    customerName = Trim(CStr(GetInputValue(wsIn, "Customer")))
    approver2Name = GetApprover2Name(wsIn)

    ReplaceTemplateToken wsOut, "{ELITE_FOREMAN}", CStr(GetInputValue(wsIn, "Elite Foreman"))
    ReplaceTemplateToken wsOut, "{ELITE_REP}", CStr(GetInputValue(wsIn, "Elite Representative"))
    ReplaceTemplateToken wsOut, "{ELITE_REPRESENTATIVE}", CStr(GetInputValue(wsIn, "Elite Representative"))

    ReplaceTemplateToken wsOut, "{CLIENT_REP}", clientSiteRep
    ReplaceTemplateToken wsOut, "{CLIENT_REPRESENTATIVE}", clientSiteRep
    ReplaceTemplateToken wsOut, "{CLIENT_SITE_REP}", clientSiteRep
    ReplaceTemplateToken wsOut, "{CLIENT_QA}", clientSiteRep
    ReplaceTemplateToken wsOut, "{CLIENT_SITE_QA}", clientSiteRep

    SetCellBesideLabel wsOut, "Elite Foreman", CStr(GetInputValue(wsIn, "Elite Foreman"))
    SetCellBesideLabel wsOut, "Elite Representative", CStr(GetInputValue(wsIn, "Elite Representative"))

    ApplyFooterApproverLabels wsOut, customerName, approver2Name

    If Not IsTwoApproverWorksheet(wsOut) Then
        SetMergedRangeValue wsOut.Range("AD3:AI3"), wsIn.Range("E9").Value
    End If

End Sub

Private Sub ApplyFooterApproverLabels(ByVal wsOut As Worksheet, ByVal approver1Name As String, ByVal approver2Name As String)

    Dim approver1Label As String
    Dim footerTitleRow As Long
    Dim approverLabelRow As Long

    approver1Label = BuildRepLabelText(approver1Name)
    footerTitleRow = FindRowContains(wsOut, "Final Acceptance of ITP after completion of work.", 1, 1, 250)

    If footerTitleRow > 0 Then
        approverLabelRow = footerTitleRow + 1
    ElseIf IsTwoApproverWorksheet(wsOut) Then
        approverLabelRow = 67
    Else
        approverLabelRow = 161
    End If

    If IsTwoApproverWorksheet(wsOut) Then
        SetMergedRangeValue wsOut.Range("Q" & approverLabelRow), approver1Label
        SetMergedRangeValue wsOut.Range("AE" & approverLabelRow), BuildRepLabelText(approver2Name)
        ClearDuplicateFooterLabelRows wsOut, "Q", approverLabelRow, approver1Label
        ClearDuplicateFooterLabelRows wsOut, "AE", approverLabelRow, BuildRepLabelText(approver2Name)
    Else
        SetMergedRangeValue wsOut.Range("V" & approverLabelRow), approver1Label
        ClearDuplicateFooterLabelRows wsOut, "V", approverLabelRow, approver1Label
    End If

End Sub

Private Sub ClearDuplicateFooterLabelRows(ByVal wsOut As Worksheet, ByVal colLetter As String, _
                                          ByVal intendedRow As Long, ByVal labelText As String)

    Dim rowNum As Long

    If Len(labelText) = 0 Then Exit Sub

    For rowNum = intendedRow + 1 To intendedRow + 8
        If StrComp(Trim$(CStr(wsOut.Range(colLetter & rowNum).Value)), labelText, vbTextCompare) = 0 Then
            SetMergedRangeValue wsOut.Range(colLetter & rowNum), vbNullString
        End If
    Next rowNum

End Sub

Private Function BuildRepLabelText(ByVal companyName As String) As String

    companyName = Trim$(companyName)
    If Len(companyName) = 0 Then companyName = "Approver"

    BuildRepLabelText = companyName & " Rep's Name:"

End Function

Private Sub SetMergedRangeValue(ByVal rng As Range, ByVal valueText As Variant)

    If rng.Cells(1, 1).MergeCells Then
        rng.Cells(1, 1).MergeArea.Cells(1, 1).Value = valueText
    Else
        rng.Cells(1, 1).Value = valueText
    End If

End Sub

Private Sub ClearRowMergeAreaContents(ByVal ws As Worksheet, ByVal rowNum As Long)

    Dim clearArea As Range

    Set clearArea = ws.Range("A" & rowNum)
    If clearArea.MergeCells Then
        clearArea.MergeArea.ClearContents
    Else
        clearArea.EntireRow.ClearContents
    End If

End Sub

Private Sub WriteCommentsValue(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal comments As String, _
                               Optional ByVal commentIsFlagged As Boolean = False, _
                               Optional ByVal sourceCommentCell As Range = Nothing)

    Dim targetCell As Range
    Dim wrapArea As Range

    Set targetCell = GetCommentsTargetCell(ws, rowNum)
    SetMergedRangeValue targetCell, comments

    If targetCell.MergeCells Then
        Set wrapArea = targetCell.MergeArea
    Else
        Set wrapArea = targetCell.Resize(1, 1)
    End If

    wrapArea.WrapText = True
    wrapArea.VerticalAlignment = xlTop

    ' Preserve bold/italic/underline/font color on individual words when text length still matches.
    CopyRichTextFormatting sourceCommentCell, targetCell

    ' Reviewer trigger: yellow + bold. Track row so we can re-apply after final formatting.
    If commentIsFlagged Then
        TrackFlaggedCommentRow rowNum
        ApplyFlaggedCommentStyle ws, rowNum
    End If

End Sub

Private Sub ResetFlaggedCommentRows()

    Set flaggedCommentRows = New Collection

End Sub

Private Sub TrackFlaggedCommentRow(ByVal rowNum As Long)

    If flaggedCommentRows Is Nothing Then Set flaggedCommentRows = New Collection
    flaggedCommentRows.Add rowNum

End Sub

Private Sub ApplyAllFlaggedCommentStyles(ByVal ws As Worksheet)

    Dim rowItem As Variant

    If ws Is Nothing Then Exit Sub
    If flaggedCommentRows Is Nothing Then Exit Sub

    For Each rowItem In flaggedCommentRows
        ' Comments area only (section rows are split so this is never the title band).
        ApplyFlaggedCommentStyle ws, CLng(rowItem), True
    Next rowItem

End Sub

Private Sub ApplyFlaggedCommentStyle(ByVal ws As Worksheet, ByVal rowNum As Long, _
                                     Optional ByVal forceBold As Boolean = True)

    Dim targetCell As Range
    Dim wrapArea As Range

    If ws Is Nothing Or rowNum < 1 Then Exit Sub

    Set targetCell = GetCommentsTargetCell(ws, rowNum)
    If targetCell.MergeCells Then
        Set wrapArea = targetCell.MergeArea
    Else
        Set wrapArea = targetCell.Resize(1, 1)
    End If

    With wrapArea.Interior
        .Pattern = xlSolid
        .Color = FLAGGED_COMMENT_FILL_COLOR
    End With
    If forceBold Then wrapArea.Font.Bold = True

End Sub

Private Function CopyRichTextFormatting(ByVal sourceCell As Range, ByVal targetCell As Range) As Boolean

    Dim destCell As Range

    If sourceCell Is Nothing Or targetCell Is Nothing Then Exit Function

    If targetCell.MergeCells Then
        Set destCell = targetCell.MergeArea.Cells(1, 1)
    Else
        Set destCell = targetCell
    End If

    CopyRichTextFormatting = CopyRichTextFormattingToSubstring(sourceCell, destCell, CStr(sourceCell.Value))

End Function

' Library text is often uniformly bold. Generated text stays non-bold unless a word stands
' out (partial bold/unbold, italic, underline, or different color). Those differentiator
' words become bold + red on the generated ITP (section titles and activity descriptions).
Private Sub ApplyLibraryTextDifferentiatorFormatting(ByVal sourceCell As Range, ByVal targetCell As Range, _
                                                     ByVal mappedText As String)

    Dim sourceText As String
    Dim targetText As String
    Dim destCell As Range
    Dim mapText As String
    Dim sourceStart As Long
    Dim destStart As Long
    Dim textLen As Long
    Dim i As Long
    Dim styleKey As String
    Dim styleCounts As Object
    Dim isDiff() As Boolean
    Dim majorityKey As String
    Dim majorityCount As Long
    Dim keyItem As Variant
    Dim hasDiff As Boolean
    Dim runStart As Long
    Dim srcFont As Font

    If sourceCell Is Nothing Or targetCell Is Nothing Then Exit Sub

    If targetCell.MergeCells Then
        Set destCell = targetCell.MergeArea.Cells(1, 1)
    Else
        Set destCell = targetCell
    End If

    sourceText = CStr(sourceCell.Value)
    targetText = CStr(destCell.Value)
    mapText = mappedText
    If Len(mapText) = 0 Then mapText = sourceText
    If Len(sourceText) = 0 Or Len(targetText) = 0 Or Len(mapText) = 0 Then Exit Sub

    destStart = InStr(1, targetText, mapText, vbBinaryCompare)
    If destStart = 0 Then destStart = InStr(1, targetText, mapText, vbTextCompare)
    If destStart = 0 Then Exit Sub

    sourceStart = InStr(1, sourceText, mapText, vbBinaryCompare)
    If sourceStart = 0 Then sourceStart = InStr(1, sourceText, mapText, vbTextCompare)
    If sourceStart = 0 Then
        If StrComp(StripLeadingSectionNumber(sourceText), mapText, vbBinaryCompare) = 0 Then
            sourceStart = InStr(1, sourceText, mapText, vbTextCompare)
        End If
    End If
    If sourceStart = 0 Then Exit Sub
    If StrComp(Mid$(sourceText, sourceStart, Len(mapText)), mapText, vbTextCompare) <> 0 Then Exit Sub

    textLen = Len(mapText)
    On Error GoTo CleanFail

    ' Generated headings are non-bold by default (library all-bold is not a differentiator).
    With destCell.Characters(destStart, textLen).Font
        .Bold = False
        .Italic = False
        .Underline = xlUnderlineStyleNone
    End With

    Set styleCounts = CreateObject("Scripting.Dictionary")
    ReDim isDiff(1 To textLen)

    For i = 1 To textLen
        Set srcFont = sourceCell.Characters(sourceStart + i - 1, 1).Font
        styleKey = BuildLibraryTextStyleKey(srcFont)
        If styleCounts.Exists(styleKey) Then
            styleCounts(styleKey) = CLng(styleCounts(styleKey)) + 1
        Else
            styleCounts.Add styleKey, 1&
        End If
    Next i

    ' One uniform style (all bold / all plain / same color) => no differentiators.
    If styleCounts.Count <= 1 Then Exit Sub

    majorityCount = 0
    majorityKey = vbNullString
    For Each keyItem In styleCounts.Keys
        If CLng(styleCounts(keyItem)) > majorityCount Then
            majorityCount = CLng(styleCounts(keyItem))
            majorityKey = CStr(keyItem)
        End If
    Next keyItem

    hasDiff = False
    For i = 1 To textLen
        Set srcFont = sourceCell.Characters(sourceStart + i - 1, 1).Font
        If StrComp(BuildLibraryTextStyleKey(srcFont), majorityKey, vbBinaryCompare) <> 0 Then
            isDiff(i) = True
            hasDiff = True
        Else
            isDiff(i) = False
        End If
    Next i

    If Not hasDiff Then Exit Sub

    runStart = 0
    For i = 1 To textLen
        If isDiff(i) Then
            If runStart = 0 Then runStart = i
        ElseIf runStart > 0 Then
            With destCell.Characters(destStart + runStart - 1, i - runStart).Font
                .Bold = True
                .Color = SECTION_DIFF_HIGHLIGHT_COLOR
            End With
            runStart = 0
        End If
    Next i
    If runStart > 0 Then
        With destCell.Characters(destStart + runStart - 1, textLen - runStart + 1).Font
            .Bold = True
            .Color = SECTION_DIFF_HIGHLIGHT_COLOR
        End With
    End If

    Exit Sub

CleanFail:
    On Error GoTo 0

End Sub

Private Function BuildLibraryTextStyleKey(ByVal sourceFont As Font) As String

    Dim boldPart As String
    Dim italicPart As String
    Dim underlinePart As String
    Dim colorPart As String

    On Error GoTo CleanFail

    boldPart = IIf(IsFontPropertyTrue(sourceFont.Bold), "1", "0")
    italicPart = IIf(IsFontPropertyTrue(sourceFont.Italic), "1", "0")
    underlinePart = IIf(IsFontUnderlineApplied(sourceFont.Underline), "1", "0")

    If IsNull(sourceFont.Color) Then
        colorPart = "N"
    Else
        colorPart = CStr(CLng(sourceFont.Color))
    End If

    BuildLibraryTextStyleKey = boldPart & "|" & italicPart & "|" & underlinePart & "|" & colorPart
    Exit Function

CleanFail:
    BuildLibraryTextStyleKey = "0|0|0|N"

End Function

Private Function CopyRichTextFormattingToSubstring(ByVal sourceCell As Range, ByVal targetCell As Range, _
                                                   ByVal mappedText As String) As Boolean

    Dim sourceText As String
    Dim targetText As String
    Dim destCell As Range
    Dim mapText As String
    Dim sourceStart As Long
    Dim destStart As Long
    Dim textLen As Long
    Dim i As Long
    Dim runStart As Long
    Dim srcFont As Font
    Dim sourceIndex As Long
    Dim destIndex As Long

    If sourceCell Is Nothing Or targetCell Is Nothing Then Exit Function

    If targetCell.MergeCells Then
        Set destCell = targetCell.MergeArea.Cells(1, 1)
    Else
        Set destCell = targetCell
    End If

    sourceText = CStr(sourceCell.Value)
    targetText = CStr(destCell.Value)
    mapText = mappedText
    If Len(mapText) = 0 Then mapText = sourceText
    If Len(sourceText) = 0 Or Len(targetText) = 0 Or Len(mapText) = 0 Then Exit Function

    destStart = InStr(1, targetText, mapText, vbBinaryCompare)
    If destStart = 0 Then destStart = InStr(1, targetText, mapText, vbTextCompare)
    If destStart = 0 Then Exit Function

    sourceStart = InStr(1, sourceText, mapText, vbBinaryCompare)
    If sourceStart = 0 Then sourceStart = InStr(1, sourceText, mapText, vbTextCompare)
    If sourceStart = 0 Then
        If StrComp(StripLeadingSectionNumber(sourceText), mapText, vbBinaryCompare) = 0 Then
            sourceStart = InStr(1, sourceText, mapText, vbTextCompare)
        End If
    End If
    If sourceStart = 0 Then Exit Function
    If StrComp(Mid$(sourceText, sourceStart, Len(mapText)), mapText, vbTextCompare) <> 0 Then Exit Function

    textLen = Len(mapText)
    On Error GoTo CleanFail

    runStart = 1
    For i = 1 To textLen
        If i = textLen Or Not CommentFontLooksSame(sourceCell.Characters(sourceStart + runStart - 1, 1).Font, _
                                                   sourceCell.Characters(sourceStart + i, 1).Font) Then
            Set srcFont = sourceCell.Characters(sourceStart + runStart - 1, i - runStart + 1).Font
            sourceIndex = sourceStart + runStart - 1
            destIndex = destStart + runStart - 1
            With destCell.Characters(destIndex, i - runStart + 1).Font
                If Not IsNull(srcFont.Bold) Then .Bold = srcFont.Bold
                If Not IsNull(srcFont.Italic) Then .Italic = srcFont.Italic
                If Not IsNull(srcFont.Underline) Then .Underline = srcFont.Underline
                If Not IsNull(srcFont.Color) Then .Color = srcFont.Color
            End With
            runStart = i + 1
        End If
    Next i

    CopyRichTextFormattingToSubstring = True
    Exit Function

CleanFail:
    On Error GoTo 0

End Function

Private Function CommentFontLooksSame(ByVal fontA As Font, ByVal fontB As Font) As Boolean

    On Error GoTo CleanFail

    If IsFontPropertyTrue(fontA.Bold) <> IsFontPropertyTrue(fontB.Bold) Then Exit Function
    If IsFontPropertyTrue(fontA.Italic) <> IsFontPropertyTrue(fontB.Italic) Then Exit Function
    If IsFontUnderlineApplied(fontA.Underline) <> IsFontUnderlineApplied(fontB.Underline) Then Exit Function

    If IsNull(fontA.Color) Xor IsNull(fontB.Color) Then Exit Function
    If Not IsNull(fontA.Color) Then
        If CLng(fontA.Color) <> CLng(fontB.Color) Then Exit Function
    End If

    CommentFontLooksSame = True
    Exit Function

CleanFail:
    CommentFontLooksSame = False

End Function

Private Function GetApprover1TargetCell(ByVal ws As Worksheet, ByVal rowNum As Long) As Range

    Set GetApprover1TargetCell = ws.Range("Y" & rowNum)

End Function

Private Function GetApprover2TargetCell(ByVal ws As Worksheet, ByVal rowNum As Long) As Range

    Set GetApprover2TargetCell = ws.Range("AD" & rowNum)

End Function

Private Function GetCommentsTargetCell(ByVal ws As Worksheet, ByVal rowNum As Long) As Range

    If IsTwoApproverWorksheet(ws) Then
        Set GetCommentsTargetCell = ws.Range("AI" & rowNum)
    Else
        Set GetCommentsTargetCell = ws.Range("AD" & rowNum)
    End If

End Function

Private Function IsTwoApproverWorksheet(ByVal ws As Worksheet) As Boolean

    IsTwoApproverWorksheet = (StrComp(Trim$(CStr(ws.Range("AI8").Value)), "Comments", vbTextCompare) = 0)

End Function

Private Sub ReplaceTemplateToken(ByVal ws As Worksheet, ByVal tokenText As String, ByVal replacementText As String)

    ws.Cells.Replace What:=tokenText, _
                     Replacement:=replacementText, _
                     LookAt:=xlPart, _
                     SearchOrder:=xlByRows, _
                     MatchCase:=False

End Sub

Private Sub SetCellBesideLabel(ByVal ws As Worksheet, ByVal labelText As String, ByVal valueText As String)

    Dim f As Range
    Dim targetCell As Range

    Set f = ws.Range("A1:AJ250").Find( _
        What:=labelText, _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        MatchCase:=False)

    If f Is Nothing Then Exit Sub

    Set targetCell = f.Offset(0, 1)

    If targetCell.MergeCells Then
        targetCell.MergeArea.Cells(1, 1).Value = valueText
    Else
        targetCell.Value = valueText
    End If

End Sub

Private Function UniqueChapterValueSummary(ByVal wsIn As Worksheet, ByVal valueCol As Long, ByVal fallbackText As String) As String

    Dim d As Object
    Dim r As Long
    Dim v As String
    Dim scopeName As String

    Set d = CreateObject("Scripting.Dictionary")

    For r = INPUT_CHAPTER_FIRST_ROW To GetInputChapterLastRow(wsIn)

        scopeName = Trim(CStr(wsIn.Cells(r, INPUT_SCOPE_COL).Value))

        If Len(scopeName) > 0 Then
            v = Trim(CStr(wsIn.Cells(r, valueCol).Value))
            If Len(v) > 0 Then d(v) = True
        End If

    Next r

    If d.Count = 0 Then
        UniqueChapterValueSummary = fallbackText
    ElseIf d.Count = 1 Then
        UniqueChapterValueSummary = CStr(d.Keys()(0))
    Else
        UniqueChapterValueSummary = "Multiple - Refer to applicable ITP sections"
    End If

End Function

Private Function GetSelectedWPSList(ByVal wsIn As Worksheet) As String

    Dim r As Long
    Dim selected As String
    Dim activeText As String
    Dim wpsNo As String
    Dim revNo As String
    Dim entryText As String
    Dim result As String

    For r = WPS_FIRST_ROW To WPS_LAST_ROW

        selected = UCase$(Trim(CStr(wsIn.Cells(r, "N").Value)))
        activeText = UCase$(Trim(CStr(wsIn.Cells(r, "S").Value)))
        wpsNo = Trim(CStr(wsIn.Cells(r, "O").Value))
        revNo = Trim(CStr(wsIn.Cells(r, "P").Value))

        If (selected = "YES" Or selected = "Y" Or selected = "TRUE" Or selected = "1") And _
           (activeText <> "NO") And Len(wpsNo) > 0 Then

            entryText = wpsNo
            If Len(revNo) > 0 Then entryText = entryText & " (Rev " & revNo & ")"

            If Len(result) > 0 Then result = result & "; "
            result = result & entryText

        End If

    Next r

    If Len(result) = 0 Then result = ""

    GetSelectedWPSList = result

End Function

Private Function ValidateGenerateITPInputs(ByVal wsIn As Worksheet) As Boolean

    Dim requiredMissing As String
    Dim optionalMissing As String
    Dim firstRequiredAddress As String
    Dim firstOptionalAddress As String
    Dim response As VbMsgBoxResult

    AddMissingField wsIn, "B5", "Customer", False, optionalMissing, firstOptionalAddress
    AddMissingField wsIn, "B6", "Project Title", False, optionalMissing, firstOptionalAddress
    AddMissingField wsIn, "B7", "Tag Number", False, optionalMissing, firstOptionalAddress
    AddMissingField wsIn, "B8", "Description", False, optionalMissing, firstOptionalAddress
    AddMissingField wsIn, "B9", "Location", False, optionalMissing, firstOptionalAddress
    AddMissingField wsIn, "B10", "Elite Engineering", False, optionalMissing, firstOptionalAddress
    AddMissingField wsIn, "B11", "Elite QC", False, optionalMissing, firstOptionalAddress
    AddMissingField wsIn, "B12", "Elite Foreman", False, optionalMissing, firstOptionalAddress
    AddMissingField wsIn, "E8", "Contract Number", False, optionalMissing, firstOptionalAddress
    AddMissingField wsIn, "E9", "Client QC", False, optionalMissing, firstOptionalAddress
    AddMissingField wsIn, "E10", "Inspection Type (Shop/Field)", True, requiredMissing, firstRequiredAddress
    AddMissingField wsIn, "E5", "ITP Number", True, requiredMissing, firstRequiredAddress
    AddMissingField wsIn, "E6", "Revision", True, requiredMissing, firstRequiredAddress
    AddMissingField wsIn, "E7", "Code Basis", True, requiredMissing, firstRequiredAddress

    If Len(requiredMissing) > 0 Then
        MsgBox "Cannot generate the ITP until these required fields are filled:" & vbCrLf & vbCrLf & requiredMissing, _
               vbExclamation, "Elite ITP Builder"
        ShowInputCell wsIn, firstRequiredAddress
        Exit Function
    End If

    If Len(optionalMissing) > 0 Then
        response = MsgBox("These fields are blank:" & vbCrLf & vbCrLf & optionalMissing & vbCrLf & vbCrLf & _
                          "Do you want to go back and fill them before generating?", _
                          vbQuestion + vbYesNo, "Elite ITP Builder")
        If response = vbYes Then
            ShowInputCell wsIn, firstOptionalAddress
            Exit Function
        End If
    End If

    ValidateGenerateITPInputs = True

End Function

Private Sub AddMissingField(ByVal ws As Worksheet, ByVal cellAddress As String, ByVal fieldName As String, _
                            ByVal isRequired As Boolean, ByRef missingList As String, ByRef firstMissingAddress As String)

    If Len(Trim(CStr(ws.Range(cellAddress).Value))) = 0 Then
        If Len(missingList) > 0 Then missingList = missingList & vbCrLf
        missingList = missingList & fieldName & " (" & cellAddress & ")"
        If Len(firstMissingAddress) = 0 Then firstMissingAddress = cellAddress
    End If

End Sub

Private Sub ShowInputCell(ByVal ws As Worksheet, ByVal cellAddress As String)

    If Len(cellAddress) = 0 Then Exit Sub

    ws.Activate
    ws.Range(cellAddress).Select

End Sub

Private Sub FormatFinalOutput(ByVal ws As Worksheet)

    Dim lastRow As Long
    Dim lastCell As Range

    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookIn:=xlFormulas, _
                                 LookAt:=xlPart, SearchOrder:=xlByRows, _
                                 SearchDirection:=xlPrevious, MatchCase:=False)

    If lastCell Is Nothing Then
        lastRow = 1
    Else
        lastRow = lastCell.Row
    End If

    If lastRow < ws.Rows.Count Then lastRow = lastRow + 1

    GetCommentsFormatRange(ws).WrapText = True
    GetCommentsFormatRange(ws).VerticalAlignment = xlTop
    GetHeaderFormatRange(ws).VerticalAlignment = xlBottom

    ws.ResetAllPageBreaks

    ws.PageSetup.PrintArea = "$A$2:$" & GetPrintLastColumn(ws) & "$" & lastRow
    ws.PageSetup.Orientation = xlLandscape

    ws.PageSetup.PrintTitleRows = "$7:$9"
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False

    ApplyContinuationPageHeader ws

End Sub

Private Sub ApplyContinuationPageHeader(ByVal ws As Worksheet)

    Const HEADER_LOGO_HEIGHT As Double = HEADER_LOGO_LOCKED_HEIGHT

    Dim titleLines() As String
    Dim titleText As String
    Dim mainTitle As String
    Dim itpNumber As String
    Dim logoPath As String
    Dim clientLogoPath As String
    Dim existingLogoPath As String
    Dim existingClientLogoPath As String
    Dim logoDisplayWidth As Double
    Dim clientLogoDisplayWidth As Double
    Dim hasSeparateClientLogo As Boolean

    If Not IsITPWorksheetForHeader(ws) Then Exit Sub

    ApplyEliteLogoToRowHeader ws
    If CLIENT_LOGO_FEATURE_ENABLED Then ApplyClientLogoToRowHeader ws

    titleText = Replace(CStr(ws.Range("A1").Value), vbCr, vbNullString)
    titleLines = Split(titleText, vbLf)
    mainTitle = Trim$(titleLines(0))
    If UBound(titleLines) >= 1 Then itpNumber = Trim$(titleLines(UBound(titleLines)))

    logoPath = GetEliteHeaderLogoPath()
    hasSeparateClientLogo = CLIENT_LOGO_FEATURE_ENABLED And IsLogoEnabled(CLIENT_LOGO_ENABLED_REF)
    If hasSeparateClientLogo Then clientLogoPath = BuildPositionedClientHeaderLogoPath(IsTwoApproverWorksheet(ws))

    With ws.PageSetup
        existingLogoPath = vbNullString
        On Error Resume Next
        existingLogoPath = .LeftHeaderPicture.fileName
        existingClientLogoPath = .CenterHeaderPicture.fileName
        On Error GoTo 0

        .DifferentFirstPageHeaderFooter = False

        .LeftHeader = vbNullString
        .CenterHeader = vbNullString
        If IsTwoApproverWorksheet(ws) Then
            .RightHeader = Chr$(10) & Chr$(10) & "&""Calibri,Regular""&18" & EscapePageHeaderText(mainTitle)
        Else
            .RightHeader = Chr$(10) & "&""Calibri,Regular""&18" & EscapePageHeaderText(mainTitle)
        End If
        If Len(itpNumber) > 0 Then .RightHeader = .RightHeader & Chr$(10) & "&""Calibri,Regular""&18" & EscapePageHeaderText(itpNumber)

        .HeaderMargin = Application.InchesToPoints(0.17)
        .TopMargin = Application.InchesToPoints(0.9)
        .LeftFooter = BuildRevisionFooterText(ws)
        .CenterFooter = "&P of &N"
        .RightFooter = vbNullString

        If Len(logoPath) > 0 Then
            On Error Resume Next
            .LeftHeaderPicture.fileName = logoPath
            logoDisplayWidth = GetScaledLogoWidth(logoPath, HEADER_LOGO_HEIGHT, 126)
            .LeftHeaderPicture.Height = HEADER_LOGO_HEIGHT
            .LeftHeaderPicture.Width = logoDisplayWidth
            .LeftHeader = "&G"
            On Error GoTo 0
        End If

        If hasSeparateClientLogo And Len(clientLogoPath) > 0 Then
            On Error Resume Next
            .CenterHeaderPicture.fileName = clientLogoPath
            clientLogoDisplayWidth = GetScaledLogoWidth(clientLogoPath, HEADER_LOGO_HEIGHT, IIf(IsTwoApproverWorksheet(ws), 760, 620))
            .CenterHeaderPicture.Height = HEADER_LOGO_HEIGHT
            .CenterHeaderPicture.Width = clientLogoDisplayWidth
            On Error GoTo 0
            .CenterHeader = "&G"
        End If
    End With

End Sub

Private Function BuildRevisionFooterText(ByVal ws As Worksheet) As String

    Dim revisionText As String
    Dim f As Range
    Dim targetCell As Range

    If ws Is Nothing Then Exit Function

    Set f = ws.Range("A1:AJ250").Find( _
        What:="Revision", _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        MatchCase:=False)

    If Not f Is Nothing Then
        Set targetCell = f.Offset(0, 1)
        If targetCell.MergeCells Then
            revisionText = Trim$(CStr(targetCell.MergeArea.Cells(1, 1).Value))
        Else
            revisionText = Trim$(CStr(targetCell.Value))
        End If
    End If

    If Len(revisionText) = 0 Then
        On Error Resume Next
        revisionText = Trim$(CStr(ThisWorkbook.Worksheets("Inputs").Range("E6").Value))
        On Error GoTo 0
    End If

    BuildRevisionFooterText = FormatRevisionFooterText(revisionText)

End Function

Private Function FormatRevisionFooterText(ByVal revisionText As String) As String

    revisionText = Trim$(revisionText)
    If Len(revisionText) = 0 Then Exit Function

    If UCase$(Left$(revisionText, 4)) = "REV-" Or UCase$(Left$(revisionText, 4)) = "REV " Then
        revisionText = Trim$(Mid$(revisionText, 5))
    End If

    If Len(revisionText) = 0 Then Exit Function
    FormatRevisionFooterText = "Rev " & revisionText

End Function

Private Sub ApplyEliteLogoToRowHeader(ByVal ws As Worksheet)

    Dim logoPath As String
    Dim eliteLogo As Shape
    Dim newLogo As Shape
    Dim logoLeft As Double
    Dim logoTop As Double
    Dim logoHeight As Double
    Dim logoWidth As Double
    Dim logoName As String

    logoPath = GetEliteHeaderLogoPath()
    If Len(logoPath) = 0 Then
        ClearEliteRowHeaderLogo ws
        Exit Sub
    End If

    Set eliteLogo = GetEliteRowHeaderLogoShape(ws)
    If eliteLogo Is Nothing Then
        logoLeft = ws.Range("A1").Left
        logoTop = ws.Range("A1").Top
        logoHeight = ws.Rows(1).Height - 1.5
        If logoHeight <= 0 Then logoHeight = HEADER_LOGO_LOCKED_HEIGHT
        logoWidth = GetScaledLogoWidth(logoPath, logoHeight, 110)
        logoName = "Picture 1"
    Else
        logoLeft = eliteLogo.Left
        logoTop = eliteLogo.Top
        logoHeight = eliteLogo.Height
        logoWidth = GetScaledLogoWidth(logoPath, logoHeight, eliteLogo.Width)
        logoName = eliteLogo.Name
    End If

    On Error Resume Next
    If Not eliteLogo Is Nothing Then eliteLogo.Delete
    Set newLogo = ws.Shapes.AddPicture(logoPath, msoFalse, msoTrue, logoLeft, logoTop, logoWidth, logoHeight)
    On Error GoTo 0
    If newLogo Is Nothing Then Exit Sub

    newLogo.Name = logoName
    newLogo.Placement = xlMoveAndSize

End Sub

Private Sub ClearEliteRowHeaderLogo(ByVal ws As Worksheet)

    Dim eliteLogo As Shape

    Set eliteLogo = GetEliteRowHeaderLogoShape(ws)

    On Error Resume Next
    If Not eliteLogo Is Nothing Then eliteLogo.Delete
    On Error GoTo 0

End Sub

Private Sub ApplyClientLogoToRowHeader(ByVal ws As Worksheet)

    Const CLIENT_LOGO_SHAPE_NAME As String = "ITP_Client_Row1_Logo"
    Const CLIENT_LOGO_GAP As Double = 12

    Dim clientLogoPath As String
    Dim eliteLogo As Shape
    Dim clientLogo As Shape
    Dim clientLogoWidth As Double
    Dim clientLogoHeight As Double
    Dim logoLeft As Double
    Dim logoTop As Double
    Dim existingClientLogo As Shape

    If Not CLIENT_LOGO_FEATURE_ENABLED Then Exit Sub

    clientLogoPath = GetClientHeaderLogoPath()
    If Len(clientLogoPath) = 0 Then
        On Error Resume Next
        ws.Shapes(CLIENT_LOGO_SHAPE_NAME).Delete
        On Error GoTo 0
        Exit Sub
    End If

    Set eliteLogo = GetEliteRowHeaderLogoShape(ws)
    On Error Resume Next
    Set existingClientLogo = ws.Shapes(CLIENT_LOGO_SHAPE_NAME)
    On Error GoTo 0

    If eliteLogo Is Nothing Then
        If existingClientLogo Is Nothing Then Exit Sub

        clientLogoHeight = existingClientLogo.Height
        logoLeft = existingClientLogo.Left
        logoTop = existingClientLogo.Top
    Else
        clientLogoHeight = eliteLogo.Height
        logoLeft = eliteLogo.Left + eliteLogo.Width + CLIENT_LOGO_GAP
        logoTop = eliteLogo.Top
    End If

    clientLogoWidth = GetScaledLogoWidth(clientLogoPath, clientLogoHeight, 90)

    On Error Resume Next
    ws.Shapes(CLIENT_LOGO_SHAPE_NAME).Delete
    Set clientLogo = ws.Shapes.AddPicture(clientLogoPath, msoFalse, msoTrue, logoLeft, logoTop, clientLogoWidth, clientLogoHeight)
    On Error GoTo 0
    If clientLogo Is Nothing Then Exit Sub

    clientLogo.Name = CLIENT_LOGO_SHAPE_NAME
    clientLogo.Placement = xlMoveAndSize

End Sub

Private Function GetEliteRowHeaderLogoShape(ByVal ws As Worksheet) As Shape

    Dim candidate As Shape

    For Each candidate In ws.Shapes
        If candidate.TopLeftCell.Row = 1 And candidate.TopLeftCell.Column = 1 Then
            Set GetEliteRowHeaderLogoShape = candidate
            Exit Function
        End If
    Next candidate

End Function

Private Function IsITPWorksheetForHeader(ByVal ws As Worksheet) As Boolean

    If ws Is Nothing Then Exit Function

    IsITPWorksheetForHeader = (InStr(1, CStr(ws.Range("A1").Value), "INSPECTION TEST PLAN", vbTextCompare) > 0)

End Function

Private Function GetWorkbookFolderPath() As String

    Dim rawPath As String
    Dim localFullName As String
    Dim fso As Object

    rawPath = ThisWorkbook.Path
    If Len(rawPath) = 0 Then Exit Function

    ' Normal local/network path
    If InStr(1, rawPath, "://") = 0 Then
        GetWorkbookFolderPath = rawPath
        Exit Function
    End If

    ' OneDrive/SharePoint cloud URL -> local synced folder
    localFullName = ResolveCloudWorkbookLocalFullName()
    If Len(localFullName) = 0 Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    GetWorkbookFolderPath = fso.GetParentFolderName(localFullName)

End Function

Private Function ResolveCloudWorkbookLocalFullName() As String

    Dim urlPath As String
    Dim documentsPos As Long
    Dim relativePath As String
    Dim oneDriveRoot As String
    Dim candidate As String
    Dim roots As Variant
    Dim i As Long
    Dim fso As Object

    urlPath = Replace(ThisWorkbook.FullName, "%20", " ")
    documentsPos = InStr(1, urlPath, "/Documents/", vbTextCompare)
    If documentsPos = 0 Then Exit Function

    relativePath = Mid$(urlPath, documentsPos + Len("/Documents/"))
    relativePath = Replace(relativePath, "/", Application.PathSeparator)

    roots = Array(Environ$("OneDriveCommercial"), Environ$("OneDrive"), Environ$("OneDriveConsumer"))
    Set fso = CreateObject("Scripting.FileSystemObject")

    For i = LBound(roots) To UBound(roots)
        oneDriveRoot = Trim$(CStr(roots(i)))
        If Len(oneDriveRoot) > 0 Then
            candidate = oneDriveRoot & Application.PathSeparator & relativePath
            If fso.FileExists(candidate) Then
                ResolveCloudWorkbookLocalFullName = candidate
                Exit Function
            End If
            If fso.FolderExists(fso.GetParentFolderName(candidate)) Then
                ResolveCloudWorkbookLocalFullName = candidate
                Exit Function
            End If
        End If
    Next i

End Function

Private Function FileExistsPath(ByVal filePath As String) As Boolean

    Dim fso As Object

    If Len(filePath) = 0 Then Exit Function
    If InStr(1, filePath, "://") > 0 Then Exit Function

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExistsPath = fso.FileExists(filePath)
    On Error GoTo 0

End Function

Private Sub DeleteFileIfExists(ByVal filePath As String)

    On Error Resume Next
    If FileExistsPath(filePath) Then Kill filePath
    On Error GoTo 0

End Sub

Private Function GetEliteHeaderLogoPath() As String

    Dim filePath As String
    Dim folderPath As String

    If Not IsLogoEnabled(ELITE_LOGO_ENABLED_REF) Then Exit Function

    folderPath = GetWorkbookFolderPath()
    If Len(folderPath) = 0 Then Exit Function

    filePath = folderPath & Application.PathSeparator & ELITE_HEADER_LOGO_FILE
    If Not FileExistsPath(filePath) Then EnsureEliteHeaderLogoFile filePath
    If Not FileExistsPath(filePath) Then Exit Function

    GetEliteHeaderLogoPath = filePath

End Function

Private Sub EnsureEliteHeaderLogoFile(ByVal targetFile As String)

    Dim wsAssets As Worksheet
    Dim wsWork As Worksheet
    Dim wsPrevious As Worksheet
    Dim logoShape As Shape
    Dim chartHolder As ChartObject
    Dim priorVisibility As XlSheetVisibility

    If Len(targetFile) = 0 Then Exit Sub

    On Error GoTo CleanFail
    Set wsPrevious = ActiveSheet
    Set wsAssets = ThisWorkbook.Worksheets(LOGO_ASSET_SHEET)
    Set wsWork = ThisWorkbook.Worksheets("Inputs")
    priorVisibility = wsAssets.Visible
    wsAssets.Visible = xlSheetVisible
    Set logoShape = wsAssets.Shapes(LOGO_ASSET_ELITE_SHAPE)

    If FileExistsPath(targetFile) Then
        If FileLen(targetFile) > 500 Then
            wsAssets.Visible = priorVisibility
            Exit Sub
        End If
    End If
    DeleteFileIfExists targetFile

    wsWork.Activate
    Set chartHolder = wsWork.ChartObjects.Add(0, 0, logoShape.Width, logoShape.Height)
    chartHolder.Activate
    chartHolder.Chart.ChartArea.Format.Fill.Visible = msoFalse
    chartHolder.Chart.ChartArea.Format.Line.Visible = msoFalse
    chartHolder.Chart.ChartArea.Select
    logoShape.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
    DoEvents
    ActiveChart.Paste
    DoEvents
    If ActiveChart.Shapes.Count = 0 Then
        logoShape.Copy
        DoEvents
        ActiveChart.Paste
        DoEvents
    End If
    chartHolder.Chart.Export fileName:=targetFile, FilterName:="PNG"
    chartHolder.Delete
    If Not FileExistsPath(targetFile) Then GoTo CleanFail
    If FileLen(targetFile) <= 500 Then GoTo CleanFail
    wsAssets.Visible = priorVisibility
    If Not wsPrevious Is Nothing Then wsPrevious.Activate

    Exit Sub

CleanFail:
    On Error Resume Next
    If Not chartHolder Is Nothing Then chartHolder.Delete
    If Not wsAssets Is Nothing Then wsAssets.Visible = priorVisibility
    If Not wsPrevious Is Nothing Then wsPrevious.Activate
    On Error GoTo 0

End Sub

Private Sub UpdateEmbeddedEliteLogoAsset(ByVal logoFile As String)

    Dim wsAssets As Worksheet
    Dim logoShape As Shape

    If Not FileExistsPath(logoFile) Then Exit Sub

    On Error Resume Next
    Set wsAssets = ThisWorkbook.Worksheets(LOGO_ASSET_SHEET)
    On Error GoTo 0

    If wsAssets Is Nothing Then
        Set wsAssets = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsAssets.Name = LOGO_ASSET_SHEET
    End If

    wsAssets.Visible = xlSheetVisible

    On Error Resume Next
    wsAssets.Shapes(LOGO_ASSET_ELITE_SHAPE).Delete
    On Error GoTo 0

    Set logoShape = wsAssets.Shapes.AddPicture(logoFile, msoFalse, msoTrue, 0, 0, -1, -1)
    logoShape.Name = LOGO_ASSET_ELITE_SHAPE
    logoShape.Placement = xlFreeFloating

    wsAssets.Visible = xlSheetVeryHidden

End Sub

Private Function BuildPositionedClientHeaderLogoPath(Optional ByVal isTwoApprover As Boolean = False) As String

    Const LOGO_HEIGHT As Double = HEADER_LOGO_LOCKED_HEIGHT
    Const CLIENT_CANVAS_WIDTH As Double = 620
    Const CLIENT_CANVAS_WIDTH_TYPE_2 As Double = 760

    Dim ws As Worksheet
    Dim chartHolder As ChartObject
    Dim clientLogoPath As String
    Dim filePath As String
    Dim clientLogoWidth As Double
    Dim canvasWidth As Double
    Dim exportScale As Double

    If isTwoApprover Then
        canvasWidth = CLIENT_CANVAS_WIDTH_TYPE_2
        filePath = GetWorkbookFolderPath() & Application.PathSeparator & BuildLogoCacheFileName("Client_Header_Positioned_Type2")
    Else
        canvasWidth = CLIENT_CANVAS_WIDTH
        filePath = GetWorkbookFolderPath() & Application.PathSeparator & BuildLogoCacheFileName("Client_Header_Positioned")
    End If

    clientLogoPath = GetClientHeaderLogoPath()
    If Len(clientLogoPath) = 0 Then
        Exit Function
    End If

    clientLogoWidth = GetScaledLogoWidth(clientLogoPath, LOGO_HEIGHT, 90)
    If clientLogoWidth > canvasWidth Then canvasWidth = clientLogoWidth
    exportScale = HEADER_LOGO_EXPORT_SCALE
    If exportScale < 1 Then exportScale = 1

    On Error GoTo CleanFail
    Set ws = ThisWorkbook.Worksheets("Inputs")
    Set chartHolder = ws.ChartObjects.Add(0, 0, canvasWidth * exportScale, LOGO_HEIGHT * exportScale)

    chartHolder.Chart.ChartArea.Format.Fill.Visible = msoFalse
    chartHolder.Chart.ChartArea.Format.Line.Visible = msoFalse
    chartHolder.Chart.Shapes.AddPicture clientLogoPath, msoFalse, msoTrue, 0, 0, clientLogoWidth * exportScale, LOGO_HEIGHT * exportScale
    chartHolder.Chart.Export fileName:=filePath, FilterName:="PNG"
    chartHolder.Delete

    BuildPositionedClientHeaderLogoPath = filePath
    Exit Function

CleanFail:
    On Error Resume Next
    If Not chartHolder Is Nothing Then chartHolder.Delete
    On Error GoTo 0

    If FileExistsPath(filePath) Then BuildPositionedClientHeaderLogoPath = filePath

End Function

Private Function BuildLogoCacheFileName(ByVal filePrefix As String) As String

    BuildLogoCacheFileName = filePrefix & "_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Timer * 100)) & ".png"

End Function

Private Function GetClientHeaderLogoPath() As String

    Dim filePath As String
    Dim folderPath As String

    folderPath = GetWorkbookFolderPath()
    If Len(folderPath) = 0 Then Exit Function

    filePath = folderPath & Application.PathSeparator & CLIENT_HEADER_LOGO_FILE
    If Not FileExistsPath(filePath) Then Exit Function

    GetClientHeaderLogoPath = filePath

End Function

Private Function IsLogoEnabled(ByVal stateName As String) As Boolean

    Dim nameRef As Name
    Dim stateValue As String

    On Error Resume Next
    Set nameRef = ThisWorkbook.Names(stateName)
    On Error GoTo 0

    If nameRef Is Nothing Then Exit Function

    stateValue = CStr(nameRef.RefersTo)
    If Left$(stateValue, 1) = "=" Then stateValue = Mid$(stateValue, 2)
    stateValue = Replace(stateValue, Chr$(34), vbNullString)

    IsLogoEnabled = (UCase$(Trim$(stateValue)) = "TRUE")

End Function

Private Sub SetLogoEnabled(ByVal stateName As String, ByVal enabled As Boolean)

    On Error Resume Next
    ThisWorkbook.Names(stateName).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:=stateName, RefersTo:="=" & IIf(enabled, "TRUE", "FALSE")

End Sub

Private Function BuildCombinedHeaderLogoPath(ByVal eliteLogoPath As String, ByVal clientLogoPath As String) As String

    Const LOGO_HEIGHT As Double = 72
    Const DEFAULT_ELITE_LOGO_WIDTH As Double = 126
    Const DEFAULT_CLIENT_LOGO_WIDTH As Double = 116
    Const LOGO_GAP As Double = 8
    Const MAX_COMBINED_LOGO_WIDTH As Double = 285

    Dim ws As Worksheet
    Dim chartHolder As ChartObject
    Dim filePath As String
    Dim eliteLogoWidth As Double
    Dim clientLogoWidth As Double
    Dim clientLogoHeight As Double
    Dim clientLogoTop As Double
    Dim maxClientLogoWidth As Double
    Dim clientScale As Double

    If (Not FileExistsPath(eliteLogoPath)) Or (Not FileExistsPath(clientLogoPath)) Then Exit Function

    filePath = GetWorkbookFolderPath() & Application.PathSeparator & COMBINED_HEADER_LOGO_FILE
    eliteLogoWidth = GetScaledLogoWidth(eliteLogoPath, LOGO_HEIGHT, DEFAULT_ELITE_LOGO_WIDTH)
    clientLogoWidth = GetScaledLogoWidth(clientLogoPath, LOGO_HEIGHT, DEFAULT_CLIENT_LOGO_WIDTH)
    clientLogoHeight = LOGO_HEIGHT

    maxClientLogoWidth = MAX_COMBINED_LOGO_WIDTH - eliteLogoWidth - LOGO_GAP
    If maxClientLogoWidth > 0 And clientLogoWidth > maxClientLogoWidth Then
        clientScale = maxClientLogoWidth / clientLogoWidth
        clientLogoWidth = clientLogoWidth * clientScale
        clientLogoHeight = clientLogoHeight * clientScale
        clientLogoTop = (LOGO_HEIGHT - clientLogoHeight) / 2
    End If

    On Error GoTo CleanFail
    Set ws = ThisWorkbook.Worksheets("Inputs")
    Set chartHolder = ws.ChartObjects.Add(0, 0, eliteLogoWidth + LOGO_GAP + clientLogoWidth, LOGO_HEIGHT)

    chartHolder.Chart.ChartArea.Format.Fill.Visible = msoFalse
    chartHolder.Chart.ChartArea.Format.Line.Visible = msoFalse
    chartHolder.Chart.Shapes.AddPicture eliteLogoPath, msoFalse, msoTrue, 0, 0, eliteLogoWidth, LOGO_HEIGHT
    chartHolder.Chart.Shapes.AddPicture clientLogoPath, msoFalse, msoTrue, eliteLogoWidth + LOGO_GAP, clientLogoTop, clientLogoWidth, clientLogoHeight
    chartHolder.Chart.Export fileName:=filePath, FilterName:="PNG"
    chartHolder.Delete

    BuildCombinedHeaderLogoPath = filePath
    Exit Function

CleanFail:
    On Error Resume Next
    If Not chartHolder Is Nothing Then chartHolder.Delete
    On Error GoTo 0

End Function

Private Function GetScaledLogoWidth(ByVal logoPath As String, ByVal targetHeight As Double, ByVal fallbackWidth As Double) As Double

    Dim aspectRatio As Double
    Dim logoShape As Shape
    Dim originalWidth As Double
    Dim originalHeight As Double

    aspectRatio = GetImageAspectRatio(logoPath)
    If aspectRatio > 0 Then
        GetScaledLogoWidth = targetHeight * aspectRatio
        Exit Function
    End If

    On Error GoTo CleanFail
    Set logoShape = ThisWorkbook.Worksheets("Inputs").Shapes.AddPicture(logoPath, msoFalse, msoFalse, 0, 0, -1, -1)
    originalWidth = logoShape.Width
    originalHeight = logoShape.Height
    logoShape.Delete

    If originalHeight <= 0 Then GoTo CleanFail

    GetScaledLogoWidth = targetHeight * (originalWidth / originalHeight)
    Exit Function

CleanFail:
    On Error Resume Next
    If Not logoShape Is Nothing Then logoShape.Delete
    On Error GoTo 0
    GetScaledLogoWidth = fallbackWidth

End Function

Private Function GetImageAspectRatio(ByVal logoPath As String) As Double

    Dim imageFile As Object

    If Not FileExistsPath(logoPath) Then Exit Function

    On Error GoTo CleanFail
    Set imageFile = CreateObject("WIA.ImageFile")
    imageFile.LoadFile logoPath
    If imageFile.Height > 0 Then GetImageAspectRatio = CDbl(imageFile.Width) / CDbl(imageFile.Height)
    Exit Function

CleanFail:
    GetImageAspectRatio = 0

End Function

Public Sub AddClientHeaderLogo()

    Dim selectedFile As String
    Dim targetFile As String
    Dim fd As FileDialog

    If Not CLIENT_LOGO_FEATURE_ENABLED Then
        MsgBox "Client logo is currently disabled. Use Add Logo instead.", vbInformation, "Elite ITP Builder"
        Exit Sub
    End If

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select Client Logo"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif"
        If .Show <> -1 Then Exit Sub
        selectedFile = .SelectedItems(1)
    End With

    targetFile = GetWorkbookFolderPath() & Application.PathSeparator & CLIENT_HEADER_LOGO_FILE
    SaveLogoFile selectedFile, targetFile

    SetLogoEnabled CLIENT_LOGO_ENABLED_REF, True
    DeleteCombinedHeaderLogo
    DeletePositionedClientHeaderLogos
    ApplyHeaderLogoToType2Sheets

    MsgBox "Client logo added. Generated ITP headers will show it beside the Elite logo.", vbInformation, "Elite ITP Builder"

End Sub

Public Sub ClearClientHeaderLogo()

    Dim targetFile As String

    If Not CLIENT_LOGO_FEATURE_ENABLED Then
        MsgBox "Client logo is currently disabled. Use Clear Logo instead.", vbInformation, "Elite ITP Builder"
        Exit Sub
    End If

    targetFile = GetWorkbookFolderPath() & Application.PathSeparator & CLIENT_HEADER_LOGO_FILE
    DeleteFileIfExists targetFile

    SetLogoEnabled CLIENT_LOGO_ENABLED_REF, False
    DeleteCombinedHeaderLogo
    DeletePositionedClientHeaderLogos
    ApplyHeaderLogoToType2Sheets

    MsgBox "Client logo cleared. Generated ITP headers will use the Elite logo only.", vbInformation, "Elite ITP Builder"

End Sub

Public Sub AddEliteHeaderLogo()

    Dim selectedFile As String
    Dim targetFile As String
    Dim folderPath As String
    Dim fd As FileDialog
    Dim response As VbMsgBoxResult

    folderPath = GetWorkbookFolderPath()
    If Len(folderPath) = 0 Then
        MsgBox "Could not resolve a local folder for this workbook." & vbCrLf & vbCrLf & _
               "Open the file from your local OneDrive folder (not the SharePoint/web path), then try again.", _
               vbExclamation, "Elite ITP Builder"
        Exit Sub
    End If

    response = MsgBox("Do you want to select a new logo file?" & vbCrLf & vbCrLf & _
                      "Yes = browse for a new logo." & vbCrLf & _
                      "No = restore the embedded Elite logo." & vbCrLf & _
                      "Cancel = do nothing.", _
                      vbQuestion + vbYesNoCancel, "Elite ITP Builder")
    If response = vbCancel Then Exit Sub

    targetFile = folderPath & Application.PathSeparator & ELITE_HEADER_LOGO_FILE

    If response = vbNo Then
        DeleteFileIfExists targetFile
        EnsureEliteHeaderLogoFile targetFile
        If Not FileExistsPath(targetFile) Then
            MsgBox "Could not restore the embedded Elite logo.", vbExclamation, "Elite ITP Builder"
            Exit Sub
        End If

        SetLogoEnabled ELITE_LOGO_ENABLED_REF, True
        DeleteCombinedHeaderLogo
        ApplyHeaderLogoToType2Sheets

        MsgBox "Embedded Elite logo restored. Generated ITP headers will use the embedded Elite logo.", vbInformation, "Elite ITP Builder"
        Exit Sub
    End If

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select Logo"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif"
        If .Show <> -1 Then Exit Sub
        selectedFile = .SelectedItems(1)
    End With

    If Not SaveLogoFile(selectedFile, targetFile) Then
        MsgBox "Could not copy the selected logo file to:" & vbCrLf & vbCrLf & targetFile, _
               vbExclamation, "Elite ITP Builder"
        Exit Sub
    End If

    UpdateEmbeddedEliteLogoAsset targetFile
    SetLogoEnabled ELITE_LOGO_ENABLED_REF, True
    DeleteCombinedHeaderLogo
    ApplyHeaderLogoToType2Sheets

    MsgBox "Logo added. Generated ITP headers will use the selected logo.", vbInformation, "Elite ITP Builder"

End Sub

Private Function SaveLogoFile(ByVal sourceFile As String, ByVal targetFile As String) As Boolean

    Dim fso As Object

    If Len(sourceFile) = 0 Or Len(targetFile) = 0 Then Exit Function
    If InStr(1, targetFile, "://") > 0 Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(sourceFile) Then Exit Function

    On Error Resume Next
    If fso.FileExists(targetFile) Then fso.DeleteFile targetFile, True
    fso.CopyFile sourceFile, targetFile, True
    On Error GoTo 0

    SaveLogoFile = fso.FileExists(targetFile)

End Function

Private Sub SaveLogoAtLockedHeight(ByVal sourceFile As String, ByVal targetFile As String, ByVal lockedHeight As Double)

    Dim ws As Worksheet
    Dim chartHolder As ChartObject
    Dim tempFile As String
    Dim targetWidth As Double
    Dim aspectRatio As Double
    Dim exportScale As Double

    If lockedHeight <= 0 Then lockedHeight = 72

    On Error GoTo CleanFail
    Set ws = ThisWorkbook.Worksheets("Inputs")

    aspectRatio = GetImageAspectRatio(sourceFile)
    If aspectRatio <= 0 Then GoTo CleanFail
    targetWidth = lockedHeight * aspectRatio
    exportScale = HEADER_LOGO_EXPORT_SCALE
    If exportScale < 1 Then exportScale = 1

    tempFile = GetWorkbookFolderPath() & Application.PathSeparator & "Elite_Header_Logo_Import.png"
    DeleteFileIfExists tempFile

    Set chartHolder = ws.ChartObjects.Add(0, 0, targetWidth * exportScale, lockedHeight * exportScale)
    chartHolder.Chart.ChartArea.Format.Fill.Visible = msoFalse
    chartHolder.Chart.ChartArea.Format.Line.Visible = msoFalse
    chartHolder.Chart.Shapes.AddPicture sourceFile, msoFalse, msoTrue, 0, 0, targetWidth * exportScale, lockedHeight * exportScale
    chartHolder.Chart.Export fileName:=tempFile, FilterName:="PNG"
    chartHolder.Delete
    Set chartHolder = Nothing

    DeleteFileIfExists targetFile
    Name tempFile As targetFile
    Exit Sub

CleanFail:
    On Error Resume Next
    If Not chartHolder Is Nothing Then chartHolder.Delete
    If Len(tempFile) > 0 Then DeleteFileIfExists tempFile
    SaveLogoFile sourceFile, targetFile
    On Error GoTo 0

End Sub

Private Function GetLogoImageHeight(ByVal logoPath As String, ByVal fallbackHeight As Double) As Double

    Dim logoShape As Shape

    If Not FileExistsPath(logoPath) Then
        GetLogoImageHeight = fallbackHeight
        Exit Function
    End If

    On Error GoTo CleanFail
    Set logoShape = ThisWorkbook.Worksheets("Inputs").Shapes.AddPicture(logoPath, msoFalse, msoFalse, 0, 0, -1, -1)
    GetLogoImageHeight = logoShape.Height
    logoShape.Delete
    Exit Function

CleanFail:
    On Error Resume Next
    If Not logoShape Is Nothing Then logoShape.Delete
    On Error GoTo 0
    GetLogoImageHeight = fallbackHeight

End Function

Public Sub ClearEliteHeaderLogo()

    Dim targetFile As String
    Dim folderPath As String

    folderPath = GetWorkbookFolderPath()
    If Len(folderPath) = 0 Then
        SetLogoEnabled ELITE_LOGO_ENABLED_REF, False
        MsgBox "Logo cleared in settings, but the local workbook folder could not be resolved.", vbExclamation, "Elite ITP Builder"
        Exit Sub
    End If

    targetFile = folderPath & Application.PathSeparator & ELITE_HEADER_LOGO_FILE
    DeleteFileIfExists targetFile

    SetLogoEnabled ELITE_LOGO_ENABLED_REF, False
    DeleteCombinedHeaderLogo
    ApplyHeaderLogoToType2Sheets

    MsgBox "Logo cleared. Generated ITP headers will not show a logo until one is added.", vbInformation, "Elite ITP Builder"

End Sub

Private Sub DeleteCombinedHeaderLogo()

    Dim folderPath As String

    folderPath = GetWorkbookFolderPath()
    If Len(folderPath) = 0 Then Exit Sub

    DeleteFileIfExists folderPath & Application.PathSeparator & COMBINED_HEADER_LOGO_FILE

End Sub

Private Sub DeletePositionedClientHeaderLogos()

    Dim positionedFile As String
    Dim folderPath As String

    folderPath = GetWorkbookFolderPath()
    If Len(folderPath) = 0 Then Exit Sub

    positionedFile = Dir$(folderPath & Application.PathSeparator & "Client_Header_Positioned*.png")

    On Error Resume Next
    Do While Len(positionedFile) > 0
        Kill folderPath & Application.PathSeparator & positionedFile
        positionedFile = Dir$()
    Loop
    On Error GoTo 0

End Sub

Private Sub ApplyHeaderLogoToType2Sheets()

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If IsITPWorksheetForHeader(ws) Then ApplyContinuationPageHeader ws
    Next ws

    ThisWorkbook.Save

End Sub

Private Function EscapePageHeaderText(ByVal headerText As String) As String

    EscapePageHeaderText = Replace(headerText, "&", "&&")

End Function

Private Function GetDeleteButtonColumn(ByVal ws As Worksheet) As String

    If IsTwoApproverWorksheet(ws) Then
        GetDeleteButtonColumn = "AQ"
    Else
        GetDeleteButtonColumn = "AL"
    End If

End Function

Private Function GetAddButtonColumn(ByVal ws As Worksheet) As String

    If IsTwoApproverWorksheet(ws) Then
        GetAddButtonColumn = "AR"
    Else
        GetAddButtonColumn = "AM"
    End If

End Function

Private Function GetMoveUpButtonColumn(ByVal ws As Worksheet) As String

    If IsTwoApproverWorksheet(ws) Then
        GetMoveUpButtonColumn = "AS"
    Else
        GetMoveUpButtonColumn = "AN"
    End If

End Function

Private Function GetMoveDownButtonColumn(ByVal ws As Worksheet) As String

    If IsTwoApproverWorksheet(ws) Then
        GetMoveDownButtonColumn = "AT"
    Else
        GetMoveDownButtonColumn = "AO"
    End If

End Function

Private Function GetPrintSafetyColumn(ByVal ws As Worksheet) As String

    If IsTwoApproverWorksheet(ws) Then
        GetPrintSafetyColumn = "AP"
    Else
        GetPrintSafetyColumn = "AK"
    End If

End Function

Private Function GetCommentsFormatRange(ByVal ws As Worksheet) As Range

    If IsTwoApproverWorksheet(ws) Then
        Set GetCommentsFormatRange = ws.Columns("AI:AO")
    Else
        Set GetCommentsFormatRange = ws.Columns("AD:AJ")
    End If

End Function

Private Function GetHeaderFormatRange(ByVal ws As Worksheet) As Range

    If IsTwoApproverWorksheet(ws) Then
        Set GetHeaderFormatRange = ws.Range("AI2:AO4")
    Else
        Set GetHeaderFormatRange = ws.Range("AD2:AJ4")
    End If

End Function

Private Function GetPrintLastColumn(ByVal ws As Worksheet) As String

    If IsTwoApproverWorksheet(ws) Then
        GetPrintLastColumn = "AO"
    Else
        GetPrintLastColumn = "AJ"
    End If

End Function

Private Function FindRowContains(ByVal ws As Worksheet, ByVal txt As String, _
                                 ByVal colNum As Long, ByVal firstRow As Long, ByVal lastRow As Long) As Long

    Dim r As Long

    For r = firstRow To lastRow
        If InStr(1, CStr(ws.Cells(r, colNum).Value), txt, vbTextCompare) > 0 Then
            FindRowContains = r
            Exit Function
        End If
    Next r

    FindRowContains = 0

End Function

Private Function GetTemplateWorksheet(ByVal wsIn As Worksheet) As Worksheet

    If IsTwoApproverTemplate(wsIn) Then
        Set GetTemplateWorksheet = ThisWorkbook.Worksheets(TEMPLATE_SHEET_TYPE_2)
    Else
        Set GetTemplateWorksheet = ThisWorkbook.Worksheets(TEMPLATE_SHEET)
    End If

End Function

Private Function IsTwoApproverTemplate(ByVal wsIn As Worksheet) As Boolean

    Dim approver2Name As String

    approver2Name = UCase$(GetApprover2Name(wsIn))

    Select Case approver2Name
        Case "", "N/A", "NA", "NO"
            IsTwoApproverTemplate = False
        Case Else
            IsTwoApproverTemplate = True
    End Select

End Function

Private Function GetApprover2Name(ByVal wsIn As Worksheet) As String

    Dim approver2Name As String

    approver2Name = Trim(CStr(wsIn.Range("E11").Value))

    GetApprover2Name = approver2Name

End Function

Private Function GetInputValue(ByVal wsIn As Worksheet, ByVal labelText As String) As Variant

    Dim f As Range

    Set f = wsIn.Range("A1:S100").Find( _
        What:=labelText, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False)

    If f Is Nothing Then
        GetInputValue = vbNullString
    Else
        GetInputValue = f.Offset(0, 1).Value
    End If

End Function

Private Function FirstListValue(ByVal wsIn As Worksheet, ByVal addressText As String) As String

    Dim c As Range

    For Each c In wsIn.Range(addressText).Cells
        If Len(Trim(CStr(c.Value))) > 0 Then
            FirstListValue = Trim(CStr(c.Value))
            Exit Function
        End If
    Next c

    FirstListValue = ""

End Function

Private Function FormatPrefixedValue(ByVal valueText As String, ByVal prefixText As String) As String

    Dim values() As String
    Dim formattedValues() As String
    Dim itemText As String
    Dim i As Long
    Dim formattedCount As Long

    valueText = Trim(valueText)

    If Len(valueText) = 0 Then
        FormatPrefixedValue = ""
        Exit Function
    End If

    values = Split(valueText, ",")
    ReDim formattedValues(LBound(values) To UBound(values))

    For i = LBound(values) To UBound(values)
        itemText = Trim$(CStr(values(i)))
        If Len(itemText) > 0 Then
            If UCase$(Left$(itemText, Len(prefixText))) <> UCase$(prefixText) Then
                itemText = prefixText & itemText
            End If
            formattedValues(formattedCount) = itemText
            formattedCount = formattedCount + 1
        End If
    Next i

    If formattedCount = 0 Then Exit Function

    ReDim Preserve formattedValues(0 To formattedCount - 1)
    FormatPrefixedValue = Join(formattedValues, ", ")

End Function

Private Function FormatSectionHeading(ByVal headingText As String, ByVal scopeWO As String, ByVal scopeDrawing As String) As String

    Dim hasWO As Boolean
    Dim hasDrawing As Boolean
    Dim headingSuffix As String

    headingText = Trim$(headingText)
    scopeWO = Trim$(scopeWO)
    scopeDrawing = Trim$(scopeDrawing)

    hasWO = (Len(scopeWO) > 0 And InStr(1, headingText, scopeWO, vbTextCompare) > 0)
    hasDrawing = (Len(scopeDrawing) > 0)

    If hasWO Or hasDrawing Then
        If hasWO Then
            headingText = Replace(headingText, "(" & scopeWO & ")", "")
            headingText = Replace(headingText, scopeWO, "")
        End If
        If hasDrawing Then
            headingText = Replace(headingText, "(" & scopeDrawing & ")", "")
            headingText = Replace(headingText, scopeDrawing, "")
        End If

        headingSuffix = BuildHeadingScopeDisplay(IIf(hasWO, scopeWO, ""), IIf(hasDrawing, scopeDrawing, ""))
    End If

    headingText = Replace(headingText, "(), ()", "")
    headingText = Replace(headingText, "(),", "")
    headingText = Replace(headingText, ", ()", "")
    headingText = Replace(headingText, "()", "")

    Do While InStr(headingText, "  ") > 0
        headingText = Replace(headingText, "  ", " ")
    Loop

    headingText = Trim$(headingText)
    If Right$(headingText, 1) = "," Then headingText = Trim$(Left$(headingText, Len(headingText) - 1))
    If Right$(headingText, 1) = "-" Then headingText = Trim$(Left$(headingText, Len(headingText) - 1))

    If Len(headingSuffix) > 0 Then
        If Len(headingText) > 0 Then
            headingText = headingText & headingSuffix
        Else
            headingText = Trim$(headingSuffix)
        End If
    End If

    FormatSectionHeading = headingText

End Function

Private Function BuildHeadingScopeDisplay(ByVal scopeWO As String, ByVal scopeDrawing As String) As String

    Dim headingSuffix As String

    If Len(scopeWO) > 0 Then headingSuffix = "(" & scopeWO & ")"
    If Len(scopeDrawing) > 0 Then
        If Len(headingSuffix) > 0 Then
            headingSuffix = headingSuffix & ", (" & scopeDrawing & ")"
        Else
            headingSuffix = "(" & scopeDrawing & ")"
        End If
    End If

    If Len(headingSuffix) > 0 Then BuildHeadingScopeDisplay = " " & headingSuffix

End Function

Public Function GetDynamicInputListRange(ByVal ws As Worksheet, ByVal colLetter As String, ByVal firstRow As Long) As Range

    Dim lastRow As Long

    lastRow = ws.Cells(GetInputListLastRow(colLetter), colLetter).End(xlUp).Row
    If lastRow < firstRow Then lastRow = firstRow
    If lastRow > GetInputListLastRow(colLetter) Then lastRow = GetInputListLastRow(colLetter)

    Set GetDynamicInputListRange = ws.Range(colLetter & firstRow & ":" & colLetter & lastRow)

End Function

Private Function GetDynamicInputListClearRange(ByVal ws As Worksheet, ByVal colLetter As String, _
                                               ByVal firstRow As Long, ByVal lastRow As Long) As Range

    Dim endColLetter As String

    If lastRow < firstRow Then lastRow = firstRow
    If lastRow > GetInputListLastRow(colLetter) Then lastRow = GetInputListLastRow(colLetter)
    endColLetter = GetInputListEndColumn(colLetter)

    Set GetDynamicInputListClearRange = ws.Range(colLetter & firstRow & ":" & endColLetter & lastRow)

End Function

Private Function GetInputListEndColumn(ByVal startColLetter As String) As String

    Select Case UCase$(startColLetter)
        Case "A"
            GetInputListEndColumn = "B"
        Case "D"
            GetInputListEndColumn = "E"
        Case Else
            GetInputListEndColumn = startColLetter
    End Select

End Function

Private Function GetInputListLastRow(ByVal startColLetter As String) As Long

    Select Case UCase$(startColLetter)
        Case "A"
            GetInputListLastRow = INPUT_WO_LIST_LAST_ROW
        Case "D"
            GetInputListLastRow = INPUT_DWG_LIST_LAST_ROW
        Case Else
            GetInputListLastRow = ThisWorkbook.Worksheets("Inputs").Rows.Count
    End Select

End Function

Private Function RemovePrefixIfPresent(ByVal valueText As String, ByVal prefixText As String) As String

    Dim values() As String
    Dim cleanedValues() As String
    Dim itemText As String
    Dim i As Long
    Dim cleanedCount As Long

    valueText = Trim(valueText)

    If Len(valueText) = 0 Then Exit Function

    values = Split(valueText, ",")
    ReDim cleanedValues(LBound(values) To UBound(values))

    For i = LBound(values) To UBound(values)
        itemText = Trim$(CStr(values(i)))
        If Len(itemText) > 0 Then
            If UCase$(Left$(itemText, Len(prefixText))) = UCase$(prefixText) Then
                itemText = Mid$(itemText, Len(prefixText) + 1)
            End If
            cleanedValues(cleanedCount) = itemText
            cleanedCount = cleanedCount + 1
        End If
    Next i

    If cleanedCount = 0 Then Exit Function

    ReDim Preserve cleanedValues(0 To cleanedCount - 1)
    RemovePrefixIfPresent = Join(cleanedValues, ", ")

End Function

Private Function NormalizeITPNumber(ByVal itpNumber As String) As String

    itpNumber = Trim(itpNumber)
    If Len(itpNumber) = 0 Then Exit Function

    If UCase$(Left$(itpNumber, 4)) = "ITP-" Then
        NormalizeITPNumber = itpNumber
    ElseIf UCase$(itpNumber) = "ITP" Then
        NormalizeITPNumber = "ITP"
    ElseIf UCase$(Left$(itpNumber, 3)) = "ITP" Then
        NormalizeITPNumber = "ITP-" & Trim(Mid$(itpNumber, 4))
    Else
        NormalizeITPNumber = "ITP-" & itpNumber
    End If

End Function

Private Function RemoveBlankDrawingPhrase(ByVal txt As String) As String

    txt = Replace(txt, "IFC'ed Drawing {DRAWING}, ", "")
    txt = Replace(txt, "IFC'ed Drawing {DRAWING},", "")
    txt = Replace(txt, "IFC'ed Drawing {DRAWING}", "")
    txt = Replace(txt, "{DRAWING}", "")

    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop

    txt = Replace(txt, " ,", ",")
    txt = Replace(txt, ", ,", ",")
    txt = Replace(txt, ",,", ",")

    Do While Left$(Trim$(txt), 1) = ","
        txt = Trim$(Mid$(Trim$(txt), 2))
    Loop

    txt = Trim$(txt)
    If Right$(txt, 1) = "," Then txt = Trim$(Left$(txt, Len(txt) - 1))

    RemoveBlankDrawingPhrase = txt

End Function

Private Function ReplaceTokens(ByVal txt As String, ByVal wsIn As Worksheet, _
                               Optional ByVal scopeWO As String = "", _
                               Optional ByVal scopeDrawing As String = "", _
                               Optional ByVal wpsList As String = "", _
                               Optional ByVal allowDrawingListFallback As Boolean = True) As String

    Dim clientSiteRep As String

    clientSiteRep = CStr(wsIn.Range("E12").Value)

    If Len(scopeWO) = 0 Then scopeWO = FirstListValue(wsIn, GetDynamicInputListRange(wsIn, "A", INPUT_WO_LIST_FIRST_ROW).Address(False, False))
    If Len(wpsList) = 0 Then wpsList = GetSelectedWPSList(wsIn)

    scopeWO = FormatPrefixedValue(scopeWO, "WO-")
    scopeDrawing = FormatPrefixedValue(scopeDrawing, "DWG-")

    txt = Replace(txt, "{WO}", scopeWO)
    If Len(scopeDrawing) = 0 Then
        txt = RemoveBlankDrawingPhrase(txt)
    Else
        txt = Replace(txt, "{DRAWING}", scopeDrawing)
    End If
    txt = Replace(txt, "{WPS_LIST}", wpsList)
    txt = Replace(txt, "{ITP}", NormalizeITPNumber(CStr(GetInputValue(wsIn, "ITP Number"))))
    txt = Replace(txt, "{CUSTOMER}", CStr(GetInputValue(wsIn, "Customer")))
    txt = Replace(txt, "{CLIENT}", CStr(GetInputValue(wsIn, "Client")))
    txt = Replace(txt, "{TAG}", CStr(GetInputValue(wsIn, "Tag Number")))
    txt = Replace(txt, "{DESCRIPTION}", CStr(GetInputValue(wsIn, "Description")))
    txt = Replace(txt, "{REV}", CStr(GetInputValue(wsIn, "Revision")))

    txt = Replace(txt, "{ELITE_FOREMAN}", CStr(GetInputValue(wsIn, "Elite Foreman")))
    txt = Replace(txt, "{ELITE_REP}", CStr(GetInputValue(wsIn, "Elite Representative")))
    txt = Replace(txt, "{ELITE_REPRESENTATIVE}", CStr(GetInputValue(wsIn, "Elite Representative")))

    txt = Replace(txt, "{CLIENT_REP}", clientSiteRep)
    txt = Replace(txt, "{CLIENT_REPRESENTATIVE}", clientSiteRep)
    txt = Replace(txt, "{CLIENT_SITE_REP}", clientSiteRep)
    txt = Replace(txt, "{CLIENT_QA}", clientSiteRep)
    txt = Replace(txt, "{CLIENT_SITE_QA}", clientSiteRep)

    ReplaceTokens = txt

End Function

Private Function StripLeadingSectionNumber(ByVal txt As String) As String

    Dim p As Long
    Dim firstPart As String

    txt = Trim(txt)
    p = InStr(1, txt, " ")

    If p > 0 Then
        firstPart = Left(txt, p - 1)

        If firstPart Like "#.#" Or _
           firstPart Like "##.#" Or _
           firstPart Like "#.##" Or _
           firstPart Like "##.##" Then
            txt = Mid(txt, p + 1)
        End If
    End If

    StripLeadingSectionNumber = Trim(txt)

End Function

Private Sub DeleteSheetIfExists(ByVal sheetName As String)

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            ws.Delete
            Exit Sub
        End If
    Next ws

End Sub

Private Function GetGeneratedITPWorksheet() As Worksheet

    Dim ws As Worksheet
    Dim storedSheetName As String

    storedSheetName = GetStoredGeneratedITPSheetName()
    If Len(storedSheetName) > 0 Then
        On Error Resume Next
        Set GetGeneratedITPWorksheet = ThisWorkbook.Worksheets(storedSheetName)
        On Error GoTo 0
        If Not GetGeneratedITPWorksheet Is Nothing Then Exit Function
    End If

    For Each ws In ThisWorkbook.Worksheets
        If IsGeneratedITPSheet(ws) Then
            Set GetGeneratedITPWorksheet = ws
            Exit Function
        End If
    Next ws

    On Error Resume Next
    Set GetGeneratedITPWorksheet = ThisWorkbook.Worksheets(OUTPUT_SHEET)
    On Error GoTo 0

End Function

Private Function IsGeneratedITPSheet(ByVal ws As Worksheet) As Boolean

    Dim markerValue As String

    On Error Resume Next
    markerValue = CStr(ws.CustomProperties(OUTPUT_SHEET_MARKER).Value)
    On Error GoTo 0

    IsGeneratedITPSheet = (StrComp(markerValue, OUTPUT_SHEET, vbTextCompare) = 0)

End Function

Private Sub MarkGeneratedITPSheet(ByVal ws As Worksheet)

    On Error Resume Next
    ws.CustomProperties(OUTPUT_SHEET_MARKER).Delete
    On Error GoTo 0

    ws.CustomProperties.Add Name:=OUTPUT_SHEET_MARKER, Value:=OUTPUT_SHEET

End Sub

Private Sub ClearGeneratedITPSheetMarker(ByVal ws As Worksheet)

    On Error Resume Next
    ws.CustomProperties(OUTPUT_SHEET_MARKER).Delete
    On Error GoTo 0

End Sub

Private Function GetStoredGeneratedITPSheetName() As String

    Dim nameRef As Name
    Dim storedValue As String

    On Error Resume Next
    Set nameRef = ThisWorkbook.Names(OUTPUT_SHEET_NAME_REF)
    On Error GoTo 0

    If nameRef Is Nothing Then Exit Function

    storedValue = CStr(nameRef.RefersTo)
    If Left$(storedValue, 1) = "=" Then storedValue = Mid$(storedValue, 2)
    If Left$(storedValue, 1) = Chr$(34) And Right$(storedValue, 1) = Chr$(34) Then
        storedValue = Mid$(storedValue, 2, Len(storedValue) - 2)
    End If

    GetStoredGeneratedITPSheetName = Replace(storedValue, Chr$(34) & Chr$(34), Chr$(34))

End Function

Private Sub SetStoredGeneratedITPSheetName(ByVal sheetName As String)

    Dim quotedName As String

    quotedName = "=" & Chr$(34) & Replace(sheetName, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)

    On Error Resume Next
    ThisWorkbook.Names(OUTPUT_SHEET_NAME_REF).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:=OUTPUT_SHEET_NAME_REF, RefersTo:=quotedName

End Sub

Private Sub ClearStoredGeneratedITPSheetName()

    On Error Resume Next
    ThisWorkbook.Names(OUTPUT_SHEET_NAME_REF).Delete
    On Error GoTo 0

End Sub

Private Function GetGeneratedITPSheetIndex() As Long

    Dim ws As Worksheet

    Set ws = GetGeneratedITPWorksheet()
    If Not ws Is Nothing Then GetGeneratedITPSheetIndex = ws.Index

End Function

Private Sub DeleteGeneratedITPSheetIfExists()

    Dim ws As Worksheet

    Set ws = GetGeneratedITPWorksheet()
    If Not ws Is Nothing Then ws.Delete
    ClearStoredGeneratedITPSheetName

End Sub

Private Sub RenameWorksheetForArchivedState(ByVal wsOut As Worksheet)

    Dim archivedName As String

    archivedName = BuildArchivedWorksheetName(wsOut.Name)
    archivedName = EnsureUniqueWorksheetName(archivedName, wsOut)

    On Error Resume Next
    wsOut.Name = archivedName
    On Error GoTo 0

End Sub

Private Sub RenameGeneratedITPSheet(ByVal wsOut As Worksheet, ByVal wsIn As Worksheet)

    Dim desiredName As String
    Dim finalName As String

    desiredName = BuildGeneratedITPSheetName(wsIn)
    If Len(desiredName) = 0 Then desiredName = OUTPUT_SHEET
    finalName = EnsureUniqueWorksheetName(desiredName, wsOut)

    On Error Resume Next
    wsOut.Name = finalName
    On Error GoTo 0

    SetStoredGeneratedITPSheetName wsOut.Name

End Sub

Private Function BuildGeneratedITPSheetName(ByVal wsIn As Worksheet, Optional ByVal includeWorkingSuffix As Boolean = True) As String

    Dim baseName As String
    Dim revisionText As String

    baseName = NormalizeITPNumber(CStr(GetInputValue(wsIn, "ITP Number")))
    revisionText = Trim(CStr(wsIn.Range("E6").Value))

    If Len(baseName) = 0 Then Exit Function
    If Len(revisionText) > 0 Then baseName = baseName & " (Rev-" & revisionText & ")"

    baseName = SanitizeWorksheetName(baseName)

    If includeWorkingSuffix Then
        BuildGeneratedITPSheetName = AddWorkingSheetSuffix(baseName)
    Else
        BuildGeneratedITPSheetName = baseName
    End If

End Function

Private Function AddWorkingSheetSuffix(ByVal sheetName As String) As String

    Dim maxBaseLength As Long

    sheetName = RemoveWorkingSheetSuffix(sheetName)
    maxBaseLength = 31 - Len(WORKING_SHEET_SUFFIX)

    If Len(sheetName) > maxBaseLength Then sheetName = Left$(sheetName, maxBaseLength)

    AddWorkingSheetSuffix = sheetName & WORKING_SHEET_SUFFIX

End Function

Private Function RemoveWorkingSheetSuffix(ByVal sheetName As String) As String

    If Len(sheetName) >= Len(LEGACY_WORKING_SHEET_SUFFIX) Then
        If StrComp(Right$(sheetName, Len(LEGACY_WORKING_SHEET_SUFFIX)), LEGACY_WORKING_SHEET_SUFFIX, vbTextCompare) = 0 Then
            RemoveWorkingSheetSuffix = Left$(sheetName, Len(sheetName) - Len(LEGACY_WORKING_SHEET_SUFFIX))
            Exit Function
        End If
    End If

    If Len(sheetName) >= Len(WORKING_SHEET_SUFFIX) Then
        If StrComp(Right$(sheetName, Len(WORKING_SHEET_SUFFIX)), WORKING_SHEET_SUFFIX, vbTextCompare) = 0 Then
            RemoveWorkingSheetSuffix = Left$(sheetName, Len(sheetName) - Len(WORKING_SHEET_SUFFIX))
            Exit Function
        End If
    End If

    RemoveWorkingSheetSuffix = sheetName

End Function

Private Function BuildArchivedWorksheetName(ByVal sheetName As String) As String

    sheetName = SanitizeWorksheetName(RemoveWorkingSheetSuffix(sheetName))
    If Len(sheetName) = 0 Then sheetName = OUTPUT_SHEET

    BuildArchivedWorksheetName = sheetName

End Function

Private Function SanitizeWorksheetName(ByVal sheetName As String) As String

    Dim invalidChars As Variant
    Dim i As Long

    sheetName = Replace(sheetName, vbCr, " ")
    sheetName = Replace(sheetName, vbLf, " ")
    invalidChars = Array(":", "\\", "/", "?", "*", "[", "]")

    For i = LBound(invalidChars) To UBound(invalidChars)
        sheetName = Replace(sheetName, CStr(invalidChars(i)), "-")
    Next i

    sheetName = Trim(sheetName)
    If Len(sheetName) > 31 Then sheetName = Left$(sheetName, 31)

    SanitizeWorksheetName = sheetName

End Function

Private Function EnsureUniqueWorksheetName(ByVal baseName As String, ByVal targetSheet As Worksheet) As String

    Dim tryName As String
    Dim suffixNo As Long

    tryName = baseName

    Do While WorksheetNameExists(tryName, targetSheet)
        suffixNo = suffixNo + 1
        tryName = Left$(baseName, 31 - Len(CStr(suffixNo)) - 3) & " (" & CStr(suffixNo) & ")"
    Loop

    EnsureUniqueWorksheetName = tryName

End Function

Private Function WorksheetNameExists(ByVal sheetName As String, ByVal excludeSheet As Worksheet) As Boolean

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            If excludeSheet Is Nothing Then
                WorksheetNameExists = True
                Exit Function
            ElseIf ws.Index <> excludeSheet.Index Then
                WorksheetNameExists = True
                Exit Function
            End If
        End If
    Next ws

End Function

Private Function GetSheetIndex(ByVal sheetName As String) As Long

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            GetSheetIndex = ws.Index
            Exit Function
        End If
    Next ws

    GetSheetIndex = 0

End Function

Public Sub ExportGeneratedITPToPDF()

    Dim wsOut As Worksheet
    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String

    Set wsOut = GetGeneratedITPWorksheet()

    folderPath = GetWorkbookFolderPath()
    If Len(folderPath) = 0 Then
        MsgBox "Could not resolve a local folder for PDF export." & vbCrLf & vbCrLf & _
               "Open/save the workbook from your local OneDrive folder, then try again.", _
               vbExclamation, "Elite ITP Builder"
        Exit Sub
    End If

    fileName = CleanFileName( _
        NormalizeITPNumber(CStr(ThisWorkbook.Worksheets("Inputs").Range("E5").Value)) & "_" & _
        CStr(ThisWorkbook.Worksheets("Inputs").Range("B7").Value) & ".pdf")

    fullPath = folderPath & Application.PathSeparator & fileName

    wsOut.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=fullPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

End Sub

Private Function CleanFileName(ByVal s As String) As String

    Dim badChars As Variant
    Dim i As Long

    badChars = Array("<", ">", ":", Chr(34), "/", "\", "|", "?", "*")

    For i = LBound(badChars) To UBound(badChars)
        s = Replace(s, CStr(badChars(i)), "-")
    Next i

    CleanFileName = Trim(s)

End Function

Public Sub SetupITPButtons()

    Dim ws As Worksheet
    Dim shp As Shape

    Set ws = ThisWorkbook.Worksheets("Inputs")

    For Each shp In ws.Shapes
        If shp.Name Like "ITP_*Button*" Then shp.Delete
    Next shp

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("A16").Left, ws.Range("A16").Top, ws.Range("A16:B17").Width, ws.Range("A16:B17").Height)

    shp.Name = "ITP_Generate_Button"
    shp.TextFrame.Characters.Text = "GENERATE ITP"
    shp.Fill.ForeColor.RGB = RGB(112, 173, 71)
    shp.Line.ForeColor.RGB = RGB(80, 130, 50)
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.Characters.Font.Bold = True
    shp.OnAction = "GenerateITP"

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("D16").Left, ws.Range("D16").Top, ws.Range("D16:E17").Width, ws.Range("D16:E17").Height)

    shp.Name = "ITP_PDF_Button"
    shp.TextFrame.Characters.Text = "EXPORT PDF"
    shp.Fill.ForeColor.RGB = RGB(68, 114, 196)
    shp.Line.ForeColor.RGB = RGB(47, 84, 150)
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.Characters.Font.Bold = True
    shp.OnAction = "ExportGeneratedITPToPDF"

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("A18").Left, ws.Range("A18").Top, ws.Range("A18:B19").Width, ws.Range("A18:B19").Height)

    shp.Name = "ITP_ClearInputs_Button"
    shp.TextFrame.Characters.Text = "CLEAR ALL INPUT"
    shp.Fill.ForeColor.RGB = RGB(192, 0, 0)
    shp.Line.ForeColor.RGB = RGB(128, 0, 0)
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.Characters.Font.Bold = True
    shp.OnAction = "RefreshInputsSheet"

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("D18").Left, ws.Range("D18").Top, ws.Range("D18:E19").Width, ws.Range("D18:E19").Height)

    shp.Name = "ITP_RefreshScope_Button"
    shp.TextFrame.Characters.Text = "REFRESH BLOCK LIST"
    shp.Fill.ForeColor.RGB = RGB(112, 48, 160)
    shp.Line.ForeColor.RGB = RGB(80, 35, 120)
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.Characters.Font.Bold = True
    shp.OnAction = "RefreshScopeList"

    ' Client logo UI disabled for now (code kept for possible restore):
    ' Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("A20").Left, ws.Range("A20").Top, ws.Range("A20:B21").Width, ws.Range("A20:B21").Height)
    ' shp.Name = "ITP_AddClientLogo_Button"
    ' shp.TextFrame.Characters.Text = "ADD CLIENT LOGO"
    ' shp.Fill.ForeColor.RGB = RGB(237, 125, 49)
    ' shp.Line.ForeColor.RGB = RGB(197, 90, 17)
    ' shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    ' shp.TextFrame.Characters.Font.Bold = True
    ' shp.OnAction = "AddClientHeaderLogo"
    '
    ' Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("D20").Left, ws.Range("D20").Top, ws.Range("D20:E21").Width, ws.Range("D20:E21").Height)
    ' shp.Name = "ITP_ClearClientLogo_Button"
    ' shp.TextFrame.Characters.Text = "CLEAR CLIENT LOGO"
    ' shp.Fill.ForeColor.RGB = RGB(127, 127, 127)
    ' shp.Line.ForeColor.RGB = RGB(89, 89, 89)
    ' shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    ' shp.TextFrame.Characters.Font.Bold = True
    ' shp.OnAction = "ClearClientHeaderLogo"

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("A20").Left, ws.Range("A20").Top, ws.Range("A20:B21").Width, ws.Range("A20:B21").Height)

    shp.Name = "ITP_AddEliteLogo_Button"
    shp.TextFrame.Characters.Text = "ADD LOGO"
    shp.Fill.ForeColor.RGB = RGB(237, 125, 49)
    shp.Line.ForeColor.RGB = RGB(197, 90, 17)
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.Characters.Font.Bold = True
    shp.OnAction = "AddEliteHeaderLogo"

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("D20").Left, ws.Range("D20").Top, ws.Range("D20:E21").Width, ws.Range("D20:E21").Height)

    shp.Name = "ITP_ClearEliteLogo_Button"
    shp.TextFrame.Characters.Text = "CLEAR LOGO"
    shp.Fill.ForeColor.RGB = RGB(127, 127, 127)
    shp.Line.ForeColor.RGB = RGB(89, 89, 89)
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.Characters.Font.Bold = True
    shp.OnAction = "ClearEliteHeaderLogo"

End Sub




