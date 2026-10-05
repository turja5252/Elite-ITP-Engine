Private Const CUSTOM_INPUT_ENABLED_COLOR As Long = 16777215
Private Const CUSTOM_INPUT_DISABLED_COLOR As Long = 14277081
Private Const INPUT_FIRST_ROW As Long = 6
Private Const INPUT_LAST_ROW As Long = 79

Private Sub Worksheet_Activate()

    Module1.RefreshScopeList
    Module1.RefreshInputScopeValidation Me
    RefreshCustomInputFormatting

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If Target.CountLarge > 1 Then Exit Sub
    If Not IsRestrictedCustomInputCell(Target) Then Exit Sub

    Application.EnableEvents = False
    Me.Cells(Target.Row, "H").Select
    Application.EnableEvents = True

    MsgBox "Custom Chapter Title and Custom Rows can only be edited when Scope Block is Custom for that row.", vbInformation, "Elite ITP Builder"

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim OldValue As String
    Dim NewValue As String
    Dim Separator As String

    Separator = ", "

    If Target.CountLarge > 1 Then Exit Sub

    On Error GoTo SafeExit

    If StrComp(Target.Address(False, False), "E10", vbTextCompare) = 0 Then
        Application.EnableEvents = False
        Module1.RefreshScopeList
        Module1.RefreshInputScopeValidation Me
        GoTo SafeExit
    End If

    If Target.Validation.Type <> 3 Then Exit Sub

    NewValue = Trim(CStr(Target.Value))

    If Not Intersect(Target, Me.Range("K" & INPUT_FIRST_ROW & ":L" & INPUT_LAST_ROW)) Is Nothing Then
        If Not IsCustomScopeRow(Target.Row) Then
            Application.EnableEvents = False
            Application.Undo
            Me.Cells(Target.Row, "H").Select
            MsgBox "Custom Chapter Title and Custom Rows can only be edited when Scope Block is Custom for that row.", vbExclamation, "Elite ITP Builder"
            GoTo SafeExit
        End If
        GoTo SafeExit
    End If

    If Not Intersect(Target, Me.Range("H" & INPUT_FIRST_ROW & ":H" & INPUT_LAST_ROW)) Is Nothing Then
        If UCase$(NewValue) <> "CUSTOM" Then
            Application.EnableEvents = False
            Me.Cells(Target.Row, "K").Resize(1, 2).ClearContents
            Application.EnableEvents = True
        End If

        UpdateCustomInputFormatting Target.Row

        If Len(NewValue) > 0 Then
            Module1.ShowScopeBlockPreview Me, NewValue, _
                                          Trim(CStr(Me.Cells(Target.Row, "I").Value)), _
                                          Trim(CStr(Me.Cells(Target.Row, "J").Value))
            If UCase$(NewValue) = "CUSTOM" Then
                Module1.PromptForCustomChapterInputs Me, Target.Row
            End If
        End If
        GoTo SafeExit
    End If

    If Intersect(Target, Me.Range("I" & INPUT_FIRST_ROW & ":J" & INPUT_LAST_ROW)) Is Nothing Then Exit Sub

    Application.EnableEvents = False

    'CLEAR ALL option
    If UCase(NewValue) = "CLEAR ALL" Then
        Target.Value = ""
        GoTo SafeExit
    End If

    If NewValue = "" Then GoTo SafeExit

    Application.Undo
    OldValue = Target.Value

    If OldValue = "" Then
        Target.Value = NewValue
    Else
        If InStr(1, OldValue, NewValue, vbTextCompare) = 0 Then
            Target.Value = OldValue & Separator & NewValue
        Else
            Target.Value = OldValue
        End If
    End If

SafeExit:
    Application.EnableEvents = True

End Sub

Private Function IsCustomScopeRow(ByVal rowNum As Long) As Boolean

    IsCustomScopeRow = (UCase$(Trim$(CStr(Me.Cells(rowNum, "H").Value))) = "CUSTOM")

End Function

Private Sub RefreshCustomInputFormatting()

    Dim rowNum As Long

    For rowNum = INPUT_FIRST_ROW To INPUT_LAST_ROW
        UpdateCustomInputFormatting rowNum
    Next rowNum

End Sub

Private Sub UpdateCustomInputFormatting(ByVal rowNum As Long)

    With Me.Cells(rowNum, "K").Resize(1, 2).Interior
        If IsCustomScopeRow(rowNum) Then
            .Pattern = xlSolid
            .Color = CUSTOM_INPUT_ENABLED_COLOR
        Else
            .Pattern = xlSolid
            .Color = CUSTOM_INPUT_DISABLED_COLOR
        End If
    End With

End Sub

Private Function IsRestrictedCustomInputCell(ByVal Target As Range) As Boolean

    If Target Is Nothing Then Exit Function
    If Target.Row < INPUT_FIRST_ROW Or Target.Row > INPUT_LAST_ROW Then Exit Function
    If Intersect(Target, Me.Range("K" & INPUT_FIRST_ROW & ":L" & INPUT_LAST_ROW)) Is Nothing Then Exit Function

    IsRestrictedCustomInputCell = Not IsCustomScopeRow(Target.Row)

End Function








