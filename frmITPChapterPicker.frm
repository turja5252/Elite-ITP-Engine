VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmITPChapterPicker 
   Caption         =   "UserForm1"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9465.001
   OleObjectBlob   =   "frmITPChapterPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmITPChapterPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private Sub UserForm_Initialize()
    Dim wsIn As Worksheet
    Dim wsLists As Worksheet
    Dim scopeRange As Range

    Me.Caption = "Add Chapter Block"
    Me.Height = 246
    Me.StartUpPosition = 1

    ApplyLayout

    Set wsIn = ThisWorkbook.Worksheets("Inputs")
    Set wsLists = ThisWorkbook.Worksheets("Lists")

    Module1.RefreshScopeList

    Set scopeRange = Module1.GetScopeBlockRangeForCurrentInspection(wsIn, wsLists)
    If Not scopeRange Is Nothing Then AddChoicesFromRange Me.cboScope, scopeRange

    AddChoicesFromRange Me.cboWO, Module1.GetDynamicInputListRange(wsIn, "A", 22)
    AddChoicesFromRange Me.cboDWG, Module1.GetDynamicInputListRange(wsIn, "D", 22)

    Me.txtCustomRows.Value = "5"
    Me.lblHint.Caption = "Adding new chapter after:" & vbCrLf & Module1.ChapterPickerAfterChapterText & vbCrLf & vbCrLf & "WO and Drawing can be selected or typed. Custom rows apply only when Scope Block is Custom."
    ToggleCustomFields
End Sub

Private Sub UserForm_Activate()
    ApplyLayout
End Sub

Private Sub cboScope_Change()
    ToggleCustomFields
End Sub

Private Sub btnOK_Click()
    If Len(Trim(Me.cboScope.Value)) = 0 Then
        MsgBox "Select a scope block.", vbExclamation, "Elite ITP Builder"
        Exit Sub
    End If

    If UCase$(Trim(Me.cboScope.Value)) = "CUSTOM" Then
        If Len(Trim(Me.txtCustomTitle.Value)) = 0 Then Me.txtCustomTitle.Value = "Custom"
        If Not IsNumeric(Trim(Me.txtCustomRows.Value)) Or CLng(Trim(Me.txtCustomRows.Value)) <= 0 Then
            MsgBox "Enter a whole number greater than zero for blank rows.", vbExclamation, "Elite ITP Builder"
            Exit Sub
        End If
    End If

    Module1.ChapterPickerScopeName = Trim(Me.cboScope.Value)
    Module1.ChapterPickerWOText = Trim(Me.cboWO.Value)
    Module1.ChapterPickerDWGText = Trim(Me.cboDWG.Value)
    Module1.ChapterPickerCustomTitle = Trim(Me.txtCustomTitle.Value)
    Module1.ChapterPickerCustomRows = Val(Me.txtCustomRows.Value)
    If Module1.ChapterPickerCustomRows <= 0 Then Module1.ChapterPickerCustomRows = 5
    Module1.ChapterPickerConfirmed = True
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub AddChoicesFromRange(ByVal combo As Object, ByVal sourceRange As Range)
    Dim c As Range

    combo.Clear
    For Each c In sourceRange.Cells
        If Len(Trim(CStr(c.Value))) > 0 Then combo.AddItem Trim(CStr(c.Value))
    Next c
End Sub

Private Sub ToggleCustomFields()
    Dim isCustom As Boolean

    isCustom = (UCase$(Trim(Me.cboScope.Value)) = "CUSTOM")
    Me.lblCustomTitle.Visible = isCustom
    Me.txtCustomTitle.Visible = isCustom
    Me.lblCustomRows.Visible = isCustom
    Me.txtCustomRows.Visible = isCustom
End Sub

Private Sub ApplyLayout()
    Const FORM_WIDTH As Single = 460
    Const FIELD_LEFT As Single = 108
    Const FIELD_RIGHT_PADDING As Single = 18
    Const BUTTON_SPACING As Single = 12
    Const MAX_FIELD_WIDTH As Single = 300

    Dim fieldWidth As Single

    Me.Width = FORM_WIDTH

    fieldWidth = Me.InsideWidth - FIELD_LEFT - FIELD_RIGHT_PADDING
    If fieldWidth < 180 Then fieldWidth = 180
    If fieldWidth > MAX_FIELD_WIDTH Then fieldWidth = MAX_FIELD_WIDTH

    Me.cboScope.Left = FIELD_LEFT
    Me.cboScope.Width = fieldWidth
    Me.cboScope.ListWidth = fieldWidth

    Me.cboWO.Left = FIELD_LEFT
    Me.cboWO.Width = fieldWidth
    Me.cboWO.ListWidth = fieldWidth

    Me.cboDWG.Left = FIELD_LEFT
    Me.cboDWG.Width = fieldWidth
    Me.cboDWG.ListWidth = fieldWidth

    Me.txtCustomTitle.Left = FIELD_LEFT
    Me.txtCustomTitle.Width = fieldWidth

    Me.txtCustomRows.Left = FIELD_LEFT

    Me.lblHint.Width = Me.InsideWidth - Me.lblHint.Left - FIELD_RIGHT_PADDING

    Me.btnCancel.Left = Me.InsideWidth - FIELD_RIGHT_PADDING - Me.btnCancel.Width
    Me.btnOK.Left = Me.btnCancel.Left - BUTTON_SPACING - Me.btnOK.Width
End Sub

