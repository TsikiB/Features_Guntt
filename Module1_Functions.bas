Attribute VB_Name = "Module1_Functions"
Public Sub ClickToData()

Dim ErrMsg As String

    ActiveRow = Mid(ActiveCell.Address, InStrRev(ActiveCell.Address, "$") + 1)
    If ActiveRow < 3 Then
        ErrMsg = "You must select a cell with valid data"
        MsgBox ErrMsg, vbExclamation, "Au10tix - Features Guntt"
        Exit Sub
    End If
    
    If UCase(ActiveWorkbook.ActiveSheet.Name) = "FEATURE TIMELINE" Then
        If GoToAddress(ActiveCell.Address) Then
        End If
    ElseIf UCase(ActiveWorkbook.ActiveSheet.Name) = "TFS DATA" Then
        If BackToReference(ActiveCell.Address) Then
        End If
    Else
        ErrMsg = "This operation valid only from 'Feature Timeline' worksheet"
        MsgBox ErrMsg, vbExclamation, "Au10tix - Features Guntt"
    End If
 
End Sub
Public Sub RefreshTFSData()

    Call RefreshData(ActiveCell.Address)
 
End Sub

Function GoToAddress(Address) As String
On Error GoTo MissedData

Dim GoToRow, ErrMsg As String
Dim GoToRange As Range

GoToRow = 0
 Set GoToRange = Worksheets(UCase("FEATURE TIMELINE")).Range(Address)
 
 GoToRow = Application.WorksheetFunction.Match(GoToRange.Value, Worksheets(UCase("TFS DATA")).Columns(1), 0)
 If GoToRow <> 0 Or IsEmpty(GoToRow) Then
    Application.Goto Reference:=Worksheets(UCase("TFS DATA")).Range("A" & GoToRow), scroll:=True
    Set GoToRange = Nothing
    GoToAddress = True
 End If
 
MissedData:
    If Err.Number = 1004 Then
        ErrMsg = "Selected feature not found on TFS data"
        MsgBox ErrMsg, vbExclamation, "Au10tix - Features Guntt"
        GoToAddress = False
    End If

End Function

Function BackToReference(Address) As String
On Error GoTo MissedData

Dim GoToRow, ErrMsg As String
Dim GoToRange As Range

GoToRow = 0
 Set GoToRange = Worksheets(UCase("TFS DATA")).Range(Address)
 
 GoToRow = Application.WorksheetFunction.Match(GoToRange.Value, Worksheets(UCase("FEATURE TIMELINE")).Columns(1), 0)
 If GoToRow <> 0 Or IsEmpty(GoToRow) Then
    Application.Goto Reference:=Worksheets(UCase("FEATURE TIMELINE")).Range("A" & GoToRow), scroll:=True
    Set GoToRange = Nothing
    BackToReference = True
 End If
 
MissedData:
    If Err.Number = 1004 Then
        ErrMsg = "Selected feature not found on Feature Timeline"
        MsgBox ErrMsg, vbExclamation, "Au10tix - Features Guntt"
        BackToReference = False
    End If

End Function

Function RefreshData(Address) As Boolean

On Error Resume Next

'-- Refresh TFS data
    Sheets("TFS Data").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("TFS Data").ListObjects( _
        "VSTS_1767b646_5ecb_4441_83ba_052a656d849c").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("TFS Data").ListObjects( _
        "VSTS_1767b646_5ecb_4441_83ba_052a656d849c").Sort.SortFields.Add2 Key:=Range( _
        "VSTS_1767b646_5ecb_4441_83ba_052a656d849c[ID]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("TFS Data").ListObjects( _
        "VSTS_1767b646_5ecb_4441_83ba_052a656d849c").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'-- Return focus to starting worksheet/cell
    Sheets("Feature Timeline").Select
    Range(Address).Select

End Function

