Attribute VB_Name = "Module1_Functions"
Public Sub ClickToData()

Dim ErrMsg As String

    ActiveRow = Mid(ActiveCell.Address, InStrRev(ActiveCell.Address, "$") + 1)
    If ActiveRow < 3 Then
        ErrMsg = "You must select a cell with valid data"
        MsgBox ErrMsg, vbExclamation, "Au10tix - Features Guntt"
        Exit Sub
    End If
    
    If ActiveWorkbook.ActiveSheet.Name = "Feature Timeline" Then
        Call GoToAddress(ActiveCell.Address)
    Else
        ErrMsg = "This operation valid only from 'Feature Timeline' worksheet"
        MsgBox ErrMsg, vbExclamation, "Au10tix - Features Guntt"
    End If
 
End Sub
Public Sub RefreshTFSData()

    Call RefreshData(ActiveCell.Address)
 
End Sub

Function GoToAddress(Address) As String

Dim GoToRow As String
Dim GoToRange As Range

 Set GoToRange = Worksheets("Feature Timeline").Range(Address)
 
 GoToRow = Application.WorksheetFunction.Match(GoToRange.Value, Worksheets("TFS Data").Columns(1), 0)
 Application.Goto Reference:=Worksheets("TFS Data").Range("A" & GoToRow), scroll:=True
 Set GoToRange = Nothing
 
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

