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

Dim Address As String

If UCase(Application.ActiveSheet.Name) = UCase("FEATURE TIMELINE") Then
    Address = ActiveCell.Address
Else
    Address = "$A$2"
End If
    Call RefreshData(Address)

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
RefreshTFSQuery ("TFS Data")

'-- Return focus to starting worksheet/cell
    Sheets(UCase("FEATURE TIMELINE")).Select
    Range(Address).Select

End Function

Public Function RefreshTFSQuery(wshName)

On Error GoTo Err

Dim PauseCounter As Integer
Dim Flag_RefreshTFS As Boolean
Dim RefreshButton As CommandBarControl
    
PauseCounter = 0
Flag_RefreshTFS = False

   ActiveWorkbook.Sheets(wshName).Visible = xlSheetVisible
   Set RefreshButton = Application.CommandBars.FindControl(Tag:="IDC_REFRESH")

    With Application.Worksheets(wshName).Select
        Do While PauseCounter < 5
            If RefreshButton.Enabled Then
                RefreshButton.Execute
                Flag_RefreshTFS = True
                Exit Do
            Else
                Application.Wait (Now + TimeValue("0:00:02"))
                PauseCounter = PauseCounter + 1
            End If
        Loop
        
        Set RefreshButton = Nothing
        If Not Flag_RefreshTFS Then
            MsgBox "Warning: Refresh Button is not avilable !!"
            Exit Function
        End If

    End With

RefreshTFSQuery = True

Err:
    Application.StatusBar = Err.Description
    TheErrorText = Err.Description

End Function
