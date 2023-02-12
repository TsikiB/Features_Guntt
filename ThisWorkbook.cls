VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub Workbook_RigntClickMenu()

[COLOR=#008080]  ' Dim ContextMenu As CommandBar[/COLOR] <-not needed but left in the code in case any errors occur and I can turn it back on

  Set ContextMenu = Application.CommandBars("cell")

  With Application.CommandBars("cell")
        .Reset
        With ContextMenu.Controls.Add(Type:=msoControlButton, before:=1)
                   .FaceId = 338
                   .Caption = "Go To Reference"
                   .OnAction = "ClickToData"
        End With
        With ContextMenu.Controls.Add(Type:=msoControlButton, before:=2)
                   .FaceId = 37
                   .Caption = "Refresh TFS Data"
                   .OnAction = "RefreshTFSData"
        End With
        With ContextMenu.Controls.Add(Type:=msoControlButton, before:=3)
                   .Caption = "------------------"
        End With
  End With
End Sub



'Private Sub Workbook_Open()

'Dim MyMenu As Object
 
'    Set MyMenu = Application.ShortcutMenus(xlWorksheetCell) _
'        .MenuItems.AddMenu("Program Management", 1)
'
'    With MyMenu.MenuItems
'        .Add "GoTo Data", "ClickToData", , 1, , ""
'        .Add "Back To Reference", "BackToReference", , 2, , ""
'        .Add "Refresh TFS Data", "RefreshTFSData", , 3, , ""
'    End With
    
'    Set MyMenu = Nothing

'End Sub

