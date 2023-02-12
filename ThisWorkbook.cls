VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

Dim MyMenu As Object
 
    Set MyMenu = Application.ShortcutMenus(xlWorksheetCell) _
        .MenuItems.AddMenu("Program Management", 1)

    With MyMenu.MenuItems
        .Add "GoTo Data", "ClickToData", , 1, , ""
        .Add "Back To Reference", "BackToReference", , 2, , ""
        .Add "Refresh TFS Data", "RefreshTFSData", , 3, , ""
    End With
    
    Set MyMenu = Nothing

End Sub

