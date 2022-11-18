### Shortcut Menus
```vb
Private Sub Workbook_Open()
Dim MyMenu As Object
    Set MyMenu = Application.ShortcutMenus(xlWorksheetCell) _
        .MenuItems.AddMenu("Custom Shortcut Menu", 1)
                
        With MyMenu.MenuItems
            .Add "MassageUOB_SG", "MassageUOB_SG", , 1, , ""
            .Add "MassageBDO_PH", "MassageBDO_PH", , 2, , ""
            .Add "GPay6421", "GPay6421", , 3, , ""
            .Add "Normal Mode", "NormalMode", , 4, , ""
            ' First Arguement: Display Name
            ' 2nd Arguement: Subroutine Name
        End With
    Set MyMenu = Nothing
End Sub
```
---
### Screen Updating & Etc.
```vb
Sub ExcelIsBusy(bUpdate As Boolean)
    If bUpdate = True Then
        StartRoutine
    Else
        EndRoutine
    End If
End Sub
```
```vb
Public Sub StartRoutine()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.Calculation = xlCalculationManual
End Sub
```
```vb
Public Sub EndRoutine()
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = False
    Application.Calculation = xlCalculationAutomatic
End Sub
```
### Toggle Reference Style
```vb
Sub ToggleReferenceStyle()
    With Application
        If .ReferenceStyle = xlR1C1 Then
            .ReferenceStyle = xlA1
        Else
            .ReferenceStyle = xlR1C1
        End If
    End With
End Sub
```

### File Dialog
[[FileDialog]]

