# Best Practices

## Turn off unnecessary Excel actions

```vb
  Sub ExcelBusy(bUpdate As Boolean)
      If bUpdate = True Then
          StartRoutine
      Else
          EndRoutine
      End If
  End Sub
  Public Sub StartRoutine()
      Application.DisplayAlerts = False
      Application.ScreenUpdating = False
      Application.DisplayStatusBar = False
      Application.EnableEvents = False
      ActiveSheet.DisplayPageBreaks = False
      Application.Calculation = xlCalculationManual
  End Sub
  Public Sub EndRoutine()
      Application.DisplayAlerts = True
      Application.ScreenUpdating = True
      Application.DisplayStatusBar = True
      Application.EnableEvents = True
      ActiveSheet.DisplayPageBreaks = False
      Application.Calculation = xlCalculationAutomatic
  End Sub
```
