# Frequently Used VBA Codes
## Contents
  1. [ExcelBusy - Function to Turn off/on Screen Updating & etc when running subroutines](https://github.com/amirmirmirdan/Excel-VBA/edit/main/Best_Practices.md#excelbusy---function-to-turn-offon-screen-updating--etc-when-running-subroutines).
  2. 

---


## ExcelBusy - Function to Turn off/on Screen Updating & etc when running subroutines.

  Purpose: To increase the speed of the subroutine process when it is running.

  Input: Boolean (True/False)

  Output: 


```vb

  Sub ExcelBusy(bUpdate As Boolean)
      If bUpdate = True Then
          StartRoutine
      Else
          EndRoutine
      End If
  End Sub

  Private Sub StartRoutine()
      Application.DisplayAlerts = False
      Application.ScreenUpdating = False
      Application.DisplayStatusBar = False
      Application.EnableEvents = False
      ActiveSheet.DisplayPageBreaks = False
      Application.Calculation = xlCalculationManual
  End Sub
  
  Private Sub EndRoutine()
      Application.DisplayAlerts = True
      Application.ScreenUpdating = True
      Application.DisplayStatusBar = True
      Application.EnableEvents = True
      ActiveSheet.DisplayPageBreaks = False
      Application.Calculation = xlCalculationAutomatic
  End Sub

```
