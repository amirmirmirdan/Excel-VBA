*Displays the standard **Open** dialog box and gets a file name from the user without actually opening any files.*

**Syntax**

*expression*.**GetOpenFilename** (*FileFilter*, *FilterIndex*, *Title*, *ButtonText*, *MultiSelect*)

#### Example:
```vb
fileToOpen = Application.GetOpenFilename("csv file (*.csv), *.csv") 
If fileToOpen <> False Then 
 MsgBox "Open " & fileToOpen 
End If
```
