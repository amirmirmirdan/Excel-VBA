## Listbox

### List Collection
```vb
Private Function ListFiles() As Collection
Dim cRows As Long, i As Long
Dim coll As Collection
    Set coll = New Collection
    i = 1
    cRows = Sheet1.Range("A1").CurrentRegion.Offset(1, 0).Rows.Count
    For i = 1 To cRows
        coll.Add Sheet1.Range("A1").Offset(i, 0).Value
    Next i
    Set ListFiles = coll
End Function
```
### Add Collection to Listbox
```vb
Private Sub UserForm_Initialize()
Dim coll As Collection
Set coll = New Collection
    Set coll = ListFiles
Dim i As Long, cList As Long
i = 1
cList = coll.Count
    For i = 1 To cList    
        With Me.ListBox1
            .AddItem coll.Item(i)
        End With
    Next i
Me.Show
End Sub
```
