### Passing Arrays Into a Function
**Example:**
```vb
Function DynamicArray(ByRef myInternalArray() As Long) As Long
' Code here
End Function
```
---
### Returning Array Out from a Function
**Example:**
```vb
Public Function GetNumbers_ReturnLong() As Long()   
   Dim mynumbers() as Long   
   ReDim mynumbers(1 to 2)   
   mynumbers(1) = 10   
   mynumbers(2) = 20   
   GetNumbers_ReturnLong = mynumbers   
End Function
```
---
