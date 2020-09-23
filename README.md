<div align="center">

## Forms Unloader


</div>

### Description

This code unloads all the forms of the program returning the resources back to the computer
 
### More Info
 
Optional force input as a boolean

This code is pretty straight forward and an understanding of loops and arrays is will help.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[rbennett](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rbennett.md)
**Level**          |Intermediate
**User Rating**    |4.0 (16 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rbennett-forms-unloader__1-4963/archive/master.zip)





### Source Code

```
Private Sub unloader(Optional ByVal ForceClose As Boolean = False)
  Dim i As Long
On Error Resume Next
  For i = Forms.Count - 1 To 0 Step -1
    Unload Forms(i)
    Set Forms(i) = Nothing
    If Not ForceClose Then
      If Forms.Count > i Then
        Exit Sub
      End If
    End If
  Next i
  If ForceClose Or (Forms.Count = 0) Then Close
  If ForceClose Or (Forms.Count > 0) Then End
End Sub
```

