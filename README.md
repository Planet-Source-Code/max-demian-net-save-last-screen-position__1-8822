<div align="center">

## Save last screen position


</div>

### Description

Save the last screen position of your form.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Max \- Demian Net](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/max-demian-net.md)
**Level**          |Intermediate
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/max-demian-net-save-last-screen-position__1-8822/archive/master.zip)





### Source Code

```
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Function GetFromINI(Section As String, Key As String, Directory As String) As String
  Dim strBuffer As String
  strBuffer = String(750, Chr(0))
  Key$ = LCase$(Key$)
  GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function
Private Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
  Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub
Private Sub Form_Load()
On Error Resume Next
Form1.Top = GetFromINI("SCREEN", "TOP", App.Path & "\screen.ini")
Form1.Left = GetFromINI("SCREEN", "LEFT", App.Path & "\screen.ini")
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
WriteToINI "SCREEN", "TOP", Form1.Top, App.Path & "\screen.ini"
WriteToINI "SCREEN", "LEFT", Form1.Left, App.Path & "\screen.ini"
End Sub
```

