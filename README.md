<div align="center">

## Get User Name


</div>

### Description

Returns the current user name using a dll call
 
### More Info
 
The current user that is logged on


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Grove](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-grove.md)
**Level**          |Unknown
**User Rating**    |4.0 (12 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-grove-get-user-name__1-891/archive/master.zip)

### API Declarations

```
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long
```


### Source Code

```
Function GetUser()
 ' This function uses a windows dll to query the registry automatically ti return the user name
 Dim sBuffer As String
 Dim lSize As Long
 ' Parameters for the dll declaration are set
 sBuffer = Space$(255)
 lSize = Len(sBuffer)
 Call GetUserName(sBuffer, lSize)   ' Call the declared dll function
If lSize > 0 Then
 GetUser = Left$(sBuffer, lSize)   ' Remove empty spaces
Else
 GetUser = vbNullString   ' Return empty if no user is found
End If
End Function
```

