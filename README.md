<div align="center">

## Copy Files Using Copy Progress Dialog


</div>

### Description

Copy a file the using SHFileOperation API call so that Windows copy progress dialog appears...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Munim\.VIP](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/munim-vip.md)
**Level**          |Intermediate
**User Rating**    |4.5 (59 globes from 13 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/munim-vip-copy-files-using-copy-progress-dialog__1-28983/archive/master.zip)

### API Declarations

```
'// Please Vote 4 Me, If You Like This Code...
Private Type SHFILEOPSTRUCT
   hWnd As Long
   wFunc As Long
   pFrom As String
   pTo As String
   fFlags As Integer
   fAnyOperationsAborted As Boolean
   hNameMappings As Long
   lpszProgressTitle As String
End Type
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Const FO_COPY = &H2
Private Const FOF_ALLOWUNDO = &H40
```


### Source Code

```
Public Sub CopyFileWindowsWay(SourceFile As String, DestinationFile As String)
   Dim lngReturn As Long
   Dim typFileOperation As SHFILEOPSTRUCT
   With typFileOperation
    .hWnd = 0
    .wFunc = FO_COPY
    .pFrom = SourceFile & vbNullChar & vbNullChar 'source file
    .pTo = DestinationFile & vbNullChar & vbNullChar 'destination file
    .fFlags = FOF_ALLOWUNDO
   End With
   lngReturn = SHFileOperation(typFileOperation)
   If lngReturn <> 0 Then 'Operation failed
     MsgBox Err.LastDllError, vbCritical Or vbOKOnly
   Else 'Aborted
     If typFileOperation.fAnyOperationsAborted = True Then
        MsgBox "Operation Failed", vbCritical Or vbOKOnly
     End If
   End If
End Sub
```

