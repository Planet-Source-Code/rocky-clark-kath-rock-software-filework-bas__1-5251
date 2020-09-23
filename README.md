<div align="center">

## FileWork\.bas


</div>

### Description

A collection of file associated functions such as GetLongFilename(), GetShortFilename(), GetFilePath(), GetFileTitle(), Exists(), etc.
 
### More Info
 
Great for beginners or power coders. Simply create a new Module, name it something like basFileWork(FileWork.bas), then copy and paste the entire code sample into the 'General/Declarations' section.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rocky Clark \(Kath\-Rock Software\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rocky-clark-kath-rock-software.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rocky-clark-kath-rock-software-filework-bas__1-5251/archive/master.zip)

### API Declarations

Uses GetLongPathname() and GetShortPathname() API functions.


### Source Code

```
'API declarations
Public Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Function AddBackSlash(ByVal sPath As String) As String
'Returns sPath with a trailing backslash if sPath does not
'already have a trailing backslash. Otherwise, returns sPath.
 sPath = Trim$(sPath)
 If Len(sPath) > 0 Then
  sPath = sPath & IIf(Right$(sPath, 1) <> "\", "\", "")
 End If
 AddBackSlash = sPath
End Function
Public Function GetLongFilename(ByVal sShortFilename As String) As String
'Returns the Long Filename associated with sShortFilename
Dim lRet As Long
Dim sLongFilename As String
 'First attempt using 1024 character buffer.
 sLongFilename = String$(1024, " ")
 lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
 'If buffer is too small lRet contains buffer size needed.
 If lRet > Len(sLongFilename) Then
  'Increase buffer size...
  sLongFilename = String$(lRet + 1, " ")
  'and try again.
  lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
 End If
 'lRet contains the number of characters returned.
 If lRet > 0 Then
  GetLongFilename = Left$(sLongFilename, lRet)
 End If
End Function
Public Function GetShortFilename(ByVal sLongFilename As String) As String
'Returns the Short Filename associated with sLongFilename
Dim lRet As Long
Dim sShortFilename As String
 'First attempt using 1024 character buffer.
 sShortFilename = String$(1024, " ")
 lRet = GetShortPathName(sLongFilename, sShortFilename, Len(sShortFilename))
 'If buffer is too small lRet contains buffer size needed.
 If lRet > Len(sShortFilename) Then
  'Increase buffer size...
  sShortFilename = String$(lRet + 1, " ")
  'and try again.
  lRet = GetShortPathName(sLongFilename, sShortFilename, Len(sShortFilename))
 End If
 'lRet contains the number of characters returned.
 If lRet > 0 Then
  GetShortFilename = Left$(sShortFilename, lRet)
 End If
End Function
Public Function RemoveBackSlash(ByVal sPath As String) As String
'Returns sPath without a trailing backslash if sPath
'has one. Otherwise, returns sPath.
 sPath = Trim$(sPath)
 If Len(sPath) > 0 Then
  sPath = Left$(sPath, Len(sPath) - IIf(Right$(sPath, 1) = "\", 1, 0))
 End If
 RemoveBackSlash = sPath
End Function
Public Function AppPath() As String
'Returns App.Path with backslash "\"
Dim sPath As String
 sPath = App.Path
 AppPath = sPath & IIf(Right$(sPath, 1) <> "\", "\", "")
End Function
Public Function Exists(ByVal sFilename As String) As Boolean
'Returns True if File Exists.
'Else returns False.
 If Len(Trim$(sFilename)) > 0 Then
  On Error Resume Next
  sFilename = Dir$(sFilename)
  Exists = ((Err.Number = 0) And (Len(sFilename) > 0))
 Else
  Exists = False
 End If
End Function
Public Function GetFilePath(ByVal sFilename As String, Optional ByVal bAddBackslash As Boolean) As String
'Returns Path Without FileTitle
Dim lPos As Long
 lPos = InStrRev(sFilename, "\")
 If lPos > 0 Then
  GetFilePath = Left$(sFilename, lPos - 1) _
   & IIf(bAddBackslash, "\", "")
 Else
  GetFilePath = ""
 End If
End Function
Public Function GetFileTitle(ByVal sFilename As String) As String
'Returns FileTitle Without Path
Dim lPos As Long
 lPos = InStrRev(sFilename, "\")
 If lPos > 0 Then
  If lPos < Len(sFilename) Then
   GetFileTitle = Mid$(sFilename, lPos + 1)
  Else
   GetFileTitle = ""
  End If
 Else
  GetFileTitle = sFilename
 End If
End Function
```

