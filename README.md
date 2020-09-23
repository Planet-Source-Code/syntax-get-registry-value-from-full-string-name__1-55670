<div align="center">

## Get Registry Value From Full String Name


</div>

### Description

Get a REG_SZ (string) registry value from the full key name, such as "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\DirectX". Functions for both default and named values.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[syntax\.](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/syntax.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Registry](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/registry__1-36.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/syntax-get-registry-value-from-full-string-name__1-55670/archive/master.zip)

### API Declarations

```
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long   ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" Alias "RegCloseKey" (ByVal hKey As Long) As Long
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_NOTIFY = &H10
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const READ_CONTROL = &H20000
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const ERROR_SUCCESS = 0&
```


### Source Code

```
'Retrieves a REG_SZ registry value
Public Function sGetNamedRegValue(sKey As String, sValue As String) As String
 Dim sHive As String
 Dim hKey As Long
 Dim hHive As Long
 Dim sData As String
 Dim lLenData As Long
 lLenData = 255
 sData = String$(255, 0)
 sHive = Left(sKey, InStr(sKey, "\") - 1)
 sKey = Replace(sKey, sHive & "\", "")
 Select Case sHive
  Case "HKEY_CLASSES_ROOT"
   hHive = HKEY_CLASSES_ROOT
  Case "HKEY_CURRENT_CONFIG"
   hHive = HKEY_CURRENT_CONFIG
  Case "HKEY_CURRENT_USER"
   hHive = HKEY_CURRENT_USER
  Case "HKEY_DYN_DATA"
   hHive = HKEY_DYN_DATA
  Case "HKEY_LOCAL_MACHINE"
   hHive = HKEY_LOCAL_MACHINE
  Case "HKEY_PERFORMANCE_DATA"
   hHive = HKEY_PERFORMANCE_DATA
  Case "HKEY_USERS"
   hHive = HKEY_USERS
 End Select
 Dim lKeyType As Long
 If RegOpenKeyEx(hHive, sKey, 0, KEY_READ, hKey) = ERROR_SUCCESS Then
  If RegQueryValueEx(hKey, sValue, 0, lKeyType, ByVal sData, lLenData) = ERROR_SUCCESS Then
   sGetNamedRegValue = Left(sData, InStr(sData, Chr(0)) - 1)
   RegCloseKey hKey
   Exit Function
  End If
  RegCloseKey hKey
 End If
 sGetNamedRegValue = ""
End Function
'Retrieves a default REG_SZ registry value
Public Function sGetRegValue(sKey As String) As String
 Dim sHive As String
 Dim hKey As Long
 Dim hHive As Long
 Dim sData As String
 Dim lLenData As Long
 lLenData = 255
 sData = String$(255, 0)
 sHive = Left(sKey, InStr(sKey, "\") - 1)
 sKey = Replace(sKey, sHive & "\", "")
 Select Case sHive
  Case "HKEY_CLASSES_ROOT"
   hHive = HKEY_CLASSES_ROOT
  Case "HKEY_CURRENT_CONFIG"
   hHive = HKEY_CURRENT_CONFIG
  Case "HKEY_CURRENT_USER"
   hHive = HKEY_CURRENT_USER
  Case "HKEY_DYN_DATA"
   hHive = HKEY_DYN_DATA
  Case "HKEY_LOCAL_MACHINE"
   hHive = HKEY_LOCAL_MACHINE
  Case "HKEY_PERFORMANCE_DATA"
   hHive = HKEY_PERFORMANCE_DATA
  Case "HKEY_USERS"
   hHive = HKEY_USERS
 End Select
 If RegQueryValue(hHive, sKey, sData, lLenData) = ERROR_SUCCESS Then
  sGetRegValue = Left(sData, InStr(sData, Chr(0)) - 1)
  Exit Function
 End If
 sGetRegValue = ""
End Function
```

