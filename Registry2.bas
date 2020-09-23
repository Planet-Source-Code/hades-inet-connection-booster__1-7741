Attribute VB_Name = "Registry"
Option Explicit
Global Const REG_NONE = 0
Global Const REG_SZ = 1
Global Const REG_EXPAND_SZ = 2
Global Const REG_BINARY = 3
Global Const REG_DWORD = 4
Global Const REG_DWORD_LITTLE_ENDIAN = 4
Global Const REG_DWORD_BIG_ENDIAN = 5
Global Const REG_LINK = 6
Global Const REG_MULTI_SZ = 7
Global Const REG_RESOURCE_LIST = 8
Global Const REG_FULL_RESOURCE_DESCRIPTOR = 9
Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003
Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_INVALID_PARAMETER = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259
Global Const KEY_ALL_ACCESS = &H3F
Global Const SYNCHRONIZE = &H100000
Global Const STANDARD_RIGHTS_ALL = &H1F0000
Global Const KEY_QUERY_VALUE = &H1
Global Const KEY_SET_VALUE = &H2
Global Const KEY_CREATE_SUB_KEY = &H4
Global Const KEY_ENUMERATE_SUB_KEYS = &H8
Global Const KEY_NOTIFY = &H10
Global Const KEY_CREATE_LINK = &H20
Global Const REG_OPTION_NON_VOLATILE = 0
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)
Dim hNewKey As Long
Dim lRetVal As Long
lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, _
vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
0&, hNewKey, lRetVal)
RegCloseKey (hNewKey)
End Sub
Sub SetKeyValue(sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
Dim Zero As Long, IRetVal As Long, hKey As Long, OrigKeyNam As String
IRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyName, Zero, KEY_ALL_ACCESS, hKey)
If IRetVal Then MsgBox "RegOpenKey error - " & IRetVal
IRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
If IRetVal Then MsgBox "SetValue error - " & IRetVal
RegCloseKey (hKey)
End Sub
Sub QueryValue(sKeyName As String, sValueName As String)
Dim lRetVal As Long
Dim hKey As Long
Dim vValue As Variant
lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyName, 0, KEY_ALL_ACCESS, hKey)
lRetVal = QueryValueEx(hKey, sValueName, vValue)
MsgBox vValue
RegCloseKey (hKey)
End Sub
Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
Dim lValue As Long
Dim sValue As String
Select Case lType
Case REG_SZ
sValue = vValue & Chr$(0)
SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
Case REG_DWORD
lValue = vValue
SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
End Select
End Function
Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
Dim cch As Long
Dim lrc As Long
Dim lType As Long
Dim lValue As Long
Dim sValue As String
On Error GoTo QueryValueExError
lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
If lrc <> ERROR_NONE Then Error 5
Select Case lType
Case REG_SZ:
sValue = String(cch, 0)
lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
If lrc = ERROR_NONE Then
vValue = Left$(sValue, cch)
Else
vValue = Empty
End If
Case REG_DWORD:
lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
If lrc = ERROR_NONE Then vValue = lValue
Case Else
lrc = -1
End Select
QueryValueExExit:
QueryValueEx = lrc
Exit Function
QueryValueExError:
Resume QueryValueExExit
End Function
