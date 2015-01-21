Attribute VB_Name = "Registry"
Option Explicit

Const REG_SZ As Long = 1
Const REG_DWORD As Long = 4

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006

Public Const KEY_ALL_ACCESS = &H3F

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Public Sub RegWrite(Path As String, Key As String, Value As String)
    Dim Result As Long
    RegOpenKeyEx HKEY_CURRENT_USER, Path, 0&, KEY_ALL_ACCESS, Result
    RegSetValueEx Result, Key, 0&, REG_SZ, ByVal Value, Len(Value)
    RegCloseKey Result
End Sub

Public Sub RegDelete(Path As String, Key As String)
    Dim Result As Long
    RegOpenKeyEx HKEY_CURRENT_USER, Path, 0&, KEY_ALL_ACCESS, Result
    RegDeleteValue Result, Key
    RegCloseKey Result
End Sub

Public Function RegGetValue(Path As String, Key As String) As String
    Dim Result As Long
    Dim STResult As String
    STResult = Space(255)
    RegOpenKeyEx HKEY_CURRENT_USER, Path, 0, KEY_ALL_ACCESS, Result
    RegQueryValueEx Result, Key, 0, 0, ByVal STResult, Len(STResult)
    RegCloseKey Result
    RegGetValue = STResult
End Function
