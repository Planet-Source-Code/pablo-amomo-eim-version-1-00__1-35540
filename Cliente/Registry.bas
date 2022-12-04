Attribute VB_Name = "Registry"
Option Explicit
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
       'RegQueryValueEx: If you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
       'Remote Access Services (RAS) APIs.
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal Hkey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Declare Function GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal Buffer As String, buffersize As Long) As Long

Public Const REG_DWORD = 4                      ' 32-bit number
Public Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Public Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const ERROR_SUCCESS = 0&
Public Const APINULL = 0&
Public Const MAX_STRING_LENGTH As Integer = 256

Public ReturnCode As Long, blnOffice As Boolean
Public Security As SECURITY_ATTRIBUTES

Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
Public Sub GrabarRegistry(Hkey As Long, strPath As String, StrValue As String, strdata As String)
Dim keyhand As Long
Dim R As Long
    
    R = RegCreateKey(Hkey, strPath, keyhand)
    R = RegSetValueEx(keyhand, StrValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    R = RegCloseKey(keyhand)

End Sub
Public Function LeerRegistry(Hkey As Long, SubKey As String, Valor As String) As String
    'Dim Hkey As Long
    Dim lpSubKey As String
    Dim phkResult As Long
    Dim lpValueName As String
    Dim lpReserved As Long
    Dim lpType As Long
    Dim lpData As String
    Dim lpcbData As Long
        
    ReturnCode = RegOpenKey(Hkey, SubKey, phkResult)
    If ReturnCode = ERROR_SUCCESS Then
      Hkey = phkResult
      lpValueName = Valor
      lpReserved = APINULL
      lpType = APINULL
      lpData = APINULL
      lpcbData = APINULL
      ReturnCode = RegQueryValueEx(Hkey, lpValueName, lpReserved, lpType, ByVal lpData, lpcbData)
      lpData = String(lpcbData, 0)
      ReturnCode = RegQueryValueEx(Hkey, lpValueName, lpReserved, lpType, ByVal lpData, lpcbData)
      If ReturnCode = ERROR_SUCCESS Then
        If lpcbData - 1 >= 0 Then
          LeerRegistry = Left(lpData, lpcbData - 1)
         Else
          LeerRegistry = ""
        End If
       Else
        LeerRegistry = ""
      End If
      RegCloseKey (Hkey)
     Else
      LeerRegistry = ""
    End If
   
End Function
