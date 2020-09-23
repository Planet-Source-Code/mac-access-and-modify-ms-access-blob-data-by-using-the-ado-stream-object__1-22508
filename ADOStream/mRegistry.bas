Attribute VB_Name = "mRegistry"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"38AABC22012D"
'// THIS Code was a little bit re-designed to work with the MS Access Database
'// Feel free to use this code for your own needs
'// It would be nice if you could give a credit to me in ya AppZ
'// Please contact me for further information and questions:  marcuslauermann@gmx.net

' -----------------
' ADVAPI32
' -----------------
' function prototypes, constants, and type definitions
' for Windows 32-bit Registry API
Option Base 0

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

' Registry API prototypes

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA"
'(ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal
'lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long,
'lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition
'As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
'Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal
'hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As
'Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long,
'lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Declare Function RegGetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long,
'ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR,
'lpcbSecurityDescriptor As Long) As Long
Public Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpFile As String) As Long
Public Declare Function RegNotifyChangeKeyValue Lib "advapi32.dll" (ByVal hKey As Long, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long, ByVal hEvent As Long, ByVal fAsynchronus As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
'Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA"
'(ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, lpReserved
'As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As
'Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As
'Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long
Public Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
'Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey
'As Long, ByVal lpFile As String, lpSecurityAttributes As SECURITY_ATTRIBUTES)
'As Long
'Declare Function RegSetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long,
'ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR)
'As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Public Const REG_NONE = 0
Public Const REG_SZ = 1
Public Const REG_EXPAND_SZ = 2
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const REG_DWORD_LITTLE_ENDIAN = 4
Public Const REG_DWORD_BIG_ENDIAN = 5
Public Const REG_LINK = 6
Public Const REG_MULTI_SZ = 7
Public Const REG_RESOURCE_LIST = 8
Public Const REG_FULL_RESOURCE_DESCRIPTOR = 9
Public Const REG_RESOURCE_REQUIREMENTS_LIST = 10
Public Const REG_CREATED_NEW_KEY = &H1
Public Const REG_OPENED_EXISTING_KEY = &H2
Public Const REG_WHOLE_HIVE_VOLATILE = &H1
Public Const REG_REFRESH_HIVE = &H2
Public Const REG_NOTIFY_CHANGE_NAME = &H1
Public Const REG_NOTIFY_CHANGE_ATTRIBUTES = &H2
Public Const REG_NOTIFY_CHANGE_LAST_SET = &H4
Public Const REG_NOTIFY_CHANGE_SECURITY = &H8
Public Const REG_LEGAL_CHANGE_FILTER = (REG_NOTIFY_CHANGE_NAME Or REG_NOTIFY_CHANGE_ATTRIBUTES Or REG_NOTIFY_CHANGE_LAST_SET Or REG_NOTIFY_CHANGE_SECURITY)
'Public Const REG_LEGAL_OPTION = (REG_OPTION_RESERVED Or REG_OPTION_NON_VOLATILE Or REG_OPTION_VOLATILE Or REG_OPTION_CREATE_LINK Or REG_OPTION_BACKUP_RESTORE)

' Reg Create Type Values...
Public Const REG_OPTION_RESERVED = 0
Public Const REG_OPTION_NON_VOLATILE = 0
Public Const REG_OPTION_VOLATILE = 1
Public Const REG_OPTION_CREATE_LINK = 2
Public Const REG_OPTION_BACKUP_RESTORE = 4
Public Const STANDARD_RIGHTS_READ = &H20000
Public Const STANDARD_RIGHTS_WRITE = &H20000
Public Const STANDARD_RIGHTS_EXECUTE = &H20000
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const DELETE = &H10000
Public Const READ_CONTROL = &H20000
Public Const WRITE_DAC = &H40000
Public Const WRITE_OWNER = &H80000
Public Const SYNCHRONIZE = &H100000

' Reg Key Security Options
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
'Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or
'KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
'Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or
'KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
'Public Const KEY_EXECUTE = (KEY_READ)
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Sub savekey(hKey As Long, strpath As String)
Dim keyhand&
r = RegCreateKey(hKey, strpath, keyhand&)
r = RegCloseKey(keyhand&)
End Sub

Public Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As Variant
    Dim lResult As Long
    Dim lValueType As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    
    On Error GoTo 0
    lResult = RegQueryValueEx(hKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_SZ Then
            strBuf = String(lDataBufSize, " ")
            lResult = RegQueryValueEx(hKey, strValueName, 0&, 0&, ByVal strBuf, lDataBufSize)
            If lResult = ERROR_SUCCESS Then
                RegQueryStringValue = StripTerminator(strBuf)
            End If
        End If
    End If
End Function

Public Function getstring(hKey As Long, strpath As String, strvalue As String) As Variant
'This module allows easy access to the Windows 95 and NT Registry.
'It takes the following parameters:
'
'HKEY - one of the following constants:
'HKEY_CLASSES_ROOT
'HKEY_CURRENT_USER
'HKEY_LOCAL_MACHINE
'HKEY_USERS
'HKEY_PERFORMANCE_DATA
'
'strpath - the rest of the registry 'path' eg:
'"Control Panel\desktop"
'
'strvalue - A string that refers to the key that you want to retrieve eg:
'"Wallpaper"
'
'The Example above would be called as follows:
'
'dim strwallpaper as string
'strwallpaper = getstring(HKEY_CURRENT_USER, "Control Panel\desktop", "Wallpaper")
'
'This would return the current desktop wallpaper
'
'This module downloaded from VB-World at www.geocities.com/SiliconValley/Bay/8409/


Dim keyhand&
Dim datatype&
r = RegOpenKey(hKey, strpath, keyhand&)
getstring = RegQueryStringValue(keyhand&, strvalue)
r = RegCloseKey(keyhand&)
End Function

Public Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Public Sub savestring(hKey As Long, strpath As String, strvalue As String, strdata As String)
'This module allows easy access to the Windows 95 and NT Registry.
'It takes the following parameters:
'
'HKEY - one of the following constants:
'HKEY_CLASSES_ROOT
'HKEY_CURRENT_USER
'HKEY_LOCAL_MACHINE
'HKEY_USERS
'HKEY_PERFORMANCE_DATA
'
'strpath - the rest of the registry 'path' eg:
'"Control Panel\desktop"
'
'strvalue - A string that refers to the key that you want to save eg:
'"Wallpaper"
'
'strdata - the string you want to save against strvalue eg:
'"c:\windows\clouds.bmp"
'
'The Example above would be called as follows:
'
'dim strwallpaper as string
'call savestring(HKEY_CURRENT_USER, "Control Panel\desktop", "Wallpaper", "C:\Windows\Clouds.bmp")
'
'If this 'path' does not exist, it will be created.
'
'This module downloaded from VB-World at www.geocities.com/SiliconValley/Bay/8409/


Dim keyhand&
r = RegCreateKey(hKey, strpath, keyhand&)
r = RegSetValueEx(keyhand&, strvalue, 0, REG_SZ, ByVal strdata, Len(strdata))
r = RegCloseKey(keyhand&)
End Sub
