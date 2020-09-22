Attribute VB_Name = "modFunctions"
Private lngHKEY As Long
'API Functions
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long

'Constants used in Functions
Private Const HKEY_LOCAL_MACHINE = &H80000002

Private Const REG_BINARY = 3
Private Const REG_CREATED_NEW_KEY = &H1
Private Const REG_DWORD = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_EXPAND_SZ = 2
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9
Private Const REG_LINK = 6
Private Const REG_MULTI_SZ = 7
Private Const REG_NONE = 0
Private Const REG_RESOURCE_LIST = 8
Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10
Private Const REG_SZ = 1

Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000
Private Const KEY_READ = &H20009
Private Const KEY_WRITE = &H20006
Private Const KEY_READ_WRITE = ( _
KEY_READ _
And _
KEY_WRITE _
)
Private Const KEY_ALL_ACCESS = ( _
( _
STANDARD_RIGHTS_ALL Or _
KEY_QUERY_VALUE Or _
KEY_SET_VALUE Or _
KEY_CREATE_SUB_KEY Or _
KEY_ENUMERATE_SUB_KEYS Or _
KEY_NOTIFY Or _
KEY_CREATE_LINK _
) _
And _
( _
Not SYNCHRONIZE _
) _
)
Private Const REG_OPTION_NON_VOLATILE = 0&
Private Const REG_OPTION_VOLATILE = &H1


Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type



















Public Function GetUserNames(strUsers() As String) As Boolean
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Napster", KEY_READ, 0&, lngHKEY) <> 0 Then Exit Function  'Open Reg
    Dim i As Integer, strName As String, strKeys() As String, Temp As FILETIME, intUsers As Integer, strClass As String
    Do  'Get all Usernames stored in Registry
        Dim inL As Long
        inL = 256
        strClass = Space(256)
        strName = Space(256)
        If RegEnumKeyEx(lngHKEY, i, strName, inL, 0&, strClass, 0, Temp) <> 0 Then Exit Do 'Get Username
        ReDim Preserve strKeys(i)
        strKeys(i) = Trim(Left(strName, InStr(1, strName, Chr(0)) - 1)) 'Add to array
        i = i + 1
    Loop
    For i = LBound(strKeys) To UBound(strKeys)
        If LCase(strKeys(i)) <> LCase("File Types") Then 'Make Sure its a username
            ReDim Preserve strUsers(intUsers)
            strUsers(intUsers) = strKeys(i)
            intUsers = intUsers + 1
        End If
    Next i
    RegCloseKey lngHKEY 'Close Registry
    GetUserNames = True
End Function

Public Function SetUser(ByVal strUser As Variant) As Boolean
    Dim lngTemp As Long, strKey As String
    strKey = strUser & Chr(0)
    lngTemp = Len(strKey)
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Napster", KEY_READ, 0&, lngHKEY) <> 0 Then Exit Function 'Open Reg
    If RegSetValueEx(lngHKEY, "CurrentUser", 0&, REG_SZ, strKey, lngTemp) <> 0 Then Exit Function 'Set Key to User
    SetUser = True
    RegCloseKey lngHKEY 'Close Reg
    
End Function



Public Function GetNapPath() As String
    Dim strTemp As String, lngTemp As Long, retval, i As Long

    retval = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Napster", KEY_READ, 0&, lngHKEY) 'Open Registry
    If retval <> 0 Then Exit Function
    retval = RegQueryValueEx(lngHKEY, "InstallPath", 0&, i, ByVal 0&, lngTemp) 'Get String Length
    strTemp = Space(lngTemp)    'Pad String
    retval = RegQueryValueEx(lngHKEY, "InstallPath", 0&, i, ByVal strTemp, lngTemp) 'Get String
    GetNapPath = Left(strTemp, Len(strTemp) - 1)
    RegCloseKey lngHKEY 'Close Registry
End Function

Public Function GetCurUser()
    Dim strTemp As String, lngTemp As Long, retval, i As Long
    retval = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Napster", KEY_READ, 0&, lngHKEY) 'Open Registry
    If retval <> 0 Then Exit Function
    retval = RegQueryValueEx(lngHKEY, "CurrentUser", 0&, i, ByVal 0&, lngTemp) 'Get String Length
    strTemp = Space(lngTemp) 'Pad String
    retval = RegQueryValueEx(lngHKEY, "CurrentUser", 0&, i, ByVal strTemp, lngTemp) 'Get String
    GetCurUser = Left(strTemp, Len(strTemp) - 1) 'Trim String
    RegCloseKey lngHKEY 'Close Registry
End Function
