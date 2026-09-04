Attribute VB_Name = "RegistryGrip"
    Option Explicit
    
    Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long
    Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
    Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
    Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
    
    Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
    Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
    Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
    
    Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
    Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
    Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
    Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
    Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
    Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
    Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
    Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
    
    
    
    Public Const READ_CONTROL = &H20000
    Public Const SYNCHRONIZE = &H100000
    Public Const STANDARD_RIGHTS_ALL = &H1F0000
    Public Const STANDARD_RIGHTS_READ = READ_CONTROL
    Public Const STANDARD_RIGHTS_WRITE = READ_CONTROL
    Public Const KEY_QUERY_VALUE = &H1
    Public Const KEY_SET_VALUE = &H2
    Public Const KEY_CREATE_SUB_KEY = &H4
    Public Const KEY_ENUMERATE_SUB_KEYS = &H8
    Public Const KEY_NOTIFY = &H10
    Public Const KEY_CREATE_LINK = &H20
    Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
    Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
    Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
    Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
    
    Public Const HKEY_CLASSES_ROOT = &H80000000
    Public Const HKEY_CURRENT_USER = &H80000001
    Public Const HKEY_LOCAL_MACHINE = &H80000002
    Public Const HKEY_USERS = &H80000003
    Public Const HKEY_PERFORMANCE_DATA = &H80000004
    
    Public Const ERROR_SUCCESS = 0&
    Public Const ERROR_BADDB = 1009&
    Public Const ERROR_BADKEY = 1010&
    Public Const ERROR_CANTOPEN = 1011&
    Public Const ERROR_CANTREAD = 1012&
    Public Const ERROR_CANTWRITE = 1013&
    Public Const ERROR_OUTOFMEMORY = 14&
    Public Const ERROR_INVALID_PARAMETER = 87&
    Public Const ERROR_ACCESS_DENIED = 5&
    
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
    
    Public Const REG_OPTION_NON_VOLATILE = 0&
    Public Const REG_OPTION_VOLATILE = &H1
    
    Type SECURITY_ATTRIBUTES
      nLength As Long
      lpSecurityDescriptor As Variant
      bInheritHandle As Long
    End Type

Public Sub gAPIDisplayError(Code&)
    Dim RegResult As String
    If Code& = 0 Then Exit Sub Else
    GoTo ErrorDetected
ErrorDetected:
    If Code& <> 2 Then ' 2 = Key not found
      Log "Registry Read Error"
    End If
    Exit Sub
End Sub
Public Sub RegistryWriteValue(WriteThis As Variant, hWndWriteToThisKey&, WriteToThisValueName$, ValueDataType&, Optional Multi_SZ_AddtlStrings As Variant)
    Dim lpData As Variant
    Dim cbData As Long
    Dim lReturn As Long
    Dim Str As Variant
    
    Select Case ValueDataType&
        Case REG_SZ, REG_EXPAND_SZ
            lpData = WriteThis & Chr(0)
            cbData = Len(lpData)
            lReturn = RegSetValueEx(hWndWriteToThisKey&, WriteToThisValueName$, 0&, ValueDataType&, lpData, cbData)
            Call gAPIDisplayError(lReturn)
        Case REG_MULTI_SZ
            lpData = WriteThis & Chr(0)
            If Not IsMissing(Multi_SZ_AddtlStrings) Then
                If IsArray(Multi_SZ_AddtlStrings) Then
                    For Each Str In Multi_SZ_AddtlStrings
                        If Str <> "" And Str <> Chr(0) And Not IsNull(Str) Then
                            lpData = lpData & Str & Chr(0)
                        End If
                    Next Str
                Else
                    If Multi_SZ_AddtlStrings <> "" And Multi_SZ_AddtlStrings <> Chr(0) And Not IsNull(Multi_SZ_AddtlStrings) Then
                        lpData = lpData & Multi_SZ_AddtlStrings & Chr(0)
                    End If
                End If
            End If
            lpData = lpData & Chr(0)
            cbData = Len(lpData)
            lReturn = RegSetValueEx(hWndWriteToThisKey&, WriteToThisValueName$, 0&, ValueDataType&, lpData, cbData)
            Call gAPIDisplayError(lReturn)
        Case REG_DWORD
            lpData = CLng(WriteThis)
            cbData = 4
            lReturn = RegSetValueEx(hWndWriteToThisKey&, WriteToThisValueName$, 0&, ValueDataType&, lpData, cbData)
            Call gAPIDisplayError(lReturn)
        Case Else
            MsgBox "Unable to process that type of data. Please contact support."
    End Select
End Sub

Public Sub SaveSettingByte(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, bData() As Byte)
    ' Make sure that the array starts with element 0 before passing it!
    ' (otherwise it will not be saved!)
    Dim lRegResult As Long
    Dim hCurKey As Long
    
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    ' Pass the first array element and length of array
    lRegResult = RegSetValueEx(hCurKey, strValueName, 0&, REG_BINARY, bData(0), UBound(bData()) + 1)
    lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function RegistryOpenKey(hWndAlreadyOpenedKey&, OpenThisKey$) As Long
    Dim hWndOpenedKey&
    Dim lReturn As Long
    lReturn = RegOpenKeyEx(hWndAlreadyOpenedKey&, OpenThisKey$, 0&, KEY_ALL_ACCESS, hWndOpenedKey&)
    Call gAPIDisplayError(lReturn)
    RegistryOpenKey = hWndOpenedKey&
End Function

Public Function RegistryQueryValue(RegistryKey&, ValueName$, ValueDataType&, Optional ByRef aryMultiSZ As Variant) As String
    Dim lpData As String
    Dim lpcbData As Long
    Dim lReturn As Long
    
    Select Case ValueDataType&
    
        Case REG_SZ, REG_EXPAND_SZ
            lpcbData = 255
            lpData = Space$(lpcbData)
            lReturn = RegQueryValueEx(RegistryKey&, ValueName$, 0&, ValueDataType&, lpData, lpcbData)
            Call gAPIDisplayError(lReturn)
            RegistryQueryValue = Left$(lpData, lpcbData - 1)
        Case REG_BINARY
            lpcbData = 255
            lpData = Space$(lpcbData)
            lReturn = RegQueryValueEx(RegistryKey&, ValueName$, 0&, ValueDataType&, lpData, lpcbData)
            Call gAPIDisplayError(lReturn)
            RegistryQueryValue = BinaryInStringToHexInString(Left$(lpData, lpcbData))
        Case REG_MULTI_SZ
            lpcbData = 255
            lpData = Space$(lpcbData)
            lReturn = RegQueryValueEx(RegistryKey&, ValueName$, 0&, ValueDataType&, lpData, lpcbData)
            Call gAPIDisplayError(lReturn)
            Dim i As Integer
            Dim BeginNullPos As Integer
            Dim EndNullPos As Integer
            Dim DoubleNull As Boolean
            BeginNullPos = 0
            i = 0
            DoubleNull = False
            
            EndNullPos = InStr(BeginNullPos + 1, lpData, Chr(0))
            If EndNullPos = BeginNullPos + 1 Then
                RegistryQueryValue = ""
                DoubleNull = True
            Else
                RegistryQueryValue = Mid(lpData, BeginNullPos + 1, EndNullPos - BeginNullPos - 1)
            End If
            
            If Not IsMissing(aryMultiSZ) Then
                ReDim aryMultiSZ(0 To i)
                EndNullPos = 0
                If DoubleNull = True Then aryMultiSZ(i) = ""
                Do Until DoubleNull = True
                    ReDim Preserve aryMultiSZ(0 To i)
                    BeginNullPos = EndNullPos
                    EndNullPos = InStr(BeginNullPos + 1, lpData, Chr(0))
                    aryMultiSZ(i) = Mid(lpData, BeginNullPos + 1, EndNullPos - BeginNullPos - 1)
                    If EndNullPos = BeginNullPos + 1 Then
                        DoubleNull = True
                        ReDim Preserve aryMultiSZ(i - 1)
                    End If
                    i = i + 1
                Loop
            End If
        Case Else
            MsgBox "Unable to query that type of data. Please contact Support."
    End Select
End Function

Public Function BinaryInStringToHexInString(ConvertThis$) As String
    Dim i%
    Dim Buffer$
    Dim HexPacket$
        
    For i = 1 To Len(ConvertThis$) Step 1
        HexPacket$ = Hex$(Asc(Mid$(ConvertThis$, i, 1)))
        If Len(HexPacket$) = 1 Then
            HexPacket$ = "0" & HexPacket$
        End If
        Buffer$ = Buffer$ + HexPacket$
    Next i
    BinaryInStringToHexInString = Buffer$
End Function

Public Function RegistryConnectRegistry(ConnectToThisComputer$, AlreadyOpenedKey&) As Long
    Dim ConnectedRegistryKey&
    Dim lReturn As Long
    lReturn = RegConnectRegistry(ConnectToThisComputer$, AlreadyOpenedKey&, ConnectedRegistryKey&)
    Call gAPIDisplayError(lReturn)
    RegistryConnectRegistry = ConnectedRegistryKey&
End Function

Public Sub RegistryCloseKey(CloseThisKey&)
    Dim lReturn As Long
    lReturn = RegCloseKey(CloseThisKey&)
    Call gAPIDisplayError(lReturn)
End Sub

Public Function RegistryCreateKey(hWndAlreadyOpenedKey&, CreateThisKey$) As Long
    Dim hWndCreatedKey&
    Dim lpdwDisposition As Long
    Dim lpSecurityAttributes As SECURITY_ATTRIBUTES
    Dim lReturn As Long
    lReturn = RegCreateKeyEx(hWndAlreadyOpenedKey&, CreateThisKey$, 0&, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpSecurityAttributes, hWndCreatedKey&, lpdwDisposition)
    Call gAPIDisplayError(lReturn)
    RegistryCreateKey = hWndCreatedKey&
End Function

Public Sub RegistryFlushKey(FlushThisKey&)
    Dim lReturn As Long
    lReturn = RegFlushKey(FlushThisKey&)
    Call gAPIDisplayError(lReturn)
End Sub


