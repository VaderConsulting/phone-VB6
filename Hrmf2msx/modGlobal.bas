Attribute VB_Name = "modGlobal"
Option Explicit
Option Base 1
' Global variables

' General
Global gError As String
Global Const gSQL_DATE_FORMAT As String = "dd-mmm-yy hh:nn:ss"
' Database
Global gDb_Version As Integer
Global gDbServerName As String, gDbName As String
Global gDbUsername As String, gDbPassword As String, gDbOwner As String
Global gModifiedSP As String, gUpdatesSP As String, gAdminUsers As String
' Exchange
Global gMSXServer As String, gMSXOrganisation As String
Global gMSXUsername As String, gMSXPassword As String
' Network
Global gDomainName As String, gAdminGroupName As String

' Types and Enums
Public Enum ADS_SECURITY
  ADS_SECURE_AUTHENTICATION = 1
  ADS_USE_ENCRYPTION = 2
  ADS_USE_SSL = 2
  ADS_READONLY_SERVER = 4
  ADS_PROMPT_CREDENTIALS = 8
  ADS_NO_AUTHENTICATION = 16
  ADS_FAST_BIND = 32
  ADS_USE_SIGNING = 64
  ADS_USE_SEALING = 128
  ADS_USE_DELEGATION = 256
  ADS_SERVER_BIND = 512
End Enum

' ************************************************************
' Log specified text to application log file
' ************************************************************
' Version  Date         Who Reason
'   1     21 May 2001   DR  Initial
' ************************************************************
Sub Log(strText As String)
  Dim intHandle As Integer, Prefix As String
  intHandle = FreeFile
  Prefix = Format(Date & " " & Time, "YYYYMMDD HH:NN:SS")
  On Error GoTo LogError
  Open App.Path & "\HRMF2MSX-" & Format(Date & " " & Time, "YYYYMMDD") & ".log" For Append As #intHandle
    Print #intHandle, Prefix & " " & strText
  Close intHandle
  ' Removed frmMain, so this line no longer valid
  'frmMain.lstStatus.AddItem Prefix & " " & strText, 0
  Exit Sub
LogError:
  ' Error saving/appending to file
  ' Removed frmMain, so this line no longer valid
  ' frmMain.lstStatus.AddItem "Unexpected error " & Err.Number & " writing to log file.", 0
  Err.Clear
End Sub

' ************************************************************
' Get Application defaults from Registry
' ************************************************************
' Version  Date         Who Reason
'   1     21 May 2001   DR  Initial
' ************************************************************

Public Sub GetDefaults()
  Dim strTemp As String
  Dim strRegPath As String
  Dim Hostname As String
  
  strRegPath = "Software\CSC\HRMF2MSX"
  Hostname = Environ$("computername")
  
  strTemp = GetValue("", strRegPath, "gDb_Version", False)
  If strTemp <> "" Then
    gDb_Version = CInt(strTemp)
  Else
    gDb_Version = 7
    SetValueString "", strRegPath, "gDb_Version", CStr(gDb_Version), REG_SZ
  End If
  
  strTemp = GetValue("", strRegPath, "gDbServerName", False)
  If strTemp <> "" Then
    gDbServerName = strTemp
  Else
    gDbServerName = "SVNTDEV-51"
    SetValueString "", strRegPath, "gDbServerName", gDbServerName, REG_SZ
  End If
  
  strTemp = GetValue("", strRegPath, "gDbName", False)
  If strTemp <> "" Then
    gDbName = strTemp
  Else
    gDbName = "HRMF_D"
    SetValueString "", strRegPath, "gDbName", gDbName, REG_SZ
  End If
  
  strTemp = GetValue("", strRegPath, "gDbUsername", False)
  If strTemp <> "" Then
    gDbUsername = strTemp
  Else
    gDbUsername = "hrmf"
    SetValueString "", strRegPath, "gDbUsername", gDbUsername, REG_SZ
  End If
  
  strTemp = GetValue("", strRegPath, "gDbPassword", False)
  If strTemp <> "" Then
    gDbPassword = strTemp
  Else
    gDbPassword = "hrmf"
    SetValueString "", strRegPath, "gDbPassword", gDbPassword, REG_SZ
  End If
  
  strTemp = GetValue("", strRegPath, "gDbOwner", False)
  If strTemp <> "" Then
    gDbOwner = strTemp
  Else
    gDbOwner = "hrmf"
    SetValueString "", strRegPath, "gDbOwner", gDbOwner, REG_SZ
  End If
  
  strTemp = GetValue("", strRegPath, "gModifiedSP", False)
  If strTemp <> "" Then
    gModifiedSP = strTemp
  Else
    gModifiedSP = "sp_Get_HRMF_Modifications"
    SetValueString "", strRegPath, "gModifiedSP", gModifiedSP, REG_SZ
  End If
  
  strTemp = GetValue("", strRegPath, "gUpdatesSP", False)
  If strTemp <> "" Then
    gUpdatesSP = strTemp
  Else
    gUpdatesSP = "sp_HRMF_Catchup"
    SetValueString "", strRegPath, "gUpdatesSP", gUpdatesSP, REG_SZ
  End If
  
  strTemp = GetValue("", strRegPath, "gAdminUsers", False)
  If strTemp <> "" Then
    gAdminUsers = strTemp
  Else
    gAdminUsers = "HRMF_Admin_User"
    SetValueString "", strRegPath, "gAdminUsers", gAdminUsers, REG_SZ
  End If
  
  strTemp = GetValue("", strRegPath, "gMSXServer", False)
  If strTemp <> "" Then
    gMSXServer = strTemp
  Else
    gMSXServer = "SVNTDEV-54"
    SetValueString "", strRegPath, "gMSXServer", gMSXServer, REG_SZ
  End If
  
  strTemp = GetValue("", strRegPath, "gMSXOrganisation", False)
  If strTemp <> "" Then
    gMSXOrganisation = strTemp
  Else
    gMSXOrganisation = "Water Corporation"
    SetValueString "", strRegPath, "gMSXOrganisation", gMSXOrganisation, REG_SZ
  End If
  
  strTemp = GetValue("", strRegPath, "gMSXUsername", False)
  If strTemp <> "" Then
    gMSXUsername = strTemp
  Else
    gMSXUsername = "JTC\srvmsxapps"
    SetValueString "", strRegPath, "gMSXUsername", gMSXUsername, REG_SZ
  End If
  
  strTemp = GetValue("", strRegPath, "gMSXPassword", False)
  If strTemp <> "" Then
    gMSXPassword = strTemp
  Else
    gMSXPassword = "monkey88"
    SetValueString "", strRegPath, "gMSXPassword", gMSXPassword, REG_SZ
  End If
  
  strTemp = GetValue("", strRegPath, "gDomainName", False)
  If strTemp <> "" Then
    gDomainName = strTemp
  Else
    gDomainName = "wadev"
    SetValueString "", strRegPath, "gDomainName", gDomainName, REG_SZ
  End If
  
  strTemp = GetValue("", strRegPath, "gAdminGroupName", False)
  If strTemp <> "" Then
    gAdminGroupName = strTemp
  Else
    gAdminGroupName = "pb_Helpdesk"
    SetValueString "", strRegPath, "gAdminGroupName", gAdminGroupName, REG_SZ
  End If
  
End Sub

Public Function GetValue(ByVal Hostname As String, ByVal Key As String, ByVal Value As String, ByRef ECode As Boolean) As String
    Dim OpenKeyVal As Long
    Dim OpenHiveVal As Long
    Dim RResult As Long
    Dim InfoTextStr As String
    
    'Init
    ECode = False
    Hostname = Trim(Hostname)
    Key = Trim(Key)
    Value = Trim(Value)
    GetValue = ""
        
    RResult = RegConnectRegistry("", HKEY_LOCAL_MACHINE, OpenHiveVal)
    If (RResult <> ERROR_SUCCESS) Then
      ECode = False
      Log "Error connecting to Registry"
      Exit Function
    End If
    
    OpenKeyVal = RegistryOpenKey(OpenHiveVal, Key)
    InfoTextStr = RegistryQueryValue(OpenKeyVal, Value, REG_SZ)
    
    RegCloseKey (OpenKeyVal)
    RegCloseKey (OpenHiveVal)
    
    GetValue = Trim(InfoTextStr)
    ECode = True
End Function

Public Function CreateKey(ByVal Hostname As String, ByVal Key As String, NewKey As String) As Boolean
    Dim OpenKeyVal As Long
    Dim OpenHiveVal As Long
    Dim RResult As Long
    Dim lCreateResult As Long
    
    'Init
    Hostname = Trim(Hostname)
    Key = Trim(Key)
        
    RResult = RegConnectRegistry("", HKEY_LOCAL_MACHINE, OpenHiveVal)
    If (RResult <> ERROR_SUCCESS) Then
      CreateKey = False
      Log "Error creating Registry Key"
      Exit Function
    End If
    
    OpenKeyVal = RegistryOpenKey(OpenHiveVal, Key)
    lCreateResult = RegistryCreateKey(OpenKeyVal, NewKey)
    
    RegCloseKey (OpenKeyVal)
    RegCloseKey (OpenHiveVal)
    
    CreateKey = True
    
End Function

Public Function SetValueString(Hostname As String, Key As String, Value As String, Data As String, regType As Long) As Long
    Dim OpenKeyVal As Long
    Dim RResult As Long
    Dim OpenHiveVal As Long
    Dim strValue, CMPTRName, keytogo As String, InfoTextStr As String
    Dim x As Integer
    
    If Hostname = "" Then
      Hostname = Environ$("computername")
    End If
    'GetIPCConnection (Trim(Hostname))
    CMPTRName = Trim(Hostname)
    
    RResult = RegConnectRegistry(CMPTRName, HKEY_LOCAL_MACHINE, OpenHiveVal)
    keytogo = Trim(Key)
    OpenKeyVal = RegistryOpenKey(OpenHiveVal, keytogo)
    RegistryWriteValue Data, OpenKeyVal, Value, regType
    
    RegCloseKey (OpenKeyVal)
    RegCloseKey (OpenHiveVal)
    
    'DisIPCConnection (Trim(Hostname))
        
End Function


