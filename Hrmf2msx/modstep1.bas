Attribute VB_Name = "modStep1"
Option Explicit
Option Base 1
' **********************************************************************
' Extract list of users in specified NT Global group
' Update specified table with this new list
' **********************************************************************
' Version  Date         Who Reason
'   1     21 May 2001   DR  Initial
' **********************************************************************

' Start in this function.
Public Function Step1() As Boolean
  On Error GoTo Step1Error
  Dim strTableName As String
  Dim strDbServerName As String
  Dim strDatabaseName As String
  Dim strDomainName As String
  Dim strDBOwner As String
  Dim strSQLString As String
  Dim strDBUsername As String, strDBPassword As String
  Dim strUsers() As String, strDBUsers() As String, i As Integer
  Dim strDSN As String, strSQL As String
  Dim adoConnection As ADODB.Connection
  Dim adoRecordset As ADODB.Recordset
  
  Set adoConnection = CreateObject("ADODB.Connection")
  Set adoRecordset = CreateObject("ADODB.Recordset")
  
  ' Set Database defaults
  
  strDBUsername = gDbUsername     '"hrmf"
  strDBPassword = gDbPassword     '"hrmf"
  
  strDatabaseName = gDbName       '"HRMF_D"
  strDbServerName = gDbServerName '"SVNTDEV-51"
  strDBOwner = gDbOwner           '"hrmf"
  strTableName = gAdminUsers      '"HRMF_Admin_User"
  
  'Make Database defaults as per DB version
  Select Case gDb_Version
    Case 65
      strTableName = strDBOwner & "." & strTableName
      strDSN = "Provider=SQLOLEDB;Data Source=" & strDbServerName & ";User ID=" & strDBUsername & ";password=;" & strDBPassword
    Case 7
    strTableName = strTableName
    strDSN = "Provider=SQLOLEDB.1;Persist Security Info=False;Initial Catalog=" & strDatabaseName & ";Data Source=" & strDbServerName
  End Select
  
  ' Open Connection to nominated Database
  Log "Opening ADO Connection"
  adoConnection.Open strDSN, strDBUsername, strDBPassword
  
  ' Use ADSI to retrieve list of admin users, and store in strUsers()
  ' NOTE.. caters for development environment:
  '    When in development, create a local group on your workstation called
  '    the same as the contents of gAdminGroupName.  Then place local users in this
  '    group.  After this, change the line below from A000677 to your workstation name,
  '    and make sure that DEBUG is included in your command line arguments
  '    After this is done, all Group membership functions will refer to your local machine.
  '
  If InStr(UCase(Command$), "DEBUG") <> 0 Then
    strUsers() = GetUsers(gAdminGroupName, "A000677")  ' REFER TO NOTE ABOVE
  Else
    strUsers() = GetUsers(gAdminGroupName, gDomainName)
  End If
  If gError <> "" Then
    Log "GetUsers(" & gAdminGroupName & "," & gDomainName & ") returned an error!"
    Step1 = False
  Else
    ' Delete existing administrators of phonebook
    Log "Deleting Admin Users from " & strTableName ' The table name changes to reflect SQL 6.5 or 7
    strSQL = "DELETE FROM " & strTableName
    adoConnection.Execute strSQL
    
    ' Update Table with correct list of users
    Log "Updating " & strTableName & " with Admin users"
    For i = LBound(strUsers()) To UBound(strUsers())
      strSQL = "INSERT INTO " & strTableName & "(Username) VALUES ('" & strUsers(i) & "')"
      adoConnection.Execute strSQL
    Next
    Step1 = True
  End If
  
  Set adoConnection = Nothing
  Set adoRecordset = Nothing
  Exit Function
Step1Error:
  Log Err.Description & " (Error " & Err.Number & ") whilst in step1 Function."
  Err.Clear
  Step1 = False
End Function

' Get a list of all users in the specified Global Group in the specified Domain
Function GetUsers(Groupname As String, strDomainName As String) As String()
  Dim Group As IADsGroup, Member As Object, i As Integer
  Dim lUsers() As String
  Log "Retrieving members of " & Groupname & " in Domain " & strDomainName
  On Error GoTo UserError
  Set Group = GetObject("WinNT://" & strDomainName & "/" + Groupname)
  For Each Member In Group.Members
    If Member.Class <> "Group" Then
      i = i + 1
      ReDim Preserve lUsers(i)
      lUsers(i) = Member.Name
    End If
  Next
  GetUsers = lUsers()
  Set Group = Nothing
  Set Member = Nothing
  gError = ""
  Exit Function
UserError:
  Log Err.Description & " (Error " & Err.Number & ") whilst retrieving members of " & Groupname & "."
  Err.Clear
  Set Group = Nothing
  Set Member = Nothing
  gError = "GetUsers"
End Function
