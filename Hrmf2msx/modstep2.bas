Attribute VB_Name = "modStep2"
Option Explicit
Option Base 1


' Contains workaround to problem of not being able to query LDAP:
' 1.  Get list of NT Logon ID's to modify
' 2.  Get list of sites (ie ou's in organisation) - place in Collection
' 3.  For each Logon ID, attempt to connect using each ou in turn (discarding errors)
' 4.  When a valid reference is returned (ie err.number =0), modify mailbox
' 5.  Continue through Logon's

Public Function Step2() As Boolean
  'Dim strTableName As String
  Dim strDbServerName As String
  Dim strDatabaseName As String
  Dim strDomainName As String
  Dim strDBOwner As String
  Dim strSQLString As String
  Dim strDBUsername As String, strDBPassword As String
  Dim strUsers() As String, strDBUsers() As String, i As Integer
  Dim strDSN As String, strSQL As String, ADSSQL As String
  
  Dim adoADSConnection As ADODB.Connection
  Dim adoADSRecordset As ADODB.Recordset
  
  Dim NameSpace As IADsOpenDSObject, colSites As New Collection
  Dim Organisation As IADsO, Site As Object, Path As String
  Dim Mailbox, Person, ou As Variant
  Dim NTAccount As String, HRMFField As String
  
  ' Exchange info
  Dim MSXTelephone As String, MSXFax As String
  Dim MSXMobile As String, MSXEmail As String, MSXPager As String
  Dim MSXCompany As String, MSXDepartment As String, MSXOffice As String
  Dim MSXTitle As String, MSXAddress As String, MSXCity As String
  Dim MSXState As String, MSXZipCode As String, MSXCountry As String
  
  ' New Exchange Mailbox info
  Dim NewTelephone As String, NewFax As String, NewMobile As String
  Dim NewEmail As String, NewPager As String
  Dim NewCompany As String, NewDepartment As String, NewOffice As String
  Dim NewTitle As String, NewAddress As String, NewCity As String
  Dim NewState As String, NewZipCode As String, NewCountry As String
  
  Set adoADSConnection = CreateObject("ADODB.Connection")
  
  Dim objMailbox As Object          ' Mailbox object
  Dim strMSXUsername As String      ' Service Account Username
  Dim strMSXPassword As String      ' Service Account Password
  Dim strMSXServer As String        ' Exchange Server
  Dim strMSXMailbox As String       ' Mailbox to modify
  Dim strMSXMailboxPath As String   ' AD Path to Mailbox
  
  Dim adoConnection As ADODB.Connection
  Dim adoRecordset As ADODB.Recordset
  Dim ADSUID As Object
  
  Dim lngReturnVal As Long
  Dim SQLServerDateTime As String
  Dim MSXError As Long
  
  Set adoConnection = CreateObject("ADODB.Connection")
  Set adoRecordset = CreateObject("ADODB.Recordset")
  
  ' Set Database defaults
  
  strDBUsername = gDbUsername     '"hrmf"
  strDBPassword = gDbPassword     '"hrmf"
  
  strDatabaseName = gDbName       '"HRMF_D"
  strDbServerName = gDbServerName '"SVNTDEV-51"
  strDBOwner = gDbOwner           '"hrmf"
  'strTableName = gdbtable         '"HRMF_SAP_Person"
  
  ' Set MSX Defaults
  strMSXServer = gMSXServer       '"SVNTDEV-54"
  strMSXUsername = gMSXUsername   '"jtc\srvmsxapps"
  strMSXPassword = gMSXPassword   '"monkey88"
  
  ' ADSI Defaults
  adoADSConnection.Provider = "ADSDSOObject"
  
  'Make Database defaults as per DB version
  Select Case gDb_Version
    Case 65
      'strTableName = strDBOwner & "." & strTableName
      strDSN = "Provider=SQLOLEDB;Data Source=" & strDbServerName & ";User ID=" & strDBUsername & ";password=;" & strDBPassword
    Case 7
    'strTableName = strTableName
    strDSN = "Provider=SQLOLEDB.1;Persist Security Info=False;Initial Catalog=" & strDatabaseName & ";Data Source=" & strDbServerName
  End Select
  
  On Error GoTo Step2Error
  
  ' Open Connection to nominated Database
  adoConnection.Open strDSN, strDBUsername, strDBPassword
  
  ' Open ADO Connection to ADSI
  adoADSConnection.Open "ADs Provider"
  
  ' Get list of Sites (ie Exchange ou's)
  Set NameSpace = GetObject("LDAP:")
  Set Organisation = NameSpace.OpenDSObject("LDAP://" & strMSXServer & "/o=" & gMSXOrganisation, strMSXUsername, strMSXPassword, ADS_SECURE_AUTHENTICATION)
  For Each Site In Organisation
    colSites.Add Right(Site.Name, Len(Site.Name) - 3), Right(Site.Name, Len(Site.Name) - 3)
  Next
  
  ' Get modified users
  Log "Retrieving modified users from " & strDatabaseName
  
  ' Get time on SQL Server
  Log "Getting time from SQL Server"
  strSQL = "select getdate() as SQLDateTime"
  adoRecordset.Open strSQL, adoConnection
  SQLServerDateTime = adoRecordset.Fields("SQLDateTime")
  adoRecordset.Close
  
  ' Remove / character.. use - instead
  SQLServerDateTime = Format(SQLServerDateTime, gSQL_DATE_FORMAT)
  SQLServerDateTime = Replace(SQLServerDateTime, "/", "-")
  
  ' Retrieve all modified records older than this time - pass the returned date back
  Log "Retrieving modified records from Recordset"
  strSQL = "Exec " & gModifiedSP & " '" & SQLServerDateTime & "'"
  adoRecordset.Open strSQL, adoConnection
  
  If adoRecordset.EOF Or adoRecordset.BOF Then
    Log "Nothing to do"
  Else
    ' Go through this list and send updates to appropriate Exchange Server
    Log "Modifying mailboxes on " & strMSXServer
    Do Until adoRecordset.EOF
      ' Get the NT logon according to SAP
      ' ***************************************************************************
      ' NOTE the use of ' & "" ' when creating strMailbox.. used to get around NULL
      ' ***************************************************************************
      strMSXMailbox = Trim(adoRecordset("NTLogin_Id") & "")
      HRMFField = "NTLogin_Id"
      ' If the NT Logon is unknown, use the SAP Personnel Number instead
      If Trim(strMSXMailbox) = "" Then
        strMSXMailbox = Trim(adoRecordset("Personnel_Id") & "")
        HRMFField = "Personnel_Id"
      End If
      If Trim(strMSXMailbox) <> "" Then
        For Each ou In colSites
          Path = "LDAP://" & strMSXServer & "/cn=" & strMSXMailbox & ",cn=Recipients,ou=" & ou & ",o=" & gMSXOrganisation
          Err.Clear
          On Error Resume Next
          Set Mailbox = NameSpace.OpenDSObject(Path, strMSXUsername, strMSXPassword, ADS_SECURE_AUTHENTICATION)
          MSXError = Err.Number
          On Error GoTo Step2Error
          Select Case MSXError
            Case 0 ' ERROR_SUCCESS or ERROR_NOERROR
              Log "Modifying " & strMSXMailbox & " using path " & Path
              If Mailbox.Class = "organizationalPerson" Then ' should be always.. but just to make sure
                ' Clear all pre-existing values
                ' Internal Name       LDAP Property name
                NTAccount = ""      ' **Not referenced in Exchange (would be Assoc-NT-Account)**
                MSXTelephone = ""   ' telephoneNumber
                MSXFax = ""         ' facsimileTelephoneNumber
                MSXMobile = ""      ' mobile
                MSXEmail = ""       ' email OR smtp... NOT MODIFIED!!!!
                MSXPager = ""       ' pager
                MSXCompany = ""     ' Company
                MSXDepartment = ""  ' department
                MSXOffice = ""      ' physicalDeliveryOfficeName
                MSXTitle = ""       ' Positiontitle
                MSXAddress = ""     ' postalAddress
                MSXCity = ""        ' l
                MSXState = ""       ' st
                MSXZipCode = ""     ' postalCode
                MSXCountry = ""     ' co
                
                NewTelephone = ""
                NewFax = ""
                NewMobile = ""
                NewEmail = ""
                NewPager = ""
                
                NTAccount = Mailbox.Name
                NTAccount = Replace(NTAccount, "cn=", "")
                On Error GoTo Step2Error
                ' Now retrieve new mailbox details
                NewTelephone = adoRecordset("Telephone") & ""
                NewFax = adoRecordset("Fax") & ""
                NewMobile = adoRecordset("Mobile") & ""
                NewPager = adoRecordset("Pager") & ""
                ' DONT TOUCH EMAIL ADDRESS!!!!
                NewCompany = adoRecordset("Division_Desc") & ""
                NewDepartment = adoRecordset("Branch_Desc") & ""
                NewOffice = adoRecordset("Section_Desc") & ""
                NewTitle = adoRecordset("Substantive_Desc") & ""
                NewAddress = adoRecordset("Location_Desc") & ""
                ' Blank out the following fields in Exchange
                NewCity = " "
                NewState = " "
                NewZipCode = " "
                NewCountry = " "
                
                ' Modify the **CACHED** LDAP properties
                If NewTelephone <> "" Then
                  Mailbox.Put "telephoneNumber", CStr(NewTelephone)
                Else
                  Mailbox.Put "telephoneNumber", " "
                End If
                If NewFax <> "" Then
                  Mailbox.Put "facsimileTelephoneNumber", CStr(NewFax)
                Else
                  Mailbox.Put "facsimileTelephoneNumber", " "
                End If
                If NewMobile <> "" Then
                  Mailbox.Put "mobile", CStr(NewMobile)
                Else
                  Mailbox.Put "mobile", " "
                End If
                If NewPager <> "" Then
                  Mailbox.Put "pager", CStr(NewPager)
                Else
                  Mailbox.Put "pager", " "
                End If
                If NewCompany <> "" Then
                  Mailbox.Put "Company", CStr(NewCompany)
                Else
                  Mailbox.Put "Company", " "
                End If
                If NewDepartment <> "" Then
                  Mailbox.Put "department", CStr(NewDepartment)
                Else
                  Mailbox.Put "department", " "
                End If
                If NewOffice <> "" Then
                  Mailbox.Put "physicalDeliveryOfficeName", CStr(NewOffice)
                Else
                  Mailbox.Put "physicalDeliveryOfficeName", " "
                End If
                If NewTitle <> "" Then
                  Mailbox.Put "Positiontitle", CStr(NewTitle)
                Else
                  Mailbox.Put "Positiontitle", " "
                End If
                If NewAddress <> "" Then
                  Mailbox.Put "postalAddress", CStr(NewAddress)
                Else
                  Mailbox.Put "postalAddress", " "
                End If
                If NewCity <> "" Then
                  Mailbox.Put "l", CStr(NewCity)
                Else
                  Mailbox.Put "l", " "
                End If
                If NewState <> "" Then
                  Mailbox.Put "st", CStr(NewState)
                Else
                  Mailbox.Put "st", " "
                End If
                If NewZipCode <> "" Then
                  Mailbox.Put "postalCode", CStr(NewZipCode)
                Else
                  Mailbox.Put "postalCode", " "
                End If
                If NewCountry <> "" Then
                  Mailbox.Put "co", CStr(NewCountry)
                Else
                  Mailbox.Put "co", " "
                End If
                On Error Resume Next ' - possible errors... continue!!!!!
                  ' Save **CACHED** Properties to LDAP Directory
                  Err.Clear
                  Mailbox.SetInfo
                  If Err.Number <> 0 Then
                    Log "Error modifying '" & strMSXMailbox & "' (" & Err.Description & ")"
                    strSQL = "INSERT INTO HRMF_Task_Errors (Message) VALUES ('HRMF2MSX: Error modifying " & strMSXMailbox & " (" & Err.Description & ")')"
                    adoConnection.Execute strSQL
                    Err.Clear
                  Else
                    Log "Modified " & strMSXMailbox
                  End If
                  ' Delete this record from HRMF_Modifications
                  strSQL = "DELETE FROM HRMF_Modifications WHERE " & HRMFField & " = '" & strMSXMailbox & "'"
                  adoConnection.Execute strSQL
                On Error GoTo Step2Error
              End If
              Err.Clear
              Exit For
            Case -2147016656 ' Incorrect path - normal error for all but one site
            Case 13          ' Incorrect data
              Log "INCORRECT DATA - " & Path
              strSQL = "INSERT INTO HRMF_Task_Errors (Message) VALUES ('HRMF2MSX: INCORRECT DATA - " & Path & "')"
              adoConnection.Execute strSQL
            Case -2147463168
              Log "Failed to connect to LDAP server - " & Path
              strSQL = "INSERT INTO HRMF_Task_Errors (Message) VALUES ('HRMF2MSX: Failed to connect to LDAP server - " & Path & "')"
              adoConnection.Execute strSQL
            Case Else
              Log "UNKNOWN ERROR " & Err.Number & " - " & Path
              strSQL = "INSERT INTO HRMF_Task_Errors (Message) VALUES ('HRMF2MSX: UNKNOWN ERROR " & Err.Number & " - " & Path & "')"
              adoConnection.Execute strSQL
          End Select
        Next
      Else
        Log "NULL found for NT Logon ID and SAP Personnel ID - unable to reference mailbox"
        strSQL = "INSERT INTO HRMF_Task_Errors (Message) VALUES ('HRMF2MSX: NULL found for NT Logon ID and SAP Personnel ID - unable to reference mailbox')"
        adoConnection.Execute strSQL
      End If
      adoRecordset.MoveNext
    Loop
  End If
  ' Finish up
  Set Mailbox = Nothing
  Set Organisation = Nothing
  Set NameSpace = Nothing
  Set adoConnection = Nothing
  Set adoRecordset = Nothing
  Step2 = True
  Exit Function
  
Step2Error:
  Log Err.Description & " (Error " & Err.Number & ") whilst in step2 Function."
  Err.Clear
  Resume Next
End Function
