Attribute VB_Name = "modMain"
' **********************************************************************
' Process HRMF Data and distribute chnages to HR data to MS Exchange
' 1 command line argument supported (auto) - used to start automatically
' Usage:  HRMF2MSX.EXE AUTO
' **********************************************************************
' Version  Date         Who Reason
'   1     21 May 2001   DR  Initial
' **********************************************************************
Sub Main()
  Dim boolTemp As Boolean
  Dim strDir As String, strDate As String
  ' Set up registry as required
  ' Create registry keys
  boolTemp = CreateKey("", "Software", "CSC")
  boolTemp = CreateKey("", "Software\CSC", "HRMF2MSX")
  
  ' Get current values from Registry
  GetDefaults
  
  ' Ensure only one copy of App is running
  If App.PrevInstance Then
    Log "Application closed due to previous instance already running"
    End
  End If
  
  Log "Application started"
  
  ' Remove log files 14 days and older
  strDir = Dir(App.Path & "\HRMF2MSX*.log")
    Do Until strDir = ""
      strDate = FileDateTime(App.Path & "\" & strDir)
      If DateDiff("d", Format(strDate, "DD/MM/YYYY"), Date) >= 14 Then
        Kill App.Path & "\" & strDir
      End If
      strDir = Dir
    Loop
  
  Start
  If Err.Number = 0 Or gError <> "" Then
    Log "Application ended cleanly."
  Else
    Log "Application ended with one or more errors."
  End If
  End
End Sub


Sub Start()
  Dim StepStatus As Boolean
  On Error GoTo StartError
  Log "Transfer started"
  
  ' Two steps to perform:
  ' 1.  Update Table (HRMF_Admin_User) with names of users that may administer phonebook
  ' 2.  Update exchange servers with mailbox details
  StepStatus = Step1
  If Not StepStatus Then Log "Step 1 error"
  StepStatus = Step2
  If Not StepStatus Then Log "Step 2 error"
  
  Exit Sub
StartError:
  Log Err.Description & " (Error " & Err.Number & ") whilst in extract process"
  Err.Clear
  Resume Next
End Sub
