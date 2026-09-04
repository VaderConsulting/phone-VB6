VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HRMF to Microsoft Exchange Transfer"
   ClientHeight    =   2235
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6120
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   6120
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstStatus 
      Height          =   1620
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   5775
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   9360
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   9360
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuStart 
         Caption         =   "Start"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
' **********************************************************************
' Process HRMF Data and distribute chnages to HR data to MS Exchange
' 1 command line argument supported (auto) - used to start automatically
' Usage:  HRMF2MSX.EXE AUTO
' **********************************************************************
' Version  Date         Who Reason
'   1     21 May 2001   DR  Initial
' **********************************************************************

Private Sub Form_Load()
  Dim boolTemp As Boolean
  ' Set up registry as required
  ' Create registry keys
  boolTemp = CreateKey("", "Software", "CSC")
  boolTemp = CreateKey("", "Software\CSC", "HRMF2MSX")
  
  ' Get current values from Registry
  GetDefaults
  
  ' Ensure only one copy of App is running
  If App.PrevInstance Then
    Unload Me
    End
  End If
  
  Log "Application started"
  ' Set global variable to indicate running mode
  gAuto = False
  If InStr(1, UCase(Command$), "AUTO") > 0 Then
    gAuto = True
    mnuStart_Click
    Unload Me
    End
  End If
  Me.Show
  Me.Refresh
End Sub

Private Sub mnuExit_Click()
  Dim intRetval As Integer
  If gAuto Then
    intRetval = vbYes ' if automatic mode, then default to yes
  Else
    intRetval = MsgBox("Exit HRMF2MSX ?", vbYesNo + vbQuestion, "Confirm exit")
  End If
  If intRetval = vbYes Then
    Log "Application ended"
    Unload Me
    End
  End If
End Sub

Private Sub mnuOptions_Click()
  frmOptions.Show vbModal
End Sub

Private Sub mnuStart_Click()
  Dim StepStatus As Boolean
  On Error GoTo StartError
  Log "Transfer started"
  Screen.MousePointer = vbHourglass
  ' Two steps to perform:
  ' 1.  Update Table (HRMF_Admin_User) with names of users that may administer phonebook
  ' 2.  Update exchange servers with mailbox details
  StepStatus = Step1
  If StepStatus Then ' The step was successful, so continue
    StepStatus = Step2
    If Not StepStatus Then Log "Step 2 error"
  Else
    Log "Step 1 error"
  End If
  Screen.MousePointer = vbDefault
  Exit Sub
StartError:
  Log Err.Description & " (Error " & Err.Number & ") whilst in extract process"
  Err.Clear
  Screen.MousePointer = vbDefault
  Resume Next
End Sub
