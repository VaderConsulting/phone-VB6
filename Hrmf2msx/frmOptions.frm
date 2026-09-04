VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5985
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab tabOptions 
      Height          =   3735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6588
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Database"
      TabPicture(0)   =   "frmOptions.frx":0742
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDbVersion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDbServerName"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDbName"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDbUsername"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDbPassword"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblDbOwner"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtServerName"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtDatabaseName"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "optServerVersion(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "optServerVersion(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtDbUsername"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtDbPassword"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtDbOwner"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "fmeSP"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Exchange"
      TabPicture(1)   =   "frmOptions.frx":075E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtOrgName"
      Tab(1).Control(1)=   "txtMSXPassword"
      Tab(1).Control(2)=   "txtMSXUsername"
      Tab(1).Control(3)=   "txtMSXServer"
      Tab(1).Control(4)=   "lblOrgName"
      Tab(1).Control(5)=   "lblMSXPassword"
      Tab(1).Control(6)=   "lblMSXUsername"
      Tab(1).Control(7)=   "lblMSXServerName"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Network"
      TabPicture(2)   =   "frmOptions.frx":077A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtGroupName"
      Tab(2).Control(1)=   "txtDomainName"
      Tab(2).Control(2)=   "lblAdminGroupName"
      Tab(2).Control(3)=   "lblDomain"
      Tab(2).ControlCount=   4
      Begin VB.Frame fmeSP 
         Caption         =   "Stored Procedures and Tables"
         Height          =   1695
         Left            =   120
         TabIndex        =   28
         Top             =   1920
         Width           =   5535
         Begin VB.TextBox txtAdminTable 
            Height          =   285
            Left            =   120
            TabIndex        =   34
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtUpdates 
            Height          =   285
            Left            =   2880
            TabIndex        =   32
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox txtModified 
            Height          =   285
            Left            =   120
            TabIndex        =   30
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label lblAdminTable 
            Alignment       =   2  'Center
            Caption         =   "Admin Users"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label lblCleanup 
            Alignment       =   2  'Center
            Caption         =   "Remove Pending Updates"
            Height          =   255
            Left            =   2880
            TabIndex        =   31
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label lblModified 
            Alignment       =   2  'Center
            Caption         =   "Retrieve Pending Updates"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.TextBox txtGroupName 
         Height          =   285
         Left            =   -72960
         TabIndex        =   27
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtOrgName 
         Height          =   285
         Left            =   -72960
         TabIndex        =   25
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtDomainName 
         Height          =   285
         Left            =   -74880
         TabIndex        =   23
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtMSXPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -72960
         PasswordChar    =   "*"
         TabIndex        =   22
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtMSXUsername 
         Height          =   285
         Left            =   -74880
         TabIndex        =   20
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtMSXServer 
         Height          =   285
         Left            =   -74880
         TabIndex        =   18
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtDbOwner 
         Height          =   285
         Left            =   3960
         TabIndex        =   15
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtDbPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtDbUsername 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton optServerVersion 
         Caption         =   "7.0"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   9
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton optServerVersion 
         Caption         =   "6.5"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   8
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtDatabaseName 
         Height          =   285
         Left            =   3960
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtServerName 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblAdminGroupName 
         Alignment       =   2  'Center
         Caption         =   "Admin Group Name"
         Height          =   255
         Left            =   -72960
         TabIndex        =   26
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblOrgName 
         Alignment       =   2  'Center
         Caption         =   "Organisation Name"
         Height          =   255
         Left            =   -73080
         TabIndex        =   24
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblMSXPassword 
         Alignment       =   2  'Center
         Caption         =   "Password"
         Height          =   255
         Left            =   -72960
         TabIndex        =   21
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblMSXUsername 
         Alignment       =   2  'Center
         Caption         =   "Username"
         Height          =   255
         Left            =   -74880
         TabIndex        =   19
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblMSXServerName 
         Alignment       =   1  'Right Justify
         Caption         =   "Server Name"
         Height          =   255
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblDomain 
         Alignment       =   2  'Center
         Caption         =   "Account Domain Name"
         Height          =   255
         Left            =   -74880
         TabIndex        =   16
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblDbOwner 
         Alignment       =   2  'Center
         Caption         =   "Owner"
         Height          =   255
         Left            =   3960
         TabIndex        =   14
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblDbPassword 
         Alignment       =   2  'Center
         Caption         =   "Password"
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblDbUsername 
         Alignment       =   2  'Center
         Caption         =   "Username"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblDbName 
         Alignment       =   2  'Center
         Caption         =   "Database Name"
         Height          =   255
         Left            =   3960
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblDbServerName 
         Alignment       =   2  'Center
         Caption         =   "Server Name"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblDbVersion 
         Alignment       =   2  'Center
         Caption         =   "Server Version"
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   4080
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   9360
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   9360
      Y1              =   3975
      Y2              =   3975
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Dim strRegPath As String
  Dim Hostname As String, C As Control, boolInvalid As Boolean
  
  ' Ensure each textbox has data.. note the use of TypeName statement
  boolInvalid = False
  For Each C In frmOptions.Controls
    If TypeName(C) = "TextBox" Then
      If Trim(C.Text) = "" Then boolInvalid = True
    End If
  Next
  
  If boolInvalid Then
    MsgBox "Please ensure you have an entry in each Textbox", vbExclamation, "Error"
    Exit Sub
  End If
  
  'Save values in Textboxes (and Option button) to Global Variables
  gDbServerName = txtServerName
  If optServerVersion(0).Value = True Then
    gDb_Version = 65
  Else
    gDb_Version = 7
  End If
  gDbName = txtDatabaseName
  gDbUsername = txtDbUsername
  gDbPassword = txtDbPassword
  gDbOwner = txtDbOwner
  gModifiedSP = txtModified
  gUpdatesSP = txtUpdates
  gAdminUsers = txtAdminTable
  gMSXServer = txtMSXServer
  gMSXOrganisation = txtOrgName
  gMSXUsername = txtMSXUsername
  gMSXPassword = txtMSXPassword
  gDomainName = txtDomainName
  gAdminGroupName = txtGroupName
  
  
  ' Save modified values from Global Variables to registry
  strRegPath = "Software\CSC\HRMF2MSX"
  SetValueString "", strRegPath, "gDb_Version", CStr(gDb_Version), REG_SZ
  SetValueString "", strRegPath, "gDbServerName", gDbServerName, REG_SZ
  SetValueString "", strRegPath, "gDbName", gDbName, REG_SZ
  SetValueString "", strRegPath, "gDbUsername", gDbUsername, REG_SZ
  SetValueString "", strRegPath, "gDbPassword", gDbPassword, REG_SZ
  SetValueString "", strRegPath, "gDbOwner", gDbOwner, REG_SZ
  SetValueString "", strRegPath, "gModifiedSP", gModifiedSP, REG_SZ
  SetValueString "", strRegPath, "gUpdatesSP", gUpdatesSP, REG_SZ
  SetValueString "", strRegPath, "gAdminUsers", gAdminUsers, REG_SZ
  SetValueString "", strRegPath, "gMSXServer", gMSXServer, REG_SZ
  SetValueString "", strRegPath, "gMSXOrganisation", gMSXOrganisation, REG_SZ
  SetValueString "", strRegPath, "gMSXUsername", gMSXUsername, REG_SZ
  SetValueString "", strRegPath, "gMSXPassword", gMSXPassword, REG_SZ
  SetValueString "", strRegPath, "gDomainName", gDomainName, REG_SZ
  SetValueString "", strRegPath, "gAdminGroupName", gAdminGroupName, REG_SZ
  Unload Me
End Sub

Private Sub Form_Load()
  ' Load text boxes from defaults previously retrieved from Registry
  txtServerName = gDbServerName
  Select Case CInt(gDb_Version)
    Case 65
      optServerVersion(0).Value = True
    Case 7
      optServerVersion(1).Value = True
  End Select
  txtDatabaseName = gDbName
  txtDbUsername = gDbUsername
  txtDbPassword = gDbPassword
  txtDbOwner = gDbOwner
  txtModified = gModifiedSP
  txtUpdates = gUpdatesSP
  txtAdminTable = gAdminUsers
  txtMSXServer = gMSXServer
  txtOrgName = gMSXOrganisation
  txtMSXUsername = gMSXUsername
  txtMSXPassword = gMSXPassword
  txtDomainName = gDomainName
  txtGroupName = gAdminGroupName
End Sub
