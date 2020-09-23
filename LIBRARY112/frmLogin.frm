VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2565
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515.487
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label Label1 
      Caption         =   "The user name is 1234 and the password is 1234. Please send your comments to ris.riscniaga@time.net.my"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//////////////////////
'Version 1.1.1 Dated 23 January 2004 By RIS
'Copyright 2004 Raja Iskandar Shah
'//////////////////////
'Version 1.1.1
'Added frmLogin
'//////////////////////

Option Explicit

Public LoginSucceeded As Boolean
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1


Private Sub cmdCancel_Click()
    LoginSucceeded = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call dbConnect
    Set adoPrimaryRS = New Recordset
    adoPrimaryRS.Open "select EmployeeId from Employees where EmployeeId='" & txtUserName & "' and Password='" & txtPassword & "' ", db, adOpenStatic, adLockOptimistic
    If adoPrimaryRS.RecordCount > 0 Then
        gblstrUserName = txtUserName
        LoginSucceeded = True
        Unload Me
    Else
        MsgBox "Invalid User Name or Password, try again!", , "Login"
        txtUserName.SetFocus
        SendKeys "{Home}+{End}"
    End If
    Set adoPrimaryRS = Nothing
    Set db = Nothing
End Sub

Private Sub Form_Load()
    LoginSucceeded = False
End Sub
