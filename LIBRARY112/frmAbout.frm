VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H000080FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Pustaka Library System"
   ClientHeight    =   3990
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6345
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2753.968
   ScaleMode       =   0  'User
   ScaleWidth      =   5958.283
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   5250
      Top             =   3975
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2505
      TabIndex        =   0
      Top             =   2175
      Width           =   1260
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   690
      Left            =   135
      Top             =   2775
      Width           =   6015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "RISC Niaga Enterprise"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   390
      TabIndex        =   3
      ToolTipText     =   ".   windows_me@rediffmail.com   ."
      Top             =   495
      Width           =   5460
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Pustaka Library System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   315
      TabIndex        =   2
      ToolTipText     =   ".   Abdul Rafay Mansoor   ."
      Top             =   225
      Width           =   5610
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   281.716
      X2              =   5606.139
      Y1              =   507.31
      Y2              =   507.31
   End
   Begin VB.Label lblDisclaimer 
      BackColor       =   &H00FF8080&
      Caption         =   $"frmAbout.frx":6852
      ForeColor       =   &H00000000&
      Height          =   1125
      Left            =   315
      TabIndex        =   1
      Top             =   840
      Width           =   5610
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//////////////////////
'Version 1.1.2 Dated 1 February 2004 By RIS
'Copyright 2004 Raja Iskandar Shah
'//////////////////////

'//////////////////////
'Version 1.1.1 Dated 23 January 2004 By RIS
'Copyright 2004 Raja Iskandar Shah
'//////////////////////

'//////////////////////
'Version 1.1.1
'Amended copyright notice for Raja Iskandar,
'Rafay Mansoor, Dirceu Veiga, Philip Narapan,
'Joyprakah Saikia
'//////////////////////

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
Shape1.BorderColor = vbBlack

End Sub

Private Sub Timer1_Timer()
If Shape1.BorderColor = vbBlack Then
Shape1.BorderColor = vbRed
Else
Shape1.BorderColor = vbBlack
End If
End Sub
