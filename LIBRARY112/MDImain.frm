VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.MDIForm MDImain 
   BackColor       =   &H8000000C&
   Caption         =   "Pustaka Library System 1.1.2"
   ClientHeight    =   4665
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7395
   Icon            =   "MDImain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgfrmIssue 
      Left            =   2640
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":D0B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":13916
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1A178
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":209DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1320
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "Close"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu munDB 
      Caption         =   "Maintain"
      Enabled         =   0   'False
      Begin VB.Menu munSubjects 
         Caption         =   "Subjects"
      End
      Begin VB.Menu munTitles 
         Caption         =   "Titles"
      End
      Begin VB.Menu munBooks 
         Caption         =   "Books"
      End
      Begin VB.Menu sdsd 
         Caption         =   "-"
      End
      Begin VB.Menu munMembers 
         Caption         =   "Members"
      End
      Begin VB.Menu munEmployees 
         Caption         =   "Employees"
      End
      Begin VB.Menu dd 
         Caption         =   "-"
      End
      Begin VB.Menu munOptions 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu munTransactions 
      Caption         =   "Transactions"
      Enabled         =   0   'False
      Begin VB.Menu munIssue 
         Caption         =   "Issue"
      End
      Begin VB.Menu munRenewal 
         Caption         =   "Renewal"
      End
      Begin VB.Menu munReturn 
         Caption         =   "Return"
      End
      Begin VB.Menu kkj 
         Caption         =   "-"
      End
      Begin VB.Menu munRes 
         Caption         =   "Reserve Book"
      End
      Begin VB.Menu munMiss 
         Caption         =   "Missing Book"
      End
      Begin VB.Menu oo 
         Caption         =   "-"
      End
      Begin VB.Menu munPayfine 
         Caption         =   "Pay Fine"
      End
   End
   Begin VB.Menu munReports 
      Caption         =   "Reports"
      Enabled         =   0   'False
      Begin VB.Menu munFineReport 
         Caption         =   "Fines Paid Report"
      End
      Begin VB.Menu munFineBal 
         Caption         =   "Fines Balance Report"
      End
      Begin VB.Menu munMissBook 
         Caption         =   "Missing Books Report"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu munAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MDImain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//////////////////////
'Version 1.1.1 Dated 23 January 2004 By RIS

Private Sub MDIForm_Load()
'MsgBox M.TotalIssueBook & M.RenewalCounter & M.MaxFineBal
End Sub

'///////////////////////
'Version 1.1.1
'Added mnuFileClose_Click
'///////////////////////
Private Sub mnuFileClose_Click()
    M.FileName = ""
    gblstrUserName = ""
    frmLogin.LoginSucceeded = False
    With MDImain
        .mnuFileOpen.Enabled = Not frmLogin.LoginSucceeded
        .mnuFileClose.Enabled = frmLogin.LoginSucceeded
        .munDB.Enabled = frmLogin.LoginSucceeded
        .munTransactions = frmLogin.LoginSucceeded
        .munReports = frmLogin.LoginSucceeded
    End With
    
    On Error GoTo Skip_While
    CloseAllActive = True
    While Forms.count > 1
        Unload Me.ActiveForm
    Wend
Skip_While:
    CloseAllActive = False


End Sub

'///////////////////////
'Version 1.1.1
'Added mnuFileExit_Click
'///////////////////////
Private Sub mnuFileExit_Click()
Unload Me
End Sub

'///////////////////////
'Version 1.1.1
'Added mnuFileOpen_Click
'///////////////////////
Private Sub mnuFileOpen_Click()
    With dlgCommonDialog
        .DialogTitle = "Open Working Database"
        .CancelError = False
        .InitDir = App.Path
        .Filter = "All Files (*.mdb)|*.mdb"
        .FileName = ""
        .ShowOpen
        If Len(.FileName) = 0 Then
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
        M.FileName = .FileName
    End With
    
    frmLogin.Show vbModal
            
    With MDImain
        .mnuFileOpen.Enabled = Not frmLogin.LoginSucceeded
        .mnuFileClose.Enabled = frmLogin.LoginSucceeded
        .munDB.Enabled = frmLogin.LoginSucceeded
        .munTransactions = frmLogin.LoginSucceeded
        .munReports = frmLogin.LoginSucceeded
    End With
    
    If frmLogin.LoginSucceeded = True Then
        M.LoadGlobalVariables
    End If
        

End Sub

Private Sub munAbout_Click()
frmAbout.Show
End Sub

Private Sub munBooks_Click()
frmBooks.Show
End Sub

Private Sub munEmployees_Click()
frmEmployees.Show
End Sub


Private Sub munFineBal_Click()
 Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & M.FileName & ";"
  '''''''''''''''''''''''''''''''''''''''
  
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select MemberId,BooksInHand,FineBal,Tel,Email,Address from Members where FineBal>0", db, adOpenStatic, adLockOptimistic
Set FineBalReport.DataSource = adoPrimaryRS
FineBalReport.Show
End Sub

Private Sub munFineReport_Click()
 Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & M.FileName & ";"

  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select MemberId,FineAmount,PayDate from Fine", db, adOpenStatic, adLockOptimistic
Set FineReport.DataSource = adoPrimaryRS

FineReport.Show

End Sub

Private Sub munIssue_Click()
'''''''
frmIssue.cmdIssue.Visible = True
frmIssue.cmdcharge.Visible = False
frmIssue.cmdrenewal.Visible = False
frmIssue.cmdreturn.Visible = False
'''''''
frmIssue.Caption = "Issue Book"
frmIssue.Icon = imgfrmIssue.ListImages(1).Picture
'''''
munIssue.Enabled = False
munRenewal.Enabled = True
munReturn.Enabled = True
'''''
frmIssue.Show
End Sub

Private Sub munMembers_Click()
frmMembers.Show
End Sub

Private Sub munMiss_Click()
frmIssue.cmdIssue.Visible = False
frmIssue.cmdcharge.Visible = False
frmIssue.cmdrenewal.Visible = False
frmIssue.cmdreturn.Visible = False
''''''''''''''
frmIssue.cmdMiss.Visible = True
'''''''
frmIssue.Caption = "Missing Book"
frmIssue.Icon = imgfrmIssue.ListImages(4).Picture
frmIssue.Show
End Sub

Private Sub munMissBook_Click()
Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & M.FileName & ";"
  '''''''''''''''''''''''''''''''''''''''
  
  Set adoPrimaryRS = New Recordset
   sSQL = "select Titles.TitleId,Titles.Subject,Titles.Title,Titles.Author,Books.BookId from Books,Titles where Titles.TitleId=Books.TitleID and books.condition='MISSING'"
  adoPrimaryRS.Open sSQL, db, adOpenStatic, adLockOptimistic
Set MissReport.DataSource = adoPrimaryRS
MissReport.Show

End Sub

Private Sub munOptions_Click()
frmOptions.Show
munOptions.Enabled = False
End Sub

Private Sub munPayfine_Click()
frmpayfine.Show
End Sub

Private Sub munRenewal_Click()
frmIssue.cmdrenewal.Visible = True
frmIssue.cmdcharge.Visible = False
frmIssue.cmdIssue.Visible = False
frmIssue.cmdreturn.Visible = False
frmIssue.Show
frmIssue.Caption = "Book Renewal"
frmIssue.Icon = imgfrmIssue.ListImages(2).Picture
munRenewal.Enabled = False
munIssue.Enabled = True
munReturn.Enabled = True
End Sub

Private Sub munRes_Click()
frmIssue.cmdIssue.Visible = False
frmIssue.cmdcharge.Visible = False
frmIssue.cmdrenewal.Visible = False
frmIssue.cmdreturn.Visible = False
''''''''''''''
frmIssue.cmdReserve.Visible = True
'''''''
frmIssue.Caption = "Reserve Book"
frmIssue.Icon = imgfrmIssue.ListImages(5).Picture
frmIssue.Show
End Sub

Private Sub munReturn_Click()
frmIssue.cmdreturn.Visible = True
frmIssue.cmdcharge.Visible = False
frmIssue.cmdIssue.Visible = False
frmIssue.cmdrenewal.Visible = False
frmIssue.Show
frmIssue.Caption = "Book Return"
frmIssue.Icon = imgfrmIssue.ListImages(3).Picture
munReturn.Enabled = False
munIssue.Enabled = True
munRenewal.Enabled = True
End Sub

Private Sub munSubjects_Click()
frmSubjects.Show
End Sub

Private Sub munTitles_Click()
frmTitles.Show
End Sub
