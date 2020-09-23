VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmBooks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Books"
   ClientHeight    =   4245
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "frmBooks.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7200
   Begin VB.CommandButton cmdBarCode 
      BackColor       =   &H80000002&
      Caption         =   "Bar Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtSearchText 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   5640
      TabIndex        =   12
      Top             =   360
      Width           =   1320
   End
   Begin MSDataListLib.DataCombo DComTitleId 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2025
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   675
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Options"
      Height          =   990
      Left            =   150
      TabIndex        =   28
      Top             =   4980
      Width           =   5565
      Begin VB.ComboBox comSearch 
         Height          =   315
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   375
         Width           =   1590
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   3075
         TabIndex        =   29
         Top             =   375
         Width           =   2340
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1110
      Left            =   0
      ScaleHeight     =   1110
      ScaleWidth      =   7200
      TabIndex        =   27
      Top             =   3135
      Width           =   7200
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H80000002&
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   5940
         Picture         =   "frmBooks.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H80000002&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   1275
         Picture         =   "frmBooks.frx":6D2C
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H80000002&
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   120
         Picture         =   "frmBooks.frx":7222
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H80000002&
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   4740
         Picture         =   "frmBooks.frx":773F
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H80000002&
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   3585
         Picture         =   "frmBooks.frx":7BD3
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H80000002&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   2430
         Picture         =   "frmBooks.frx":81E7
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H80000002&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   1275
         Picture         =   "frmBooks.frx":875D
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H80000002&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   120
         Picture         =   "frmBooks.frx":8CB1
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   75
      ScaleHeight     =   300
      ScaleWidth      =   7950
      TabIndex        =   21
      Top             =   4230
      Width           =   7950
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   5325
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4950
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   675
         TabIndex        =   26
         Top             =   0
         Width           =   4260
      End
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "TypeIssue"
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   11
      Top             =   2620
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ReserveId"
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   10
      Top             =   2300
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "IssueCounter"
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   9
      Top             =   1980
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ReturnDate"
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   8
      Top             =   1660
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "MemberId"
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   7
      Top             =   1340
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "IsIn"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   6
      Top             =   1020
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Condition"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   4
      Top             =   700
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BookId"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   380
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TitleId"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   60
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "TypeIssue:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   20
      Top             =   2620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ReserveId:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   19
      Top             =   2300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "IssueCounter:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   18
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ReturnDate:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   17
      Top             =   1660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "MemberId:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "IsIn:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Condition:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BookId:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TitleId:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmBooks"
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

Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS2 As Recordset
Attribute adoPrimaryRS2.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS3 As Recordset
Attribute adoPrimaryRS3.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

'//////////////////////
'Version 1.1.1
'Added cmdbarcode
'//////////////////////
Private Sub cmdBarcode_Click(Index As Integer)
    frmBarcode.PreviewBarcode (txtFields(Index).Text)
    frmBarcode.Show vbModal
End Sub

'//////////////////////
'Version 1.1.2
'Added cmdSearch as part of search function
'//////////////////////
Private Sub cmdSearch_Click()
    If adoPrimaryRS.RecordCount = 0 Then Exit Sub
    
    Dim strSearchSQL As String
    strSearchSQL = "select BookId,TitleId,Condition,IsIn,MemberId,ReturnDate,IssueCounter,ReserveId,TypeIssue from Books"
    frmSearch.FindSearchText strSearchSQL, 1
    frmSearch.Show vbModal
    txtSearchText_LostFocus (1)
    txtSearchText(1).SetFocus
End Sub

'//////////////////////
'Version 1.1.2
'amended the auto generation of book id
'so that new books on existing titles can be added from this frmBooks
'//////////////////////

Private Sub DComTitleId_Change()
Dim qunt As Integer
'adoprimaryrs2.MoveFirst
'adoprimaryrs2.Find ("TitleId" & "='" & DComTitleId & "'")
'qunt = adoprimaryrs2.Fields(1)
adoPrimaryRS3.Requery
adoPrimaryRS3.Filter = "TitleId = '" & DComTitleId & "' "
qunt = adoPrimaryRS3.RecordCount + 1
For i = 0 To qunt
adoPrimaryRS3.Filter = "TitleId = '" & DComTitleId & "' and BookId = '" & DComTitleId & "/" & qunt & "'"
If adoPrimaryRS3.RecordCount > 0 Then
    qunt = qunt + 1
End If
adoPrimaryRS3.Filter = "TitleId = '" & DComTitleId & "' "
Next

txtFields(1).Text = DComTitleId & "/" & qunt
adoPrimaryRS3.Filter = ""
End Sub


Private Sub Form_Load()
  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & M.FileName & ";"

  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select TitleId,BookId,Condition,IsIn,MemberId,ReturnDate,IssueCounter,ReserveId,TypeIssue from Books", db, adOpenStatic, adLockOptimistic

  Set adoPrimaryRS3 = New Recordset
  adoPrimaryRS3.Open "select TitleId,BookId,Condition,IsIn,MemberId,ReturnDate,IssueCounter,ReserveId,TypeIssue from Books", db, adOpenStatic, adLockOptimistic
  
  Set adoPrimaryRS2 = New Recordset
  adoPrimaryRS2.Open "select TitleId,Quantity from Titles", db, adOpenStatic, adLockOptimistic
  
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
    oText.Locked = True
  Next
  Dim oCheck As CheckBox
  'Bind the check boxes to the data provider
  For Each oCheck In Me.chkFields
    Set oCheck.DataSource = adoPrimaryRS
  Next

  mbDataChanged = False
  '''''''''''''''''''''''''''''''''''''
  Combo1.AddItem ("EXCELLENT")
  Combo1.AddItem ("GOOD")
  Combo1.AddItem ("POOR")
  Combo1.AddItem ("WORST")
  Combo1.AddItem ("MISSING")
  Combo1.ListIndex = 0
  comSearch.AddItem ("BookId")
  comSearch.AddItem ("TitleId")
  comSearch.AddItem ("MemberId")
  comSearch.AddItem ("ReserveId")
  comSearch.ListIndex = 0
  ''''''''''''''''''''''''''''''''''''''
txtFields(0).Locked = True
txtFields(1).Locked = True
''''''''''''''''''''''''''''''''''''''''''
Set DComTitleId.DataSource = adoPrimaryRS2
Set DComTitleId.RowSource = adoPrimaryRS2
 DComTitleId.ListField = "TitleId"
''''''''''''''''''''''''''''''''''''''''''
 If M.BooksCallByTitle Then
 'cmdAdd_Click     ' ris - something that i need to look at
 End If
 
  End Sub




Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  '''''''''''''''''''
  DComTitleId.Refresh
  '''''''''''''''''''
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    lblStatus.Caption = "Add record"
    mbAddNewFlag = True
    SetButtons False
  End With
txtFields(6).Text = "0"
txtFields(4) = "0"
txtFields(7) = "0"
DComTitleId.Text = ""
DComTitleId.SetFocus
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  Dim q As Integer
  On Error GoTo DeleteErr
  With adoPrimaryRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  '''''''''''''''''''''''''''''''''''
'update the title table '''''''''''''
adoPrimaryRS2.MoveFirst
adoPrimaryRS2.Find ("TitleId" & "='" & txtFields(0).Text & "'")
q = adoPrimaryRS2.Fields(1)
q = q - 1
adoPrimaryRS2.Fields(1) = q
adoPrimaryRS2.Update
''''''''''''''''''''''''''''''''''''
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  adoPrimaryRS.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr
  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
'''''''''''''''''''''''''''''
'hide the combos cause no edit allowed to title id and bookid
Combo2.Visible = False
DComTitleId.Visible = False
'''''''''''''''''''''''''''''
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
  ''''''''''''''''''''''''''''''''
txtFields(2).Text = Combo1.Text
txtFields(0).Text = DComTitleId.Text
'txtFields(1).Text = Combo2.Text

''''''''''''''''''''''''''''''''''
  adoPrimaryRS.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'move to the new record
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False

  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError
'i dont want to make canges while moving << or >>
'just use the add update buttons
adoPrimaryRS.CancelUpdate

  adoPrimaryRS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError
'i dont want to make canges while moving << or >>
'just use the add update buttons
adoPrimaryRS.CancelUpdate

  adoPrimaryRS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError
'i dont want to make canges while moving << or >>
'just use the add update buttons
adoPrimaryRS.CancelUpdate

  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    adoPrimaryRS.MoveLast
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError
'i dont want to make canges while moving << or >>
'just use the add update buttons
adoPrimaryRS.CancelUpdate

  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoPrimaryRS.MoveFirst
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
  ''extra code for our frame
  Frame1.Enabled = bVal
  Combo1.Visible = Not bVal
  DComTitleId.Visible = Not bVal
    For Each oText In Me.txtFields
        oText.Locked = bVal
    Next

  'Combo2.Visible = Not bVal


End Sub

Private Sub txtFields_Change(Index As Integer)
txtFields(Index).Text = UCase(Trim(txtFields(Index).Text))
End Sub

'//////////////////////
'Version 1.1.2
'Added txtSearchText as part of Search function
'//////////////////////
Private Sub txtSearchText_LostFocus(Index As Integer)
    If txtSearchText(Index) = "" Then Exit Sub
    If CheckSearchText("Books", txtFields(Index).DataField, txtSearchText(Index).Text) = True Then
        adoPrimaryRS.MoveFirst
        adoPrimaryRS.Find (txtFields(Index).DataField & "='" & txtSearchText(Index).Text & "' ")
        txtSearchText(Index).Text = ""
        Else
        txtSearchText(Index).SetFocus
    End If
End Sub

