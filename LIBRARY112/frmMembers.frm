VERSION 5.00
Begin VB.Form frmMembers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Members"
   ClientHeight    =   6015
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7080
   Icon            =   "frmMembers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7080
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
      Index           =   0
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   480
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
      Index           =   0
      Left            =   5520
      TabIndex        =   38
      Top             =   120
      Width           =   1320
   End
   Begin VB.TextBox txtFields 
      DataField       =   "State"
      Height          =   285
      Index           =   13
      Left            =   2040
      TabIndex        =   23
      Top             =   3720
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Postcode"
      Height          =   285
      Index           =   12
      Left            =   2040
      TabIndex        =   21
      Top             =   3360
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Town"
      Height          =   285
      Index           =   11
      Left            =   2040
      TabIndex        =   19
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Address2"
      Height          =   285
      Index           =   10
      Left            =   2040
      TabIndex        =   17
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Options"
      Height          =   990
      Left            =   120
      TabIndex        =   35
      Top             =   6120
      Width           =   5565
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   3075
         TabIndex        =   37
         Top             =   375
         Width           =   2340
      End
      Begin VB.ComboBox comSearch 
         Height          =   315
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   375
         Width           =   1590
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   0
      ScaleHeight     =   1125
      ScaleWidth      =   7080
      TabIndex        =   34
      Top             =   4890
      Width           =   7080
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
         Left            =   5880
         Picture         =   "frmMembers.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   46
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
         Picture         =   "frmMembers.frx":6D2C
         Style           =   1  'Graphical
         TabIndex        =   45
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
         Picture         =   "frmMembers.frx":7222
         Style           =   1  'Graphical
         TabIndex        =   44
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
         Picture         =   "frmMembers.frx":773F
         Style           =   1  'Graphical
         TabIndex        =   43
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
         Picture         =   "frmMembers.frx":7BD3
         Style           =   1  'Graphical
         TabIndex        =   42
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
         Picture         =   "frmMembers.frx":81E7
         Style           =   1  'Graphical
         TabIndex        =   41
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
         Picture         =   "frmMembers.frx":875D
         Style           =   1  'Graphical
         TabIndex        =   40
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
         Picture         =   "frmMembers.frx":8CB1
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5775
      TabIndex        =   28
      Top             =   5280
      Width           =   5775
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   5385
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   33
         Top             =   0
         Width           =   4320
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Email"
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   27
      Top             =   4380
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Tel"
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   25
      Top             =   4065
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Address"
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   15
      Top             =   2300
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "FineBal"
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   13
      Top             =   1980
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BooksInHand"
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   1660
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DateOfExpire"
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   1340
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DateOfJoining"
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1020
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "LastName"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   700
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "FirstName"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   380
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "MemberId"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   60
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "State:"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   22
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Postcode:"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Town / City:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Address (Line 2):"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Email:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   26
      Top             =   4380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Tel:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   24
      Top             =   4065
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Address (Line 1):"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   2300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "FineBal:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BooksInHand:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DateOfExpire:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DateOfJoining:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "LastName:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "FirstName:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "MemberId:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmMembers"
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
    strSearchSQL = "select MemberId,FirstName,LastName,DateOfJoining,DateOfExpire,BooksInHand,FineBal,Address,Tel,Email,Address2,Town,Postcode,State from Members"
    frmSearch.FindSearchText strSearchSQL, 0
    frmSearch.Show vbModal
    txtSearchText_LostFocus (0)
    txtSearchText(0).SetFocus
End Sub

'//////////////////////
'Version 1.1.2
'Amended to include additional fields:
'Address2, Town, Postcode and State
'//////////////////////
Private Sub Form_Load()
  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & M.FileName & ";"

  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select MemberId,FirstName,LastName,DateOfJoining,DateOfExpire,BooksInHand,FineBal,Address,Tel,Email,Address2,Town,Postcode,State from Members", db, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
    oText.Locked = True
  Next

  mbDataChanged = False
  ''''''''''''''''''''''''''''''''''
  comSearch.AddItem ("MemberId")
  comSearch.AddItem ("FirstName")
  comSearch.AddItem ("LastName")
  comSearch.ListIndex = 0
  
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
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
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
  
    
    .AddNew
    lblStatus.Caption = "Add record"
    mbAddNewFlag = True
    SetButtons False
  End With
  '''''''''''''''''''''''''''''''''

  txtFields(3).Text = Date
  txtFields(4).Text = DateAdd("m", M.MembershipDuration, Date)
  txtFields(5).Text = "0" 'books in hand
  txtFields(6).Text = M.MembershipFee

    
    '''''''''''''''''''''''''''''''''''
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
If adoPrimaryRS.Fields(5) > 0 Then
MsgBox "The Member should return all the books before his record is Deleted"
Exit Sub
End If
If adoPrimaryRS.Fields(6) > 0 Then
MsgBox "The Member should clear all the Fines before his record is Deleted"
Exit Sub
End If

  With adoPrimaryRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
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
  txtFields(0).Enabled = Not bVal 'member ID
    For Each oText In Me.txtFields
        oText.Locked = bVal
    Next

  End Sub


Private Sub txtFields_LostFocus(Index As Integer)
If Index = 8 And Not IsNumeric(txtFields(Index).Text) Then
MsgBox "Enter a Telephone number!!!"
txtFields(Index).Text = ""
txtFields(Index).SetFocus
End If

End Sub

'//////////////////////
'Version 1.1.2
'Added txtSearchText as part of Search function
'//////////////////////
Private Sub txtSearchText_LostFocus(Index As Integer)
    If txtSearchText(Index) = "" Then Exit Sub
    If CheckSearchText("Members", txtFields(Index).DataField, txtSearchText(Index).Text) = True Then
        adoPrimaryRS.MoveFirst
        adoPrimaryRS.Find (txtFields(Index).DataField & "='" & txtSearchText(Index).Text & "' ")
        Else
        txtSearchText(Index).SetFocus
    End If
End Sub

