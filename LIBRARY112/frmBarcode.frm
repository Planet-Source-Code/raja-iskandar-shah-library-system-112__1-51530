VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmBarcode 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preview Barcode"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   Icon            =   "frmBarcode.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1545
      Left            =   90
      ScaleHeight     =   1545
      ScaleWidth      =   8745
      TabIndex        =   12
      Top             =   3300
      Width           =   8745
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "BarCode 128 Properties"
      Height          =   825
      Left            =   60
      TabIndex        =   3
      Top             =   720
      Width           =   8805
      Begin VB.CheckBox ChkExecute 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Display after Save"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   540
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox checkBarCaption 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bar With Caption"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   210
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkTextAlignment 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bottom Align Caption"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   210
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.ComboBox cboBarSize 
         Height          =   315
         ItemData        =   "frmBarcode.frx":6852
         Left            =   7710
         List            =   "frmBarcode.frx":685F
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   180
         Width           =   975
      End
      Begin VB.ComboBox cboTextStyle 
         Height          =   315
         ItemData        =   "frmBarcode.frx":6879
         Left            =   5640
         List            =   "frmBarcode.frx":6889
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Bar Size"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6780
         TabIndex        =   7
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Text Style"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4620
         TabIndex        =   5
         Top             =   210
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Type of Bar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5370
      TabIndex        =   2
      Top             =   -30
      Width           =   3165
      Begin VB.CheckBox chkBar39 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bar 39"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1740
         TabIndex        =   11
         Top             =   180
         Value           =   1  'Checked
         Width           =   1395
      End
      Begin VB.CheckBox chkBar128 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bar 128"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   10
         Top             =   210
         Width           =   1425
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1545
      Left            =   90
      ScaleHeight     =   1545
      ScaleWidth      =   8745
      TabIndex        =   1
      Top             =   1650
      Width           =   8745
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1650
      Locked          =   -1  'True
      MaxLength       =   21
      TabIndex        =   0
      Text            =   "S123456T789"
      Top             =   90
      Width           =   3675
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   540
      Top             =   5370
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Text to Create Bar Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   180
      TabIndex        =   14
      Top             =   180
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Right Click Over Each Bar to Get Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   30
      TabIndex        =   13
      Top             =   4920
      Width           =   6375
   End
   Begin VB.Menu mnucp 
      Caption         =   "Copy Image"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyImage1 
         Caption         =   "Copy  Bar of Code 128 to Clip Board"
      End
      Begin VB.Menu mnuSaveImage1 
         Caption         =   "Save As"
      End
   End
   Begin VB.Menu mnucp1 
      Caption         =   "Copy Image2"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyImage2 
         Caption         =   "Copy This Bar of Code 39 to Clip Board"
      End
      Begin VB.Menu mnuSaveImage2 
         Caption         =   "Save As"
      End
   End
End
Attribute VB_Name = "frmBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//////////////////////
'Version 1.1.2 Dated 1 February 2004 By RIS
'Copyright 2004 Raja Iskandar Shah
'//////////////////////

'//////////////////////
'Version 1.1.2
'Added form frmBarcode to enable barcode printing of ID
'//////////////////////


'---------------------------------------------------------------------------------------
' Module    : Form1
' DateTime  : 10/7/2002 19:40
' Author    : Joyprakash Saikia
' Purpose   :  To Display the BarCode and Illustrate the use of it
'---------------------------------------------------------------------------------------
Option Explicit

Dim Target As Object
Dim LastSave As String 'To Store last Saved Directory
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim strBarcode As String


Private Sub cboBarSize_Click()
'---------------------------------------------------------------------------------------
' Procedure : cboBarSize_Click
' DateTime  : 10/7/2002 19:40
' Author    : Joyprakash
' Purpose   : This Method is basically used to resize the BarCode
'---------------------------------------------------------------------------------------

Picture2.ScaleMode = 3
Picture1.ScaleMode = 3
Select Case cboBarSize.ListIndex
    Case 0
        Picture2.Height = Picture2.Height * (1.4 * 40 / Picture2.ScaleHeight)
        Picture2.FontSize = 8
    Case 1
        Picture2.Height = Picture2.Height * (2.4 * 40 / Picture2.ScaleHeight)
        Picture2.FontSize = 10
    Case 2
        Picture2.Height = Picture2.Height * (3 * 40 / Picture2.ScaleHeight)
        Picture2.FontSize = 14
End Select
    Text1_Change
End Sub

Private Sub cboTextStyle_Click()
    Text1_Change
End Sub

Private Sub checkBarCaption_Click()
If checkBarCaption.Value = 1 Then
    checkBarCaption.Caption = "Show Text Caption"
    chkTextAlignment.Visible = True
Else
    checkBarCaption.Caption = "No Caption"
    chkTextAlignment.Value = 0
    chkTextAlignment.Visible = False
End If
End Sub

Private Sub chkBar128_Click()
    If chkBar128.Value = 0 Then
        Picture1.Visible = False
    Else
        Picture1.Visible = True
    End If

End Sub

Private Sub chkBar39_Click()
   
    If chkBar39.Value = 0 Then
        Picture2.Visible = False
    Else
        Picture2.Visible = True
    End If
End Sub

Private Sub chkTextAlignment_Click()
If chkTextAlignment.Value = 1 Then
    chkTextAlignment.Caption = "Bottom Align Caption"
Else
        chkTextAlignment.Caption = "Top Align Caption"
End If
End Sub

'//////////////////////
'Version 1.1.2
'removed function printbar128 which is not used
'//////////////////////
Private Sub printBar128()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub


Private Sub Form_Load()
Text1.Text = strBarcode
cboBarSize.ListIndex = 0
cboTextStyle.ListIndex = 0
End Sub

Private Sub mnuCopyImage1_Click()
    pCopyToClipBoard Picture1
End Sub

Private Sub mnuCopyImage2_Click()
     pCopyToClipBoard Picture2
End Sub

Private Sub mnuSaveImage1_Click()
    'Call FileSaving Routine for BarCode 128
     FileSave Picture1
     If ChkExecute.Value = 1 Then
     'if Asked to Show the exported File
        ShellExecute Me.hwnd, "open", CDialog.FileName, 0&, LastSave, vbNormalFocus
     End If
End Sub

Private Sub mnuSaveImage2_Click()
    'Call FileSaving Routine for BarCode 39
     FileSave Picture2
     If ChkExecute.Value = 1 Then
     'if Asked to Show the exported File
        ShellExecute Me.hwnd, "open", CDialog.FileName, 0&, LastSave, vbNormalFocus
     End If
    
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And Text1 <> "" Then
         
        PopupMenu mnucp
    End If
End Sub



Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 And Text1 <> "" Then
    
    PopupMenu mnucp1
End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Text1_Change
' DateTime  : 10/7/2002 19:48
' Author    : Joy
' Purpose   : This Procedure is used to Print Barcode
'            I used Change event to Show you the effects
'---------------------------------------------------------------------------------------
'
Private Sub Text1_Change()

Dim MinWidth, pw
If chkBar128.Value = 1 Then
   printBar128
 End If
 If chkBar39.Value = 1 Then
    Call DrawBarcode(Text1, Picture2)
End If


End Sub
Sub pCopyToClipBoard(ByVal PictureBox As Object)
    ' copy The Image to clipboard
    PictureBox.Picture = PictureBox.Image
    Clipboard.Clear
    Clipboard.SetData PictureBox.Image, 2
End Sub


Sub FileSave(Picbox As PictureBox)
    'This Procedure Saves the Bars to desired Format
    Dim sName, RetVal, retSave
    Dim ObjGifImg As GIF
    Screen.MousePointer = 11
    
    CDialog.FilterIndex = 1 'make Dialog as Save
    If LastSave <> "" Then
        CDialog.InitDir = LastSave
    Else
        CDialog.InitDir = App.Path & "\ExportedImages"
    End If
    
    On Error GoTo ErrHandler
    CDialog.FileName = Picbox.Name
    CDialog.CancelError = True
    CDialog.Flags = cdlOFNOverwritePrompt + cdlOFNNoReadOnlyReturn
    CDialog.Filter = "Bitmaps (*.bmp)|*.bmp|Gif (*.gif)|*.gif|Transparent Gif (*.gif)|*.gif"
    CDialog.ShowSave
    'retrive the Folder Name
    retSave = InStrRev(CDialog.FileName, "\")
    LastSave = Mid(CDialog.FileName, 1, retSave)
    DoEvents
    
    Picbox.Picture = Picbox.Image
    Select Case CDialog.FilterIndex
      Case 1: 'if BITMAP is selected
                SavePicture Picbox.Picture, CDialog.FileName
      Case 2: 'if GIF is selected
                Set ObjGifImg = New GIF
                ObjGifImg.SaveGIF Picbox.Image, CDialog.FileName, Picbox.hDc, False, Picbox.Point(0, 0)
                Set ObjGifImg = Nothing
      Case 2: 'if Transperent GIF is selected
                Set ObjGifImg = New GIF
                ObjGifImg.SaveGIF Picbox.Image, CDialog.FileName, Picbox.hDc, True, Picbox.Point(0, 0)
                Set ObjGifImg = Nothing
    End Select

Screen.MousePointer = 0
Exit Sub
ErrHandler:

    If Err.Number = 32755 Then ' Handle the Cancel error
        Screen.MousePointer = 0
        Exit Sub
    Else
            If Err.Number <> 0 Then MsgBox "Error saving file: " & Err.Number & " - " & Err.Description
            Screen.MousePointer = 0
    End If
    
End Sub

Public Sub PreviewBarcode(BarcodeNo As String)
    strBarcode = BarcodeNo
End Sub
