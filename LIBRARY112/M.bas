Attribute VB_Name = "M"
'//////////////////////
'Version 1.1.1 Dated 23 January 2004 By RIS
'Copyright 2004 Raja Iskandar Shah
'//////////////////////

Public FileName As String
Public BooksCallByTitle As Boolean
Public TotalIssueBook, RenewalCounter, MaxFineBal, MembershipDuration, MembershipFee, RenewalFees As Integer
Public db As New ADODB.Connection
Public gblstrUserName, gblstrPassword As String


Public Sub LoadGlobalVariables()
Dim adoPrimaryRS As Recordset
Call dbConnect
Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select TotalIssueBooks,RenewalCounter,MaxFineBal, MembershipDuration, MembershipFee, RenewalFees from GlobalVariables ", db, adOpenStatic, adLockOptimistic
''''''''''''''
TotalIssueBook = adoPrimaryRS.Fields(0)
RenewalCounter = adoPrimaryRS.Fields(1)
MaxFineBal = adoPrimaryRS.Fields(2)

MembershipDuration = adoPrimaryRS.Fields(3)
MembershipFee = adoPrimaryRS.Fields(4)
RenewalFees = adoPrimaryRS.Fields(5)
db.Close
End Sub

'///////////////////////
'Version 1.1.1
'Added dbConnect
'Copyright 2003 Philip V Naparan
'///////////////////////

Public Sub dbConnect()
On Error Resume Next
    Set db = New Connection
    db.CursorLocation = adUseClient
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & M.FileName & ";"

End Sub


