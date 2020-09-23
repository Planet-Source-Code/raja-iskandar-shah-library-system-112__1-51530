Attribute VB_Name = "Search"
'//////////////////////
'Version 1.1.2 Dated 1 February 2004 By RIS
'Copyright 2004 Raja Iskandar Shah
'//////////////////////

'//////////////////////
'Version 1.1.2
'Added module Search as part of search function
'//////////////////////

Option Explicit


Public Function CheckSearchText(TableName As String, FieldName As String, SearchText As String) As Boolean
    Call dbConnect
    Dim rs As Recordset
    Set rs = New Recordset
    rs.Open "select " & FieldName & " from " & TableName & " Where " & FieldName & " = '" & SearchText & "' ", db, adOpenStatic, adLockOptimistic
    If rs.RecordCount = 0 Then
        MsgBox "The text " & SearchText & " is not found. Please enter again or click on the Search button. Otherwise, you may add a new record by clicking on the Add button.", vbInformation, "Sorry."
        CheckSearchText = False
        Else
        CheckSearchText = True
    End If
    rs.Close
    Set rs = Nothing
    
End Function

