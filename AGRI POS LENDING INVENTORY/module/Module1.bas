Attribute VB_Name = "Connection"
Option Explicit
Global RS As New ADODB.Recordset
Global CN As New ADODB.Connection
Public DB As ADODB.Connection
Global CNf As New ADODB.Connection
Global RSf As New ADODB.Recordset
Global cncategory3 As New ADODB.Connection
Global rscategory3 As New ADODB.Recordset


Global UserName As String
Global password As String
Global retvalue As Integer
Global Worm As Integer
Global cnjen As New ADODB.Connection
Global rstransaction As New ADODB.Recordset
Global path As Variant



Public Function Found(ByRef Rslog As ADODB.Recordset, ByVal sField As String, ByVal sfindtext As String) As Boolean
  Rslog.Requery
    Rslog.Find sField & " = '" & sfindtext & "'"

If Rslog.EOF Then
    Found = False
Else
    Found = True
    UserName = Rslog.Fields(0)
    password = Rslog.Fields(1)
End If

End Function

''''''''''''''''''''''''access connection'''''''''''''''''''''''
Public Sub connect(ByRef Con As ADODB.Connection, ByVal dataloc As String)
            Con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dataloc & ";Persist Security Info=False"
End Sub
''''''''''''''SQL Connection'''''''''''''''''''''''''''''
'Public Sub SQLconnect(ByRef Con As ADODB.Connection)
 'Con.Open "Provider=SQLOLEDB.1;Persist Security Info=false;User Id=Administrador;Password=Admin;Initial Catalog=Tailor;Data Source=yuki" 'SQL Connection




Public Sub SetRs(ByRef sRec As ADODB.Recordset, ByRef sCon As ADODB.Connection, ByVal sSQL As String)

With sRec
       .CursorLocation = adUseClient
        .Open sSQL, sCon, adOpenKeyset, adLockPessimistic
    End With

End Sub


Public Sub developer()
retvalue = GetSetting("A", "0", "Runcount")
Worm = Val(retvalue) + 1
SaveSetting "A", "0", "RunCount", Worm

If Worm > 9999 Then
MsgBox "This is the End of the trial run....", 16, " This is a Trial Version Copy"
    MsgBox "Email the developer of this system @ venticjojo@yahoo.com or call 09186070112 Ask for the FULL VERSION", 6, "System Developer"
    Unload frmLogin
End If
'Upload
End Sub
Public Sub delete_rec(ByRef sCONN As ADODB.Connection, ByVal sTable As String, ByVal sField As String, ByVal sString As String)
        sCONN.Execute "Delete * From " & sTable & " Where " & sField & " ='" & sString & "'"
End Sub

Public Sub conek()
path = App.path & "\mydb.mdb"
 If cnjen.State = 1 Then cnjen.Close
  cnjen.CursorLocation = adUseClient
  cnjen.Open "Provider=Microsoft.Jet.OLEDB.4.0;persist security info=false;data source =" & App.path & "\myDb.mdb"

End Sub

Public Sub checkrs()
If rstransaction.State = 1 Then
    rstransaction.Close
   rstransaction.LockType = adLockPessimistic
   rstransaction.CursorType = adOpenKeyset
End If
End Sub
