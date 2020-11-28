Attribute VB_Name = "BanTemporal"
Option Explicit
Public Baneos As New Collection

Public Function ChangeBan(ByVal Name As String, ByVal Baneado As Byte) As Boolean
Dim Orden As String

Dim RS As New ADODB.Recordset
Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")

If RS.BOF Or RS.EOF Then Exit Function

Orden = "UPDATE `charflags` SET"
Orden = Orden & " IndexPJ=" & RS!indexpj
Orden = Orden & ",Nombre='" & UCase$(Name) & "'"
Orden = Orden & ",Ban=" & Baneado
Orden = Orden & " WHERE IndexPJ=" & RS!indexpj

Call Con.Execute(Orden)

Set RS = Nothing

End Function
