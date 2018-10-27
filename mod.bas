Attribute VB_Name = "mod"
Option Explicit
'Coded by: Welch Regime Marcellana
'Re-Edit by: Rhalp 10

Global Conn As New ADODB.Connection, RS As New ADODB.Recordset, Item As ListItem
Global onTop As New clsOnTop, I As Integer, SQL As String, Cancel As Boolean

Public Sub dbConnect()
  Set Conn = New ADODB.Connection
  Conn.ConnectionString = strConn
  Conn.Open
End Sub

Public Function strConn() As String
  strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database.mdb" & ";Persist Security Info=False"
End Function
