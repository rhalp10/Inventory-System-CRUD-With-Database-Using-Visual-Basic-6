Attribute VB_Name = "mod"
Option Explicit
'CODED BY:  Welch Regime Marcellana
'I hope that my code will help you
'JOIN IN MY FORUM SITE, IT'S FREE TO REGISTER!!.
'Post topic about VB Tutorials, Love/Relationships, Careers/At the Job,
'Movie, Music etc.
'www.thesacrificiallamb.com
'This is a new website and currently looking for members.
'Your registration is very much appreciated :)  Thank you.

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
