Attribute VB_Name = "modFuncAndProced"
Option Explicit
'Coded by: Welch Regime Marcellana
'Re-Edit by: Rhalp 10

Public Sub loadRecords()
  Dim maxRec As Long
  Form1.ListView.ListItems.Clear
  maxRec = countAllRec

  frmLoading.Show
  I = 0
  Call dbConnect
    SQL = "SELECT tbl_info.* FROM tbl_info order by item_ID asc;"
    RS.Open SQL, Conn, adOpenDynamic
      If Not RS.EOF Then
        RS.MoveFirst
        Do While Not RS.EOF
          With Form1.ListView.ListItems
            Set Item = .Add(, , RS!item_ID)
              Item.SubItems(1) = RS!item_Name
              Item.SubItems(2) = RS!item_Descr
          End With
          I = I + 1
          frmLoading.lblSub.Caption = "Loading records..." & I & " of " & maxRec
          RS.MoveNext
          DoEvents
        Loop
      End If
    RS.Close
  Conn.Close
  Set Conn = Nothing
  Unload frmLoading
End Sub

Public Function countAllRec() As Long
  Call dbConnect
    SQL = "SELECT tbl_info.* FROM tbl_info order by item_Name asc;"
    RS.Open SQL, Conn, adOpenDynamic
      If Not RS.EOF Then
        RS.MoveFirst
        Do While Not RS.EOF
          countAllRec = countAllRec + 1
          RS.MoveNext
          DoEvents
        Loop
      End If
    RS.Close
  Conn.Close
  Set Conn = Nothing
End Function
