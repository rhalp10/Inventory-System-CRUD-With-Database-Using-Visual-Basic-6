VERSION 5.00
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Entry"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      Height          =   330
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.TextBox txtAdd 
      Height          =   330
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.PictureBox cmdAdd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3360
      Picture         =   "frmEdit.frx":0000
      ScaleHeight     =   345
      ScaleWidth      =   1185
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Item Name:"
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   300
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Description:"
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   660
      Width           =   990
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'CODED BY:  Welch Regime Marcellana
'I hope that my code will help you
'JOIN IN MY FORUM SITE, IT'S FREE TO REGISTER!!.
'Post topic about VB Tutorials, Love/Relationships, Careers/At the Job,
'Movie, Music etc.
'www.thesacrificiallamb.com
'This is a new website and currently looking for members.
'Your registration is very much appreciated :)  Thank you.

Private Sub cmdAdd_Click()
  On Error GoTo errtrap

  If Me.txtName.Text = "" Or Me.txtAdd.Text = "" Then
    MsgBox "All fields are required!", vbExclamation, "Error"
    Exit Sub
  End If
  
  If MsgBox("This action will modify the selected record.  Proceed?", vbYesNo, "Update") = vbYes Then
    SQL = "UPDATE tbl_info SET tbl_info.item_Name = '" & Me.txtName.Text & "', tbl_info.item_Descr = '" & Me.txtAdd.Text & "' " & _
          "WHERE (((tbl_info.item_ID)=" & Val(Form1.ListView.SelectedItem.Text) & "));"
    Call dbConnect
      Conn.Execute SQL
    Conn.Close
    Set Conn = Nothing
    Unload Me
    Call updateList
  Else
    Cancel = True
  End If
  
  Exit Sub
errtrap:
  Select Case Err.Number
    Case -2147467259
      MsgBox "The name already exists in the database", vbCritical, "Error"
  
    Case Else
      MsgBox Err.Description, vbCritical, "The system encountered an error"
  End Select
End Sub

Private Sub Form_Load()
  SQL = "SELECT tbl_info.item_ID, tbl_info.* From tbl_info " & _
        "WHERE (((tbl_info.item_ID)=" & Val(Form1.ListView.SelectedItem.Text) & "));"
  Call dbConnect
    RS.Open SQL, Conn, adOpenDynamic
      If Not RS.EOF Then
        Me.txtName.Text = RS!item_Name
        Me.txtAdd.Text = RS!item_Descr
      End If
    RS.Close
  Conn.Close
  Set Conn = Nothing
End Sub

Public Sub updateList()
  frmLoading.Show
  frmLoading.lblSub.Caption = "Updating record...."
  With Form1.ListView.ListItems(Form1.ListView.SelectedItem.Index)
    SQL = "SELECT tbl_info.item_ID, tbl_info.item_ID, tbl_info.item_Name, tbl_info.item_Descr " & _
          "From tbl_info WHERE (((tbl_info.item_ID)=" & Val(Form1.ListView.SelectedItem.Text) & "));"

    Call dbConnect
      RS.Open SQL, Conn, adOpenDynamic
        If Not RS.EOF Then
          .Text = RS!item_ID
          .SubItems(1) = RS!item_Name
          .SubItems(2) = RS!item_Descr
        End If
      RS.Close
    Conn.Close
    Set Conn = Nothing
  End With
  Unload frmLoading
  MsgBox "The selected record was successfully updated!", vbInformation, "Update"
End Sub

