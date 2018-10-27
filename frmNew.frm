VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Entry"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
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
   ScaleHeight     =   1560
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox cmdAdd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3360
      Picture         =   "frmNew.frx":0000
      ScaleHeight     =   345
      ScaleWidth      =   1185
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtAdd 
      Height          =   330
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox txtName 
      Height          =   330
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Description:"
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   660
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Item Name:"
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   300
      Width           =   960
   End
End
Attribute VB_Name = "frmNew"
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
  
  Call dbConnect
    Conn.Execute "Insert into tbl_info(item_Name,item_Descr) values('" & Me.txtName.Text & "','" & Me.txtAdd.Text & "')"
  Conn.Close
  Set Conn = Nothing
  Unload Me
  Call loadNew
  
  Exit Sub
errtrap:
  Select Case Err.Number
    Case -2147467259
      MsgBox "The name already exists in the database", vbCritical, "Error"
  
    Case Else
      MsgBox Err.Description, vbCritical, "The system encountered an error"
  End Select
End Sub

Public Sub loadNew()
  frmLoading.Show
  frmLoading.lblSub.Caption = "Saving your entry...."
  With Form1.ListView.ListItems
    Call dbConnect
    SQL = "SELECT tbl_info.* FROM tbl_info order by item_ID asc;"
    RS.Open SQL, Conn, adOpenDynamic
      If Not RS.EOF Then
        RS.MoveLast
        Set Item = .Add(, , RS!item_ID)
          Item.SubItems(1) = RS!item_Name
          Item.SubItems(2) = RS!item_Descr
          Item.EnsureVisible
      End If
    RS.Close
    Conn.Close
    Set Conn = Nothing
  End With
  Unload frmLoading
  MsgBox "New entry was added successfully", vbInformation, "Save"
End Sub

