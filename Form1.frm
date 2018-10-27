VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory System"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox cmdAdd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   345
      ScaleWidth      =   1185
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1320
      Top             =   5160
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form1.frx":35FF
      OLEDBString     =   $"Form1.frx":36AF
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7435
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ITEM ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ITEM NAME"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "DESCRIPTION "
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.PictureBox cmdEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1440
      Picture         =   "Form1.frx":375F
      ScaleHeight     =   345
      ScaleWidth      =   1185
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.PictureBox cmdDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2760
      Picture         =   "Form1.frx":6DEF
      ScaleHeight     =   345
      ScaleWidth      =   1185
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.PictureBox cmdSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5400
      Picture         =   "Form1.frx":A53C
      ScaleHeight     =   345
      ScaleWidth      =   1185
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
   Begin VB.PictureBox cmdRefresh 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4080
      Picture         =   "Form1.frx":DC4E
      ScaleHeight     =   345
      ScaleWidth      =   1185
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Coded by: Welch Regime Marcellana
'Re-Edit by: Rhalp 10

Private Sub cmdAdd_Click()
  frmNew.Show 1
End Sub

Private Sub cmdDelete_Click()
  If Me.ListView.ListItems.Count = 0 Then
    MsgBox "There are no records to modify or delete!", vbExclamation, "Error"
    Exit Sub
  End If
  
  If MsgBox("Are you sure you want to delete the selected record?", vbYesNo, "Delete") = vbYes Then
    Call dbConnect
      Conn.Execute "Delete * from tbl_info where item_ID=" & Val(Me.ListView.SelectedItem.Text) & ""
    Conn.Close
    Set Conn = Nothing
    Me.ListView.ListItems.Remove (Me.ListView.SelectedItem.Index)
    MsgBox "The selected record was deleted", vbExclamation, "Delete"
  Else
    Cancel = True
  End If
End Sub

Private Sub cmdEdit_Click()
  If Me.ListView.ListItems.Count = 0 Then
    MsgBox "There are no records to modify or delete!", vbExclamation, "Error"
    Exit Sub
  End If
  frmEdit.Show 1
End Sub

Private Sub cmdRefresh_Click()
  Call loadRecords
End Sub

Private Sub cmdSearch_Click()
  frmSearch.Show 1
End Sub

Private Sub Form_Load()
  Call loadRecords
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim Form As Form
  
  For Each Form In Forms
    Unload Form
    DoEvents
  Next
End Sub
