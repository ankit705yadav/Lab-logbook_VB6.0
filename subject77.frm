VERSION 5.00
Begin VB.Form Subject 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Subject"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "subject77.frx":0000
   ScaleHeight     =   4890
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "subject77.frx":BA311
      Left            =   1920
      List            =   "subject77.frx":BA321
      TabIndex        =   2
      Text            =   "SELECT LAB"
      Top             =   960
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "subject77.frx":BA341
      Left            =   1920
      List            =   "subject77.frx":BA381
      TabIndex        =   1
      Text            =   "SYSTEM NUMBER"
      Top             =   1920
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "subject77.frx":BA3CC
      Left            =   1920
      List            =   "subject77.frx":BA3DF
      TabIndex        =   0
      Text            =   "SELECT SUBJECT"
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   4920
      Left            =   -120
      Picture         =   "subject77.frx":BA40E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10200
   End
End
Attribute VB_Name = "Subject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
'Dim rs As New ADODB.Recordset



Private Sub Command1_Click()
'MsgBox (Combo1.Text)
'MsgBox (Combo2.Text)
'MsgBox (Combo1.Text)

Dim sql As String

sql = "UPDATE stud_db.dbo.usersss SET lab='" & Combo3.Text & "',system_number='" & Combo2.Text & "',subjectr='" & Combo1.Text & "' WHERE username='" & Text1.Text & "'"
'sql = "insert into stud_db.dbo.users where  (lab,system_number,subject) values('" & Combo3.Text & "',' " & Combo2 & " ',' " & Combo1 & " ')"
con.Execute sql
FrmLogged.Show

'Dim sql As String
'sql = "insert into stud_db.dbo.selection (lab,system_number,subject) values('" & Combo3.Text & "',' " & Combo2 & " ',' " & Combo1 & " ')"
'con.Execute sql




End Sub

Private Sub Form_Load()

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

con.Open _
"Provider=sqloledb;" & _
"Data Source=LAPTOP-37DEMHK6\SQLEXPRESS;" & _
"Intial Catalog=stud_db;" & _
"Trusted_Connection=yes;"

Text1.Text = frmlogin.TxtUserId.Text

End Sub



