VERSION 5.00
Begin VB.Form Reportt 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Report"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   13080
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "report77.frx":0000
      Left            =   3000
      List            =   "report77.frx":0010
      TabIndex        =   6
      Text            =   "Select Lab"
      Top             =   2760
      Width           =   3255
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "report77.frx":0030
      Left            =   3000
      List            =   "report77.frx":0070
      TabIndex        =   5
      Text            =   "Select System Number"
      Top             =   3240
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Text            =   "Your Username"
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit report"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   7200
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   2535
      Left            =   3360
      TabIndex        =   2
      Text            =   "Descripition along with username"
      Top             =   4080
      Width           =   7815
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "report77.frx":00BB
      Left            =   3000
      List            =   "report77.frx":00C8
      TabIndex        =   1
      Text            =   "Frequency"
      Top             =   2280
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "report77.frx":00E6
      Left            =   3000
      List            =   "report77.frx":00FF
      TabIndex        =   0
      Text            =   "Type of error"
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   8040
      Left            =   0
      Picture         =   "report77.frx":0162
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13080
   End
End
Attribute VB_Name = "Reportt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection

Private Sub Command1_Click()

Dim sql As String
sql = "insert into stud_db.dbo.reportt (name,type,freq,lab,sys,descri) values ('" & Text2.Text & "','" & Combo1.Text & "','" & Combo2.Text & "','" & Combo4.Text & "','" & Combo3.Text & "','" & Text1.Text & "') "
con.Execute sql
MsgBox "Report was submitted we'll get back to you soon"
End Sub

Private Sub Form_Load()
con.Open _
"Provider=sqloledb;" & _
"Data Source=LAPTOP-37DEMHK6\SQLEXPRESS;" & _
"Intial Catalog=stud_db;" & _
"Trusted_Connection=yes;"


Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub Text1_GotFocus()
Text1.Text = ""
End Sub

Private Sub Text2_GotFocus()
Text2.Text = ""
End Sub
