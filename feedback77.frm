VERSION 5.00
Begin VB.Form feedbackk 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Feedback"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   12450
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Text            =   "Your Username"
      Top             =   1800
      Width           =   2895
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "feedback77.frx":0000
      Left            =   9120
      List            =   "feedback77.frx":0040
      TabIndex        =   2
      Text            =   "Select System Number"
      Top             =   1800
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "feedback77.frx":008B
      Left            =   4680
      List            =   "feedback77.frx":009B
      TabIndex        =   1
      Text            =   "Select Lab"
      Top             =   1800
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit Feedback"
      Height          =   615
      Left            =   5040
      TabIndex        =   4
      Top             =   6600
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   2535
      Left            =   2640
      TabIndex        =   3
      Text            =   "Write your feedback along with your Username"
      Top             =   2880
      Width           =   7695
   End
   Begin VB.Image Image1 
      Height          =   8280
      Left            =   0
      Picture         =   "feedback77.frx":00BB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12480
   End
End
Attribute VB_Name = "feedbackk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
Dim sql As String
sql = "insert into stud_db.dbo.feedback (name,lab,system,description) values('" & Text2.Text & "','" & Combo1.Text & "','" & Combo2.Text & "','" & Text1.Text & "')"
con.Execute sql

MsgBox "Your Feedback was Submitted"

End Sub

Private Sub Form_Load()
con.Open _
"Provider=sqloledb;" & _
"Data Source=LAPTOP-37DEMHK6\SQLEXPRESS;" & _
"Intial Catalog=stud_db;" & _
"Trusted_Connection=yes;"                                    'CONNECTION IS DONE/OPEN


Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub Text1_GotFocus()
Text1.Text = ""
End Sub

Private Sub Text2_GotFocus()
Text2.Text = ""
End Sub
