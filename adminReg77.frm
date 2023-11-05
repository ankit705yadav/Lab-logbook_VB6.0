VERSION 5.00
Begin VB.Form adminReg 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Admin Register"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5520
      Width           =   2655
   End
   Begin VB.CommandButton reg_btn 
      BackColor       =   &H00FFFF00&
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   2295
   End
   Begin VB.TextBox TbPass 
      Height          =   495
      Left            =   7440
      TabIndex        =   1
      Top             =   3960
      Width           =   2535
   End
   Begin VB.TextBox TbUsername 
      Height          =   495
      Left            =   7440
      TabIndex        =   0
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ADMIN  REGISTER "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   6
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   5
      Top             =   3840
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "New Username"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   4
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   7320
      Left            =   0
      Picture         =   "adminReg77.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11760
   End
End
Attribute VB_Name = "adminReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rec As New ADODB.Recordset
Dim un As New ADODB.Recordset
Private Sub Form_Load()
Image1.Width = Me.ScaleWidth
Image1.Height = Me.ScaleWidth

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

con.Open _
"Provider=sqloledb;" & _
"Data Source=LAPTOP-37DEMHK6\SQLEXPRESS;" & _
"Intial Catalog=stud_db;" & _
"Trusted_Connection=yes;"

End Sub

Private Sub Form_Resize()
Image1.Move 0, 0, Me.Width, Me.Height
End Sub

Private Sub reg_btn_Click()

If TbUsername.Text = "" And TbPass.Text = "" Then
     MsgBox "Fields Can't Be Empty"
 Else
    
    Dim sql As String

    sql = "insert into stud_db.dbo.admin (username, password) values ("
    sql = sql & "'" & TbUsername.Text & "',"
    sql = sql & "'" & TbPass.Text & "')"
    con.Execute sql
    
    MsgBox "New admin Registered"
    TbUsername.Text = ""
    TbPass.Text = ""
 End If

End Sub
