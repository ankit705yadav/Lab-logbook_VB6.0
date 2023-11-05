VERSION 5.00
Begin VB.Form FrmRegister 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   7485
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmSignup77.frx":0000
   ScaleHeight     =   7485
   ScaleWidth      =   12255
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TbUsername 
      Height          =   495
      Left            =   6480
      TabIndex        =   0
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox TbPass 
      Height          =   495
      Left            =   6480
      TabIndex        =   1
      Top             =   3960
      Width           =   2535
   End
   Begin VB.CommandButton reg_btn 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "User Register"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   4080
      TabIndex        =   6
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2520
      TabIndex        =   5
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   2400
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   9240
      Left            =   0
      Picture         =   "FrmSignup77.frx":17A4C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12240
   End
End
Attribute VB_Name = "FrmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rec As New ADODB.Recordset
Dim un As New ADODB.Recordset

Private Sub Command1_Click()
TbUsername.Text = ""
TbPass.Text = ""
End Sub

Private Sub Form_Resize()
Image1.Width = Me.ScaleWidth
Image1.Height = Me.ScaleWidth
End Sub



Private Sub reg_btn_Click()
    
 If TbUsername.Text = "" And TbPass.Text = "" Then
     MsgBox "Fields Can't Be Empty"
 Else
 
    'un.Open "select password from stud_db.dbo.users ", con, adOpenKeyset
    'MsgBox (un.RecordCount)
    
    'For Index = 1 To un.RecordCount
    'i = 0
   ' if TbUsername.Text = un.Fields(i).Value
    
    'Next Index
    
    Dim sql As String

    sql = "insert into stud_db.dbo.usersss (username, password) values ("
    sql = sql & "'" & TbUsername.Text & "',"
    sql = sql & "'" & TbPass.Text & "')"
    con.Execute sql
    
    MsgBox "Registered"
    TbUsername.Text = ""
    TbPass.Text = ""
 End If
    
End Sub


Private Sub Form_Load()
Image1.Move 0, 0, Me.Width, Me.Height


Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2


con.Open _
"Provider=sqloledb;" & _
"Data Source=LAPTOP-37DEMHK6\SQLEXPRESS;" & _
"Intial Catalog=stud_db;" & _
"Trusted_Connection=yes;"
End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub


