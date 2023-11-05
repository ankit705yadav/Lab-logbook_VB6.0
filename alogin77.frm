VERSION 5.00
Begin VB.Form aLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Admin Login"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   5760
      TabIndex        =   0
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   5760
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3000
      Width           =   2895
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00808000&
      Caption         =   "Show"
      Height          =   375
      Left            =   9360
      TabIndex        =   4
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "Sign In"
      Height          =   615
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808000&
      Caption         =   "Sign Up"
      Height          =   615
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2040
      TabIndex        =   7
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Admin Page"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3840
      TabIndex        =   6
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   6480
      Left            =   0
      Picture         =   "alogin77.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11760
   End
End
Attribute VB_Name = "aLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Check1_Click()
  If Check1.Value = 1 Then
    Text2.PasswordChar = ""
    ElseIf Check1.Value = 0 Then
    Text2.PasswordChar = "*"
  End If
End Sub

Private Sub Command1_Click()

rs.Open "select username from stud_db.dbo.admin where username = '" & Text1.Text & "' ", con, adOpenKeyset


   If rs.RecordCount = 0 Then
   MsgBox "Admin Not Exist"
   Else
    If rs(0) = Text2.Text Then
    MsgBox "Admin Password Verified"
    con.Close
    Admin.Show
    Unload Me
    Else
    MsgBox "Wrong Password"
    End If
   End If

End Sub

Private Sub Command2_Click()
Dim pin As Integer
pin = InputBox("Please Enter Admin Pin To Proceed", Testting, 0)
If pin = 705 Then
  If Text1.Text = "" And Text2.Text = "" Then
     MsgBox "Fields Can't Be Empty"
 Else
    Dim sql As String

    sql = "insert into stud_db.dbo.admin_t (username, password) values ("
    sql = sql & "'" & Text1.Text & "',"
    sql = sql & "'" & Text2.Text & "')"
    con.Execute sql
    
    MsgBox "New Admin Registered"
    Text1.Text = ""
    Text2.Text = ""
 End If

Else
MsgBox "Wrong Admin Pin"
End If

End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
Image1.Move 0, 0, Me.Width, Me.Height

con.Open _
"Provider=sqloledb;" & _
"Data Source=LAPTOP-37DEMHK6\SQLEXPRESS;" & _
"Intial Catalog=stud_db;" & _
"Trusted_Connection=yes;"                                    'CONNECTION IS DONE/OPEN
End Sub



Private Sub Form_Resize()
Image1.Width = Me.ScaleWidth
Image1.Height = Me.ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)
'con.Close
End Sub
