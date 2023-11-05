VERSION 5.00
Begin VB.Form FrmLogin66 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login Form"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   12135
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   30
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmLogin77.frx":0000
   ScaleHeight     =   7485
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Show"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox TxtUserId 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   5520
      TabIndex        =   0
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox TxtPassword 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      IMEMode         =   3  'DISABLE
      Left            =   5520
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3480
      Width           =   3375
   End
   Begin VB.CommandButton CmdSignin 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   6480
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      UseMaskColor    =   -1  'True
      Width           =   2175
   End
   Begin VB.CommandButton CmdClear 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      UseMaskColor    =   -1  'True
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "User Login"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4320
      TabIndex        =   7
      Top             =   480
      Width           =   3135
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
      Height          =   615
      Left            =   2280
      TabIndex        =   6
      Top             =   3480
      Width           =   1935
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
      Height          =   855
      Left            =   2280
      TabIndex        =   5
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   16200
      Left            =   0
      Picture         =   "FrmLogin77.frx":17A4C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   28800
   End
End
Attribute VB_Name = "FrmLogin66"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset


Private Sub CmdSignup_Click()
FrmRegister.Show
End Sub

Private Sub Command1_Click()
Form1.Show
End Sub

Private Sub Command2_Click()
deleteuser.Show
End Sub

Private Sub Command3_Click()
FrmRegister.Show
End Sub

'Establishing the DB connection
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


Private Sub CmdClear_Click()
TxtUserId.Text = ""
TxtPassword.Text = ""

End Sub

Private Sub Check1_Click()
  If Check1.Value = 1 Then
    TxtPassword.PasswordChar = ""
    ElseIf Check1.Value = 0 Then
    TxtPassword.PasswordChar = "*"
  End If

End Sub

Private Sub CmdSignin_Click()


rs.Open "select password from stud_db.dbo.usersss where username = '" & TxtUserId.Text & "' ", con, adOpenKeyset


   If rs.RecordCount = 0 Then
   MsgBox "user not found"
   Else
   'MsgBox (rs.RecordCount)
    If rs(0) = TxtPassword.Text Then
    MsgBox "Password Verified"
    'con.Close
    Subject.Show
    'FrmLogged.Show
    Unload Me
    Else
    MsgBox "Wrong Password"
    End If
   End If
   
   
   



End Sub

Private Sub Form_Resize()
Image1.Width = Me.ScaleWidth
Image1.Height = Me.ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub





