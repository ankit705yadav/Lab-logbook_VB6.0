VERSION 5.00
Begin VB.Form Welcome 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9960
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome To Digital Logbook"
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
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   4680
      Left            =   0
      Picture         =   "Welcome77.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9960
   End
End
Attribute VB_Name = "Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmlogin.Show
Welcome.Hide
FrmRegister.Hide
End Sub

Private Sub Command2_Click()
FrmRegister.Show
Welcome.Hide
frmlogin.Hide
End Sub

Private Sub Form_Load()

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2


'Image1.Move 0, 0, Me.Width, Me.Height
'Picture1.SizeMode = StretchImage

End Sub

Private Sub Form_Resize()
Image1.Width = Me.ScaleWidth
Image1.Height = Me.ScaleWidth
End Sub

