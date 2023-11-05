VERSION 5.00
Begin VB.Form FrmLogged 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "logged77.frx":0000
   ScaleHeight     =   7305
   ScaleWidth      =   13680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "NEW HORIZON COLLEGE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   13695
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   9720
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
         Caption         =   "Delete user"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5400
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "Logout"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   11
         Top             =   2640
         Width           =   135
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8880
         TabIndex        =   10
         Top             =   2640
         Width           =   135
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Uptime"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3960
         TabIndex        =   9
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   8
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Session"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6120
         TabIndex        =   7
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9000
         TabIndex        =   6
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8520
         TabIndex        =   5
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8040
         TabIndex        =   4
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8040
         TabIndex        =   3
         Top             =   4200
         Width           =   3375
      End
   End
   Begin VB.Timer Timer2 
      Left            =   15840
      Top             =   7440
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   15840
      Top             =   6840
   End
End
Attribute VB_Name = "FrmLogged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aConnection As New ADODB.Connection
Dim aRecSet As New ADODB.Recordset

Dim sec, min, hrs As Integer


Private Sub Command1_Click()

Dim uptime As String
uptime = Label2.Caption & Label9.Caption & Label3.Caption & Label8.Caption & Label4.Caption


Dim sql As String

sql = "UPDATE stud_db.dbo.usersss SET upTime='" & uptime & "' WHERE username='" & Text1.Text & "'"

'sql = "insert into [stud_db].[dbo].[time] (time) values ("
'sql = sql & "'" & Label2.Caption & "'" & "'" & Label3.Caption & "'" & "'" & Label4.Caption & "')"
aConnection.Execute sql
Timer1.Interval = 0

Dim uptime1 As String
Dim uptime2 As String
Dim uptime3 As String
uptime1 = Label2.Caption
uptime2 = Label3.Caption
uptime3 = Label4.Caption





MsgBox ("Your runtime was: " & uptime)

frmlogin.Show
'Timer1_Timer
End Sub

Private Sub Command2_Click()
MsgBox "Please Continue With Admin Login"
aLogin.Show
End Sub

Private Sub Form_Load()

Dim time As String
time = time
'MsgBox (time)
aConnection.Open _
"Provider=sqloledb; " & _
"Data Source=LAPTOP-37DEMHK6\SQLEXPRESS; " & _
"Initial Catalog=sample; " & _
"Trusted_Connection=yes; "





Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

Timer1.Interval = 1000
Timer2.Interval = 100

Text1.Text = Subject.Text1.Text

'Dim loginTime As String
'loginTime = Timer1.time

'MsgBox (loginTime)

End Sub






Private Sub Timer1_Timer()
sec = sec + 1
Label4.Caption = Format(sec, "00")
If sec = 60 Then
min = min + 1
sec = 0
Label3.Caption = Format(min, "00")
End If
If min = 60 Then
hrs = hrs + 1
min = 0
Label2.Caption = Format(hrs, "00")
End If
End Sub

Private Sub Timer2_Timer()
Label1 = time

End Sub

Private Sub Form_Unload(Cancel As Integer)
aConnection.Close
 
End Sub

