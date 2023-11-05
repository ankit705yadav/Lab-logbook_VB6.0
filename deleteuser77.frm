VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form deleteuser 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   15330
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C0C0&
      Caption         =   "Delete User"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15375
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3975
         Left            =   2640
         TabIndex        =   4
         Top             =   840
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   7011
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   5400
         TabIndex        =   2
         Top             =   5520
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Delete"
         Height          =   615
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5640
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter the Username to be deleted:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3000
         TabIndex        =   3
         Top             =   5640
         Width           =   1815
      End
   End
End
Attribute VB_Name = "deleteuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rec As New ADODB.Recordset

Private Sub Command1_Click()

Dim sql As String
rec.Open "select * from stud_db.dbo.users where username= '" & Text1.Text & "'", con, adOpenStatic, adLockReadOnly
sql = "delete from stud_db.dbo.users where username= '" & Text1.Text & "'"

If rec.EOF() Then
MsgBox "No Such Record Not Found"
rec.Close
con.Close
Else

con.Execute sql
MsgBox "Record Deleted"

rec.Close
con.Close
Unload Me
deleteuser.Show

End If

End Sub

Private Sub Form_Load()

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2


'Establishing the DB connection
con.Open _
"Provider=sqloledb;" & _
"Data Source=LAPTOP-37DEMHK6\SQLEXPRESS;" & _
"Intial Catalog=stud_db;" & _
"Trusted_Connection=yes;"                                    'CONNECTION IS DONE/OPEN


rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockBatchOptimistic

rs.Open ("Select * from stud_db.dbo.users"), con, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rs

End Sub
