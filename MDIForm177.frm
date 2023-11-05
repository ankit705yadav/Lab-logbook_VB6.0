VERSION 5.00
Begin VB.MDIForm MDIForm11 
   BackColor       =   &H008080FF&
   Caption         =   "MDIForm1"
   ClientHeight    =   9555
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   12840
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm177.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu home 
      Caption         =   "Home"
   End
   Begin VB.Menu ser 
      Caption         =   "User"
      Begin VB.Menu userLogin 
         Caption         =   "User Login"
      End
      Begin VB.Menu newUser 
         Caption         =   "New User"
      End
   End
   Begin VB.Menu admin_login 
      Caption         =   "Admin"
      Begin VB.Menu adminLogin 
         Caption         =   "Admin Login"
      End
      Begin VB.Menu adminRegister 
         Caption         =   "Admin Register"
      End
      Begin VB.Menu deregister 
         Caption         =   "Deregister"
      End
   End
   Begin VB.Menu feedback 
      Caption         =   "Feedback"
      Begin VB.Menu giveFeedback 
         Caption         =   "Give Feedback"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu contact 
         Caption         =   "contact"
      End
      Begin VB.Menu report 
         Caption         =   "Report"
      End
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
      Begin VB.Menu exitToDesktop 
         Caption         =   "Exit To Desktop"
      End
   End
End
Attribute VB_Name = "MDIForm11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub adminLogin_Click()
aLogin.Show
End Sub

Private Sub adminRegister_Click()
adminReg.Show
End Sub

Private Sub contact_Click()
MsgBox "Gmail: digitalogbook@gmail.com"
End Sub

Private Sub exitToDesktop_Click()
End
End Sub

Private Sub giveFeedback_Click()
feedbackk.Show
End Sub

Private Sub home_Click()
Welcome.Show
Welcome.SetFocus

End Sub

Private Sub MDIForm_Load()
splash.Show
End Sub

Private Sub newUser_Click()
FrmRegister.Show
End Sub

Private Sub register_Click()
FrmRegister.Show
End Sub

Private Sub report_Click()
Reportt.Show
End Sub

Private Sub userLogin_Click()
frmlogin.Show
End Sub
