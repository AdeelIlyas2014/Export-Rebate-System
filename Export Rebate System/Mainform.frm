VERSION 5.00
Begin VB.MDIForm Mainform 
   BackColor       =   &H8000000C&
   Caption         =   "R & D Systems - Kaysons Pvt Ltd"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   -1485
   ClientWidth     =   11880
   LinkTopic       =   "MDIForm1"
   Picture         =   "Mainform.frx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuForms 
      Caption         =   "&Forms"
      Begin VB.Menu RNDmnu 
         Caption         =   "&R N D Feeding"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnus 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddC 
         Caption         =   "Add New &Company"
      End
      Begin VB.Menu mnuaddp 
         Caption         =   "Add New &Party"
      End
      Begin VB.Menu mnuaddu 
         Caption         =   "Add New &Units"
      End
      Begin VB.Menu mnuports 
         Caption         =   "Add New &Ports"
      End
      Begin VB.Menu mnuAddy 
         Caption         =   "Add New &Currency"
      End
      Begin VB.Menu mnuss 
         Caption         =   "-"
      End
      Begin VB.Menu mnex 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
   End
End
Attribute VB_Name = "Mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnex_Click()
Unload Me
End Sub

Private Sub mnuAddC_Click()
addcompany.Show
End Sub

Private Sub mnuaddp_Click()
addparty.Show
End Sub

Private Sub mnuaddu_Click()
Addunits.Show
End Sub

Private Sub mnuAddy_Click()
Addcurrency.Show
End Sub

Private Sub mnuports_Click()
addports.Show
End Sub

Private Sub mnuReport2_Click()
MsgBox "Pressed 2"
End Sub

Private Sub mnuReports_Click()
rform.Show (1)
End Sub

Private Sub RNDmnu_Click()
RND.Show

End Sub
