VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form rform 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Reports"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11235
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   11235
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton bc 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bank Certificate"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   3840
      TabIndex        =   14
      Top             =   7320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   6375
      Left            =   7740
      TabIndex        =   7
      Top             =   840
      Width           =   3135
      Begin VB.CommandButton repcancel 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   720
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4320
         Width           =   1695
      End
      Begin VB.CommandButton repprint 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   720
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2430
         Width           =   1695
      End
      Begin VB.CommandButton repView 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&View"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   720
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   540
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Select Report"
      Height          =   5655
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   7095
      Begin VB.OptionButton bf 
         BackColor       =   &H00C0C0C0&
         Caption         =   "White And Solid Bifercation"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   5040
         Width           =   3975
      End
      Begin VB.OptionButton uc 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Undertaking Company"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   1880
         Width           =   2655
      End
      Begin VB.OptionButton pfa 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Performa for Association"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   3255
      End
      Begin VB.OptionButton a1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Annexture 1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   1120
         Width           =   1695
      End
      Begin VB.OptionButton usp 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Undertaking Stamp Paper"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   2640
         Width           =   3135
      End
      Begin VB.OptionButton epc 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Export Proceed Certificate On Bank Letterhead"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   3400
         Width           =   5775
      End
      Begin VB.OptionButton ucl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Undertaking Company Letterhead"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   4160
         Width           =   3975
      End
   End
   Begin MSDataListLib.DataCombo invno 
      Bindings        =   "Reportform.frx":0000
      Height          =   315
      Left            =   2040
      TabIndex        =   11
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "Invoice_No"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc RND2 
      Height          =   375
      Left            =   1440
      Top             =   7200
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\RND\RND.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\RND\RND.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from rnd ORDER BY INVOICE_NO"
      Caption         =   "RND"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "INVOICE NO:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   300
      Index           =   0
      Left            =   480
      TabIndex        =   12
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "rform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
invno.BoundText = RND.invno.BoundText
End Sub

Private Sub repcancel_Click()
Unload Me
End Sub

Private Sub repprint_Click()
repView_Click
End Sub

Private Sub repView_Click()
If invno.Text = "" Then
MsgBox "Please specify Invoice No!", vbQuestion, "Information"
Exit Sub
End If

Dim Report As Object
If pfa Then
Set Report = New pfa
Report.RecordSelectionFormula = "{rnd.Invoice_no} = '" & CStr(invno.Text) & "'"
crystalform.Caption = "Performa for Association"
crystalform.ShowReport Report
ElseIf a1 Then
Set Report = New a1
Report.RecordSelectionFormula = "{rnd.Invoice_no} = '" & CStr(invno.Text) & "'"
crystalform.Caption = "Annexture 1"
crystalform.ShowReport Report
ElseIf uc Then
Set Report = New uc
Report.RecordSelectionFormula = "{rnd.Invoice_no} = '" & CStr(invno.Text) & "'"
crystalform.Caption = "Undertaking Company"
crystalform.ShowReport Report

ElseIf usp Then
Set Report = New usp
Report.RecordSelectionFormula = "{rnd.Invoice_no} = '" & CStr(invno.Text) & "'"
crystalform.Caption = "Undertaking Company"
crystalform.ShowReport Report

ElseIf epc Then
Set Report = New epc
Report.RecordSelectionFormula = "{rnd.Invoice_no} = '" & CStr(invno.Text) & "'"
crystalform.Caption = "Export Proceed Certificate On Bank Letterhead"
crystalform.ShowReport Report
ElseIf ucl Then
Set Report = New ucl
Report.RecordSelectionFormula = "{rnd.Invoice_no} = '" & CStr(invno.Text) & "'"
crystalform.Caption = "Undertaking Company Letterhead"
crystalform.ShowReport Report
ElseIf bc Then
Set Report = New bc
Report.RecordSelectionFormula = "{rnd.Invoice_no} = '" & CStr(invno.Text) & "'"
crystalform.Caption = "Bank Certificate"
crystalform.ShowReport Report
ElseIf bf Then
Set Report = New bifercation
Report.RecordSelectionFormula = "{rnd.Invoice_no} = '" & CStr(invno.Text) & "'"
crystalform.Caption = "White And Solid Bifercation"
crystalform.ShowReport Report
End If

End Sub
