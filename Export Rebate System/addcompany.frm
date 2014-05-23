VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form addcompany 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Add New Company"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   ForeColor       =   &H80000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1455
      Left            =   480
      TabIndex        =   2
      Top             =   8400
      Width           =   14400
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10440
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1530
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000009&
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   12480
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1635
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "addcompany.frx":0000
      Height          =   7335
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   12938
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16777215
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   20
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Company_ID"
         Caption         =   "Company_ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Name"
         Caption         =   "Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Address1"
         Caption         =   "Address1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Address2"
         Caption         =   "Address2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "NTN_NO"
         Caption         =   "NTN_NO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1635.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3119.811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3119.811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3119.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   3119.811
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   345
      Left            =   5100
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
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
      RecordSource    =   "SELECT * FROM company"
      Caption         =   "Adodc1"
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
      Caption         =   "Adding to Companies list"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   330
      Left            =   480
      TabIndex        =   1
      Top             =   345
      Width           =   4200
   End
End
Attribute VB_Name = "addcompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comex As Boolean
Private Sub Command1_Click()
If Not Adodc1.Recordset.RecordCount = 0 Then
Adodc1.Recordset.Update
Adodc1.Recordset.UpdateBatch adAffectAllChapters
End If
'Unload Me
'Conn.populate_slist
End Sub

Private Sub Command2_Click()
If Not Adodc1.Recordset.RecordCount = 0 Then
Adodc1.Refresh
'Adodc1.Recordset.CancelUpdate
Adodc1.Recordset.CancelBatch adAffectAllChapters
End If
RND.Company.Refresh
Unload Me
End Sub

Private Sub DataGrid1_OnAddNew()
Dim i As Long, orow As Long
If comex Then
Exit Sub
End If
comex = True
orows = DataGrid1.Row
Adodc1.Recordset.MoveLast
i = Adodc1.Recordset!company_id + 1
DataGrid1.Row = orows
DataGrid1.Columns(0).Text = i
DataGrid1.Refresh
comex = False

End Sub

Private Sub Form_Load()
Adodc1.Recordset.Close
Adodc1.LockType = adLockBatchOptimistic
Adodc1.RecordSource = "select * from company"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount = 0 Then
Adodc1.Recordset.AddNew "Company_id", 1
                       
End If
DataGrid1.Columns(0).Locked = True
DataGrid1.Columns(0).Caption = "Company ID"
DataGrid1.Columns(0).Width = 1400.189
DataGrid1.Columns(1).Caption = "Company Name"
DataGrid1.Columns(1).Width = 3000.236
DataGrid1.Columns(2).Caption = "Address 1"
DataGrid1.Columns(2).Width = 4000.835
DataGrid1.Columns(3).Caption = "Address 2"
DataGrid1.Columns(3).Width = 4000.835
DataGrid1.Columns(4).Caption = "NTN No"
DataGrid1.Columns(4).Width = 3500.835

End Sub

