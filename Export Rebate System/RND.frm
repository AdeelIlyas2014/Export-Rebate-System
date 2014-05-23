VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form RND 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form For Research And Development Support On Export"
   ClientHeight    =   11625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17160
   ForeColor       =   &H80000011&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11625
   ScaleWidth      =   17160
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox EformNo 
      Height          =   375
      Left            =   10560
      TabIndex        =   144
      Top             =   720
      Width           =   1575
   End
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "RND.frx":0000
      Height          =   255
      Left            =   720
      TabIndex        =   126
      Top             =   11880
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393216
      ListField       =   "Rupee_Conv"
   End
   Begin VB.Frame Frame10 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFF00&
      Height          =   5820
      Left            =   120
      TabIndex        =   108
      Top             =   5760
      Width           =   1920
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add &Banks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   5160
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add Uni&ts"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   4395
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add C&urrency"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add P&orts"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   1845
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add &Party"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   2715
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add &Company"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   960
         Width           =   1695
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   1920
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ADD NEW"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   45
         Left            =   510
         TabIndex        =   114
         Top             =   360
         Width           =   855
      End
      Begin VB.Line Line6 
         X1              =   1920
         X2              =   1920
         Y1              =   120
         Y2              =   0
      End
   End
   Begin VB.Frame Frame11 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Remarks"
      ForeColor       =   &H80000008&
      Height          =   1620
      Left            =   2160
      TabIndex        =   105
      Top             =   9720
      Width           =   14295
      Begin VB.TextBox remtext 
         Height          =   1305
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   106
         Top             =   240
         Width           =   14040
      End
   End
   Begin VB.ComboBox rdscombo 
      Height          =   315
      Left            =   12720
      Style           =   2  'Dropdown List
      TabIndex        =   86
      Top             =   5160
      Width           =   3735
   End
   Begin VB.Frame Frame8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Input Values"
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   2160
      TabIndex        =   72
      Top             =   5040
      Width           =   6255
      Begin MSDataListLib.DataCombo unitsCombo 
         Bindings        =   "RND.frx":0012
         Height          =   315
         Left            =   4560
         TabIndex        =   26
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "unit"
         BoundColumn     =   "ID"
         Text            =   ""
      End
      Begin MSMask.MaskEdBox shippieces 
         Height          =   375
         Left            =   1800
         TabIndex        =   25
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox tdvs 
         Height          =   375
         Left            =   4560
         TabIndex        =   128
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox shortship 
         Height          =   375
         Left            =   1800
         TabIndex        =   129
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox netvalship 
         Height          =   375
         Left            =   4560
         TabIndex        =   130
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Net Value:"
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
         Index           =   48
         Left            =   3240
         TabIndex        =   135
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Short Shipped:"
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
         Index           =   47
         Left            =   120
         TabIndex        =   134
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total Declared Value of Shipment:"
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
         Height          =   375
         Index           =   46
         Left            =   480
         TabIndex        =   133
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Quantities:"
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
         Index           =   39
         Left            =   360
         TabIndex        =   132
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Units:"
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
         Index           =   40
         Left            =   3720
         TabIndex        =   131
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc units 
      Height          =   375
      Left            =   10440
      Top             =   11880
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      RecordSource    =   "select * from units order by unit"
      Caption         =   "Units"
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
   Begin MSAdodcLib.Adodc party 
      Height          =   375
      Left            =   11520
      Top             =   11880
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      RecordSource    =   "select  * from party order by name"
      Caption         =   "party"
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
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Shipping Detail"
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   2160
      TabIndex        =   65
      Top             =   6840
      Width           =   6255
      Begin VB.TextBox HSCODES 
         Height          =   945
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   148
         Top             =   360
         Width           =   4440
      End
      Begin VB.TextBox GDFormNo 
         Height          =   375
         Left            =   1680
         TabIndex        =   21
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox mrno 
         Height          =   375
         Left            =   1680
         TabIndex        =   23
         Top             =   1920
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker GDFormdated 
         Height          =   375
         Left            =   4440
         TabIndex        =   22
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20447233
         CurrentDate     =   39142
      End
      Begin MSComCtl2.DTPicker MRNOdated 
         Height          =   375
         Left            =   4440
         TabIndex        =   24
         Top             =   1920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20447233
         CurrentDate     =   39142
      End
      Begin MSMask.MaskEdBox EFORMNO4 
         Height          =   345
         Left            =   3600
         TabIndex        =   127
         Top             =   2640
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "G.D Form No:"
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
         Height          =   180
         Index           =   25
         Left            =   120
         TabIndex        =   150
         Top             =   1560
         Width           =   1545
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "HS Codes:"
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
         Height          =   375
         Index           =   15
         Left            =   480
         TabIndex        =   149
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dated:"
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
         Height          =   180
         Index           =   28
         Left            =   3600
         TabIndex        =   68
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "M.R No:"
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
         Height          =   180
         Index           =   27
         Left            =   720
         TabIndex        =   67
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dated:"
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
         Height          =   180
         Index           =   26
         Left            =   3600
         TabIndex        =   66
         Top             =   1560
         Width           =   720
      End
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bank Details"
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   8520
      TabIndex        =   61
      Top             =   7920
      Width           =   7935
      Begin VB.TextBox fdbcno 
         Height          =   315
         Left            =   2280
         TabIndex        =   140
         Top             =   300
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker realdate 
         Height          =   315
         Left            =   2280
         TabIndex        =   141
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20447233
         CurrentDate     =   39142
      End
      Begin MSDataListLib.DataCombo Bankscombo 
         Bindings        =   "RND.frx":0026
         Height          =   315
         Left            =   5640
         TabIndex        =   142
         Top             =   315
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Short"
         BoundColumn     =   "ID"
         Text            =   ""
      End
      Begin MSMask.MaskEdBox bcharges 
         Height          =   315
         Left            =   5640
         TabIndex        =   143
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label edate 
         BackColor       =   &H00C0C0C0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   2280
         TabIndex        =   152
         Top             =   1200
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Expiry Date:"
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
         Index           =   52
         Left            =   690
         TabIndex        =   151
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bank Charges:"
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
         Index           =   50
         Left            =   3960
         TabIndex        =   139
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bank:"
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
         Index           =   42
         Left            =   4920
         TabIndex        =   107
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Realization Date:"
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
         Index           =   20
         Left            =   120
         TabIndex        =   63
         Top             =   780
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "FDBC Number:"
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
         Index           =   5
         Left            =   735
         TabIndex        =   62
         Top             =   315
         Width           =   1440
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Input Values"
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   8520
      TabIndex        =   56
      Top             =   3480
      Width           =   7935
      Begin MSMask.MaskEdBox fcyvalue 
         Height          =   345
         Left            =   1560
         TabIndex        =   27
         Top             =   285
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo curcombo 
         Bindings        =   "RND.frx":003F
         Height          =   315
         Left            =   2760
         TabIndex        =   28
         Top             =   315
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Currency"
         BoundColumn     =   "Currency_ID"
         Text            =   ""
      End
      Begin MSMask.MaskEdBox exrate 
         Bindings        =   "RND.frx":0054
         DataField       =   "Rupee_Conv"
         DataSource      =   "curren"
         Height          =   345
         Left            =   4080
         TabIndex        =   29
         Top             =   285
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox insurance 
         Height          =   345
         Left            =   6480
         TabIndex        =   32
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox valpkr 
         Height          =   345
         Left            =   6480
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   285
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox freight 
         Height          =   345
         Left            =   1560
         TabIndex        =   31
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox commission 
         Height          =   345
         Left            =   1560
         TabIndex        =   74
         Top             =   1170
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox nongarment 
         Height          =   345
         Left            =   6480
         TabIndex        =   123
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Value in PKR:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   225
         Index           =   1
         Left            =   4800
         TabIndex        =   138
         Top             =   405
         Width           =   1560
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Insurance:"
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
         Height          =   180
         Index           =   18
         Left            =   5160
         TabIndex        =   137
         Top             =   885
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Non-Garment Items:"
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
         Index           =   49
         Left            =   4200
         TabIndex        =   136
         Top             =   1170
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   " @"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   17
         Left            =   3720
         TabIndex        =   60
         Top             =   315
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Commission:"
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
         Height          =   180
         Index           =   19
         Left            =   120
         TabIndex        =   59
         Top             =   1245
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "FCY Value:"
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
         Height          =   210
         Index           =   16
         Left            =   240
         TabIndex        =   58
         Top             =   405
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Freight:"
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
         Height          =   180
         Index           =   7
         Left            =   480
         TabIndex        =   57
         Top             =   885
         Width           =   990
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Party Details"
      ForeColor       =   &H80000008&
      Height          =   1860
      Left            =   2160
      TabIndex        =   50
      Top             =   3120
      Width           =   6255
      Begin VB.TextBox PCountry 
         DataField       =   "Country"
         DataSource      =   "party"
         Height          =   345
         Left            =   1800
         TabIndex        =   7
         Top             =   1440
         Width           =   4215
      End
      Begin VB.TextBox paddr2 
         DataField       =   "Address2"
         DataSource      =   "party"
         Height          =   345
         Left            =   1800
         TabIndex        =   6
         Top             =   1040
         Width           =   4215
      End
      Begin VB.TextBox PAddr1 
         DataField       =   "Address1"
         DataSource      =   "party"
         Height          =   345
         Left            =   1800
         TabIndex        =   5
         Top             =   640
         Width           =   4215
      End
      Begin MSDataListLib.DataCombo PartyCombo 
         Bindings        =   "RND.frx":0079
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   270
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Name"
         BoundColumn     =   "Party_ID"
         Text            =   ""
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Address:"
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
         Height          =   270
         Index           =   31
         Left            =   720
         TabIndex        =   53
         Top             =   690
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Country:"
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
         Height          =   285
         Index           =   30
         Left            =   720
         TabIndex        =   52
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Party Name:"
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
         Height          =   285
         Index           =   29
         Left            =   360
         TabIndex        =   51
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Company Details"
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   2160
      TabIndex        =   46
      Top             =   1080
      Width           =   6255
      Begin VB.TextBox NTN 
         DataField       =   "NTN_NO"
         DataSource      =   "Company"
         Height          =   345
         Left            =   1800
         TabIndex        =   3
         Top             =   1440
         Width           =   4215
      End
      Begin VB.TextBox cAddr2 
         DataField       =   "Address2"
         DataSource      =   "Company"
         Height          =   345
         Left            =   1800
         TabIndex        =   2
         Top             =   1040
         Width           =   4215
      End
      Begin VB.TextBox cAddr1 
         DataField       =   "Address1"
         DataSource      =   "Company"
         Height          =   345
         Left            =   1800
         TabIndex        =   1
         Top             =   640
         Width           =   4215
      End
      Begin MSDataListLib.DataCombo Companycombo 
         Bindings        =   "RND.frx":008D
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Top             =   270
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Name"
         BoundColumn     =   "Company_ID"
         Text            =   ""
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Company Name:"
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
         Index           =   3
         Left            =   135
         TabIndex        =   49
         Top             =   360
         Width           =   1590
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NTN No:"
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
         Height          =   285
         Index           =   6
         Left            =   840
         TabIndex        =   48
         Top             =   1560
         Width           =   885
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Address:"
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
         Height          =   270
         Index           =   4
         Left            =   720
         TabIndex        =   47
         Top             =   720
         Width           =   915
      End
   End
   Begin MSAdodcLib.Adodc curren 
      Height          =   375
      Left            =   13680
      Top             =   11880
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   2
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
      RecordSource    =   "select * from c_currency order by currency"
      Caption         =   "curren"
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
   Begin MSAdodcLib.Adodc aport 
      Height          =   330
      Left            =   14640
      Top             =   11880
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
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
      RecordSource    =   "select * from ports where port_type = 1"
      Caption         =   "aport"
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFF00&
      Height          =   4695
      Left            =   120
      TabIndex        =   45
      Top             =   720
      Width           =   1920
      Begin VB.CommandButton exitbut 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   3960
         Width           =   1695
      End
      Begin VB.CommandButton repbut 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reports"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   3360
         Width           =   1695
      End
      Begin VB.CommandButton Delbut 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CommandButton Newbut 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton findbut 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton savebut 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2160
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   1920
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OPERATIONS"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   41
         Left            =   435
         TabIndex        =   81
         Top             =   360
         Width           =   1215
      End
      Begin VB.Line Line2 
         X1              =   1920
         X2              =   1920
         Y1              =   120
         Y2              =   0
      End
   End
   Begin VB.ComboBox Shipmentby 
      Height          =   315
      Left            =   12000
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1320
      Width           =   4455
   End
   Begin MSAdodcLib.Adodc RNDa 
      Height          =   375
      Left            =   4920
      Top             =   11880
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
      RecordSource    =   "select * from rnd order by invoice_no"
      Caption         =   "RNDA"
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
   Begin MSAdodcLib.Adodc Company 
      Height          =   330
      Left            =   2040
      Top             =   11880
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
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
      RecordSource    =   "select * from company order by name"
      Caption         =   "Company"
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
   Begin MSAdodcLib.Adodc Sport 
      Height          =   330
      Left            =   8160
      Top             =   11880
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
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
      RecordSource    =   "select * from ports where port_type = 0 order by port"
      Caption         =   "sport"
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
   Begin MSDataListLib.DataCombo invno 
      Bindings        =   "RND.frx":00A3
      Height          =   315
      Left            =   3960
      TabIndex        =   82
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Invoice_No"
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker invdate 
      Height          =   375
      Left            =   6840
      TabIndex        =   83
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   20447233
      CurrentDate     =   39142
   End
   Begin MSAdodcLib.Adodc Bank_addoc 
      Height          =   375
      Left            =   6840
      Top             =   11880
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      RecordSource    =   "select * from Banks order by Short"
      Caption         =   "Banks"
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
   Begin MSAdodcLib.Adodc USD 
      Height          =   375
      Left            =   0
      Top             =   11760
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   2
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
      RecordSource    =   "select * from c_currency where currency = ""USD"""
      Caption         =   "usd"
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
   Begin MSComCtl2.DTPicker Eformdated 
      Height          =   375
      Left            =   13920
      TabIndex        =   145
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   20447233
      CurrentDate     =   39142
   End
   Begin VB.Frame Sea 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Shipment By Sea"
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   8520
      TabIndex        =   37
      Top             =   1680
      Visible         =   0   'False
      Width           =   7935
      Begin VB.TextBox scountry 
         DataField       =   "Country"
         DataSource      =   "Sport"
         Height          =   345
         Left            =   5280
         TabIndex        =   19
         Top             =   1245
         Width           =   1575
      End
      Begin VB.TextBox cont_no 
         Height          =   345
         Left            =   1800
         TabIndex        =   17
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox BLNO 
         Height          =   345
         Left            =   1800
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker blnodated 
         Height          =   375
         Left            =   5280
         TabIndex        =   16
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20447233
         CurrentDate     =   39142
      End
      Begin MSDataListLib.DataCombo Seacombo 
         Bindings        =   "RND.frx":00B6
         Height          =   315
         Left            =   1800
         TabIndex        =   18
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Port"
         BoundColumn     =   "Destination_ID"
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker sarrival_date 
         Height          =   375
         Left            =   5280
         TabIndex        =   20
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20447233
         CurrentDate     =   39142
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Arrival Date:"
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
         Index           =   37
         Left            =   3600
         TabIndex        =   70
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Country:"
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
         Index           =   36
         Left            =   4200
         TabIndex        =   69
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Seaport:"
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
         Index           =   33
         Left            =   720
         TabIndex        =   55
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Container No:"
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
         Index           =   11
         Left            =   120
         TabIndex        =   40
         Top             =   810
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dated:"
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
         Index           =   10
         Left            =   4440
         TabIndex        =   39
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "B/L No:"
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
         Index           =   9
         Left            =   720
         TabIndex        =   38
         Top             =   285
         Width           =   855
      End
   End
   Begin VB.Frame hosiery 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Output Values For Hoseiry"
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   8520
      TabIndex        =   97
      Top             =   5520
      Visible         =   0   'False
      Width           =   7935
      Begin MSMask.MaskEdBox rds_charges_h 
         Height          =   345
         Left            =   2740
         TabIndex        =   98
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox netvalpkr_h 
         Height          =   345
         Left            =   2740
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox rds_amt_h 
         Height          =   345
         Left            =   2740
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   795
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TOT_WEIGHT 
         Height          =   345
         Left            =   2740
         TabIndex        =   124
         Top             =   1620
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total Net Weight:"
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
         Index           =   51
         Left            =   600
         TabIndex        =   125
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label rds_label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "RDS 6% Amt in PKR:"
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
         Index           =   1
         Left            =   480
         TabIndex        =   103
         Top             =   800
         Width           =   2160
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "RDS S.Charges:"
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
         Index           =   21
         Left            =   960
         TabIndex        =   102
         Top             =   1240
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Net Value in PKR:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   300
         Index           =   44
         Left            =   600
         TabIndex        =   101
         Top             =   360
         Width           =   2115
      End
   End
   Begin VB.Frame homet 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Output Values For Home Textile"
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   8520
      TabIndex        =   75
      Top             =   5520
      Visible         =   0   'False
      Width           =   7935
      Begin MSMask.MaskEdBox rdscharges 
         Height          =   345
         Left            =   2745
         TabIndex        =   76
         Top             =   1740
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox netpkrval 
         Height          =   345
         Left            =   2745
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox rdsamt 
         Height          =   345
         Left            =   2745
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   1365
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox rds_solid 
         Height          =   345
         Left            =   2745
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   990
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox rds_white 
         Height          =   345
         Left            =   2745
         TabIndex        =   95
         Top             =   615
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox white_weight 
         Height          =   345
         Left            =   5280
         TabIndex        =   118
         Top             =   660
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox solid_weight 
         Height          =   345
         Left            =   5280
         TabIndex        =   119
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox total_weight 
         Height          =   345
         Left            =   5280
         TabIndex        =   121
         Top             =   1545
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label rds_label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "KGS"
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
         Height          =   225
         Index           =   3
         Left            =   5640
         TabIndex        =   122
         Top             =   360
         Width           =   375
      End
      Begin VB.Line Line7 
         X1              =   4320
         X2              =   6540
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Label rds_label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total:"
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
         Index           =   2
         Left            =   4440
         TabIndex        =   120
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label rds_label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Weight:"
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
         Index           =   1
         Left            =   4320
         TabIndex        =   117
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label rds_label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Weight:"
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
         Left            =   4320
         TabIndex        =   116
         Top             =   720
         Width           =   975
      End
      Begin VB.Label rds_label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "RDS Amt in PKR:"
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
         Index           =   2
         Left            =   810
         TabIndex        =   96
         Top             =   1350
         Width           =   1800
      End
      Begin VB.Label rds_label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Net Value for White:"
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
         Index           =   21
         Left            =   240
         TabIndex        =   93
         Top             =   690
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "RDS S.Charges:"
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
         Index           =   22
         Left            =   885
         TabIndex        =   90
         Top             =   1725
         Width           =   1695
      End
      Begin VB.Label rds_label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Net Value for Solid:"
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
         Left            =   225
         TabIndex        =   89
         Top             =   1020
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Net Value in PKR:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   300
         Index           =   34
         Left            =   600
         TabIndex        =   88
         Top             =   300
         Width           =   2115
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   8520
      TabIndex        =   91
      Top             =   5520
      Width           =   7935
   End
   Begin VB.Frame Flight 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Shipment By Air"
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   8520
      TabIndex        =   41
      Top             =   1680
      Visible         =   0   'False
      Width           =   7935
      Begin VB.TextBox Flightno 
         Height          =   345
         Left            =   2040
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox AWBNO 
         Height          =   345
         Left            =   2040
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox acountry 
         DataField       =   "Country"
         DataSource      =   "aport"
         Height          =   345
         Left            =   5400
         TabIndex        =   13
         Top             =   1200
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo airportcombo 
         Bindings        =   "RND.frx":00CA
         DataField       =   "Port"
         DataSource      =   "port"
         Height          =   315
         Left            =   2040
         TabIndex        =   12
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Port"
         BoundColumn     =   "Destination_ID"
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker AWBdated 
         Height          =   375
         Left            =   5400
         TabIndex        =   10
         Top             =   225
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20447233
         CurrentDate     =   39142
      End
      Begin MSComCtl2.DTPicker aArrival_date 
         Height          =   375
         Left            =   5400
         TabIndex        =   14
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20447233
         CurrentDate     =   39142
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Arrival Date:"
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
         Index           =   38
         Left            =   3720
         TabIndex        =   71
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Country:"
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
         Index           =   35
         Left            =   4320
         TabIndex        =   64
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Airport:"
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
         Index           =   32
         Left            =   840
         TabIndex        =   54
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "AWB No:"
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
         Index           =   14
         Left            =   960
         TabIndex        =   44
         Top             =   285
         Width           =   885
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dated:"
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
         Index           =   13
         Left            =   4560
         TabIndex        =   43
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Flight No:"
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
         Index           =   12
         Left            =   600
         TabIndex        =   42
         Top             =   810
         Width           =   1215
      End
   End
   Begin VB.Frame Frame9 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   8520
      TabIndex        =   73
      Top             =   1680
      Width           =   7935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "E.Form No:"
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
      Height          =   375
      Index           =   23
      Left            =   8880
      TabIndex        =   147
      Top             =   720
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Dated:"
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
      Index           =   24
      Left            =   12960
      TabIndex        =   146
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "RDS For:"
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
      Height          =   180
      Index           =   43
      Left            =   11640
      TabIndex        =   87
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Dated:"
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
      Height          =   375
      Index           =   2
      Left            =   6000
      TabIndex        =   85
      Top             =   720
      Width           =   735
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
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   84
      Top             =   720
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   17160
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   17160
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Shipment By:"
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
      Index           =   8
      Left            =   10440
      TabIndex        =   36
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Form For Research And Development Support On Export"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   104
      Top             =   240
      Width           =   7335
   End
   Begin VB.Menu mnufeeding 
      Caption         =   "Research And Development"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnufind 
         Caption         =   "&Find"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnucut 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "&Cancel"
      End
   End
End
Attribute VB_Name = "RND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Myrset1 As New ADODB.Recordset
Public Mycon1 As New ADODB.Connection
Public Myrset2 As New ADODB.Recordset
Public Mycon2 As New ADODB.Connection
Public Myrset3 As New ADODB.Recordset
Public Mycon3 As New ADODB.Connection
Public newentry As Boolean
Public querystat As Boolean
Public chmd As Boolean
Dim oldinvno As String
Public lastkpress As Integer
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VK_CAPITAL = &H14
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

' API declarations:

Private Declare Function GetVersionEx Lib "kernel32" _
   Alias "GetVersionExA" _
   (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Sub keybd_event Lib "user32" _
   (ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Declare Function GetKeyboardState Lib "user32" _
   (pbKeyState As Byte) As Long

Private Declare Function SetKeyboardState Lib "user32" _
   (lppbKeyState As Byte) As Long

Private Sub aArrival_date_Change()
rout.choccur
End Sub

Private Sub acountry_Change()
rout.choccur
End Sub

Private Sub airportcombo_Change()
rout.choccur
End Sub

Private Sub airportcombo_Click(Area As Integer)
On Error GoTo err
If RTrim(airportcombo.BoundText) <> "" Then
aport.Recordset.MoveFirst
aport.Recordset.Find "Destination_ID = '" & RTrim(airportcombo.BoundText) & "'"
End If
err:
End Sub



Private Sub AWBdated_Change()
rout.choccur
End Sub

Private Sub AWBNO_Change()
rout.choccur
End Sub

Private Sub Bankscombo_Change()
rout.choccur
End Sub

Private Sub bcharges_Change()
rout.choccur
rout.RCALC
End Sub

Private Sub BLNO_Change()
rout.choccur
End Sub

Private Sub blnodated_Change()
rout.choccur
End Sub

Private Sub cAddr1_Change()
rout.choccur

End Sub

Private Sub cAddr2_Change()
rout.choccur
End Sub

Private Sub Command1_Click()
AddBanks.Show
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Command2_Click()
addcompany.Show
End Sub

Private Sub Command3_Click()
addparty.Show
End Sub

Private Sub Command4_Click()
addports.Show
End Sub

Private Sub Command5_Click()
Addcurrency.Show
End Sub

Private Sub Command6_Click()
Addunits.Show
End Sub

Private Sub commission_Change()
rout.choccur
rout.RCALC
End Sub

Private Sub commission_KeyPress(KeyAscii As Integer)

If (KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii > 47 And KeyAscii < 58)) Then
Else
KeyAscii = 0
End If
End Sub

Private Sub Companycombo_Change()
rout.choccur
End Sub

Public Sub Companycombo_Click(Area As Integer)
On Error GoTo err
If RTrim(Companycombo.BoundText) <> "" Then
Company.Recordset.MoveFirst
Company.Recordset.Find "Company_Id = '" & RTrim(Companycombo.BoundText) & "'"
End If
err:
End Sub




Private Sub cont_no_Change()
rout.choccur
End Sub

Private Sub curcombo_Change()
rout.choccur
End Sub

Private Sub curcombo_Click(Area As Integer)
'curren.Recordset.Update
On Error GoTo err
If RTrim(curcombo.BoundText) <> "" Then
curren.Recordset.MoveFirst
curren.Recordset.Find "Currency_id = '" & RTrim(curcombo.BoundText) & "'"
valpkr.Text = CStr(Val(exrate.Text) * Val(fcyvalue.Text))
End If
err:
rout.RCALC
End Sub

Private Sub Delbut_Click()

If invno.Text <> "" Then
Set Myrset1 = rout.getrcset("RND", "Select * from RND where invoice_no = '" & LTrim(RTrim(invno.Text)) & "'", Mycon1, False)
If Myrset1.RecordCount > 0 Then
If MsgBox("Are you sure you want to delete invoice no: " & LTrim(RTrim(invno.Text)), vbYesNo, "Confirmation") = 6 Then
Set Myrset1 = rout.getrcset("RND", "delete * from RND where invoice_no = '" & invno.Text & "'", Mycon1, False)
Set Myrset2 = rout.getrcset("RND", "delete * from bill where invoice_no = '" & invno.Text & "'", Mycon2, False)
rout.clearcurrent
newentry = True
rout.obuts
Mycon1.Close
Mycon2.Close
RND.Refresh
RND.RNDa.Refresh
chsaved
End If
Else
MsgBox "Invoice not found", vbCritical, "Information"
End If
End If


End Sub

Private Sub Eformdated_Change()
rout.choccur
End Sub

Private Sub EformNo_Change()
rout.choccur
End Sub

Private Sub exitbut_Click()
If chmd Then
Dim ans As Integer
ans = MsgBox("Cancel current entry, Exit?", vbYesNo, "Confirm")
If ans = 6 Then
ToggleCapsLock (False)
Unload Me
End If
Else
ToggleCapsLock (False)
Unload Me
End If



End Sub

Private Sub exrate_Change()
rout.choccur
rout.RCALC
End Sub

Private Sub exrate_KeyPress(KeyAscii As Integer)
If (KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii > 47 And KeyAscii < 58)) Then
Else
KeyAscii = 0
End If
End Sub

Private Sub fcyvalue_Change()
'rout.choccur
'rout.RCALC
End Sub

Private Sub fcyvalue_KeyPress(KeyAscii As Integer)
'MsgBox KeyAscii
If (KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii > 47 And KeyAscii < 58)) Then
Else
KeyAscii = 0
End If
End Sub

Private Sub fdbcno_Change()
rout.choccur
End Sub

Private Sub findbut_Click()
If chmd Then
If MsgBox("Cancel current entry, and find?", vbYesNo, "Confirmation") = 6 Then
querystat = True
invno.SetFocus
End If
Else
querystat = True
invno.SetFocus
End If

End Sub

Private Sub Flightno_Change()
rout.choccur
End Sub


Private Sub Form_Load()
ToggleCapsLock (True)
curren.LockType = adLockPessimistic
curren.Refresh

lastkpress = 0
Shipmentby.AddItem "By Sea"
Shipmentby.AddItem "By Air"
rdscombo.AddItem "Home Textile"  ' 3% and 5%
rdscombo.AddItem "Hosiery"       '  6%

newentry = True
rout.clearcurrent
chmd = False

querystat = False
rout.obuts
End Sub

Private Sub freight_Change()
rout.choccur
rout.RCALC
End Sub

Private Sub freight_KeyPress(KeyAscii As Integer)

If (KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii > 47 And KeyAscii < 58)) Then
Else
KeyAscii = 0
End If

End Sub

Private Sub GDFormdated_Change()
rout.choccur
End Sub

Private Sub GDFormNo_Change()
rout.choccur
End Sub

Private Sub HSCODES_Change()
rout.choccur
End Sub

Private Sub insurance_Change()
rout.choccur
rout.RCALC
End Sub

Private Sub insurance_KeyPress(KeyAscii As Integer)

If (KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii > 47 And KeyAscii < 58)) Then
Else
KeyAscii = 0
End If

End Sub


Private Sub invno_Change()
rout.choccur

End Sub


Private Sub invno_DblClick(Area As Integer)
invno_LostFocus
End Sub

Private Sub invno_GotFocus()
If querystat Then
oldinvno = invno.Text
invno.Text = ""
rout.obuts
End If
End Sub

Private Sub invno_KeyPress(KeyAscii As Integer)
If querystat Then
'MsgBox KeyAscii
If (KeyAscii = 13 Or KeyAscii = 27) Then
RND.lastkpress = KeyAscii

invno_LostFocus
End If
End If
End Sub

Private Sub invno_LostFocus()
invno.Text = UCase(invno.Text)
If querystat Then
  If invno.Text = "" Then
  Else
    If Not rout.findinv Then
    MsgBox "Not found!", vbCritical, "Information"
    invno.Text = oldinvno
      Else
    'savebut.Caption = "&Save Changes"
    newentry = False
    End If
  End If
 querystat = False
  rout.obuts
 
End If

End Sub

Private Sub mnuCancel_Click()
Unload Me
End Sub



Private Sub mrno_Change()
rout.choccur
End Sub

Private Sub MRNOdated_Change()
rout.choccur
End Sub

Private Sub netpkrval_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub netvalpkr_h_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub netvalship_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Newbut_Click()
If invno.Text <> "" Then
Dim ans As Integer
ans = MsgBox("Enter New record and cancel current entry?", vbYesNo, "Confirm")
If ans = 6 Then
rout.clearcurrent
End If
Else
rout.clearcurrent
End If
newentry = True

rout.obuts

End Sub

Private Sub nongarment_Change()
rout.choccur
rout.RCALC

End Sub

Private Sub NTN_Change()
rout.choccur
End Sub

Private Sub PAddr1_Change()
rout.choccur

End Sub

Private Sub paddr2_Change()
rout.choccur
End Sub

Private Sub PartyCombo_Change()
rout.choccur
End Sub

Public Sub PartyCombo_Click(Area As Integer)
On Error GoTo err
If RTrim(PartyCombo.BoundText) <> "" Then
party.Recordset.MoveFirst
party.Recordset.Find "Party_Id = '" & RTrim(PartyCombo.BoundText) & "'"
End If
err:
End Sub

Private Sub PCountry_Change()
rout.choccur
End Sub







Private Sub rds_charges_h_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub rds_solid_KeyPress(KeyAscii As Integer)
'MsgBox KeyAscii

 'f (KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii > 47 And KeyAscii < 58)) Then
'rout.choccur
'Else
KeyAscii = 0
'End If

End Sub



Private Sub rds_white_KeyPress(KeyAscii As Integer)
'MsgBox KeyAscii
'If (KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii > 47 And KeyAscii < 58)) Then
'rout.choccur
'Else
KeyAscii = 0
'End If

End Sub

Private Sub rdsamt_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub





Private Sub rdscharges_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub rdscombo_Click()
rout.choccur
Select Case rdscombo.ListIndex
Case 0
homet.Visible = True
hosiery.Visible = False
rout.RCALC
Exit Sub
Case 1
homet.Visible = False
hosiery.Visible = True
rout.RCALC
Exit Sub
Case Else
homet.Visible = False
hosiery.Visible = False
rout.RCALC
End Select
End Sub

Private Sub realdate_Change()
rout.choccur
End Sub

Private Sub remtext_Change()
rout.choccur
End Sub

Private Sub repbut_Click()
rform.Show (1)
End Sub

Private Sub rework_Click()

End Sub

Private Sub sarrival_date_Change()
rout.choccur
End Sub

Private Sub savebut_Click()
If Not rout.fval Then
MsgBox "Please enter necessary fields", vbQuestion, "Information"
Else
Set Myrset1 = rout.getrcset("RND", "Select * from RND where invoice_no = '" & LTrim(RTrim(invno.Text)) & "'", Mycon1, False)
'If Myrset1.RecordCount > 0 And savebut.Caption <> "&Save Changes" Then
If Myrset1.RecordCount > 0 And invno.Locked = False Then
MsgBox "Invoice Already Exist. Cannot save!", vbCritical, "Warning"
invno.SetFocus
Exit Sub
End If
rout.saveinvoice
End If

End Sub

Private Sub scountry_Change()
rout.choccur
End Sub

Private Sub Seacombo_Change()
rout.choccur
End Sub

Private Sub Seacombo_Click(Area As Integer)
On Error GoTo err
If RTrim(Seacombo.BoundText) <> "" Then
Sport.Recordset.MoveFirst
Sport.Recordset.Find "Destination_ID = '" & RTrim(Seacombo.BoundText) & "'"
End If
err:
End Sub

Private Sub Shipmentby_Click()
'MsgBox Shipmentby.ListIndex
rout.choccur
Select Case Shipmentby.ListIndex
Case 0
Sea.Visible = True
Flight.Visible = False
Exit Sub
Case 1
Sea.Visible = False
Flight.Visible = True
Exit Sub
Case Else
Sea.Visible = False
Flight.Visible = False
End Select

End Sub

Private Sub shippieces_Change()
rout.choccur
End Sub

Private Sub shippieces_KeyPress(KeyAscii As Integer)
If (KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii > 47 And KeyAscii < 58)) Then
Else
KeyAscii = 0
End If

End Sub

Private Sub shortship_Change()
rout.choccur
rout.RCALC

End Sub

Private Sub solid_weight_Change()
rout.choccur
rout.RCALC
End Sub

Private Sub solid_weight_KeyPress(KeyAscii As Integer)
If (KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii > 47 And KeyAscii < 58)) Then
rout.choccur
Else
KeyAscii = 0
End If
End Sub

Private Sub tdvs_Change()
rout.choccur
rout.RCALC

End Sub

Private Sub TOT_WEIGHT_Change()
rout.choccur
rout.RCALC

End Sub

Private Sub total_weight_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub unitsCombo_Change()
rout.choccur
End Sub


Private Sub valpkr_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub white_weight_Change()
rout.choccur
rout.RCALC
End Sub

Private Sub white_weight_KeyPress(KeyAscii As Integer)
If (KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii > 47 And KeyAscii < 58)) Then
rout.choccur
Else
KeyAscii = 0
End If
End Sub

Public Sub ToggleCapsLock(TurnOn As Boolean)

    'To turn capslock on, set turnon to true
    'To turn capslock off, set turnon to false
    
      Dim bytKeys(255) As Byte
      Dim bCapsLockOn As Boolean
      
'Get status of the 256 virtual keys
      GetKeyboardState bytKeys(0)
      
      bCapsLockOn = bytKeys(VK_CAPITAL)
      Dim typOS As OSVERSIONINFO
      
      If bCapsLockOn <> TurnOn Then 'if current state <>
                                     'requested stae
        
       If typOS.dwPlatformId = _
           VER_PLATFORM_WIN32_WINDOWS Then  '=== Win95/98

          bytKeys(VK_CAPITAL) = 1
          SetKeyboardState bytKeys(0)

        Else    '=== WinNT/2000

        'Simulate Key Press
          keybd_event VK_CAPITAL, &H45, _
             KEYEVENTF_EXTENDEDKEY Or 0, 0
        'Simulate Key Release
          keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY _
             Or KEYEVENTF_KEYUP, 0
        End If
      End If

     
End Sub


