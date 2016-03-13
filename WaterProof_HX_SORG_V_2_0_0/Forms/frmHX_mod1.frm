VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHX 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WaterProof HX - Heat Exchangers Performance Monitoring - Data Input"
   ClientHeight    =   11040
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000080&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10797.6
   ScaleMode       =   0  'User
   ScaleWidth      =   14352.46
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Data Data13 
      Caption         =   "Data13"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   12240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_s_fluid"
      Top             =   10680
      Width           =   1935
   End
   Begin VB.Data Data1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   390
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_test"
      Top             =   9240
      Width           =   6735
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Height          =   330
      Left            =   120
      TabIndex        =   170
      Top             =   9240
      Visible         =   0   'False
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20981
            Text            =   "Stato"
            TextSave        =   "Stato"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "07/03/2016"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "13.47"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame_Search 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5715
      Left            =   1440
      TabIndex        =   282
      ToolTipText     =   "Double click the unit name or date."
      Top             =   3480
      Visible         =   0   'False
      Width           =   13695
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   195
         Left            =   240
         TabIndex        =   303
         Top             =   5400
         Visible         =   0   'False
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Comm_search_close 
         Caption         =   "Close"
         Height          =   270
         Left            =   5520
         TabIndex        =   285
         ToolTipText     =   "Close the grid"
         Top             =   5340
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Bindings        =   "frmHX_mod1.frx":0000
         Height          =   4575
         Left            =   240
         TabIndex        =   283
         Top             =   300
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   8070
         _Version        =   393216
         Rows            =   7
         Cols            =   3
         FixedCols       =   0
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Double click Unit or  Date or Test_n°. The first item found for Unit or Date will be set."
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   180
         TabIndex        =   286
         Top             =   4980
         Width           =   6375
      End
   End
   Begin VB.Frame Frame_remarks 
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3555
      Left            =   1440
      TabIndex        =   291
      Top             =   5760
      Visible         =   0   'False
      Width           =   6735
      Begin RichTextLib.RichTextBox RichTextBox_REMARKS 
         DataField       =   "REMARKS"
         DataSource      =   "Data1"
         Height          =   3195
         Left            =   120
         TabIndex        =   292
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5636
         _Version        =   393217
         BackColor       =   12648447
         Enabled         =   -1  'True
         TextRTF         =   $"frmHX_mod1.frx":0015
      End
   End
   Begin VB.CommandButton Comm_search_grid 
      BackColor       =   &H00C0FFC0&
      Caption         =   "    Searching   all tests"
      Height          =   555
      Left            =   120
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   284
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox Date_Y 
      Height          =   255
      Left            =   360
      TabIndex        =   281
      Text            =   "Date_Y"
      Top             =   10080
      Width           =   1095
   End
   Begin VB.Data Data12 
      Caption         =   "Data12"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_Unit_sort"
      Top             =   10680
      Width           =   1815
   End
   Begin VB.Frame Frame_Mechanical 
      Caption         =   "Mechanical design data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5835
      Left            =   1440
      TabIndex        =   142
      Top             =   3420
      Width           =   6675
      Begin VB.ComboBox Combo_cooling_type 
         BackColor       =   &H00C0FFFF&
         DataField       =   "COOLING_TYPE"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         ItemData        =   "frmHX_mod1.frx":0091
         Left            =   4440
         List            =   "frmHX_mod1.frx":009E
         TabIndex        =   326
         Text            =   "COOLING TYPE"
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox Combo_CURRENT 
         BackColor       =   &H80000018&
         DataField       =   "CURRENT"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   330
         ItemData        =   "frmHX_mod1.frx":00C7
         Left            =   1620
         List            =   "frmHX_mod1.frx":00D7
         TabIndex        =   265
         Text            =   "Combo_CURRENT"
         ToolTipText     =   "Chose the flow pattern"
         Top             =   960
         Width           =   2475
      End
      Begin VB.Frame Frame_TUBES 
         Caption         =   "Tubes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   4515
         Left            =   0
         TabIndex        =   213
         Top             =   1320
         Width           =   3495
         Begin VB.TextBox Mat_factor 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            DataField       =   "TUBES_MAT_FACT"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   226
            ToolTipText     =   "Enter the value (suggested: 0.5 - 1.0)"
            Top             =   4020
            Width           =   795
         End
         Begin VB.CheckBox Check_U 
            Alignment       =   1  'Right Justify
            Caption         =   """U"""
            DataField       =   "Check_X"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   3060
            TabIndex        =   300
            ToolTipText     =   "Check if ""U"" tubes type"
            Top             =   480
            Width           =   195
         End
         Begin VB.CheckBox Check_MAT_FACTOR 
            BackColor       =   &H000000C0&
            DataField       =   "Check_T_mat_fact"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   120
            Left            =   3240
            MaskColor       =   &H0000FF00&
            Style           =   1  'Graphical
            TabIndex        =   290
            ToolTipText     =   "Check to activate the cursor and enter the value."
            Top             =   4040
            UseMaskColor    =   -1  'True
            Width           =   195
         End
         Begin VB.HScrollBar Spin_T_LEN 
            Height          =   195
            LargeChange     =   100
            Left            =   2220
            Max             =   2000
            Min             =   1
            TabIndex        =   278
            Top             =   756
            Value           =   1000
            Width           =   855
         End
         Begin VB.TextBox T_OD 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            DataField       =   "TUBES_OD"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   21
            ToolTipText     =   "Use the cursors to enter the OUTER DIAMETER of tubes."
            Top             =   1200
            Width           =   915
         End
         Begin VB.TextBox T_PASS 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            DataField       =   "TUBES_PASSES"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   19
            ToolTipText     =   "Use the cursors to enter the nimber of tube-side passes."
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox T_len 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            DataField       =   "TUBES_LE"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   18
            ToolTipText     =   "Use the cursors to enter the LENGHT of the exchanger."
            Top             =   720
            Width           =   915
         End
         Begin VB.TextBox TUBES_SECTION 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   231
            Top             =   3780
            Width           =   795
         End
         Begin VB.TextBox Mat_cond 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   227
            Top             =   3540
            Width           =   795
         End
         Begin VB.ComboBox Combo_MAT_SHEET 
            BackColor       =   &H80000018&
            DataField       =   "TUBES_SHEET_MAT"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   330
            ItemData        =   "frmHX_mod1.frx":0112
            Left            =   1080
            List            =   "frmHX_mod1.frx":018E
            Sorted          =   -1  'True
            TabIndex        =   26
            ToolTipText     =   "Select from the list the material of tube-sheet. Use TAB Key to save the new enter"
            Top             =   3120
            Width           =   2355
         End
         Begin VB.TextBox T_NO 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            DataField       =   "TUBES_NO"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   16
            ToolTipText     =   "Use the scrollbar to enter the NUMBER of TUBES."
            Top             =   480
            Width           =   915
         End
         Begin VB.HScrollBar HScroll_T_NO 
            Height          =   195
            LargeChange     =   100
            Left            =   2220
            Max             =   20000
            Min             =   10
            TabIndex        =   17
            Top             =   510
            Value           =   1000
            Width           =   855
         End
         Begin VB.TextBox T_ID 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   214
            ToolTipText     =   "Inner diameter of tubes."
            Top             =   2220
            Width           =   975
         End
         Begin VB.ComboBox Combo_BWG 
            BackColor       =   &H80000018&
            DataField       =   "TUBES_BWG"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   330
            ItemData        =   "frmHX_mod1.frx":0441
            Left            =   1320
            List            =   "frmHX_mod1.frx":0481
            TabIndex        =   24
            Text            =   "Combo_BWG"
            ToolTipText     =   "Select from the list the BWG. Use TAB Key to save the new enter."
            Top             =   1860
            Width           =   1215
         End
         Begin VB.ComboBox Combo_TUBES_Mat 
            BackColor       =   &H80000018&
            DataField       =   "TUBES_MAT"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   330
            ItemData        =   "frmHX_mod1.frx":04D2
            Left            =   1080
            List            =   "frmHX_mod1.frx":056D
            Sorted          =   -1  'True
            TabIndex        =   25
            ToolTipText     =   "Select from the list the material of tubes.Use TAB Key to save the new enter."
            Top             =   2760
            Width           =   2355
         End
         Begin VB.ComboBox ComboT_OD 
            BackColor       =   &H80000018&
            ForeColor       =   &H000000C0&
            Height          =   330
            ItemData        =   "frmHX_mod1.frx":0820
            Left            =   1320
            List            =   "frmHX_mod1.frx":0845
            TabIndex        =   23
            ToolTipText     =   "Select from the list the OUTER DIAMETER of tubes as fraction of inches."
            Top             =   1500
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   """U"""
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   27
            Left            =   3240
            TabIndex        =   301
            Top             =   480
            Width           =   225
         End
         Begin MSForms.SpinButton Spin_MAT_FACTOR 
            Height          =   195
            Left            =   2040
            TabIndex        =   289
            ToolTipText     =   "Enter the value (suggested: 0.5 - 1.0)"
            Top             =   4080
            Width           =   555
            Size            =   "979;344"
            Min             =   50
            Max             =   150
            Position        =   50
            Orientation     =   1
         End
         Begin VB.Label Label48 
            Caption         =   "Flow area section, m2:"
            ForeColor       =   &H00000080&
            Height          =   240
            Left            =   120
            TabIndex        =   230
            Top             =   3810
            Width           =   2415
         End
         Begin VB.Label lblLabels 
            Caption         =   "Thermal cond. Kcal/(h m^2 ºC/m):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   33
            Left            =   120
            TabIndex        =   229
            Top             =   3540
            Width           =   2535
         End
         Begin VB.Label lblLabels 
            Caption         =   "Material factor:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   34
            Left            =   120
            TabIndex        =   228
            Top             =   4080
            Width           =   1845
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Caption         =   "TUBES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   61
            Left            =   1560
            TabIndex        =   225
            Top             =   180
            Width           =   765
         End
         Begin VB.Label lblLabels 
            Caption         =   "Tubes:"
            ForeColor       =   &H00000080&
            Height          =   235
            Index           =   56
            Left            =   120
            TabIndex        =   224
            Top             =   2800
            Width           =   885
         End
         Begin VB.Label lblLabels 
            Caption         =   "Tube sheet:"
            ForeColor       =   &H00000080&
            Height          =   235
            Index           =   55
            Left            =   120
            TabIndex        =   223
            Top             =   3150
            Width           =   885
         End
         Begin VB.Label lblLabels 
            Caption         =   "Tubes Number:"
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   13
            Left            =   120
            TabIndex        =   222
            Top             =   480
            Width           =   1245
         End
         Begin VB.Label lblLabels 
            Caption         =   "Tubes lenght, m:"
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   14
            Left            =   120
            TabIndex        =   221
            Top             =   730
            Width           =   1245
         End
         Begin VB.Label lblLabels 
            Caption         =   "O.D., mm:"
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   11
            Left            =   120
            TabIndex        =   220
            Top             =   1230
            Width           =   1245
         End
         Begin VB.Label lblLabels 
            Caption         =   "BWG:"
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   12
            Left            =   120
            TabIndex        =   219
            Top             =   1900
            Width           =   1245
         End
         Begin VB.Label lblLabels 
            Caption         =   " I.D., mm:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   32
            Left            =   120
            TabIndex        =   218
            Top             =   2220
            Width           =   1245
         End
         Begin VB.Label lblLabels 
            Caption         =   "Passes, n°:"
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   15
            Left            =   120
            TabIndex        =   217
            Top             =   980
            Width           =   1245
         End
         Begin VB.Label lblLabels 
            Caption         =   "Material:"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   35
            Left            =   120
            TabIndex        =   216
            Top             =   2580
            Width           =   765
         End
         Begin MSForms.SpinButton Spin_T_PAS 
            Height          =   195
            Left            =   2220
            TabIndex        =   20
            Top             =   1002
            Width           =   855
            Size            =   "1508;344"
            Min             =   1
            Max             =   8
            Position        =   1
            Orientation     =   1
         End
         Begin MSForms.SpinButton Spin_T_OD 
            Height          =   195
            Left            =   2220
            TabIndex        =   22
            Top             =   1250
            Width           =   855
            Size            =   "1508;344"
            Max             =   10000
            Position        =   100
            Orientation     =   1
         End
         Begin VB.Label Label47 
            Caption         =   "O.D, inches:"
            ForeColor       =   &H00000080&
            Height          =   240
            Left            =   120
            TabIndex        =   215
            Top             =   1560
            Width           =   1245
         End
      End
      Begin VB.TextBox SERIES_N 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         DataField       =   "SERIES_N"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   3060
         Locked          =   -1  'True
         TabIndex        =   14
         ToolTipText     =   "Use the cursors to enter the number of units in series."
         Top             =   600
         Width           =   555
      End
      Begin VB.TextBox PARALLEL_N 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         DataField       =   "PARALLEL_N"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "Use the cursors to enter the number of units in parallel."
         Top             =   600
         Width           =   555
      End
      Begin VB.Frame Frame_Shell 
         Caption         =   "Shell"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   4540
         Left            =   3420
         TabIndex        =   193
         Top             =   1320
         Width           =   3255
         Begin VB.HScrollBar Spin_BAFFLES_SPACE 
            Height          =   195
            LargeChange     =   50
            Left            =   2280
            Max             =   2000
            Min             =   1
            TabIndex        =   280
            Top             =   1420
            Value           =   1000
            Width           =   855
         End
         Begin VB.HScrollBar Spin_SHELL_ID 
            Height          =   195
            LargeChange     =   100
            Left            =   2280
            Max             =   3000
            Min             =   1
            TabIndex        =   279
            Top             =   1720
            Value           =   1000
            Width           =   855
         End
         Begin VB.CheckBox CHECK_BAFFLES_N 
            BackColor       =   &H000000C0&
            DataField       =   "Check_BUFFLES"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   135
            Left            =   1380
            MaskColor       =   &H0000FF00&
            Style           =   1  'Graphical
            TabIndex        =   267
            ToolTipText     =   "Check to enter the value"
            Top             =   750
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.TextBox Clearance 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   234
            Top             =   2760
            Width           =   1215
         End
         Begin VB.TextBox Flow_area 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   232
            Top             =   3060
            Width           =   1215
         End
         Begin VB.ComboBox SHELL_PITCH_CONF 
            BackColor       =   &H80000018&
            DataField       =   "SHELL_PITCH_CONF"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   330
            ItemData        =   "frmHX_mod1.frx":0884
            Left            =   1380
            List            =   "frmHX_mod1.frx":088E
            Sorted          =   -1  'True
            TabIndex        =   37
            ToolTipText     =   "Select from the list.Use TAB Key to save the new enter"
            Top             =   2280
            Width           =   1755
         End
         Begin VB.TextBox SHELL_BAFFLES_N 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            DataField       =   "SHELL_BAFFLES_N"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   29
            ToolTipText     =   "Value calculated. Check the box to enter the value."
            Top             =   780
            Width           =   915
         End
         Begin VB.TextBox SHELL_TUBES_PITCH 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            DataField       =   "SHELL_TUBES_PITCH"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   35
            ToolTipText     =   "Use the cursors to enter the pitch value."
            Top             =   1980
            Width           =   915
         End
         Begin VB.TextBox SHELL_BAFFLES_SPACE 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            DataField       =   "SHELL_BAFFLES_SPACE"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   33
            ToolTipText     =   "Use the cursors to enter the inter-baffles space."
            Top             =   1380
            Width           =   915
         End
         Begin VB.TextBox SHELL_BAFFLES_CUT 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            DataField       =   "SHELL_BAFFLES_CUT"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   31
            ToolTipText     =   "Use the cursors to enter the percent cut-value."
            Top             =   1080
            Width           =   915
         End
         Begin VB.TextBox SHELL_ID 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            DataField       =   "SHELL_ID"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   34
            ToolTipText     =   "Use the cursors to enter the inner shell diameter."
            Top             =   1680
            Width           =   915
         End
         Begin VB.TextBox SHELL_PASS 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            DataField       =   "SHELL_PASSES"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   27
            ToolTipText     =   "Use the cursors to enter the shell-side passes."
            Top             =   480
            Width           =   915
         End
         Begin VB.ComboBox Combo_SHELL_MAT 
            BackColor       =   &H80000018&
            DataField       =   "SHELL_MAT"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   330
            ItemData        =   "frmHX_mod1.frx":08A6
            Left            =   780
            List            =   "frmHX_mod1.frx":091F
            Sorted          =   -1  'True
            TabIndex        =   38
            ToolTipText     =   "Select from the list the material of the shell.Use TAB Key to save the new enter"
            Top             =   3480
            Width           =   2355
         End
         Begin VB.Label lblLabels 
            Caption         =   "Clearance, m:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   73
            Left            =   120
            TabIndex        =   235
            Top             =   2790
            Width           =   1035
         End
         Begin VB.Label lblLabels 
            Caption         =   "Flow area section, m2:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   72
            Left            =   120
            TabIndex        =   233
            Top             =   3075
            Width           =   1665
         End
         Begin VB.Label lblLabels 
            Caption         =   "Pitch pattern.:"
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   69
            Left            =   120
            TabIndex        =   210
            Top             =   2325
            Width           =   1245
         End
         Begin VB.Label Label7 
            Caption         =   "Baffle n°:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   209
            Top             =   810
            Width           =   1155
         End
         Begin VB.Label lblLabels 
            Caption         =   " I.D., mm:"
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   59
            Left            =   120
            TabIndex        =   208
            Top             =   1695
            Width           =   1125
         End
         Begin VB.Label lblLabels 
            Caption         =   "Passes:"
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   58
            Left            =   120
            TabIndex        =   207
            Top             =   495
            Width           =   1185
         End
         Begin VB.Label lblLabels 
            Caption         =   "Material shell:"
            ForeColor       =   &H00000080&
            Height          =   375
            Index           =   57
            Left            =   120
            TabIndex        =   206
            Top             =   3420
            Width           =   765
         End
         Begin VB.Label Label5 
            Caption         =   "Baffle cut, %:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   205
            Top             =   1110
            Width           =   1155
         End
         Begin VB.Label Label2 
            Caption         =   "Tube Pitch, mm:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   204
            Top             =   2000
            Width           =   1215
         End
         Begin VB.Label Label30 
            Caption         =   "Baffle space,mm:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Left            =   120
            TabIndex        =   203
            Top             =   1410
            Width           =   1275
         End
         Begin MSForms.SpinButton Spin_BAFFLES_N 
            Height          =   195
            Left            =   2280
            TabIndex        =   30
            Top             =   820
            Width           =   855
            Size            =   "1508;344"
            Position        =   10
            Orientation     =   1
         End
         Begin MSForms.SpinButton Spin_TUBES_PITCH 
            Height          =   195
            Left            =   2280
            TabIndex        =   36
            Top             =   2020
            Width           =   855
            Size            =   "1508;344"
            Max             =   1000
            Position        =   25
            Orientation     =   1
         End
         Begin MSForms.SpinButton Spin_BAFFLES_CUT 
            Height          =   195
            Left            =   2280
            TabIndex        =   32
            Top             =   1120
            Width           =   855
            Size            =   "1508;344"
            Position        =   25
            Orientation     =   1
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Caption         =   "SHELL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   60
            Left            =   1380
            TabIndex        =   194
            Top             =   180
            Width           =   825
         End
         Begin MSForms.SpinButton Spin_S_PASS 
            Height          =   195
            Left            =   2280
            TabIndex        =   28
            Top             =   520
            Width           =   855
            Size            =   "1508;344"
            Min             =   1
            Max             =   8
            Position        =   1
            Orientation     =   1
         End
      End
      Begin VB.ComboBox Combo_POSITION 
         BackColor       =   &H80000018&
         DataField       =   "POSITION"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   330
         ItemData        =   "frmHX_mod1.frx":0BC4
         Left            =   3060
         List            =   "frmHX_mod1.frx":0BCE
         TabIndex        =   11
         Text            =   "Combo_POSITION"
         ToolTipText     =   "Chose if the exchanger is horizontal or vertical."
         Top             =   240
         Width           =   1155
      End
      Begin VB.ComboBox Combo_TEMA 
         BackColor       =   &H80000018&
         DataField       =   "TEMA"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   330
         ItemData        =   "frmHX_mod1.frx":0BE8
         Left            =   1080
         List            =   "frmHX_mod1.frx":0C0D
         TabIndex        =   10
         Text            =   "Combo_TYPE"
         ToolTipText     =   "Enter the TEMA type of the exchanger"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Area 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   5340
         Locked          =   -1  'True
         TabIndex        =   144
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label lblLabels 
         Caption         =   "Flow arrangement:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   266
         Top             =   1005
         Width           =   1395
      End
      Begin VB.Label lblLabels 
         Caption         =   "Series, n°:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   71
         Left            =   2280
         TabIndex        =   212
         Top             =   645
         Width           =   765
      End
      Begin MSForms.SpinButton Spin_SERIES_N 
         Height          =   195
         Left            =   3600
         TabIndex        =   15
         Top             =   600
         Width           =   615
         Size            =   "1085;344"
         Min             =   1
         Max             =   8
         Position        =   1
         Orientation     =   1
      End
      Begin VB.Label lblLabels 
         Caption         =   "Parallel, n°:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   70
         Left            =   180
         TabIndex        =   211
         Top             =   600
         Width           =   825
      End
      Begin MSForms.SpinButton Spin_PARALLEL_N 
         Height          =   195
         Left            =   1620
         TabIndex        =   13
         Top             =   620
         Width           =   555
         Size            =   "979;344"
         Min             =   1
         Max             =   8
         Position        =   1
         Orientation     =   1
      End
      Begin VB.Label Label39 
         Caption         =   "Position:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2280
         TabIndex        =   167
         Top             =   300
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "TEMA type:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   52
         Left            =   180
         TabIndex        =   158
         Top             =   300
         Width           =   825
      End
      Begin VB.Label lblLabels 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Surface, m2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   36
         Left            =   4260
         TabIndex        =   145
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Process"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1395
      Left            =   1440
      TabIndex        =   190
      Top             =   2160
      Width           =   6615
      Begin VB.ComboBox Combo_COOL_TOWER 
         BackColor       =   &H80000018&
         DataField       =   "COOL_TOWER"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   4980
         TabIndex        =   9
         Text            =   "Combo1"
         ToolTipText     =   "Enter the cooling tower used to cool this exchanger"
         Top             =   900
         Width           =   1515
      End
      Begin VB.ComboBox Combo_PROCESS_STREAM 
         BackColor       =   &H80000018&
         DataField       =   "PROCESS_STREAM"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   330
         ItemData        =   "frmHX_mod1.frx":0C48
         Left            =   1740
         List            =   "frmHX_mod1.frx":0C4A
         TabIndex        =   8
         Text            =   "Combo1"
         ToolTipText     =   "Enter the type of process stream"
         Top             =   900
         Width           =   2175
      End
      Begin VB.ComboBox Combo_PROCESS_DESCR 
         BackColor       =   &H80000018&
         DataField       =   "PROCESS_DESCR"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   1740
         TabIndex        =   7
         Text            =   "Combo1"
         ToolTipText     =   "Describe the service of this exchanger unit"
         Top             =   540
         Width           =   4755
      End
      Begin VB.ComboBox Combo_PLANT_UNIT 
         BackColor       =   &H80000018&
         DataField       =   "PLANT_UNIT"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   1740
         TabIndex        =   6
         Text            =   "Combo1"
         ToolTipText     =   "Chose or enter the plant-unit"
         Top             =   180
         Width           =   4755
      End
      Begin VB.Label lblLabels 
         Caption         =   "Cooling tower:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   3960
         TabIndex        =   264
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label lblLabels 
         Caption         =   "Plant Unit:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   63
         Left            =   180
         TabIndex        =   195
         Top             =   240
         Width           =   1605
      End
      Begin VB.Label lblLabels 
         Caption         =   "Process stream:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   54
         Left            =   180
         TabIndex        =   192
         Top             =   900
         Width           =   1185
      End
      Begin VB.Label lblLabels 
         Caption         =   "Service:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   53
         Left            =   180
         TabIndex        =   191
         Top             =   570
         Width           =   1605
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Unit identification"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2145
      Left            =   1440
      TabIndex        =   133
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton Comm_check_reset 
         Caption         =   "Reset check"
         Height          =   270
         Left            =   5460
         TabIndex        =   273
         ToolTipText     =   "Reset the selection to this unit"
         Top             =   1500
         Width           =   1095
      End
      Begin VB.CheckBox Check_ACT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Check to compare with design"
         DataField       =   "Check_Actual"
         DataSource      =   "Data1"
         ForeColor       =   &H00008000&
         Height          =   270
         Left            =   4260
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   272
         ToolTipText     =   "Check to compare this UNIT with DESIGN in the SUMMARY table"
         Top             =   1860
         UseMaskColor    =   -1  'True
         Width           =   2355
      End
      Begin VB.TextBox CHECK_ACTUAL 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   4140
         TabIndex        =   271
         Top             =   1860
         Width           =   150
      End
      Begin VB.TextBox CHECK_DESIGN 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   263
         Top             =   1860
         Width           =   150
      End
      Begin VB.CheckBox Check_des 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Check if  entering design data"
         DataField       =   "Check_Design"
         DataSource      =   "Data1"
         ForeColor       =   &H00008000&
         Height          =   270
         Left            =   1020
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Check if you enter design data"
         Top             =   1860
         UseMaskColor    =   -1  'True
         Width           =   2295
      End
      Begin VB.ComboBox Combo_UNIT 
         BackColor       =   &H80000018&
         DataField       =   "Unit_name"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   3000
         TabIndex        =   4
         Text            =   "Unit"
         ToolTipText     =   "Use TAB Key to save the new enter. Do not enter same UNIT name for the same PLANT name!"
         Top             =   1440
         Width           =   2235
      End
      Begin VB.ComboBox Combo_Country 
         BackColor       =   &H80000018&
         DataField       =   "Country"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   3000
         TabIndex        =   3
         Text            =   "Country"
         ToolTipText     =   "Use TAB Key to save the new enter"
         Top             =   1140
         Width           =   2235
      End
      Begin VB.ComboBox Combo_LOC 
         BackColor       =   &H80000018&
         DataField       =   "Location"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   3000
         TabIndex        =   2
         Text            =   "Location"
         ToolTipText     =   "Use TAB Key to save the new enter"
         Top             =   840
         Width           =   2235
      End
      Begin VB.ComboBox Combo_PLANT 
         BackColor       =   &H80000018&
         DataField       =   "Plant"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         ItemData        =   "frmHX_mod1.frx":0C4C
         Left            =   3000
         List            =   "frmHX_mod1.frx":0C4E
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "Plant"
         ToolTipText     =   "Use TAB Key to save the new enter"
         Top             =   540
         Width           =   2235
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "Date_test"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   3000
         TabIndex        =   0
         ToolTipText     =   "Enter the date (cannot be duplicated for the same unit)"
         Top             =   180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   -2147483633
         CalendarForeColor=   128
         CalendarTitleBackColor=   -2147483629
         CalendarTitleForeColor=   16512
         CalendarTrailingForeColor=   4210816
         CustomFormat    =   "01/01/01"
         Format          =   52428801
         CurrentDate     =   37920
         MinDate         =   36526
      End
      Begin VB.TextBox txt_num 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         DataField       =   "Test_NO"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   5940
         Locked          =   -1  'True
         TabIndex        =   157
         Text            =   "Record"
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox Plant 
         BackColor       =   &H8000000F&
         DataField       =   "Plant"
         DataSource      =   "Data1"
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   5580
         MaxLength       =   50
         TabIndex        =   159
         Text            =   "Plant"
         Top             =   555
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox Location 
         BackColor       =   &H8000000F&
         DataField       =   "Location"
         DataSource      =   "Data1"
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   5580
         MaxLength       =   50
         TabIndex        =   160
         Text            =   "Location"
         Top             =   855
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox Country 
         BackColor       =   &H8000000F&
         DataField       =   "Country"
         DataSource      =   "Data1"
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   5580
         MaxLength       =   50
         TabIndex        =   161
         Text            =   "Country"
         Top             =   1140
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox Unit 
         BackColor       =   &H8000000F&
         DataField       =   "Unit_name"
         DataSource      =   "Data1"
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   5580
         MaxLength       =   50
         TabIndex        =   162
         Text            =   "Unit"
         Top             =   1440
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   5160
         X2              =   5520
         Y1              =   1620
         Y2              =   1620
      End
      Begin VB.Label Label4 
         Caption         =   "Actual:"
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   3600
         TabIndex        =   274
         Top             =   1860
         Width           =   495
      End
      Begin VB.Line Line2 
         BorderStyle     =   3  'Dot
         BorderWidth     =   2
         X1              =   6000
         X2              =   6000
         Y1              =   1560
         Y2              =   1920
      End
      Begin VB.Label Label35 
         Caption         =   "Design:"
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   240
         TabIndex        =   171
         Top             =   1860
         Width           =   555
      End
      Begin VB.Label lblLabels 
         Caption         =   "Test_n°:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   21
         Left            =   5340
         TabIndex        =   143
         Top             =   180
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "Date  (must be unique for same unit!):"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   0
         Left            =   225
         TabIndex        =   138
         Top             =   285
         Width           =   2805
      End
      Begin VB.Label lblLabels 
         Caption         =   "Plant name:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   20
         Left            =   225
         TabIndex        =   137
         Top             =   615
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Location:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   19
         Left            =   225
         TabIndex        =   136
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Country:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   18
         Left            =   225
         TabIndex        =   135
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Heat transfer Unit I.D.:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   17
         Left            =   225
         TabIndex        =   134
         Top             =   1500
         Width           =   2715
      End
   End
   Begin VB.Data Data11 
      Caption         =   "Data11"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_Tower_list"
      Top             =   10680
      Width           =   1815
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3900
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_PROCESS_STREAM"
      Top             =   10680
      Width           =   1815
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2100
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_PROCESS_DESCR"
      Top             =   10680
      Width           =   1755
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_Plant_Unit_List"
      Top             =   10680
      Width           =   1755
   End
   Begin VB.Frame Frame3 
      Caption         =   "Output data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   6075
      Left            =   8040
      TabIndex        =   141
      Top             =   3540
      Width           =   7260
      Begin VB.CheckBox Comm_property 
         Caption         =   "Exist.properties"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   3600
         TabIndex        =   342
         ToolTipText     =   "Find existing properties in the data base for same fluid"
         Top             =   50
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Thermal balance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1455
         Left            =   5400
         TabIndex        =   321
         ToolTipText     =   "Click the preferred way compute the thermal balance"
         Top             =   3240
         Width           =   1695
         Begin VB.OptionButton Thermal_bal_tubes 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Tubes fn(t2) Fcost"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   324
            Top             =   1150
            Width           =   1575
         End
         Begin VB.OptionButton Thermal_bal_shell_T 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Shell   fn(T2) Fcost "
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   323
            Top             =   846
            Width           =   1575
         End
         Begin VB.OptionButton Thermal_bal_shell 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Shell   fn(F) T2cost"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   322
            Top             =   543
            Width           =   1575
         End
         Begin VB.OptionButton No_balance 
            BackColor       =   &H00C0FFFF&
            Caption         =   "No Balance"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   325
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.HScrollBar HScroll_WATER_FF 
         Height          =   195
         LargeChange     =   100
         Left            =   3660
         Max             =   1000
         TabIndex        =   297
         Top             =   5280
         Value           =   400
         Width           =   1035
      End
      Begin VB.TextBox WATER_FF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "WATER_FF"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   294
         Top             =   5220
         Width           =   1215
      End
      Begin VB.TextBox Wet_steam 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "Wet_steam"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   287
         Top             =   240
         Width           =   795
      End
      Begin VB.CheckBox Check_LATENT 
         BackColor       =   &H000000C0&
         DataField       =   "Check_LATENT"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   120
         Left            =   3900
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   276
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   2640
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.HScrollBar HScroll_LATENT 
         Height          =   195
         LargeChange     =   100
         Left            =   4800
         Max             =   3000
         TabIndex        =   275
         Top             =   2700
         Value           =   1000
         Width           =   855
      End
      Begin VB.CheckBox Check_T_TC 
         BackColor       =   &H000000C0&
         DataField       =   "Check_T_TC"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   120
         Left            =   1980
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   268
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.CheckBox Check_S_TC 
         BackColor       =   &H000000C0&
         DataField       =   "Check_S_TC"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   120
         Left            =   3900
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.HScrollBar HScroll_SHELL_TC 
         Height          =   195
         LargeChange     =   100
         Left            =   4800
         Max             =   1000
         TabIndex        =   103
         Top             =   540
         Value           =   150
         Width           =   795
      End
      Begin VB.CheckBox Check_CT 
         BackColor       =   &H000000C0&
         DataField       =   "Check_CT"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   120
         Left            =   3900
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   121
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   2880
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.HScrollBar HScroll_C_TEMP 
         Height          =   195
         LargeChange     =   100
         Left            =   4800
         Max             =   10000
         TabIndex        =   122
         Top             =   2940
         Value           =   2500
         Width           =   855
      End
      Begin VB.HScrollBar HScroll_TUBES_SPH 
         Height          =   195
         LargeChange     =   100
         Left            =   2880
         Max             =   1500
         TabIndex        =   86
         Top             =   780
         Value           =   150
         Width           =   795
      End
      Begin VB.HScrollBar HScroll_TUBES_DENS 
         Height          =   195
         LargeChange     =   100
         Left            =   2880
         Max             =   15000
         TabIndex        =   89
         Top             =   1020
         Value           =   150
         Width           =   795
      End
      Begin VB.HScrollBar HScroll_TUBES_VISC 
         Height          =   195
         LargeChange     =   100
         Left            =   2880
         Max             =   5000
         TabIndex        =   92
         Top             =   1260
         Value           =   150
         Width           =   795
      End
      Begin VB.CheckBox Check_T_DENS 
         BackColor       =   &H000000C0&
         DataField       =   "Check_T_DENS"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   120
         Left            =   1980
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.CheckBox Check_T_VISC 
         BackColor       =   &H000000C0&
         DataField       =   "Check_T_VISC"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   120
         Left            =   1980
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   91
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.CheckBox Check_T_SPH 
         BackColor       =   &H000000C0&
         DataField       =   "Check_T_SPH"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   120
         Left            =   1980
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.HScrollBar HScroll_TUBES_TC 
         Height          =   195
         LargeChange     =   100
         Left            =   2880
         Max             =   1500
         TabIndex        =   99
         Top             =   540
         Value           =   150
         Width           =   795
      End
      Begin VB.CheckBox Check_P_DROP_S 
         BackColor       =   &H000000C0&
         DataField       =   "Check_P_DROP_S"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   120
         Left            =   3900
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   118
         ToolTipText     =   "Check to see the allowed pressure drop."
         Top             =   2400
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.CheckBox Check_P_DROP_T 
         BackColor       =   &H000000C0&
         DataField       =   "Check_P_DROP_T"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   120
         Left            =   1980
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   "Check to see the allowed pressure drop."
         Top             =   2400
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.TextBox SHELL_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "Temp_COND"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   11
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   120
         ToolTipText     =   "Check the button and use the scroll bar to enter the value."
         Top             =   2880
         Width           =   915
      End
      Begin VB.TextBox SHELL_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "SHELL_LATENT"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   10
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   119
         ToolTipText     =   "Latent heat corresponding to the fraction condensing."
         Top             =   2640
         Width           =   915
      End
      Begin VB.TextBox SHELL_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "SHELL_PRESS_DROP"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   9
         Left            =   3900
         TabIndex        =   117
         Top             =   2400
         Width           =   915
      End
      Begin VB.TextBox SHELL_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   8
         Left            =   3900
         TabIndex        =   116
         Top             =   2160
         Width           =   915
      End
      Begin VB.TextBox SHELL_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "SHELL_DUTY"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   7
         Left            =   3900
         TabIndex        =   115
         Top             =   1920
         Width           =   915
      End
      Begin VB.TextBox SHELL_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "SHELL_REYNOLDS"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   6
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   114
         Top             =   1680
         Width           =   915
      End
      Begin VB.TextBox TUBES_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "TUBES_PRESS_DROP"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   9
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   97
         Top             =   2400
         Width           =   915
      End
      Begin VB.TextBox TUBES_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   8
         Left            =   1980
         TabIndex        =   96
         Top             =   2160
         Width           =   915
      End
      Begin VB.TextBox TUBES_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "TUBES_DUTY"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   7
         Left            =   1980
         TabIndex        =   95
         Top             =   1920
         Width           =   915
      End
      Begin VB.TextBox TUBES_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "TUBES_REYNOLDS"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   6
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   94
         Top             =   1680
         Width           =   915
      End
      Begin VB.TextBox TUBES_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "TUBES_VEL"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   2
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   93
         Top             =   1440
         Width           =   915
      End
      Begin VB.TextBox SHELL_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "SHELL_VEL"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   2
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   113
         Top             =   1440
         Width           =   915
      End
      Begin VB.CheckBox Check_U_CLEAN 
         BackColor       =   &H000000C0&
         DataField       =   "Check_U_CLEAN"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   120
         Left            =   2460
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   128
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   4140
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.HScrollBar HScroll_U_CLEAN 
         Height          =   195
         LargeChange     =   100
         Left            =   3660
         Max             =   10000
         TabIndex        =   129
         Top             =   4200
         Value           =   2000
         Width           =   1515
      End
      Begin VB.HScrollBar HScroll_SHELL_VISC 
         Height          =   195
         LargeChange     =   100
         Left            =   4800
         Max             =   1000
         TabIndex        =   112
         Top             =   1260
         Value           =   150
         Width           =   795
      End
      Begin VB.HScrollBar HScroll_SHELL_DENS 
         Height          =   195
         LargeChange     =   100
         Left            =   4800
         Max             =   10000
         TabIndex        =   109
         Top             =   1020
         Value           =   150
         Width           =   795
      End
      Begin VB.HScrollBar HScroll_SHELL_SPH 
         Height          =   195
         LargeChange     =   100
         Left            =   4800
         Max             =   2000
         TabIndex        =   106
         Top             =   780
         Value           =   150
         Width           =   795
      End
      Begin VB.CheckBox Check_S_SPH 
         BackColor       =   &H000000C0&
         DataField       =   "Check_S_SPH"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   120
         Left            =   3900
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   105
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.CheckBox Check_S_VISC 
         BackColor       =   &H000000C0&
         DataField       =   "Check_S_VISC"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   120
         Left            =   3900
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   111
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.CheckBox Check_S_DENS 
         BackColor       =   &H000000C0&
         DataField       =   "Check_S_DENS"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   120
         Left            =   3900
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   108
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.TextBox MTDc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "MTDc"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   125
         Top             =   3660
         Width           =   1215
      End
      Begin VB.TextBox U_COEFF_CLEAN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "Clean"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   127
         Top             =   4140
         Width           =   1215
      End
      Begin VB.TextBox SHELL_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "SHELL_VISC"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   5
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   110
         ToolTipText     =   "If blue and not checked,, the value is calculated."
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox TUBES_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "TUBES_VISC"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   5
         Left            =   1980
         TabIndex        =   90
         ToolTipText     =   "If blue and not checked,, the value is calculated."
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox SHELL_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "SHELL_DENS"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   4
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   107
         ToolTipText     =   "If blue and not checked,, the value is calculated."
         Top             =   960
         Width           =   915
      End
      Begin VB.TextBox TUBES_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "TUBES_DENS"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   4
         Left            =   1980
         TabIndex        =   87
         ToolTipText     =   "If blue and not checked,, the value is calculated."
         Top             =   960
         Width           =   915
      End
      Begin VB.TextBox SHELL_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "SHELL_SPH"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   3
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   104
         ToolTipText     =   "If blue and not checked,, the value is calculated."
         Top             =   720
         Width           =   915
      End
      Begin VB.TextBox TUBES_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "TUBES_SPH"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   3
         Left            =   1980
         TabIndex        =   84
         ToolTipText     =   "If blue and not checked,, the value is calculated."
         Top             =   720
         Width           =   915
      End
      Begin VB.TextBox TUBES_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "TUBES_T_COND"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   1
         Left            =   1980
         TabIndex        =   83
         ToolTipText     =   "If blue and not checked,, the value is calculated."
         Top             =   480
         Width           =   915
      End
      Begin VB.TextBox TUBES_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   82
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox TTD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "TTD"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   123
         Top             =   3180
         Width           =   1215
      End
      Begin VB.TextBox LMTD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "LMTD"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   124
         Top             =   3420
         Width           =   1215
      End
      Begin VB.TextBox C_Factor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "C_Factor"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   131
         Top             =   4620
         Width           =   1215
      End
      Begin VB.TextBox U_COEFF_DIRTY 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "U_DIRTY"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   130
         Top             =   4380
         Width           =   1215
      End
      Begin VB.TextBox SKIN_TEMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "SKIN_TEMP"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   126
         Top             =   3900
         Width           =   1215
      End
      Begin VB.TextBox TUBES_FF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000C0&
         DataField       =   "TUBES_FF"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   132
         Top             =   4920
         Width           =   1215
      End
      Begin VB.TextBox SHELL_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "SHELL_T_COND"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   1
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   101
         ToolTipText     =   "If blue and not checked,, the value is calculated."
         Top             =   480
         Width           =   915
      End
      Begin VB.TextBox SHELL_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   100
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label13 
         Caption         =   "Overall CLEAN Ucoeff.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   195
         TabIndex        =   182
         Top             =   4140
         Width           =   2235
      End
      Begin VB.Label lblLabels 
         Caption         =   "Allowed water side fouling factor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   26
         Left            =   180
         TabIndex        =   296
         Top             =   5240
         Width           =   2235
      End
      Begin VB.Label Label12 
         Caption         =   "[( hm^2 ºC ) / kcal] 10^-4"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   4725
         TabIndex        =   295
         Top             =   5280
         Width           =   1875
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "Wet steam"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   17
         Left            =   6000
         TabIndex        =   288
         ToolTipText     =   "Wet steam fraction"
         Top             =   50
         Width           =   1035
      End
      Begin VB.Label Label22 
         Caption         =   "Approach"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   15
         Left            =   4020
         TabIndex        =   277
         Top             =   3180
         Width           =   795
      End
      Begin VB.Label Label22 
         Caption         =   "(Allowable)"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   14
         Left            =   5220
         TabIndex        =   270
         Top             =   2430
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label22 
         Caption         =   "(Allowable)"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   12
         Left            =   2940
         TabIndex        =   269
         Top             =   2430
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label22 
         Caption         =   "m / s (liquid fraction)"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   26
         Left            =   4920
         TabIndex        =   241
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "[( hm^2 ºC ) / kcal] 10^-4"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   3780
         TabIndex        =   253
         Top             =   4950
         Width           =   1875
      End
      Begin VB.Label Label22 
         Caption         =   "kg/h"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   13
         Left            =   4860
         TabIndex        =   250
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label22 
         Caption         =   "°C"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   29
         Left            =   3780
         TabIndex        =   249
         Top             =   3660
         Width           =   435
      End
      Begin VB.Label Label22 
         Caption         =   "kJ/Kg"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   5760
         TabIndex        =   248
         Top             =   2700
         Width           =   495
      End
      Begin VB.Label lblLabels 
         Caption         =   "centipoise"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   80
         Left            =   5640
         TabIndex        =   245
         Top             =   1260
         Width           =   930
      End
      Begin VB.Label lblLabels 
         Caption         =   "Kg / m3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   78
         Left            =   5640
         TabIndex        =   243
         Top             =   1020
         Width           =   930
      End
      Begin VB.Label lblLabels 
         Caption         =   "Kcal / (Kg "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   76
         Left            =   5640
         TabIndex        =   240
         Top             =   780
         Width           =   930
      End
      Begin VB.Label lbl_tubes 
         Caption         =   "Kcal / h m ºC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   5640
         TabIndex        =   238
         Top             =   540
         Width           =   885
      End
      Begin VB.Label lblLabels 
         Caption         =   "kW/m2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   50
         Left            =   4920
         TabIndex        =   181
         Top             =   2190
         Width           =   765
      End
      Begin VB.Label lblLabels 
         Caption         =   "Water side fouling factor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   23
         Left            =   195
         TabIndex        =   177
         Top             =   4920
         Width           =   2235
      End
      Begin VB.Label Label40 
         Caption         =   "KW"
         ForeColor       =   &H00404080&
         Height          =   195
         Left            =   4920
         TabIndex        =   168
         Top             =   1980
         Width           =   375
      End
      Begin VB.Label Label22 
         Caption         =   "*C"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   11
         Left            =   5760
         TabIndex        =   155
         Top             =   2940
         Width           =   255
      End
      Begin VB.Label Label22 
         Caption         =   "kcal / (h m^2 ºC)"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   10
         Left            =   3780
         TabIndex        =   154
         Top             =   4400
         Width           =   1155
      End
      Begin VB.Label Label22 
         Caption         =   "kcal/h m^2ºC"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   9
         Left            =   4260
         TabIndex        =   153
         Top             =   4000
         Width           =   1035
      End
      Begin VB.Label Label22 
         Caption         =   "°C"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   6
         Left            =   3780
         TabIndex        =   152
         Top             =   3180
         Width           =   555
      End
      Begin VB.Label Label22 
         Caption         =   "°C"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   5
         Left            =   3780
         TabIndex        =   151
         Top             =   3420
         Width           =   435
      End
      Begin VB.Label Label22 
         Caption         =   "°C"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   4
         Left            =   3780
         TabIndex        =   150
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label Label22 
         Caption         =   "kPa"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   4920
         TabIndex        =   149
         Top             =   2430
         Width           =   435
      End
      Begin VB.Label Label24 
         Caption         =   "m3 / h / kPa^(1/2) - Tubes-side"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   3780
         TabIndex        =   148
         Top             =   4650
         Width           =   2355
      End
      Begin VB.Label lbl_tubes 
         Caption         =   "Flow calculated.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   236
         Top             =   240
         Width           =   1755
      End
      Begin VB.Label lbl_tubes 
         Caption         =   "Thermal conductivity: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   9
         Left            =   180
         TabIndex        =   251
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lbl_tubes 
         Caption         =   "Specific Heat:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   242
         Top             =   720
         Width           =   1755
      End
      Begin VB.Label lbl_tubes 
         Caption         =   "Density:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   244
         Top             =   960
         Width           =   1755
      End
      Begin VB.Label lbl_tubes 
         Caption         =   "Viscosity:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   246
         Top             =   1200
         Width           =   1755
      End
      Begin VB.Label lbl_tubes 
         Caption         =   "Flow velocity:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   174
         Top             =   1440
         Width           =   1755
      End
      Begin VB.Label lbl_tubes 
         Caption         =   "Reynolds number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   6
         Left            =   180
         TabIndex        =   172
         Top             =   1680
         Width           =   1755
      End
      Begin VB.Label lbl_tubes 
         Caption         =   "Duty:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   7
         Left            =   180
         TabIndex        =   180
         Top             =   1920
         Width           =   1755
      End
      Begin VB.Label lbl_tubes 
         Caption         =   "Heat flux:"
         ForeColor       =   &H00404080&
         Height          =   255
         Index           =   8
         Left            =   180
         TabIndex        =   169
         Top             =   2160
         Width           =   1755
      End
      Begin VB.Label lblLabels 
         Caption         =   "Total press.drop (clean):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   175
         Top             =   2400
         Width           =   1755
      End
      Begin VB.Label lblLabels 
         Caption         =   "Delta enthalpy(out-in):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   247
         Top             =   2640
         Width           =   3735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Condensing  film temperature:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   22
         Left            =   180
         TabIndex        =   176
         Top             =   2880
         Width           =   3735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Approach temperature:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   25
         Left            =   195
         TabIndex        =   179
         Top             =   3180
         Width           =   2235
      End
      Begin VB.Label lblLabels 
         Caption         =   "LMTD:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   24
         Left            =   195
         TabIndex        =   178
         Top             =   3420
         Width           =   2235
      End
      Begin VB.Label Label22 
         Caption         =   "MTDc"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   16
         Left            =   195
         TabIndex        =   239
         Top             =   3660
         Width           =   2235
      End
      Begin VB.Label Label22 
         Caption         =   "Skin temp."
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   195
         TabIndex        =   237
         Top             =   3900
         Width           =   2235
      End
      Begin VB.Label Label16 
         Caption         =   "Overall DIRTY coefficient:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   195
         TabIndex        =   183
         Top             =   4380
         Width           =   2235
      End
      Begin VB.Label lblLabels 
         Caption         =   "C Factor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   195
         TabIndex        =   173
         Top             =   4620
         Width           =   2235
      End
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_Date"
      Top             =   10320
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7620
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_Country"
      Top             =   10320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5700
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_LOC"
      Top             =   10320
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query UNIT_LIST"
      Top             =   10320
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_Plant"
      Top             =   10320
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_test"
      Top             =   10320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox PLANT_X 
      DataField       =   "PLANT_Z"
      DataSource      =   "Data2"
      Height          =   255
      Left            =   8880
      TabIndex        =   166
      Text            =   "PLANT_X"
      Top             =   10680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame7 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3555
      Left            =   0
      TabIndex        =   163
      Top             =   3180
      Width           =   1455
      Begin VB.ComboBox Combo_Date_X 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00004080&
         Height          =   330
         Left            =   60
         TabIndex        =   189
         Text            =   "Search Date"
         ToolTipText     =   "Chose the test-date to see."
         Top             =   2280
         Width           =   1335
      End
      Begin VB.ComboBox Combo_UNIT_1 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00004080&
         Height          =   330
         Left            =   60
         TabIndex        =   185
         Text            =   "Search unit"
         ToolTipText     =   "Chose the unit of the selected plant"
         Top             =   1620
         Width           =   1335
      End
      Begin VB.ComboBox Combo_Plant_1 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00004080&
         Height          =   330
         Left            =   60
         TabIndex        =   184
         Text            =   "Search plant"
         ToolTipText     =   "Chose the plant in the database"
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label46 
         Caption         =   "Date:"
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   120
         TabIndex        =   188
         Top             =   2040
         Width           =   1155
      End
      Begin VB.Label Label45 
         Caption         =   "Unit:"
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   120
         TabIndex        =   187
         Top             =   1380
         Width           =   795
      End
      Begin VB.Label Label44 
         Caption         =   "Plant:"
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   120
         TabIndex        =   186
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label38 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   60
         TabIndex        =   165
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Caption         =   "Search by Plant, Unit and Date"
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   60
         TabIndex        =   164
         Top             =   180
         Width           =   1275
      End
   End
   Begin VB.TextBox UNIT_X 
      DataField       =   "UNIT_Z"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7620
      TabIndex        =   156
      Text            =   "UNIT_X"
      Top             =   10680
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2475
      Left            =   120
      TabIndex        =   147
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   4366
      MultiRow        =   -1  'True
      Style           =   1
      HotTracking     =   -1  'True
      Placement       =   1
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   7
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Update"
            Object.ToolTipText     =   "Update the last changes/selections"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Reset"
            Object.ToolTipText     =   "Reset this record to the default values"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Get design"
            Object.ToolTipText     =   "Get design data from same existing unit and plant. "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
            Object.ToolTipText     =   "Add new record"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Object.ToolTipText     =   "Delete this record"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Object.ToolTipText     =   "Print this form"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Object.ToolTipText     =   "Close this form"
            ImageVarType    =   2
         EndProperty
      EndProperty
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Operating conditions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3675
      Left            =   8040
      TabIndex        =   139
      Top             =   0
      Width           =   7215
      Begin VB.CheckBox Check_T_OUT 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   120
         Left            =   1740
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   308
         ToolTipText     =   "Check to enter outlet temp. value and calculate flow rate"
         Top             =   2220
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.TextBox PROCESS_TARGET_T_OUT 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         DataField       =   "PROCESS_TARGET_TEMP"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   60
         TabIndex        =   298
         ToolTipText     =   "Process target outlet temp."
         Top             =   1500
         Width           =   495
      End
      Begin VB.CheckBox Check_CP 
         BackColor       =   &H000000C0&
         DataField       =   "Check_CP"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   120
         Left            =   3720
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   2940
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.TextBox FACT_FLOW 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         DataField       =   "FACT_FLOW"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   6240
         TabIndex        =   39
         Text            =   "FACT_FLOW"
         ToolTipText     =   "Span factor of flow"
         Top             =   240
         Width           =   615
      End
      Begin VB.HScrollBar Spin_PF 
         Height          =   195
         LargeChange     =   10
         Left            =   4620
         Max             =   10000
         TabIndex        =   81
         Top             =   3210
         Value           =   400
         Width           =   855
      End
      Begin VB.CheckBox Check_PF 
         BackColor       =   &H000000C0&
         DataField       =   "Check_PF"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   120
         Left            =   3720
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Uncheck to to have process fouling calculated"
         Top             =   3180
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.TextBox SHELL_FF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "SHELL_FF"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   3180
         Width           =   915
      End
      Begin VB.HScrollBar Spin_S_PRESS 
         Height          =   195
         LargeChange     =   10
         Left            =   4620
         Max             =   10000
         TabIndex        =   78
         Top             =   2970
         Value           =   100
         Width           =   855
      End
      Begin VB.TextBox S_press_KP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "Press_COND"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   76
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   2940
         Width           =   915
      End
      Begin VB.TextBox SHELL_P_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "SHELL_P_OUT"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   74
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   2700
         Width           =   915
      End
      Begin VB.HScrollBar HScroll_SHELL_P_OUT 
         Height          =   195
         LargeChange     =   10
         Left            =   4620
         Max             =   10000
         TabIndex        =   75
         Top             =   2730
         Value           =   5000
         Width           =   855
      End
      Begin VB.TextBox TUBES_P_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "TUBES_P_OUT"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   58
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   2700
         Width           =   915
      End
      Begin VB.HScrollBar HScroll_TUBES_P_OUT 
         Height          =   195
         LargeChange     =   10
         Left            =   2760
         Max             =   10000
         TabIndex        =   59
         Top             =   2740
         Value           =   5000
         Width           =   855
      End
      Begin VB.TextBox SHELL_P_IN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "SHELL_P_IN"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   72
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   2460
         Width           =   915
      End
      Begin VB.HScrollBar HScroll_SHELL_P_IN 
         Height          =   195
         LargeChange     =   100
         Left            =   4620
         Max             =   10000
         TabIndex        =   73
         Top             =   2490
         Value           =   5000
         Width           =   855
      End
      Begin VB.TextBox TUBES_P_IN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "TUBES_P_IN"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   56
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   2460
         Width           =   915
      End
      Begin VB.HScrollBar HScroll_TUBES_P_IN 
         Height          =   195
         LargeChange     =   10
         Left            =   2760
         Max             =   10000
         TabIndex        =   57
         Top             =   2497
         Value           =   5000
         Width           =   855
      End
      Begin VB.TextBox SHELL_TEMP_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "SHELL_TEMP_OUT"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   70
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   2205
         Width           =   915
      End
      Begin VB.HScrollBar HScroll_SHELL_T_OUT 
         Height          =   195
         LargeChange     =   100
         Left            =   4620
         Max             =   10000
         Min             =   5
         TabIndex        =   71
         Top             =   2250
         Value           =   3500
         Width           =   855
      End
      Begin VB.HScrollBar Spin_TUBES_T_OUT 
         Height          =   195
         LargeChange     =   100
         Left            =   2760
         Max             =   30000
         Min             =   5
         TabIndex        =   55
         Top             =   2256
         Value           =   3500
         Width           =   855
      End
      Begin VB.TextBox TUBES_TEMP_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "TUBES_TEMP_OUT"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   54
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   2205
         Width           =   915
      End
      Begin VB.TextBox SHELL_TEMP_IN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "SHELL_TEMP_IN"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   68
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   1980
         Width           =   915
      End
      Begin VB.HScrollBar HScroll_SHELL_T_IN 
         Height          =   195
         LargeChange     =   100
         Left            =   4620
         Max             =   30000
         Min             =   5
         TabIndex        =   69
         Top             =   2010
         Value           =   2500
         Width           =   855
      End
      Begin VB.HScrollBar Spin_TUBES_T_IN 
         Height          =   195
         LargeChange     =   100
         Left            =   2760
         Max             =   30000
         Min             =   5
         TabIndex        =   53
         Top             =   2015
         Value           =   2500
         Width           =   855
      End
      Begin VB.TextBox TUBES_TEMP_IN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "TUBES_TEMP_IN"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   52
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   1980
         Width           =   915
      End
      Begin VB.TextBox TUBES_NON_COND 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "TUBES_NON_COND"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   50
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   1740
         Width           =   915
      End
      Begin VB.HScrollBar HScroll_TUBES_NON_COND 
         Height          =   195
         LargeChange     =   10
         Left            =   2760
         Max             =   10000
         TabIndex        =   51
         Top             =   1774
         Width           =   855
      End
      Begin VB.HScrollBar HScroll_SHELL_NON_COND 
         Height          =   195
         LargeChange     =   10
         Left            =   4620
         Max             =   10000
         TabIndex        =   67
         Top             =   1770
         Width           =   855
      End
      Begin VB.TextBox SHELL_NON_COND 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "SHELL_NON_COND"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   66
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox TUBES_WATER 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "TUBES_WATER"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   48
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   1500
         Width           =   915
      End
      Begin VB.HScrollBar HScroll_TUBES_WATER 
         Height          =   195
         LargeChange     =   10
         Left            =   2760
         Max             =   10000
         TabIndex        =   49
         Top             =   1533
         Width           =   855
      End
      Begin VB.TextBox SHELL_WATER 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "SHELL_WATER"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   65
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   1500
         Width           =   915
      End
      Begin VB.TextBox TUBES_LIQUID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "TUBES_LIQUID"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   46
         ToolTipText     =   "Includes water."
         Top             =   1260
         Width           =   915
      End
      Begin VB.HScrollBar HScroll_TUBES_LIQUID 
         Height          =   195
         LargeChange     =   10
         Left            =   2760
         Max             =   10000
         TabIndex        =   47
         Top             =   1292
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox SHELL_LIQUID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "SHELL_LIQUID"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   64
         ToolTipText     =   "Includes water."
         Top             =   1260
         Width           =   915
      End
      Begin VB.TextBox TUBES_VAPOR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "TUBES_VAPOR"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   44
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   1020
         Width           =   915
      End
      Begin VB.HScrollBar HScroll_TUBES_VAPOR 
         Height          =   195
         LargeChange     =   10
         Left            =   2760
         Max             =   10000
         TabIndex        =   45
         Top             =   1051
         Width           =   855
      End
      Begin VB.TextBox SHELL_VAPOR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "SHELL_VAPOR"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   63
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   1020
         Width           =   915
      End
      Begin VB.TextBox SHELL_FLOW 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "SHELL_FLOW"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   61
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   780
         Width           =   915
      End
      Begin VB.HScrollBar Spin_SHELL_FLOW 
         Height          =   195
         LargeChange     =   100
         Left            =   4620
         Max             =   10000
         TabIndex        =   62
         Top             =   810
         Width           =   855
      End
      Begin VB.HScrollBar HScroll_TUBES_FLOW 
         Height          =   195
         LargeChange     =   10
         Left            =   2760
         Max             =   10000
         TabIndex        =   43
         Top             =   810
         Width           =   855
      End
      Begin VB.TextBox TUBES_FLOW 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "TUBES_FLOW"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   42
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   780
         Width           =   915
      End
      Begin VB.ComboBox Combo_T_FLUID 
         BackColor       =   &H80000018&
         DataField       =   "TUBES_FLUID"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   330
         ItemData        =   "frmHX_mod1.frx":0C50
         Left            =   1860
         List            =   "frmHX_mod1.frx":0C57
         TabIndex        =   41
         ToolTipText     =   "Select from the list.Use TAB Key to save the new enter"
         Top             =   420
         Width           =   1755
      End
      Begin VB.ComboBox Combo_S_FLUID 
         BackColor       =   &H80000018&
         DataField       =   "SHELL_FLUID"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   330
         ItemData        =   "frmHX_mod1.frx":0C62
         Left            =   3720
         List            =   "frmHX_mod1.frx":0C64
         TabIndex        =   60
         ToolTipText     =   "Select from the list.Use TAB Key to save the new enter"
         Top             =   420
         Width           =   1815
      End
      Begin VB.Frame Frame_VAP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fluid Fractions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2835
         Left            =   5760
         TabIndex        =   305
         Top             =   600
         Width           =   1455
         Begin VB.TextBox Vwat_INP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            DataField       =   "Vwat_perc"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   338
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox Vorg_INP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   840
            TabIndex        =   339
            Top             =   825
            Width           =   495
         End
         Begin VB.TextBox Vtot_INP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            DataField       =   "VAP_FRACTION"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   330
            Left            =   840
            TabIndex        =   340
            ToolTipText     =   "Fraction of condensing vapor"
            Top             =   555
            Width           =   495
         End
         Begin VB.TextBox Vwat_IN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   0
            TabIndex        =   341
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox Vorg_IN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            ForeColor       =   &H000000C0&
            Height          =   270
            Left            =   0
            TabIndex        =   336
            Top             =   825
            Width           =   735
         End
         Begin VB.TextBox Vtot_IN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            ForeColor       =   &H000000C0&
            Height          =   330
            Left            =   0
            TabIndex        =   337
            Top             =   555
            Width           =   735
         End
         Begin VB.TextBox Ltot_IN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   0
            TabIndex        =   316
            Top             =   2040
            Width           =   735
         End
         Begin VB.TextBox Lwat_IN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   0
            TabIndex        =   311
            Top             =   1800
            Width           =   735
         End
         Begin VB.TextBox Ltot_INP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            DataField       =   "LIQ_FRACTION"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   317
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox Lwat_INP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            DataField       =   "Lwat_perc"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   306
            ToolTipText     =   "Fraction of liquid + non condensable"
            Top             =   1800
            Width           =   495
         End
         Begin VB.TextBox Ftot_INP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   315
            Top             =   2520
            Width           =   495
         End
         Begin VB.TextBox Lorg_INP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   314
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox Lorg_IN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   0
            TabIndex        =   313
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox Ftot_IN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   0
            TabIndex        =   312
            Top             =   2520
            Width           =   735
         End
         Begin MSForms.SpinButton Spin_WAT_LIQ_IN 
            Height          =   255
            Left            =   720
            TabIndex        =   328
            ToolTipText     =   "Percent of water liquid"
            Top             =   1800
            Width           =   135
            Size            =   "238;450"
            Max             =   1000
         End
         Begin MSForms.SpinButton Spin_WAT_VAP_IN 
            Height          =   255
            Left            =   720
            TabIndex        =   327
            ToolTipText     =   "Percent of water vapor"
            Top             =   1080
            Width           =   135
            Size            =   "238;450"
            Max             =   1000
         End
         Begin VB.Line Line7 
            X1              =   0
            X2              =   0
            Y1              =   1560
            Y2              =   1080
         End
         Begin VB.Line Line6 
            X1              =   0
            X2              =   0
            Y1              =   1560
            Y2              =   1080
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "Vapor IN"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   270
            TabIndex        =   320
            Top             =   150
            Width           =   855
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "Total"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   360
            TabIndex        =   319
            Top             =   2310
            Width           =   855
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "%"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   870
            TabIndex        =   318
            Top             =   330
            Width           =   375
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "kg/h"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   190
            TabIndex        =   310
            Top             =   330
            Width           =   375
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "Liquid IN"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   360
            TabIndex        =   309
            Top             =   1350
            Width           =   855
         End
         Begin MSForms.SpinButton Spin_VAP_P 
            Height          =   345
            Left            =   690
            TabIndex        =   307
            ToolTipText     =   "Fraction of condensing vapor"
            Top             =   510
            Width           =   195
            Size            =   "344;609"
            Max             =   1000
            Position        =   1
         End
      End
      Begin VB.Label Label23 
         Caption         =   "FT"
         Height          =   255
         Index           =   6
         Left            =   5520
         TabIndex        =   335
         ToolTipText     =   "Total inlet fluid"
         Top             =   3165
         Width           =   255
      End
      Begin VB.Label Label23 
         Caption         =   "LT"
         Height          =   255
         Index           =   5
         Left            =   5520
         TabIndex        =   334
         ToolTipText     =   "Total inlet liquid fluid"
         Top             =   2685
         Width           =   255
      End
      Begin VB.Label Label23 
         Caption         =   "Lw"
         Height          =   255
         Index           =   4
         Left            =   5520
         TabIndex        =   333
         ToolTipText     =   "Inlet liquid water fluid"
         Top             =   2445
         Width           =   255
      End
      Begin VB.Label Label23 
         Caption         =   "Lo"
         Height          =   255
         Index           =   3
         Left            =   5520
         TabIndex        =   332
         ToolTipText     =   "Inlet liquid organic fluid"
         Top             =   2205
         Width           =   255
      End
      Begin VB.Label Label23 
         Caption         =   "Vw"
         Height          =   255
         Index           =   2
         Left            =   5490
         TabIndex        =   331
         ToolTipText     =   "Inlet water vapor fluid"
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label23 
         Caption         =   "Vo"
         Height          =   255
         Index           =   1
         Left            =   5490
         TabIndex        =   330
         ToolTipText     =   "Inlet organic vapor fluid"
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label23 
         Caption         =   "VT"
         Height          =   255
         Index           =   0
         Left            =   5490
         TabIndex        =   329
         ToolTipText     =   "Inlet total vapor fluid "
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lblLabels 
         Caption         =   "Condensation  pressure:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   10
         Left            =   60
         TabIndex        =   261
         Top             =   2910
         Width           =   1905
      End
      Begin VB.Label Label22 
         Caption         =   "Pressure OUT (design), bar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   28
         Left            =   60
         TabIndex        =   254
         Top             =   2685
         Width           =   1905
      End
      Begin VB.Label Label22 
         Caption         =   "Pressure IN (design), bar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   1
         Left            =   60
         TabIndex        =   255
         Top             =   2460
         Width           =   1905
      End
      Begin VB.Label Label17 
         Caption         =   "Span flow"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6180
         TabIndex        =   304
         Top             =   60
         Width           =   795
      End
      Begin MSForms.SpinButton Spin_TARGET_T 
         Height          =   255
         Left            =   540
         TabIndex        =   302
         ToolTipText     =   "Target T out"
         Top             =   1500
         Width           =   195
         Size            =   "344;450"
         Max             =   1000
         Position        =   1
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Target T2"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   0
         TabIndex        =   299
         Top             =   1320
         Width           =   735
      End
      Begin MSForms.SpinButton Spin_FACT_FLOW 
         Height          =   255
         Left            =   6840
         TabIndex        =   40
         ToolTipText     =   "Span factor of flow"
         Top             =   240
         Width           =   195
         Size            =   "344;450"
         Min             =   1
         Max             =   2000
         Position        =   1
      End
      Begin VB.Label lblLabels 
         Caption         =   "Processr side fouling factor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   262
         Top             =   3240
         Width           =   2010
      End
      Begin VB.Label lblLabels 
         Caption         =   "Temperature OUT, °C"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   8
         Left            =   60
         TabIndex        =   260
         Top             =   2220
         Width           =   1905
      End
      Begin VB.Label lblLabels 
         Caption         =   "Temperature IN, °C"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   7
         Left            =   60
         TabIndex        =   259
         Top             =   1980
         Width           =   1905
      End
      Begin VB.Label lblLabels 
         Caption         =   "Total INLET flow rate, kg/h:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   6
         Left            =   60
         TabIndex        =   258
         Top             =   780
         Width           =   1905
      End
      Begin VB.Label lblLabels 
         Caption         =   "Fluid type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   38
         Left            =   60
         TabIndex        =   257
         Top             =   480
         Width           =   1845
      End
      Begin VB.Label lblLabels 
         Caption         =   "Non-cond.,kg/h:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   67
         Left            =   870
         TabIndex        =   256
         Top             =   1740
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "[(hm^2 ºC)/kcal]*10^-4"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   2040
         TabIndex        =   252
         Top             =   3240
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "kPa(a)"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   10
         Left            =   3120
         TabIndex        =   146
         Top             =   2955
         Width           =   555
      End
      Begin VB.Label Label52 
         Caption         =   "kg/h"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4680
         TabIndex        =   202
         Top             =   1530
         Width           =   315
      End
      Begin VB.Label lblLabels 
         Caption         =   "Water (out):"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   68
         Left            =   855
         TabIndex        =   201
         Top             =   1500
         Width           =   1110
      End
      Begin VB.Label Label50 
         Caption         =   "kg/h"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4680
         TabIndex        =   200
         Top             =   1290
         Width           =   315
      End
      Begin VB.Label lblLabels 
         Caption         =   "Liquid (out):"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   66
         Left            =   855
         TabIndex        =   199
         Top             =   1260
         Width           =   1080
      End
      Begin VB.Label Label49 
         Caption         =   "kg/h"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4680
         TabIndex        =   198
         Top             =   1050
         Width           =   315
      End
      Begin VB.Label lblLabels 
         Caption         =   "Vapor (cond.):"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   9
         Left            =   855
         TabIndex        =   140
         Top             =   1020
         Width           =   1110
      End
      Begin VB.Label lblLabels 
         Caption         =   "SHELL:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   65
         Left            =   3960
         TabIndex        =   197
         Top             =   180
         Width           =   825
      End
      Begin VB.Label lblLabels 
         Caption         =   "TUBES:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   64
         Left            =   2160
         TabIndex        =   196
         Top             =   180
         Width           =   765
      End
   End
   Begin MSForms.ToggleButton Toggle_remarks 
      Height          =   375
      Left            =   106
      TabIndex        =   293
      Top             =   8400
      Width           =   1215
      BackColor       =   12648384
      ForeColor       =   192
      DisplayStyle    =   6
      Size            =   "2143;661"
      Value           =   "0"
      Caption         =   "Remarks"
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmHX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public XXX, YXY, COL, foul, rec1, lungh As Integer
Public metal, PROP As String
Public A40, D11, D13, D15, D17, D18, D19, D20, D29, D38, D40_S, D40_T, D43, D44 As Double
Public D54, D67, D68, D69, D70, D71, D72, D73, D74, D75, D76, D77, D78, D80 As Double
Public T_FLOW, S_FLOW, Range_T, Range_S As Double
Public K_PF, K_P_DROP_T, K_P_DROP_S, K_S_SPH, K_S_DENS, K_S_VISC
Public K_T_SPH, K_T_DENS, K_T_VISC, K_U_CLEAN
Public XD21, XD22, XD23, XD24, TH_C As Double
Public XPI, LN, XD6, XD7, XD8, XD9, XD10, SPH_S, SPH_T
Public XD18, XD19, XD20, XD9_S, COND, FLUID, FLUID_VL, DES_1
Public XD37, XD52M, XD52, XD54, XD55, XD56, XD57, XD58, XD59, XD61M
Public XD61, XD63, XD64M, XD64, XD66M, XD66, XD50, XD85, XD84, XD83, XD112
Public Q_S_V, Q_S_L, Q_S_W, Q_S_NC, S_press_KP_1, W_flow_IN_1, VAP_PERC_1
Public XD73, XD74, XD75, XD76, XD77, XD78, XD79
Public Vtot_IN_1, Vorg_IN_1, Vwat_IN_1, Ltot_IN_1, Lorg_IN_1
Public Lwat_IN_1, Vorg_INP_1, Vwat_INP_1, Lorg_INP_1, Lwat_INP_1
Public Vtot_INP_1, Ltot_INP_1, NC_IN_1
Dim S_side(20)
Private Sub Form_Load()
On Error Resume Next
    Width = frmMain.Width * 0.985 ' Imposta la larghezza del form.
    Height = frmMain.Height * 0.898    ' Imposta l'altezza del form.
    Left = 0 '(frmMain.Width - Width) / 2 ' Centra il form orizzontalmente.
    Top = 0 '(frmMain.Height - Height) / 2 ' Centra il form verticalmente.
    frmHX.WindowState = 2
    
    Dim Date_X As Date
    Dim Rs3 As Recordset
    Data3.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
    Data3.RecordSource = "Select * From [QUERY_Plant]"
    Data3.Refresh
    Set Rs3 = Data3.Recordset
    If Rs3.RecordCount > 0 Then
       Do Until Rs3.EOF
          PP1 = Data3.Recordset.Plant
          Combo_PLANT.AddItem PP1
          Combo_Plant_1.AddItem PP1
          Rs3.MoveNext
       Loop
    Else
       MsgBox "Date not found"
    End If
        
    Dim Rs4 As Recordset
    Data4.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
    Data4.RecordSource = "Select * From [Query UNIT_LIST]"
    Data4.Refresh
    Set Rs4 = Data4.Recordset
    If Rs4.RecordCount > 0 Then
       Do Until Rs4.EOF
            UUU1 = Data4.Recordset.Unit_name
                Combo_UNIT.AddItem UUU1
            Rs4.MoveNext
    Loop
    Else
       MsgBox "No Units found"
    End If
        
    Dim Rs5 As Recordset
    Data5.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
    Data5.RecordSource = "Select * From [Query_LOC]"
    Data5.Refresh
    Set Rs5 = Data5.Recordset
    If Rs5.RecordCount > 0 Then
       Do Until Rs5.EOF
            LLL1 = Data5.Recordset.Location
                Combo_LOC.AddItem LLL1
            Rs5.MoveNext
    Loop
    Else
       MsgBox "No Units found"
    End If
        
    Dim Rs6 As Recordset
    Data6.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
    Data6.RecordSource = "Select * From [Query_Country]"
    Data6.Refresh
    Set Rs6 = Data6.Recordset
    If Rs6.RecordCount > 0 Then
       Do Until Rs6.EOF
            CCC1 = Data6.Recordset.Country
                Combo_Country.AddItem CCC1
            Rs6.MoveNext
    Loop
    Else
       MsgBox "No Units found"
    End If
    
    Combo_PLANT.Text = Data1.Recordset.Plant
    Combo_LOC.Text = Data1.Recordset.Location
    Combo_Country.Text = Data1.Recordset.Country
    Combo_UNIT.Text = Data1.Recordset.Unit_name
    
Dim Rs8 As Recordset
    Data8.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
    Data8.RecordSource = "Select * From [Query_PLANT_UNIT_LIST]"
    Data8.Refresh
    Set Rs8 = Data8.Recordset
    If Rs8.RecordCount > 0 Then
       Do Until Rs8.EOF
          PP2 = Data8.Recordset.PLANT_UNIT
          Combo_PLANT_UNIT.AddItem PP2
          Rs8.MoveNext
       Loop
    Else
       MsgBox "Plant-unit not found"
    End If
    
Dim Rs9 As Recordset
    Data9.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
    Data9.RecordSource = "Select * From [Query_PROCESS_DESCR]"
    Data9.Refresh
    Set Rs9 = Data9.Recordset
    If Rs9.RecordCount > 0 Then
       Do Until Rs9.EOF
          PP3 = Data9.Recordset.PROCESS_DESCR
          Combo_PROCESS_DESCR.AddItem PP3
          Rs9.MoveNext
       Loop
    Else
       MsgBox "Process description not found"
    End If
    
Dim Rs10 As Recordset
    Data10.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
    Data10.RecordSource = "Select * From [Query_PROCESS_STREAM]"
    Data10.Refresh
    Set Rs10 = Data10.Recordset
    If Rs10.RecordCount > 0 Then
       Do Until Rs10.EOF
          PP4 = Data10.Recordset.PROCESS_STREAM
          Combo_PROCESS_STREAM.AddItem PP4
          Rs10.MoveNext
       Loop
    Else
       MsgBox "Process stream not found"
    End If
    
Dim Rs11 As Recordset
    Data11.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
    Data11.RecordSource = "Select * From [Query_TOWER_list]"
    Data11.Refresh
    Set Rs11 = Data11.Recordset
    If Rs11.RecordCount > 0 Then
       Do Until Rs11.EOF
          PP5 = Data11.Recordset.COOL_TOWER
          Combo_COOL_TOWER.AddItem PP5
          Rs11.MoveNext
       Loop
    Else
       MsgBox "TOWER not found"
    End If
    
 Dim Rs13 As Recordset
    Data13.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
    Data13.RecordSource = "Select * From [Query_s_fluid]"
    Data13.Refresh
    Set Rs13 = Data13.Recordset
    If Rs13.RecordCount > 0 Then
       Do Until Rs13.EOF
          FLUID = Data13.Recordset.SHELL_FLUID
          Combo_S_FLUID.AddItem FLUID
          Rs13.MoveNext
       Loop
    Else
       MsgBox "Fluid not found"
    End If
        
    Combo_PLANT_UNIT = Data1.Recordset.PLANT_UNIT
    Combo_PROCESS_DESCR = Data1.Recordset.PROCESS_DESCR
    Combo_PROCESS_STREAM = Data1.Recordset.PROCESS_STREAM
    Combo_COOL_TOWER = Data1.Recordset.COOL_TOWER
    
    FACT_FLOW = Data1.Recordset.FACT_FLOW
    Spin_FACT_FLOW.Value = FACT_FLOW
    
    Check_des.Value = Data1.Recordset.CHECK_DESIGN
    Check_ACT.Value = Data1.Recordset.CHECK_ACTUAL
    CHECK_BAFFLES_N.Value = Data1.Recordset.Check_BUFFLES
    Check_T_TC.Value = Data1.Recordset.Check_T_TC
    Check_T_SPH.Value = Data1.Recordset.Check_T_SPH
    Check_T_DENS.Value = Data1.Recordset.Check_T_DENS
    Check_T_VISC.Value = Data1.Recordset.Check_T_VISC
    Check_P_DROP_T.Value = Data1.Recordset.Check_P_DROP_T
    Check_S_TC.Value = Data1.Recordset.Check_S_TC
    Check_S_SPH.Value = Data1.Recordset.Check_S_SPH
    Check_S_DENS.Value = Data1.Recordset.Check_S_DENS
    Check_S_VISC.Value = Data1.Recordset.Check_S_VISC
    Check_LATENT.Value = Data1.Recordset.Check_LATENT
    Check_P_DROP_S.Value = Data1.Recordset.Check_P_DROP_S
    Check_CP.Value = Data1.Recordset.Check_CP
    Check_CT.Value = Data1.Recordset.Check_CT
    Check_PF.Value = Data1.Recordset.Check_PF
    Check_U_CLEAN.Value = Data1.Recordset.Check_U_CLEAN
    Check_U = Data1.Recordset.Check_X
    Check_T_mat_fact = Data1.Recordset.Check_T_mat_fact

    foul = 1
    No_balance = True
    With Grid1
        .COL = 2
        .Row = 7
        .ColWidth(0) = 1000
        .ColWidth(1) = 1100
        .DataSource = Rs12
        .AddItem (Unit_name), (date_test)
    End With
    XXX = 0
End Sub
Private Sub Grid1_DblClick()
On Error Resume Next
Dim Date_1 As Date
Dim Date_2 As Date
Bar1.Visible = True
Bar1 = 10
XXX = 1
TEST_Y = Grid1.Text
If Grid1.ColSel < 3 Then
    Data1.UpdateRecord
    Data1.Recordset.MoveLast
    n_record = Data1.Recordset.RecordCount
    Data1.Recordset.MoveFirst
    pos = Data1.Recordset.AbsolutePosition
    While pos < 0
        Data1.Recordset.MoveNext
    Wend
    If Grid1.ColSel = 0 Then
        Unit_1 = TEST_Y
    ElseIf Grid1.ColSel = 1 Then
        Date_1 = TEST_Y
    ElseIf Grid1.ColSel = 2 Then
        TEST_1 = CDbl(TEST_Y)
    End If
    Do Until Data1.Recordset.EOF
        n_rec_a = Data1.Recordset.AbsolutePosition + 1
        Date_2 = Data1.Recordset.date_test
        Unit_2 = Data1.Recordset.Unit_name
        TEST_2 = Data1.Recordset.TEST_NO
        If Grid1.ColSel = 0 Then
            If Unit_1 = Unit_2 Then
                XXX = 0
                If n_rec_a = 1 Then
                    Data1.Recordset.MoveNext
                    Data1.Recordset.MovePrevious
                Else
                    Data1.Recordset.MovePrevious
                    Data1.Recordset.MoveNext
                End If
                GoTo 10
            End If
        ElseIf Grid1.ColSel = 1 Then
            If Date_1 = Date_2 Then
                XXX = 0
                If n_rec_a = 1 Then
                    Data1.Recordset.MoveNext
                    Data1.Recordset.MovePrevious
                Else
                    Data1.Recordset.MovePrevious
                    Data1.Recordset.MoveNext
                End If
                GoTo 10
            End If
        ElseIf Grid1.ColSel = 2 Then
            If TEST_1 = TEST_2 Then
                XXX = 0
                If n_rec_a = 1 Then
                    Data1.Recordset.MoveNext
                    Data1.Recordset.MovePrevious
                Else
                    Data1.Recordset.MovePrevious
                    Data1.Recordset.MoveNext
                End If
                GoTo 10
            End If
        End If
        If n_rec_a = n_record Then
            MsgBox "Selected RECORD not found in the database"
            GoTo 10
        End If
        Data1.Recordset.MoveNext
        Bar1 = Bar1 + Bar1
        If Bar1 >= Bar1.Max Then
            Bar1 = 5
        End If
    Loop
End If
10  XXX = 0
Bar1 = 0
Frame_Search.Visible = False
Call Fluid_type
End Sub
Private Sub Data1_Error(DataErr As Integer, Response As Integer)
On Error Resume Next
  ER = DataErr
  MsgBox "Intercettato errore dei dati:" & Error$(DataErr)
  Response = 0  'Ignora l'errore
End Sub
Function Divide(numer, denom) As Variant
   Dim Msg As String
   Const mnErrDivByZero = 11, mnErrOverFlow = 6
   Const mnErrBadCall = 5
   On Error GoTo MathHandler
      Divide = numer / denom
      Exit Function
MathHandler:
   If Err.Number = mnErrDivByZero Or Err.Number = ErrOverFlow Or Err = ErrBadCall Then
      Divide = Null      ' In caso di un errore Divisione per zero,
                        ' Overflow o Chiamata di routine non valida,
                        ' restituisce Null.
   Else
      ' Visualizza il messaggio di errore non previsto.
      MsgBox "Errore non previsto " & Err.Number
      Msg = Msg & ": " & Err.Description
      MsgBox Msg, vbExclamation
   End If               ' In tutti i casi, Resume Next continua
                        ' l'esecuzione in corrispondenza
   Resume Next            ' dell'istruzione Exit Function.
End Function
Private Sub Data1_Reposition()
On Error Resume Next

If XXX = 1 Then
    Exit Sub
End If
    Screen.MousePointer = vbDefault
    Checkrec = Data1.Recordset.AbsolutePosition + 1
    Data1.Caption = Data1.Recordset("Unit_Name")
    
    NUM_REC = Data1.Recordset.RecordCount
    txt_num.Text = Data1.Recordset.AbsolutePosition + 1
    Combo_PLANT.Text = Data1.Recordset.Plant
    Combo_LOC.Text = Data1.Recordset.Location
    Combo_Country.Text = Data1.Recordset.Country
    Combo_UNIT.Text = Data1.Recordset.Unit_name
    Combo_PLANT_UNIT = Data1.Recordset.PLANT_UNIT
    Combo_PROCESS_DESCR = Data1.Recordset.PROCESS_DESCR
    Combo_PROCESS_STREAM = Data1.Recordset.PROCESS_STREAM
    Combo_TEMA = Data1.Recordset.TEMA
    Combo_POSITION = Data1.Recordset.Position
    Combo_cooling_type = Data1.Recordset.COOLING_TYPE
    PARALLEL_N = Data1.Recordset.PARALLEL_N
    SERIES_N = Data1.Recordset.SERIES_N
    Combo_CURRENT = Data1.Recordset.Current
    
    T_NO.Text = Data1.Recordset.TUBES_NO
    T_len.Text = Data1.Recordset.TUBES_LE
    T_PASS.Text = Data1.Recordset.TUBES_PASSES
    T_OD.Text = Data1.Recordset.TUBES_OD
    Combo_BWG.Text = Data1.Recordset.TUBES_BWG
    Combo_TUBES_Mat.Text = Data1.Recordset.TUBES_MAT
    D54 = Data1.Recordset.Mat_fact
    Spin_MAT_FACTOR = D54 * 100
    
    U = Data1.Recordset.Check_X
    If U = 0 Then
        lungh = 1
        Check_U = Unchecked
    ElseIf U = -1 Then
        lungh = 2
        Check_U = Checked
    End If
    If RichTextBox_REMARKS.Text <> "" Then
        Toggle_remarks.BackColor = &HFF&
        Toggle_remarks.ForeColor = &HFFFF&
    Else
        Toggle_remarks.BackColor = &HC0FFC0
        Toggle_remarks.ForeColor = &HC0&
    End If
    
    SHELL_PASS.Text = Data1.Recordset.SHELL_PASSES
    SHELL_BAFFLES_N.Text = Data1.Recordset.SHELL_BAFFLES_N
    SHELL_BAFFLES_CUT.Text = Data1.Recordset.SHELL_BAFFLES_CUT
    SHELL_BAFFLES_SPACE.Text = Data1.Recordset.SHELL_BAFFLES_SPACE
    SHELL_ID.Text = Data1.Recordset.SHELL_ID
    SHELL_TUBES_PITCH.Text = Data1.Recordset.SHELL_TUBES_PITCH
    SHELL_PITCH_CONF.Text = Data1.Recordset.SHELL_PITCH_CONF
    Combo_SHELL_MAT.Text = Data1.Recordset.SHELL_MAT
    
    FACTOR = Data1.Recordset.FACT_FLOW
    FACT_FLOW.Text = FACTOR
    Spin_FACT_FLOW.Value = FACTOR
    
    Combo_T_FLUID.Text = Data1.Recordset.TUBES_FLUID
    TUBES_FLOW_1 = Data1.Recordset.TUBES_FLOW
    TUBES_FLOW.Text = Format(TUBES_FLOW_1, "#,##0")
    TUBES_VAPOR_1 = Data1.Recordset.TUBES_VAPOR
    TUBES_VAPOR.Text = Format(TUBES_VAPOR_1, "#,##0")
    TUBES_LIQUID_1 = Data1.Recordset.TUBES_LIQUID
    TUBES_LIQUID.Text = Format(TUBES_LIQUID_1, "#,##0")
    TUBES_WATER_1 = Data1.Recordset.TUBES_WATER
    TUBES_WATER.Text = Format(TUBES_WATER_1, "#,##0")
    TUBES_NON_COND_1 = Data1.Recordset.TUBES_NON_COND
    TUBES_NON_COND.Text = Format(TUBES_NON_COND_1, "#,##0")
    TUBES_TEMP_IN.Text = Data1.Recordset.TUBES_TEMP_IN
    TUBES_TEMP_OUT.Text = Data1.Recordset.TUBES_TEMP_OUT
    TUBES_P_IN.Text = Data1.Recordset.TUBES_P_IN
    TUBES_P_OUT.Text = Data1.Recordset.TUBES_P_OUT
    TUBES_OUT(1).Text = Data1.Recordset.TUBES_T_COND
    TUBES_OUT(3).Text = Data1.Recordset.TUBES_SPH
    TUBES_OUT(4).Text = Data1.Recordset.TUBES_DENS
    TUBES_OUT(5).Text = Data1.Recordset.TUBES_VISC
    
    Combo_S_FLUID.Text = Data1.Recordset.SHELL_FLUID
    SHELL_FLOW_1 = Data1.Recordset.SHELL_FLOW
    SHELL_FLOW.Text = Format(SHELL_FLOW_1, "#,##0")
    SHELL_VAPOR_1 = Data1.Recordset.SHELL_VAPOR
    SHELL_VAPOR.Text = Format(SHELL_VAPOR_1, "#,##0")
    SHELL_LIQUID_1 = Data1.Recordset.SHELL_LIQUID
    SHELL_LIQUID.Text = Format(SHELL_LIQUID_1, "#,##0")
    SHELL_WATER_1 = Data1.Recordset.SHELL_WATER
    SHELL_WATER.Text = Format(SHELL_WATER_1, "#,##0")
    SHELL_NON_COND_1 = Data1.Recordset.SHELL_NON_COND
    SHELL_NON_COND.Text = Format(SHELL_NON_COND_1, "#,##0")
    SHELL_TEMP_IN.Text = Data1.Recordset.SHELL_TEMP_IN
    SHELL_TEMP_OUT.Text = Data1.Recordset.SHELL_TEMP_OUT
    SHELL_P_IN.Text = Data1.Recordset.SHELL_P_IN
    SHELL_P_OUT.Text = Data1.Recordset.SHELL_P_OUT
    S_press_KP_1 = Data1.Recordset.Press_COND
    S_press_KP.Text = Format(S_press_KP_1, "0.00")
    
    YXY = 1
    Spin_S_PRESS = S_press_KP_1 ^ 100
    SHELL_OUT(1) = Data1.Recordset.SHELL_T_COND
    SHELL_OUT(3) = Data1.Recordset.SHELL_SPH
    SHELL_OUT(4) = Data1.Recordset.SHELL_DENS
    SHELL_OUT(5) = Data1.Recordset.SHELL_VISC
    SHELL_OUT(11) = Data1.Recordset.Temp_COND
    SHELL_OUT(10) = Data1.Recordset.SHELL_LATENT

    TUBES_T_IN = Format(TUBES_TEMP_IN, "0.00")
    TUBES_T_OUT = Format(TUBES_TEMP_OUT, "0.00")
    Spin_TUBES_T_IN.Value = TUBES_T_IN * 100
    Spin_TUBES_T_OUT.Value = TUBES_TEMP_OUT * 100
    HScroll_SHELL_T_IN.Value = SHELL_TEMP_IN * 100
    HScroll_SHELL_T_OUT.Value = SHELL_TEMP_OUT * 100
    
    HScroll_TUBES_VAPOR.Max = TUBES_FLOW / FACTOR
    HScroll_TUBES_LIQUID.Max = TUBES_FLOW / FACTOR
    HScroll_TUBES_WATER.Max = TUBES_FLOW / FACTOR
    HScroll_TUBES_NON_COND.Max = TUBES_FLOW / FACTOR
    HScroll_TUBES_FLOW.Value = TUBES_FLOW / FACTOR
    HScroll_TUBES_VAPOR.Value = TUBES_VAPOR / FACTOR
    HScroll_TUBES_LIQUID.Value = TUBES_LIQUID / FACTOR
    HScroll_TUBES_WATER.Value = TUBES_WATER / FACTOR
    HScroll_TUBES_NON_COND.Value = TUBES_NON_COND / FACTOR
    TUBES_TC = TUBES_OUT(1)
    HScroll_TUBES_TC = TUBES_TC * 1000
    HScroll_SHELL_NON_COND.Max = SHELL_FLOW / FACTOR
    Spin_SHELL_FLOW.Value = SHELL_FLOW / FACTOR
    HScroll_SHELL_NON_COND.Value = SHELL_NON_COND / FACTOR
    
    Vtot_INP_1 = Data1.Recordset.VAP_FRACTION
    Vtot_INP.Text = Format(Vtot_INP_1, "0.0")
    Vwat_INP_1 = Data1.Recordset.Vwat_perc
    Vwat_INP.Text = Format(Vwat_INP_1, "0.0")
    Ltot_INP_1 = Data1.Recordset.LIQ_FRACTION
    Ltot_INP = Format(Ltot_INP_1, "0.0")
    Lwat_INP_1 = Data1.Recordset.Lwat_perc
    Lwat_INP.Text = Format(Lwat_INP_1, "0.0")
    YXY = 1
    Spin_VAP_P.Value = CDbl(Vtot_INP) * 10
    Spin_WAT_VAP_IN = CDbl(Vwat_INP) * 10
    Spin_WAT_LIQ_IN = CDbl(Lwat_INP) * 10
    
    Spin_PARALLEL_N.Value = PARALLEL_N
    Spin_SERIES_N.Value = SERIES_N
    HScroll_T_NO.Value = T_NO.Text
    Spin_TUBES_PITCH.Value = SHELL_TUBES_PITCH * 10
    Spin_S_PASS.Value = SHELL_PASS
    Spin_BAFFLES_N.Value = SHELL_BAFFLES_N
    Spin_BAFFLES_CUT.Value = SHELL_BAFFLES_CUT
    Spin_BAFFLES_SPACE.Value = SHELL_BAFFLES_SPACE
    Spin_SHELL_ID.Value = SHELL_ID
    Spin_TUBES_PITCH.Value = SHELL_TUBES_PITCH * 10
    Spin_T_LEN.Value = T_len.Text * 100
    Spin_T_PAS.Value = T_PASS.Text
    Spin_T_OD.Value = T_OD.Text * 100
    
    HScroll_TUBES_SPH = TUBES_OUT(3) * 1000
    HScroll_TUBES_DENS = TUBES_OUT(4) * 10
    HScroll_TUBES_VISC = TUBES_OUT(5) * 1000
    HScroll_TUBES_P_IN.Value = TUBES_P_IN * 100
    HScroll_TUBES_P_OUT.Value = TUBES_P_OUT * 100

    HScroll_SHELL_SPH = SHELL_OUT(3) * 1000
    HScroll_SHELL_DENS = SHELL_OUT(4) * 10
    HScroll_SHELL_VISC = SHELL_OUT(5) * 1000
    HScroll_SHELL_T_IN.Value = SHELL_TEMP_IN * 100
    HScroll_SHELL_T_OUT.Value = SHELL_TEMP_OUT * 100
    HScroll_SHELL_P_IN.Value = SHELL_P_IN * 100
    HScroll_SHELL_P_OUT.Value = SHELL_P_OUT * 100
    Spin_S_PRESS.Value = S_press_KP * 100
    HScroll_SHELL_TC = SHELL_OUT(1) * 1000
    SHELL_OUT(9).Text = Data1.Recordset.SHELL_PRESS_DROP
    HScroll_P_DROP_S = SHELL_OUT(9) * 10
    HScroll_C_TEMP = SHELL_OUT(11) * 100
    HScroll_LATENT.Value = SHELL_OUT(10)
    YXY = 0
    PROCESS_TARGET_T_OUT = Data1.Recordset.PROCESS_TARGET_TEMP
    Spin_TARGET_T = PROCESS_TARGET_T_OUT * 10
    XXX = 1
    
    FFX = Data1.Recordset.SHELL_FF
    SHELL_FF = Format(FFX, "0.00")
    D40_S = FFX / 10000
    W_FF = Data1.Recordset.WATER_FF
    WATER_FF = Format(W_FF, "0.00")
    If Combo_S_FLUID = "Water" And Combo_CURRENT = "Condensation" Then
        MsgBox ("    Tubes-side Condensation cannot be set ")
        Combo_CURRENT = "Counter-flow"
        Exit Sub
    End If
    Call Fluid_type
    foul = 0
    XXX = 0
    YXY = 0
100 End Sub
Private Sub Data1_Validate(Action As Integer, Save As Integer)
On Error Resume Next
  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
  End Select
  Screen.MousePointer = vbDefault
'  XXX = 1
End Sub
Private Sub Check_des_Click()
On Error Resume Next
    If XXX = 1 Then
        Exit Sub
    End If
    If Check_des = Checked Then
        CHECK_DESIGN.BackColor = &HFF&
        Label4.Visible = False
        CHECK_ACTUAL.Visible = False
        Check_ACT.Visible = False
        Line2.Visible = False
        Line5.Visible = False
        Comm_check_reset.Visible = False
        Check_T_OUT.Visible = False
    ElseIf Check_des = Unchecked Then
        CHECK_DESIGN.BackColor = &HE0E0E0
        Label4.Visible = True
        CHECK_ACTUAL.Visible = True
        Check_ACT.Visible = True
        Line2.Visible = True
        Line5.Visible = True
        Comm_check_reset.Visible = True
    End If
End Sub
Private Sub Check_act_Click()
On Error Resume Next
    If XXX = 1 Then
        Exit Sub
    End If
    If Check_ACT = Checked Then
        CHECK_ACTUAL.BackColor = &HFF&
        Label35.Visible = False
        CHECK_DESIGN.Visible = False
        Check_des.Visible = False
    ElseIf Check_ACT = Unchecked Then
        CHECK_ACTUAL.BackColor = &HE0E0E0
        Label35.Visible = True
        CHECK_DESIGN.Visible = True
        Check_des.Visible = True
    End If
End Sub
Private Sub Check_U_Click()
On Error Resume Next
    If Check_U = Checked Then
        lungh = 2
    Else
        lungh = 1
    End If
    If XXX = 1 Then
        Exit Sub
    End If
Call Fluid_type
End Sub
Private Sub Comm_check_reset_Click()
On Error Resume Next
    Call ACTUAL_CHECK_RESET
End Sub
Private Sub Combo_Plant_LostFocus()
On Error Resume Next
    Plant.Text = Combo_PLANT.Text
End Sub
Private Sub Combo_Country_LostFocus()
On Error Resume Next
    Country.Text = Combo_Country.Text
End Sub
Private Sub Combo_LOC_LostFocus()
On Error Resume Next
    Location.Text = Combo_LOC.Text
End Sub
Private Sub Combo_UNIT_LostFocus()
On Error Resume Next
    Unit.Text = Combo_UNIT.Text
End Sub
Private Sub Combo_CURRENT_LostFocus()
On Error Resume Next
    If Combo_CURRENT = "Condensation" Then
        Thermal_bal_shell_T.Visible = False
    End If
    If Combo_CURRENT = "Condensation" And Combo_cooling_type = "Cooling" Then
        Combo_cooling_type = "Condensation"
    End If
    If Combo_S_FLUID = "Water" Then
        If Combo_CURRENT = "Condensation" And Combo_S_FLUID = "Water" Then
            MsgBox ("Tubes-side Condensation cannot be set")
            Combo_CURRENT = "Counter-flow"
            Exit Sub
        End If
    End If
Call Fluid_type
End Sub
Private Sub HScroll_LATENT_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    If Check_LATENT = Checked Then
        SHELL_OUT(10) = HScroll_LATENT
    End If
Call Fluid_type
End Sub
Private Sub Check_LATENT_Click()
On Error Resume Next
    If XXX = 1 Then
        Exit Sub
    End If
    If Check_LATENT = Checked Then
        SHELL_OUT(10) = HScroll_LATENT
    End If
Call Fluid_type
End Sub
Private Sub HScroll_WATER_FF_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    W_FF = HScroll_WATER_FF / 100
    WATER_FF = Format(W_FF, "0.00")
End Sub
Private Sub Spin_FACT_FLOW_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    FACT_FLOW.Text = Spin_FACT_FLOW.Value
End Sub
Private Sub HScroll_TUBES_FLOW_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    T_FLOW = HScroll_TUBES_FLOW
    TUBES_FLOW = Format(T_FLOW * FACT_FLOW, "#,##0")
    TUBES_OUT(0) = Format(TUBES_FLOW, "#,##0")
    
    HScroll_TUBES_VAPOR.Max = TUBES_FLOW / FACT_FLOW
    HScroll_TUBES_LIQUID.Max = TUBES_FLOW / FACT_FLOW
    HScroll_TUBES_WATER.Max = TUBES_FLOW / FACT_FLOW
    HScroll_TUBES_NON_COND.Max = TUBES_FLOW / FACT_FLOW
    
    TUBES_VAPOR = Data1.Recordset.TUBES_VAPOR
    TUBES_LIQUID = Data1.Recordset.TUBES_LIQUID
    TUBES_WATER = Data1.Recordset.TUBES_WATER
    TUBES_NON_COND = Data1.Recordset.TUBES_NON_COND
    
    HScroll_TUBES_VAPOR.Value = TUBES_VAPOR / FACT_FLOW
    HScroll_TUBES_LIQUID.Value = TUBES_LIQUID / FACT_FLOW
    HScroll_TUBES_WATER.Value = TUBES_WATER / FACT_FLOW
    HScroll_TUBES_NON_COND.Value = TUBES_NON_COND / FACT_FLOW
    
    TUBES_WATER = Format(TUBES_FLOW, "#,##0")
    TUBES_LIQUID = Format(TUBES_FLOW - TUBES_VAPOR - TUBES_NON_COND - TUBES_WATER, "#,##0")
    T_LIQUID = (TUBES_LIQUID + TUBES_WATER) / FACT_FLOW
    HScroll_TUBES_LIQUID.Value = T_LIQUID
    HScroll_TUBES_WATER.Max = T_LIQUID
Call Fluid_type
End Sub
Private Sub HScroll_TUBES_VAPOR_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    T_VAPOR = HScroll_TUBES_VAPOR
    TUBES_VAPOR = Format(T_VAPOR * FACT_FLOW, "#,##0")
    TUBES_WATER = Format(TUBES_FLOW, "#,##0")
    TUBES_VAPOR = Format(0, "#,##0")
    TUBES_NON_COND = Format(0, "#,##0")
    TUBES_LIQUID = Format(TUBES_FLOW - TUBES_VAPOR - TUBES_NON_COND - TUBES_WATER, "#,##0")
    T_LIQUID = (TUBES_LIQUID + TUBES_WATER) / FACT_FLOW
    HScroll_TUBES_LIQUID.Value = T_LIQUID
    HScroll_TUBES_WATER.Max = T_LIQUID
    If COL = 0 Then
        S_FLOW = 0
        SHELL_OUT(0).BackColor = &HE0E0E0
        SHELL_OUT(0).ForeColor = &HC0&
    End If
    If COL = 0 Then
        T_FLOW = 0
        TUBES_OUT(0).BackColor = &HE0E0E0
        TUBES_OUT(0).ForeColor = &HC0&
    End If
Call Fluid_type
End Sub
Private Sub HScroll_TUBES_NON_COND_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    T_NON_COND = HScroll_TUBES_NON_COND
    TUBES_NON_COND = Format(T_NON_COND * FACT_FLOW, "#,##0")
    If Combo_T_FLUID = "Water" Then
        TUBES_WATER = Format(TUBES_FLOW, "#,##0")
        TUBES_VAPOR = Format(0, "#,##0")
        TUBES_NON_COND = Format(0, "#,##0")
    End If
    TUBES_LIQUID = Format(TUBES_FLOW - TUBES_VAPOR - TUBES_NON_COND - TUBES_WATER, "#,##0")
    T_LIQUID = (TUBES_LIQUID + TUBES_WATER) / FACT_FLOW
    HScroll_TUBES_LIQUID.Value = T_LIQUID
    HScroll_TUBES_WATER.Max = T_LIQUID
    If COL = 0 Then
        T_FLOW = 0
        TUBES_OUT(0).BackColor = &HE0E0E0
        TUBES_OUT(0).ForeColor = &HC0&
    End If
Call Fluid_type
    If COL = 0 Then
        S_FLOW = 0
        SHELL_OUT(0).BackColor = &HE0E0E0
        SHELL_OUT(0).ForeColor = &HC0&
    End If
End Sub
Private Sub HScroll_TUBES_WATER_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    T_WATER = HScroll_TUBES_WATER.Value
    TUBES_WATER = Format(T_WATER * FACT_FLOW, "#,##0")
End Sub
Private Sub Spin_SHELL_FLOW_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    SHELL_FLOW_1 = Spin_SHELL_FLOW.Value
    SHELL_FLOW = Format(SHELL_FLOW_1 * FACT_FLOW, "#,##0")
    SHELL_OUT(0) = Format(SHELL_FLOW, "#,##0")
Call Fluid_type
    AA = XD37 / 860
    If COL = 0 Then
        S_FLOW = 0
        SHELL_OUT(0).BackColor = &HE0E0E0
        SHELL_OUT(0).ForeColor = &HC0&
    End If
End Sub
Private Sub Spin_WAT_VAP_IN_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    Vwat_INP_1 = Spin_WAT_VAP_IN.Value / 10
        Vwat_INP = Format(CDbl(Vwat_INP_1), "0.0")
        Vwat_IN_1 = Vwat_INP_1 / 100 * Vtot_1
        Vwat_IN = Format(Vwat_IN_1, "#,##0")
        Vorg_IN_1 = Vtot_IN_1 - NC_IN_1 - Vwat_IN_1
        Vorg_IN = Format(Vorg_IN_1, "#,##0")
        Vorg_INP_1 = 100 - Vwat_INP_1
        Vorg_INP = Format(Vorg_INP_1, "0.0")
Call Fluid_type
End Sub
Private Sub Spin_VAP_P_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    Vtot_INP_1 = Spin_VAP_P.Value / 10
        Vtot_INP = Format(Vtot_INP_1, "0.0")
Call Fluid_type
End Sub
Private Sub Spin_WAT_LIQ_IN_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    Lwat_INP_1 = Format(Spin_WAT_LIQ_IN / 10, "0.0")
        Lwat_INP = Format(Lwat_INP_1, "0.0")
        Lwat_IN_1 = Format(Lwat_INP_1 / 100 * Ltot_1, "0.0")
        Lwat_IN = Format(Lwat_IN_1, "#,##0")
        Lorg_IN_1 = Ltot_IN_1 - Lwat_IN_1
        Lorg_IN = Format(Lorg_IN_1, "#,##0")
        Lorg_INP_1 = 100 - Lwat_INP_1
        Lorg_INP = Format(Lorg_INP_1, "0.0")
Call Fluid_type
End Sub
Private Sub HScroll_SHELL_NON_COND_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    SHELL_NON_COND_1 = HScroll_SHELL_NON_COND * FACT_FLOW
    SHELL_NON_COND = Format(SHELL_NON_COND_1, "#,##0")
    If COL = 0 Then
        S_FLOW = 0
        SHELL_OUT(0).BackColor = &HE0E0E0
        SHELL_OUT(0).ForeColor = &HC0&
    End If
Call Fluid_type
    If COL = 0 Then
        S_FLOW = 0
        SHELL_OUT(0).BackColor = &HE0E0E0
        SHELL_OUT(0).ForeColor = &HC0&
    End If
End Sub
Private Sub Spin_TUBES_T_IN_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    TUBES_TEMP_IN_1 = Spin_TUBES_T_IN / 100
    TUBES_TEMP_IN.Text = Format(TUBES_TEMP_IN_1, "0.00")
Call Fluid_type
End Sub
Private Sub Spin_TUBES_T_OUT_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    TUBES_TEMP_OUT_1 = Spin_TUBES_T_OUT / 100
    TUBES_TEMP_OUT = Format(TUBES_TEMP_OUT_1, "0.00")
Call Fluid_type
End Sub
Private Sub HScroll_TUBES_P_IN_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    TUBES_P_IN = HScroll_TUBES_P_IN / 100
Call Fluid_type
End Sub
Private Sub HScroll_TUBES_P_OUT_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    TUBES_P_OUT = HScroll_TUBES_P_OUT / 100
Call Fluid_type
End Sub
Private Sub HScroll_SHELL_T_IN_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    SHELL_TEMP_IN_1 = HScroll_SHELL_T_IN / 100
    SHELL_TEMP_IN = Format(SHELL_TEMP_IN_1, "0.00")
Call Fluid_type
End Sub
Private Sub HScroll_SHELL_T_OUT_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    SHELL_TEMP_OUT_1 = HScroll_SHELL_T_OUT / 100
    SHELL_TEMP_OUT = Format(SHELL_TEMP_OUT_1, "0.00")
    If CSng(SHELL_TEMP_OUT) > CSng(PROCESS_TARGET_T_OUT) Then
        SHELL_TEMP_OUT.ForeColor = &H80000018
        SHELL_TEMP_OUT.BackColor = 255.255
    Else
        SHELL_TEMP_OUT.BackColor = &H80000018
        SHELL_TEMP_OUT.ForeColor = &HC0&
    End If
Call Fluid_type
End Sub
Private Sub HScroll_SHELL_P_IN_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    SHELL_P_IN = HScroll_SHELL_P_IN / 100
Call Fluid_type
End Sub
Private Sub HScroll_SHELL_P_OUT_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    SHELL_P_OUT = HScroll_SHELL_P_OUT / 100
Call Fluid_type
End Sub
Private Sub Spin_TUBES_PITCH_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    SHELL_TUBES_PITCH = Spin_TUBES_PITCH / 10
Call Fluid_type
End Sub
Private Sub Spin_BAFFLES_CUT_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    SHELL_BAFFLES_CUT = Spin_BAFFLES_CUT
Call Fluid_type
End Sub
Private Sub Spin_BAFFLES_N_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    If CHECK_BAFFLES_N = Unchecked Then
        SHELL_BAFFLES_N = Int(XD55 / (XD64M / 1000))
    Else
        SHELL_BAFFLES_N = Spin_BAFFLES_N
    End If
Call Fluid_type
End Sub
Private Sub CHECK_BAFFLES_N_Click()
On Error Resume Next
    If XXX = 1 Then
        Exit Sub
    End If
    If CHECK_BAFFLES_N = Unchecked Then
        SHELL_BAFFLES_N = Int(XD55 / (XD64M / 1000))
    Else
         SHELL_BAFFLES_N = Spin_BAFFLES_N
    End If
Call Fluid_type
End Sub
Private Sub Spin_BAFFLES_SPACE_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    SHELL_BAFFLES_SPACE = Spin_BAFFLES_SPACE
    If CHECK_BAFFLES_N = Unchecked Then
        XD64M = SHELL_BAFFLES_SPACE
        SHELL_BAFFLES_N = Int(XD55 / (XD64M / 1000))
        Spin_BAFFLES_N = SHELL_BAFFLES_N
    Else
        SHELL_BAFFLES_N = Spin_BAFFLES_N
    End If
Call Fluid_type
End Sub
Private Sub Spin_PARALLEL_N_Change()
On Error Resume Next
    PARALLEL_N.Text = Spin_PARALLEL_N
Call Fluid_type
End Sub
Private Sub Spin_S_PASS_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    SHELL_PASS = Spin_S_PASS
Call Fluid_type
End Sub
Private Sub Spin_SERIES_N_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    SERIES_N.Text = Spin_SERIES_N
Call Fluid_type
End Sub
Private Sub HScroll_TUBES_SPH_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    If Check_T_SPH = Checked Then
        TUBES_OUT(3) = Format(HScroll_TUBES_SPH / 1000, "0.000")
    End If
Call Fluid_type
End Sub
Private Sub HScroll_TUBES_TC_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    If Check_T_TC = Checked Then
        TUBES_OUT(1) = HScroll_TUBES_TC / 1000
    End If
Call Fluid_type
End Sub
Private Sub Check_T_TC_Click()
On Error Resume Next
    If Check_T_TC = Checked Then
        TUBES_OUT(1) = HScroll_TUBES_TC / 1000
    End If
Call Fluid_type
End Sub
Private Sub HScroll_SHELL_TC_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    If Check_S_TC = Checked Then
        SHELL_OUT_1 = HScroll_SHELL_TC / 1000
        SHELL_OUT(1) = Format(SHELL_OUT_1, "0.000")
    End If
Call Fluid_type
End Sub
Private Sub Check_S_TC_Click()
On Error Resume Next
    If XXX = 1 Then
        Exit Sub
    End If
    If Check_S_TC = Checked Then
        SHELL_OUT_1 = HScroll_SHELL_TC / 1000
        SHELL_OUT(1) = Format(SHELL_OUT_1, "0.000")
    End If
Call Fluid_type
End Sub
Private Sub Check_T_SPH_Click()
On Error Resume Next
    If Check_T_SPH = Checked Then
        TUBES_OUT(3) = HScroll_TUBES_SPH / 1000
        SPH_T = TUBES_OUT(3)
    Else
        'TUBES_OUT(3) = Data1.Recordset.TUBES_SPH
        SPH_T = TUBES_OUT(3)
    End If
Call Fluid_type
End Sub
Private Sub HScroll_SHELL_SPH_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    If Check_S_SPH = Checked Then
        SHELL_OUT(3) = HScroll_SHELL_SPH / 1000
        SPH_S = SHELL_OUT(3)
    End If
Call Fluid_type
End Sub
Private Sub Check_S_SPH_Click()
On Error Resume Next
    If XXX = 1 Then
        Exit Sub
    End If
    If Check_S_SPH = Checked Then
        SHELL_OUT(3) = HScroll_SHELL_SPH / 1000
        SPH_S = SHELL_OUT(3)
    End If
Call Fluid_type
End Sub
Private Sub HScroll_TUBES_DENS_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    If Check_T_DENS = Checked Then
        TUBES_OUT(4) = Format(HScroll_TUBES_DENS / 10, "0.0")
    End If
Call Fluid_type
End Sub
Private Sub Check_T_DENS_Click()
On Error Resume Next
    If Check_T_DENS = Checked Then
        TUBES_OUT(4) = Format(HScroll_TUBES_DENS / 10, "0.0")
    End If
Call Fluid_type
End Sub
Private Sub HScroll_SHELL_DENS_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    If Check_S_DENS = Checked Then
        SHELL_OUT(4) = HScroll_SHELL_DENS / 10
    End If
Call Fluid_type
End Sub
Private Sub Check_S_DENS_Click()
On Error Resume Next
    If XXX = 1 Then
        Exit Sub
    End If
    If Check_S_DENS = Checked Then
        SHELL_OUT(4) = HScroll_SHELL_DENS / 10
    End If
Call Fluid_type
End Sub
Private Sub HScroll_TUBES_VISC_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    If Check_T_VISC = Checked Then
        TUBES_OUT(5) = HScroll_TUBES_VISC / 1000
    End If
Call Fluid_type
End Sub
Private Sub Check_T_VISC_Click()
On Error Resume Next
    If Check_T_VISC = Checked Then
        TUBES_OUT(5) = HScroll_TUBES_VISC / 1000
    End If
Call Fluid_type
End Sub
Private Sub HScroll_SHELL_VISC_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    If Check_S_VISC = Checked Then
        SHELL_OUT(5) = HScroll_SHELL_VISC / 1000
Call Fluid_type
    End If
End Sub
Private Sub Check_S_VISC_Click()
On Error Resume Next
    If XXX = 1 Then
        Exit Sub
    End If
    If Check_S_VISC = Checked Then
        SHELL_OUT(5) = HScroll_SHELL_VISC / 1000
    End If
Call Fluid_type
End Sub
Private Sub HScroll_U_CLEAN_Change()
On Error Resume Next
    If XXX = 1 Then
        Exit Sub
    End If
    If Check_U_CLEAN = Checked Then
        XD117 = HScroll_U_CLEAN
        U_COEFF_CLEAN = XD117
    Else
        U_COEFF_CLEAN = Data1.Recordset.Clean
    End If
Call Fluid_type
End Sub
Private Sub Check_U_CLEAN_Click()
On Error Resume Next
    If XXX = 1 Then
        Exit Sub
    End If
    If Check_U_CLEAN = Checked Then
        XD117 = HScroll_U_CLEAN
        U_COEFF_CLEAN = XD117
    Else
        U_COEFF_CLEAN = Data1.Recordset.Clean
    End If
Call Fluid_type
End Sub
Private Sub Check_P_DROP_T_Click()
On Error Resume Next
    If XXX = 1 Then
        Exit Sub
    End If
    If Check_P_DROP_T = Checked Then
        Label22(12).Visible = True
    Else
        Label22(12).Visible = False
    End If
Call Fluid_type
End Sub
Private Sub Check_P_DROP_S_Click()
On Error Resume Next
    If XXX = 1 Then
        Exit Sub
    End If
    If Check_P_DROP_S = Checked Then
        Label22(14).Visible = True
    Else
        Label22(14).Visible = False
    End If
Call Fluid_type
End Sub
Private Sub Combo_S_FLUID_lostfocus()
On Error Resume Next
    If Combo_S_FLUID = "Water" And Combo_CURRENT = "Condensation" Then
        MsgBox ("    Tubes-side Condensation cannot be set ")
        Combo_CURRENT = "Counter-flow"
        Exit Sub
    End If
Call Fluid_type
End Sub
Private Sub Combo_T_FLUID_lostfocus()
On Error Resume Next
Call Fluid_type
End Sub
Private Sub Combo_TUBES_Mat_LostFocus()
On Error Resume Next
    metal = Combo_TUBES_Mat.Text
Call Fluid_type
End Sub
Private Sub Check_PF_Click()
On Error Resume Next
    If XXX = 1 Then
        Exit Sub
    End If
    If Check_PF = 0 Then
        FFX = Data1.Recordset.SHELL_FF
        SHELL_FF = Format(FFX, "0.00")
    ElseIf Check_PF.Value = 1 Then
        FFX = Spin_PF.Value / 100
        SHELL_FF = Format(FFX, "0.00")
    End If
    D40_S = FFX / 10000
Call Fluid_type
End Sub
Private Sub Spin_PF_Change()
On Error Resume Next
    If Check_PF.Value = 1 Then
        FFX = Spin_PF.Value / 100
        SHELL_FF = Format(FFX, "0.00")
    End If
    D40_S = FFX / 10000
Call Fluid_type
End Sub
Private Sub Spin_S_FLOW_IN_Change()
On Error Resume Next
    SFIN_D = Spin_S_FLOW_IN
    S_flow_IN.Text = Format(SFIN_D * 100, "0.00")
Call Fluid_type
End Sub
Private Sub Spin_S_PRESS_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    If Check_CP = Checked Then
        S_press_KP_1 = Spin_S_PRESS / 100
        S_press_KP.Text = Format(S_press_KP_1, "0.00")
    End If
Call Fluid_type
End Sub
Private Sub Check_CP_Click()
On Error Resume Next
    If XXX = 1 Then
        Exit Sub
    End If
    If Check_CP = Checked Then
        S_press_KP_1 = Spin_S_PRESS / 100
        S_press_KP.Text = Format(S_press_KP_1, "0.00")
    End If
Call Fluid_type
End Sub
Private Sub HScroll_C_TEMP_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    If Check_CT = Checked Then
        SHELL_OUT(11) = Format(HScroll_C_TEMP / 100, "0.00")
    End If
Call Fluid_type
End Sub
Private Sub Check_CT_Click()
On Error Resume Next
    If Check_CT = Checked Then
        SHELL_OUT(11) = Format(HScroll_C_TEMP / 100, "0.00")
    End If
Call Fluid_type
End Sub
Private Sub HScroll_T_NO_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    T_NO.Text = HScroll_T_NO.Value
Call Fluid_type
End Sub
Private Sub Spin_SHELL_ID_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    SHELL_ID = Spin_SHELL_ID
Call Fluid_type
End Sub
Private Sub Spin_T_LEN_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    T_len.Text = Spin_T_LEN.Value / 100
Call Fluid_type
End Sub
Private Sub Spin_T_PAS_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    T_PASS.Text = Spin_T_PAS.Value
Call Fluid_type
End Sub
Private Sub Spin_T_OD_Change()
On Error Resume Next
    If YXY = 1 Then
        Exit Sub
    End If
    T_OD = Spin_T_OD / 100
Call Fluid_type
End Sub
Private Sub Spin_MAT_FACTOR_Change()
On Error Resume Next
    If XXX = 1 Then
        Exit Sub
    End If
    If Check_MAT_FACTOR = Checked Then
        D54 = Spin_MAT_FACTOR / 100
        Mat_factor.Text = D54
    ElseIf Check_MAT_FACTOR = Unchecked Then
        D54 = Data1.Recordset.TUBES_Mat_fact
        Mat_factor.Text = D54
    End If
End Sub
Private Sub Check_MAT_FACTOR_Click()
On Error Resume Next
    If XXX = 1 Then
        Exit Sub
    End If
    If Check_MAT_FACTOR = Checked Then
        D54 = Spin_MAT_FACTOR / 100
        Mat_factor.Text = D54
    ElseIf Check_MAT_FACTOR = Unchecked Then
        D54 = Data1.Recordset.TUBES_Mat_fact
        Mat_factor.Text = D54
    End If
    Call Steam
End Sub
Private Sub ComboT_OD_LostFocus()
On Error Resume Next
    Dim OD_inches As String
    OD_inches = ComboT_OD.Text
        Select Case OD_inches
        Case "3/8":     T_OD = 9.525      'OD  (mm)
        Case "1/2":     T_OD = 12.7       'OD  (mm)
        Case "5/8":     T_OD = 15.875     'OD  (mm)
        Case "3/4":     T_OD = 19.05      'OD  (mm)
        Case "7/8":     T_OD = 22.225     'OD  (mm)
        Case "1":       T_OD = 25.4       'OD  (mm)
        Case "1-1/8":   T_OD = 28.575     'OD  (mm)
        Case "1-1/4":   T_OD = 31.75      'OD  (mm)
        Case "1-1/2":   T_OD = 38.1       'OD  (mm)
        Case "1-3/4":   T_OD = 44.45      'OD  (mm)
        Case "1-7/8":   T_OD = 47.625     'OD  (mm)
        Case "2":       T_OD = 50.8       'OD  (mm)
        End Select
    Spin_T_OD.Value = T_OD * 100
Call Fluid_type
End Sub
Private Sub Combo_BWG_LostFocus()
On Error Resume Next
Call Fluid_type
End Sub
Private Sub Combo_cooling_type_LostFocus()
    If Combo_cooling_type = "Condensation" Or Combo_cooling_type = "Exaust steam" Then
        Combo_CURRENT = "Condensation"
    Else
        Combo_CURRENT = "Counter-flow"
    End If
    Call Fluid_type
End Sub
Private Sub Fluid_type()
XXX = 0
    If Combo_cooling_type.Text = "Cooling" Then
        Call COOLERS
    ElseIf Combo_cooling_type.Text = "Condensation" Then
        Call CONDENSER
    ElseIf Combo_cooling_type.Text = "Exaust steam" Then
        Call Steam
    End If
End Sub
Private Sub CONDENSER()
On Error Resume Next
    Dim XPI, LN As Double
    
'Flow
    SHELL_VAPOR.Visible = True
    SHELL_LIQUID.Visible = True
    SHELL_WATER.Visible = True
    SHELL_NON_COND.Visible = True
    
    HScroll_SHELL_NON_COND.Visible = True
    Spin_VAP_P.Visible = True
    Spin_WAT_VAP_IN.Visible = True
    
    Spin_TUBES_T_IN.Max = 10000
    Spin_TUBES_T_OUT.Max = 10000
    Combo_CURRENT = "Condensation"
    lblLabels(1).Caption = "Delta enthalpy(out-in):"
    lblLabels(3).Caption = "C Factor:"
    Label24(0).Caption = "m3/h/kPa^(1/2) - Tubes-side"
    Frame_VAP.Visible = True
'Wet steam
    Label22(17).Visible = False
    Wet_steam.Visible = False
    Check_T_OUT.Visible = True
'SPECIFIC HEAT
    Check_S_SPH.Visible = True
    SHELL_OUT(3).Visible = True
    HScroll_SHELL_SPH.Visible = True
'Termal conductivity SHELL
    lbl_tubes(9).Visible = True
    SHELL_OUT(1).Visible = True
    Check_S_TC.Visible = True
    HScroll_SHELL_TC.Visible = True
    lbl_tubes(1).Visible = True
'Termal conductivity TUBES
    TUBES_OUT(1).Visible = True
    Check_T_TC.Visible = True
    HScroll_TUBES_TC.Visible = True
'Condensing steam  pressure
    lblLabels(10).Visible = True
    S_press_KP.Visible = True
    Check_CP.Visible = True
    Spin_S_PRESS.Visible = True
    Label1(10).Visible = True
'Condensing steam  temperature
    lblLabels(22).Visible = True
    SHELL_OUT(11).Visible = True
    Check_CT.Visible = True
    HScroll_C_TEMP.Visible = True
    Label22(11).Visible = True
'Latent heat
    lblLabels(1).Visible = True
    Check_LATENT.Visible = True
    SHELL_OUT(10).Visible = True
    HScroll_LATENT.Visible = True
    Label22(0).Visible = True
'Vapor percent
    Frame_VAP.Visible = True
'Material factor
    lblLabels(34).Visible = False
    Mat_factor.Visible = False
    Spin_MAT_FACTOR.Visible = False
    Check_MAT_FACTOR.Visible = False
    lblLabels(16).Visible = False
'Skin temperature
    Label22(3).Visible = True
    SKIN_TEMP.Visible = True
    Label22(4).Visible = True
'Flow calculated
    SHELL_OUT(0).Visible = True
'Flow velocity
    Label22(26).Visible = True
    SHELL_OUT(2).Visible = True
'Reynolds number
    SHELL_OUT(6).Visible = True
'Shell temperatures
    SHELL_TEMP_IN.Visible = True
    SHELL_TEMP_OUT.Visible = True
    HScroll_SHELL_T_IN.Visible = True
    HScroll_SHELL_T_OUT.Visible = True
'Shell pressure
    SHELL_P_IN.Visible = True
    SHELL_P_OUT.Visible = True
    HScroll_SHELL_P_IN.Visible = True
    HScroll_SHELL_P_OUT.Visible = True
    Check_P_DROP_S.Visible = True
'Therminal temperature label
    lblLabels(25).Caption = "Terminal temperature:"
'Thermal balance
    Thermal_bal_tubes.Visible = True
    Thermal_bal_shell.Visible = True
    Thermal_bal_shell_T.Visible = False

'Flowrate, Kg/h
'IN
Ftot_IN_1 = CDbl(SHELL_FLOW)
Ftot_IN = Ftot_IN_1
XD18 = Ftot_IN
Vtot_INP_1 = CDbl(Vtot_INP)
Vtot_IN_1 = Vtot_INP_1 / 100 * Ftot_IN_1
Vtot_IN = Format(Vtot_IN_1, "#,##0")
    HScroll_SHELL_NON_COND.Max = Vtot_IN / FACT_FLOW
    NC_IN_1 = HScroll_SHELL_NON_COND * FACT_FLOW
    NC_IN = Format(NC_IN_1, "#,##0")
    If Vtot_INP_1 = 0 Then
        Vwat_INP_1 = 0
        Vwat_INP = Format(0, "0.0")
        Vwat_IN_1 = 0
        Vwat_IN = Format(0, "#,##0")
        Vorg_IN_1 = 0
        Vorg_IN = Format(0, "#,##0")
        Vorg_INP_1 = 0
        Vorg_INP = Format(Vorg_INP_1, "0.0")
    Else
        Vwat_IN_1 = Vwat_INP_1 / 100 * Vtot_IN_1
        Vwat_IN = Format(CDbl(Vwat_IN_1), "#,##0")
        Vorg_IN_1 = Vtot_IN_1 - NC_IN_1 - Vwat_IN_1
        Vorg_IN = Format(CDbl(Vorg_IN_1), "#,##0")
        Vorg_INP_1 = Format(Vorg_IN_1 * 100 / Vtot_IN_1, "0.0")
        Vorg_INP = Format(Vorg_INP_1, "0.0")
    End If
    Ltot_IN_1 = Ftot_IN_1 - Vtot_IN_1
    Ltot_IN = Format(Ltot_IN_1, "#,##0")
    Ltot_INP_1 = 100 - Vtot_INP_1
    Ltot_INP = Format(Ltot_INP_1, "0.0")
    If Ltot_IN_1 = 0 Then
        Lwat_INP_1 = 0   'Spin_WAT_LIQ_IN / 10   'Lwat_IN_1 * 100 / Ltot_IN_1
        Lwat_INP = Format(Lwat_INP_1, "0.0")
        Lwat_IN_1 = 0    'Lwat_INP_1 / 100 * Ltot_IN_1
        Lwat_IN = Format(Lwat_IN_1, "#,##0")
        Lorg_INP_1 = 0     '100 - Lwat_INP_1
        Lorg_INP = Format(Lorg_INP_1, "0.0")
        Lorg_IN_1 = 0        'Ltot_IN_1 - Lwat_IN_1
        Lorg_IN = Format(Lorg_IN_1, "#,##0")
    Else
        Lwat_IN_1 = Lwat_INP_1 / 100 * Ltot_IN_1
        Lwat_IN = Format(CDbl(Lwat_IN_1), "#,##0")
        Lorg_INP_1 = 100 - Lwat_INP_1
        Lorg_INP = Format(Lorg_INP_1, "0.0")
        Lorg_IN_1 = Ltot_IN_1 - Lwat_IN_1
        Lorg_IN = Format(Lorg_IN_1, "#,##0")
    End If
Ltot_out = Vorg_IN_1 + Vwat_IN_1 + Lorg_IN_1 + Lwat_IN_1
    Lorg_OUT = Lorg_IN_1
    Lwat_OUT = Vwat_IN_1 + Lwat_IN_1
    SHELL_VAPOR = Format(Vorg_IN_1, "#,##0")
    SHELL_LIQUID = Format(Lorg_OUT, "#,##0")
    SHELL_WATER = Format(Lwat_OUT, "#,##0")
    SHELL_NON_COND = Format(NC_IN, "#,##0")
'Ftot
    Ftot_IN = Format(Vorg_IN_1 + Vwat_IN_1 + Lorg_IN_1 + Lwat_IN_1, "#,##0")
    Ftot_INP = Format(Vtot_INP_1 + Ltot_INP_1, "0.0")
'Percent of fractions OUT shell side
    Lorg_0UTP = Lorg_OUT / XD18 * 100
    Lwat_OUTP = Lwat_OUT / XD18 * 100
    NC_OUTP = NC_IN / XD18 * 100
    Ltot_OUTP = Lorg_0UTP + Lwat_OUTP + NC_OUTP

Call Mechanical
    
'Thermal conductivity of tube material
    Mat_cond.Text = D78
'Heat transfer surface,m^2
    D79 = 3.14159 * D67 * 25.4 * 10 ^ -3 * D75 * D74 * SERIES_N * PARALLEL_N * lungh
'Heat transfer surface,inch^2
    D80 = D79 / (0.3048 ^ 2)
    Area.Text = Format(D79, "0.0")

Call FOULING

XPI = 3.141592654
LN = 2.302585093
XD6 = TUBES_FLOW                            'TUBES TOTAL flowrate,Kg/h
XD6L = (TUBES_LIQUID + TUBES_WATER)         'TUBES LIQUID flowrate,Kg/h
XD7 = TUBES_TEMP_IN                         'Water temperature in,ºC
XD8 = TUBES_TEMP_OUT                        'Water temperature out,ºC
Range_T = XD8 - XD7                         'Tubes side delta T

PROP = "TUBES"
Call Properties

XD5 = (TUBES_LIQUID + TUBES_WATER) / D19        'Water flow rate,m3/h
XD19 = SHELL_TEMP_IN                            'Shell temperature in,ºC
XD20 = SHELL_TEMP_OUT                           'Shell temperature out,ºC
Range_S = XD19 - XD20                           'Shell Delta T
XD52M = SHELL_TUBES_PITCH                       'Pitch, mm
XD52 = XD52M / 25.4                             'Pitch,inch
XD54 = T_NO                                     'Number of tubes
XD55 = T_len                                    'Tube lenght,m
XD56 = XD55 / 0.3048                            'Tube lenght, ft
XD57 = T_PASS                                   'Number of tube side passes
XD58 = Mat_cond                                 'Thermal conductivity of tube material,Kcal/(h m^2 ºC/m)
XD59 = SHELL_PASS                               'Shell passes
XD61M = SHELL_ID                                'Shell ID, mm
XD61 = XD61M / 25.4                             'Shell ID, inch
XD63 = SHELL_BAFFLES_CUT                        'Baffle cut, %
XD64M = SHELL_BAFFLES_SPACE                     'Baffle spacing  mm
XD64 = XD64M / 25.4                             'Baffle spacing  inch
XD66M = T_OD                                    'Tube Outlet diameter, mm
XD66 = XD66M / 1000                             'Tube Outlet diameter, m
XD50 = XD66 / 25.4 * 1000                       'Tube outlet diameter, inch
XD85 = T_ID / 1000                              'Tube Inlet diameter, m
XD84 = XD85 / 0.3048                            'Tube Inlet diameter,ft
XD83 = XD85 / 25.4 * 1000                       'Tube Inlet diameter,inches
XD112 = SHELL_FF                                'Process side fouling factor [(hm^2ºC)/Kcal]*10^4
XD30 = XD112
XD75 = SHELL_OUT(4)                             'Shell density
XD77 = SHELL_OUT(5)                             'SHELL VISCOSITY

'TUBES Caloric water temperature,ºC
    D17 = XD7 + (XD8 - XD7) / 2
    XD9 = D17
'TUBES Caloric water temperature,ºF
    XD10 = XD9 * 1.8 + 32
    
'Water density at TUBES CALORIC TEMP,Kg/m3 ((t1+t2)/2)
    D19 = 0.0002 * D_17 ^ 3 - 0.028 * D17 ^ 2 + 0.0873 * D17 + 999.92
    If Check_T_DENS = 0 Then
        XD11 = D19
        TUBES_OUT(4).Text = Format(D19, "0.0")
        TUBES_OUT(4).ForeColor = &HFF0000
        TUBES_OUT(4).BackColor = &HE0E0E0
    ElseIf Check_T_DENS = 1 Then
        D19 = HScroll_TUBES_DENS / 10
        XD11 = D19
        TUBES_OUT(4) = Format(D19, "0.0")
        TUBES_OUT(4).ForeColor = &HC0&
        TUBES_OUT(4).BackColor = &HE0E0E0
    End If
'Water viscosity at tubes caloric temp. (t1+t2)/2,centipoise
    D20 = (100 / (2.1482 * ((273.16 + D17 - 281.435) + Sqr(8078.4 + (273.16 + D17 - 281.435) ^ 2)) - 120))
    If Check_T_VISC = 0 Then
        XD12 = D20
        TUBES_OUT(5).Text = Format(D20, "0.000")
        TUBES_OUT(5).ForeColor = &HFF0000
        TUBES_OUT(5).BackColor = &HE0E0E0
    ElseIf Check_T_VISC = 1 Then
        D20 = HScroll_TUBES_VISC / 1000
        XD12 = D20
        TUBES_OUT(5) = Format(D20, "0.000")
        TUBES_OUT(5).ForeColor = &HC0&
        TUBES_OUT(5).BackColor = &HE0E0E0
    End If
'Water specific heat at TUBES CALORIC TEMP,Kcal/KgºC
    h1 = 1.00691354509505
    h2 = -1.19506245657282E-03
    h3 = 5.57856020013537E-05
    h4 = -9.75376157602428E-07
    h5 = 6.26080712782905E-09
    SPH_T = h1 + h2 * D17 + h3 * D17 ^ 2 + h4 * D17 ^ 3 + h5 * D17 ^ 4
    If Check_T_SPH = 0 Then
        TUBES_OUT(3).Text = Format(SPH_T, "0.000")
        TUBES_OUT(3).ForeColor = &HFF0000
        TUBES_OUT(3).BackColor = &HE0E0E0
    ElseIf Check_T_SPH = 1 Then
        SPH_T = HScroll_TUBES_SPH / 1000
        TUBES_OUT(3) = Format(SPH_T, "0.000")
        TUBES_OUT(3).ForeColor = &HC0&
        TUBES_OUT(3).BackColor = &HE0E0E0
    End If
'Water thermal conductivity at TUBES CALORIC TEMP, Kcal/h m ºC
    TH_C = 0.00000000592317 * D17 ^ 3 - 0.0000080425 * D17 ^ 2 + 0.0018262 * D17 + 0.478535
    If Check_T_TC = 0 Then
        TUBES_OUT(1).Text = Format(TH_C, "0.000")
        TUBES_OUT(1).ForeColor = &HFF0000
        TUBES_OUT(1).BackColor = &HE0E0E0
    ElseIf Check_T_TC = 1 Then
        TH_C = HScroll_TUBES_TC / 1000
        TUBES_OUT(1) = Format(TH_C, "0.000")
        TUBES_OUT(1).ForeColor = &HC0&
        TUBES_OUT(1).BackColor = &HE0E0E0
    End If

'Flow rate, m3/h
    XD5 = (TUBES_LIQUID + TUBES_WATER) / XD11
'Water specific heat, Kcal/Kg
    TUBES_OUT(3).Text = Format(SPH_T, "0.000")
'Water density at ((t1+t2)/2), Kg/m3
    TUBES_OUT(4).Text = Format(XD11, "0.0")
'Water viscosity at (t1+t2)/2, centipoise
    TUBES_OUT(5).Text = Format(XD12, "0.000")
'TUBES TOTAL flow rate,m3/h
    XD5 = (TUBES_LIQUID + TUBES_WATER) / D19
'SHELL Caloric temperature,ºC
    XD9_S = XD20 + (XD19 - XD20) / 2

FLUID_VL = 0
'AMMONIA
    If Combo_S_FLUID = "Ammonia" Then
        Call Ammonia
'BENZENE
    ElseIf Combo_S_FLUID = "Benzene" Then
        Call BENZENE
'BUTANE
    ElseIf Combo_S_FLUID = "Butane" Then
        Call Butane
'1-BUTANOL
    ElseIf Combo_S_FLUID = "1-Butanol" Then
        Call Butanol_1
'Cyclohexane
    ElseIf Combo_S_FLUID = "Cyclohexane" Then
        Call Cyclohexane
'ETHANOL
    ElseIf Combo_S_FLUID = "Ethanol" Then
        Call Ethanol
'HEPTANE
    ElseIf Combo_S_FLUID = "Heptane" Then
        Call Heptane
'HEXANE
    ElseIf Combo_S_FLUID = "Hexane" Then
        Call Hexane
'Isobutane
    ElseIf Combo_S_FLUID = "Isobutane" Then
        Call Isobutane
'ISOPROPANOL
    ElseIf Combo_S_FLUID = "Isopropanol" Then
        Call Isopropanol
'METHANOL
    ElseIf Combo_S_FLUID = "Methanol" Then
        Call Methanol
'PROPYLENE
    ElseIf Combo_S_FLUID = "Propylene" Then
        Call Propylene
'PROPYLENE GLYCOL
    ElseIf Combo_S_FLUID = "Propylene glycol" Then
        Call Propylene_glycol
'TOLUENE
    ElseIf Combo_S_FLUID = "Toluene" Then
        Call Toluene
'VCM
    ElseIf Combo_S_FLUID = "VCM" Then
        Call VCM
Naphtalene
    ElseIf Combo_S_FLUID = "Naphtalene" Then
        Call Naphtalene
    End If

If Combo_CURRENT = "Condensation" And FLUID_VL = 0 Then
    'Condensation pressure
        If Check_CP = 0 Then
            XD21 = S_press_KP / 100
            S_press_KP.ForeColor = &HC0&
            S_press_KP.BackColor = &HE0E0E0
        ElseIf Check_CP = 1 Then
            XD21 = Spin_S_PRESS / 10000
            S_press_KP = Format(XD21 * 100, "0.00")
            S_press_KP.ForeColor = &HC0&
            S_press_KP.BackColor = &HE0E0E0
        End If
    'Thermal conductivity at condensing film temperature
        If Check_S_TC = 1 Then
            'Thermal conductivity, Kcal/hmºC
            XD79 = SHELL_OUT(1)
            'Thermal conductivity, Btu/hftºF
            XD78 = XD79 / 1.488
            SHELL_OUT(1).ForeColor = &HC0&
            SHELL_OUT(1).BackColor = &HE0E0E0
        ElseIf Check_CP = 0 Then
            'Thermal conductivity, Kcal/hmºC
            XD79 = HScroll_SHELL_TC / 1000
            SHELL_OUT(1) = Format(XD79 * 100, "0.000")
            'Thermal conductivity, Btu/hftºF
            XD78 = XD79 / 1.488
            SHELL_OUT(1).ForeColor = &HC0&
            SHELL_OUT(1).BackColor = &HE0E0E0
        End If
    'Heat capacity,J(kg-°K)
        If Check_S_SPH = 0 Then
            SPH_S = SHELL_OUT(3)
            SHELL_OUT(3).ForeColor = &HC0&
            SHELL_OUT(3).BackColor = &HE0E0E0
        ElseIf Check_S_SPH = 1 Then
            SPH_S = HScroll_SHELL_SPH / 1000
            SHELL_OUT(3) = Format(SPH_S, "0.000")
            SHELL_OUT(3).ForeColor = &HC0&
            SHELL_OUT(3).BackColor = &HE0E0E0
        End If
    'Density, kg/m3
        If Check_S_DENS = 0 Then
            XD75 = SHELL_OUT(4)
            SHELL_OUT(4).ForeColor = &HC0&
            SHELL_OUT(4).BackColor = &HE0E0E0
        ElseIf Check_S_DENS = 1 Then
            XD75 = HScroll_SHELL_SPH / 1000
            SHELL_OUT(4) = Format(XD75, "0.000")
            SHELL_OUT(4).ForeColor = &HC0&
            SHELL_OUT(4).BackColor = &HE0E0E0
        End If
    'Viscosity, cP
        If Check_S_VISC = 0 Then
            XD77 = SHELL_OUT(5)
            SHELL_OUT(5).ForeColor = &HC0&
            SHELL_OUT(5).BackColor = &HE0E0E0
        ElseIf Check_S_VISC = 1 Then
            XD77 = HScroll_SHELL_SPH / 1000
            SHELL_OUT(5) = Format(XD77, "0.000")
            SHELL_OUT(5).ForeColor = &HC0&
            SHELL_OUT(5).BackColor = &HE0E0E0
        End If
    'Latent heat Kcal/Kg
        If Check_LATENT = 0 Then
            XD24 = SHELL_OUT(10) / 4.1868
            SHELL_OUT(10).ForeColor = &HC0&
            SHELL_OUT(10).BackColor = &HE0E0E0
        ElseIf Check_LATENT = 1 Then
            XD24 = HScroll_LATENT / 4.1868
            SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
            SHELL_OUT(10).ForeColor = &HC0&
            SHELL_OUT(10).BackColor = &HE0E0E0
        End If
    'Condensation temperature
        If Check_CT = 0 Then
            XD73 = SHELL_OUT(11)
            SHELL_OUT(11).ForeColor = &HC0&
            SHELL_OUT(11).BackColor = &HE0E0E0
        ElseIf Check_CT = 1 Then
            XD73 = HScroll_C_TEMP / 100
            SHELL_OUT(11) = Format(XD73, "0.00")
            SHELL_OUT(11).ForeColor = &HC0&
            SHELL_OUT(11).BackColor = &HE0E0E0
        End If
End If
'DUTY
'Flow rate
    TUBES_OUT(0) = Format(XD6, "#,##0")
    SHELL_OUT(0) = Format(XD18, "#,##0")
    COL = 0
'Latent heat steam, Kcal/kg
    I9 = 0.168682569821809
    J9 = -1.80896828868017E-04
    J3 = -38.2917529410035
    D38 = (-I9 - Sqr(I9 ^ 2 - 4 * J9 * (J3 - Log(0.07) / 2.3))) / (2 * J9)
'Latent heat,kJ/kg
    D38_2 = D38 * 4.1868
'Tubes side duty,Kcal/h
    XD36 = XD6 * SPH_T * (XD8 - XD7)
'Tubes side duty,KW
    DUTY_T = XD36 / 860.04
    TUBES_OUT(7) = Format(DUTY_T, "#,##0")
'Shell side duty,Kcal/h
    Q_S_Vorg = Vorg_IN_1 * XD24
    Q_S_Lorg = Lorg_IN_1 * SPH_S * (XD19 - XD20)
    Q_S_Vwat = Vwat_IN_1 * D38
    Q_S_Lwat = Lwat_IN_1 * SPH_T * (XD19 - XD20)
    Q_S_NC = NC_IN * SPH_S * (XD19 - XD20)
    
    XD37 = Q_S_Vorg + Q_S_Lorg + Q_S_Vwat + Q_S_Lwat + Q_S_NC
    xd37_1 = XD37 / 860.04

'Shell side duty,KW
    DUTY_S = XD37 / 860.04
    SHELL_OUT(7) = Format(DUTY_S, "#,##0")
    
    If Thermal_bal_tubes = True Then
        COL = 1
        T_FLW = 1
        SHELL_FLOW.ForeColor = &HC0&
        SHELL_FLOW.BackColor = &H80000018
        TUBES_OUT(0).BackColor = &HC0&
        TUBES_OUT(0).ForeColor = &HFFFFFF
        SHELL_OUT(0).BackColor = &HE0E0E0
        SHELL_OUT(0).ForeColor = &HC0&
        If Check_T_OUT = 1 Then
            TUBES_FLOW.ForeColor = &HFFFFFF
            TUBES_FLOW.BackColor = &HC0&
            TUBES_TEMP_OUT.ForeColor = &HC0&
            TUBES_TEMP_OUT.BackColor = &H80000018
            XD6 = XD37 / (SPH_T * (XD8 - XD7))
            XD6L = XD6
        ElseIf Check_T_OUT = 0 Then
            TUBES_FLOW.ForeColor = &HC0&
            TUBES_FLOW.BackColor = &H80000018
            TUBES_TEMP_OUT.ForeColor = &HFFFFFF
            TUBES_TEMP_OUT.BackColor = &HC0&
            XD6 = TUBES_FLOW
            XD8 = XD37 / XD6 / SPH_T + XD7
            Range_T = XD8 - XD7
            Spin_TUBES_T_OUT = XD8 * 100
            XD36 = XD6 * SPH_T * (XD8 - XD7)    'Tubes side duty,Kcal/h
        End If
'Tube side duty, kcal/h
        XD36 = XD6 * SPH_T * (XD8 - XD7)
        TUBES_FLOW = Format(XD6, "#,##0")
        TUBES_OUT(0) = Format(XD6, "#,##0")
        HScroll_TUBES_FLOW = XD6 / FACT_FLOW
        TUBES_WATER.Text = Format(XD6, "#,##0")
        HScroll_TUBES_WATER.Max = TUBES_WATER / FACT_FLOW
        XD5 = XD6 / D19
'TUBES side duty, KW
        TUBES_OUT(7).Text = Format(XD36 / 860.04, "#,##0")
    ElseIf Thermal_bal_shell.Value = True Then
        COL = 1
        S_FLOW = 1
        SHELL_FLOW.ForeColor = &HFFFFFF
        SHELL_FLOW.BackColor = &HC0&
        If CSng(SHELL_TEMP_OUT) > CSng(PROCESS_TARGET_T_OUT) Then
            SHELL_TEMP_OUT.ForeColor = &H80000018
            SHELL_TEMP_OUT.BackColor = 255.255
        Else
            SHELL_TEMP_OUT.BackColor = &H80000018
            SHELL_TEMP_OUT.ForeColor = &HC0&
        End If
        TUBES_FLOW.ForeColor = &HC0&
        TUBES_FLOW.BackColor = &H80000018
        TUBES_TEMP_OUT.ForeColor = &HC0&
        TUBES_TEMP_OUT.BackColor = &H80000018
        SHELL_OUT(0).BackColor = &HC0&
        SHELL_OUT(0).ForeColor = &HFFFFFF
        TUBES_OUT(0).ForeColor = &HC0&
        TUBES_OUT(0).BackColor = &HE0E0E0

        Vorg_IN_F = Q_S_Vorg / XD37
        Lorg_IN_F = Q_S_Lorg / XD37
        Vwat_IN_F = Q_S_Vwat / XD37
        Lwat_IN_F = Q_S_Lwat / XD37
        NC_IN_F = Q_S_NC / XD37
        TOT_F = Vorg_IN_F + Lorg_IN_F + Vwat_IN_F + Lwat_IN_F + NC_IN_F
        
        xd37_1 = XD37 / 860.04
        
        Vorg_OUT = XD36 * Vorg_IN_F / XD24
        Lorg_0UT = XD36 * Lorg_IN_F / SPH_S / (XD19 - XD20)
        Vwat_OUT = XD36 * Vwat_IN_F / D38
        Lwat_OUT = XD36 * Lwat_IN_F / SPH_T / (XD19 - XD20)
        NC_OUT = XD36 * NC_IN_F / SPH_S / (XD19 - XD20)
        TOT_OUT = Vorg_OUT + Lorg_0UT + Vwat_OUT + Lwat_OUT + NC_OUT
        
'40     W = 0.1
'41     j = 0.1
'42     HE = W: GoSub 48
'43     Y = X: HE = j + W
'44     GoSub 48
'45     G = W: W = G - j * Y / (X - Y)
'46     If Abs(G - W) >= 0.00001 Then GoTo 42
'47     W = HE: GoTo 60
'48     FT = HE
    
        Q_S_Vorg = Vorg_OUT * XD24
        Q_S_Vwat = Vwat_OUT * D38
        Q_S_Lorg = Lorg_0UT * SPH_S * (XD19 - XD20)
        Q_S_Lwat = Lwat_OUT * SPH_T * (XD19 - XD20)
        Q_S_NC = NC_OUT * SPH_S * (XD19 - XD20)
        XD37 = Q_S_Vorg + Q_S_Vwat + Q_S_Lorg + Q_S_Lwat + Q_S_NC
        xd37_1 = XD37 / 860.04
                
        Vorg_IN = Format(Vorg_OUT, "#,##0")
        Vwat_IN = Format(Vwat_OUT, "#,##0")
        Lorg_IN = Format(Lorg_0UT, "#,##0")
        Lwat_IN = Format(Lwat_OUT, "#,##0")
        NC_IN = Format(NC_OUT, "#,##0")
        TOT_OUT = Vorg_OUT + Vwat_OUT + Lorg_0UT + Lwat_OUT + NC_OUT
        Ftot_IN = Format(TOT_OUT, "#,##0")
        XD18 = TOT_OUT
        
        Vtot_IN = Format(CDbl(Vorg_OUT) + CDbl(Vwat_OUT) + CDbl(NC_OUT), "#,##0")
        Ltot_IN = Format(CDbl(Lorg_0UT) + CDbl(Lwat_OUT), "#,##0")
        Vtot_INP = Format(CDbl(Vtot_IN) / CDbl(Ftot_IN) * 100, "0.0")
        Ltot_INP = Format(CDbl(Ltot_IN) / CDbl(Ftot_IN) * 100, "0.0")

    'Percent of fractions IN shell side
        Vorg_INP = Format(CDbl(Vorg_OUT) / CDbl(Vtot_IN) * 100, "0.0")
        Vwat_INP = Format(CDbl(Vwat_OUT) / CDbl(Vtot_IN) * 100, "0.0")
        Lorg_INP = Format(CDbl(Lorg_OUT) / CDbl(Ltot_IN) * 100, "0.0")
        Lwat_INP = Format(CDbl(Lwat_OUT) / CDbl(Ltot_IN) * 100, "0.0")
        NC_INP = NC_IN / Vtot_IN * 100

        SHELL_VAPOR = Format(Vorg_OUT, "#,##0")
        SHELL_LIQUID = Format(Vorg_0UT + Lorg_0UT, "#,##0")
        SHELL_WATER = Format(Vwat_OUT + Lwat_OUT, "#,##0")
        SHELL_NON_COND = Format(NC_OUT, "#,##0")
        SHELL_FLOW = Format(XD18, "#,##0")
        SHELL_OUT(0) = Format(XD18, "#,##0")
                
        YXY = 1
        HScroll_SHELL_NON_COND.Value = SHELL_NON_COND / FACT_FLOW
        Spin_SHELL_FLOW = SHELL_FLOW / FACT_FLOW
        YXY = 0
60  'SHELL side duty, KW
        DUTY_S = XD37 / 860.04
        SHELL_OUT(7).Text = Format(DUTY_S, "#,##0")
    Else
        If CSng(SHELL_TEMP_OUT) > CSng(PROCESS_TARGET_T_OUT) Then
            SHELL_TEMP_OUT.ForeColor = &H80000018
            SHELL_TEMP_OUT.BackColor = 255.255
        Else
            SHELL_TEMP_OUT.BackColor = &H80000018
            SHELL_TEMP_OUT.ForeColor = &HC0&
        End If
        SHELL_OUT(7).ForeColor = &HC0&
        SHELL_OUT(7).BackColor = &HE0E0E0
        SHELL_OUT(0).ForeColor = &HC0&
        SHELL_OUT(0).BackColor = &HE0E0E0
        TUBES_OUT(0).ForeColor = &HC0&
        TUBES_OUT(0).BackColor = &HE0E0E0
        SHELL_FLOW.ForeColor = &HC0&
        SHELL_FLOW.BackColor = &H80000018
        TUBES_FLOW.ForeColor = &HC0&
        TUBES_FLOW.BackColor = &H80000018
        TUBES_TEMP_OUT.ForeColor = &HC0&
        TUBES_TEMP_OUT.BackColor = &H80000018
    End If

'CALCULATING TUBE SIDE PRESSURE DROP

'Pressure drop in tubes,psi
    'Flow area section per tube  m^2
        XD91 = (XD54 * XPI * XD85 ^ 2 / 4) / XD57
        If lungh = 2 And XD57 > 1 Then
            XD91 = XD91 * lungh
        End If
        TUBES_SECTION.Text = Format(XD91, "0.0000")
    'Flow area per tube,ft^2
        XD92 = XD91 / (0.3048 ^ 2)
    'Water velocity through tubes, m/s
        XD95 = XD5 / XD91 / 3600 / PARALLEL_N
        TUBES_OUT(2) = Format(XD95, "0.00")
    'Water velocity through tubes, ft/s
        XD96 = XD95 / 0.3048
    'Reynolds number through tubes
        XD97 = XD95 * XD85 * XD11 / (XD12 * 0.001)
        TUBES_OUT(6) = Format(XD97, "#,##0")
    'Tube side friction factor   ft^2/inch^2
        XD93 = 10 ^ ((-2.5165 - 0.263 * Log(XD97) / 2.30258))
    'Mass velocity,lb/h(ft^2)
        XD94 = XD5 * XD11 * 2.20462 / XD92
    'Pressure drop in tubes, psi
        XD98 = (XD93 * XD94 ^ 2 * XD56 * XD57) / (5.22 * 10 ^ 10 * XD84)
    'Pressure drop in tubes  bar
        XD99 = XD98 * 0.068947
    'Pressure drop in tubes  Kg/cm2
        XD100 = XD98 * 0.070307
    'Tube side pressure drop due to changes in direction,psi
        XD101 = (4 * XD57) * (XD96 ^ 2 / (2 * (9.81 / 0.3048))) * (62.5 / 144)
    'Tube side pressure drop due to changes in direction, bar
        XD102 = XD101 * 0.068947
    'Tube side pressure drop due to changes in direction Kg/cm2
        XD103 = XD101 * 0.070307
    'Total pressure drop for 100% clean tube side, psi
        XD104 = XD98 + XD101
    'Total pressure drop for 100% clean tube side, bar
        XD105 = XD99 + XD102
    'Total pressure drop for 100% clean tube side, Kg/cm2
        XD106 = XD100 + XD103
    'Total pressure drop for 100% clean tube side, kPa
        XD106_KPA = XD105 * 100
        If Check_P_DROP_T = Unchecked Then
            'Pressure drop in tubes, kPa
            TUBES_OUT(9) = Format(XD106_KPA, "0.00")
        Else
            'Pressure drop in tubes, kPa
            XD106_KPA = (TUBES_P_IN - TUBES_P_OUT) * 100
            TUBES_OUT(9).Text = Format(XD106_KPA, "0.00")
    End If
'C Factor
    C_F = XD5 / (XD106_KPA) ^ (1 / 2)
    C_Factor.Text = Format(C_F, "0.0")

S_T_OUT = XD20

'CALCULATING HEAT TRANSFER
    'Water side individual heat transfer coeficient, Btu/(h ft^2 F)
        XD108 = 150 * (1 + 0.011 * XD10) * (XD96 ^ 0.8 / XD83 ^ 0.2)
    'Water side individual heat transfer coeficient, Kcal/(h m^2 C)
        XD109 = XD108 * 4.882
    'Hio  = D58 - Water side indiv. heat transfer coeficient referred to ext. surface Kcal/(h m^2 C)
        XD110 = XD109 * (XD85 / XD66)
    'D60 - Heat transfer resistance due to the wall, [(hm^2ºC)/Kcal]*10^4
        XD111 = (XD66 * Log(XD66 / XD85) / (2 * XD58)) * 10000
    'Ho, Heat transfer resistance due to water (tube side), [(hm^2ºC)/Kcal]*10^4
        XD114 = 10 ^ 4 / XD110
    
    'Shell side crossflow area, ft^2
        XD68 = XD61 * (XD52 - XD50) * XD64 / (XD52 * 144) * lungh
    'Condensate loading  lb/h ft
        XD69 = Ftot_IN * 2.20462 / (XD56 * XD54 ^ (2 / 3))

400     W = 0.1
410     j = 0.1
420     HE = W: GoSub 480
430     Y = X: HE = j + W
440     GoSub 480
450     g = W: W = g - j * Y / (X - Y)
460     If Abs(g - W) >= 0.00001 Then GoTo 420
470     W = HE: GoTo 500
480     xd72 = HE
        
    'Condensing film temperature ºC
        XD73 = (S_T_OUT + xd72) / 2
    'Condensing film temperature ºC
        If Check_CT = Unchecked And FLUID_VL = 1 Then
            SHELL_OUT(11) = Format(XD73, "0.00")
        ElseIf Check_CT = Unchecked And FLUID_VL = 0 Then
            SHELL_OUT(11) = Format(HScroll_C_TEMP / 100, "0.00")
            XD73 = SHELL_OUT(11)
        ElseIf Check_CT = Checked Then
            XD73 = HScroll_C_TEMP / 100
        End If
    'Condensing film temperature,K
        XD74 = 273.16 + XD73
        XD20 = XD73
'AMMONIA
If Combo_S_FLUID = "Ammonia" Then
    Call Ammonia
'BENZENE
ElseIf Combo_S_FLUID = "Benzene" Then
    Call BENZENE
'BUTANE
ElseIf Combo_S_FLUID = "Butane" Then
    Call Butane
'1-BUTANOL
ElseIf Combo_S_FLUID = "1-Butanol" Then
    Call Butanol_1
'Cyclohexane
ElseIf Combo_S_FLUID = "Cyclohexane" Then
    Call Cyclohexane
'ETHANOL
ElseIf Combo_S_FLUID = "Ethanol" Then
    Call Ethanol
'HEPTANE
ElseIf Combo_S_FLUID = "Heptane" Then
    Call Heptane
'HEXANE
ElseIf Combo_S_FLUID = "Hexane" Then
    Call Hexane
'Isobutane
ElseIf Combo_S_FLUID = "Isobutane" Then
    Call Isobutane
'ISOPROPANOL
ElseIf Combo_S_FLUID = "Isopropanol" Then
    Call Isopropanol
'METHANOL
ElseIf Combo_S_FLUID = "Methanol" Then
    Call Methanol
'PROPYLENE GLYCOL
ElseIf Combo_S_FLUID = "Propylene glycol" Then
    Call Propylene_glycol
'PROPYLENE
ElseIf Combo_S_FLUID = "Propylene" Then
    Call Propylene
'TOLUENE
ElseIf Combo_S_FLUID = "Toluene" Then
    Call Toluene
ElseIf Combo_S_FLUID = "VCM" Then
    Call VCM
ElseIf Combo_S_FLUID = "Naphtalene" Then
    Call Naphtalene
    
'CONDENSATION
ElseIf Combo_CURRENT = "Condensation" Then
    'Condensation pressure
        If Check_CP = 0 Then
            XD21 = S_press_KP / 100
            S_press_KP.ForeColor = &HC0&
            S_press_KP.BackColor = &HE0E0E0
        ElseIf Check_CP = 1 Then
            XD21 = Spin_S_PRESS / 10000
            S_press_KP = Format(XD21 * 100, "0.00")
            S_press_KP.ForeColor = &HC0&
            S_press_KP.BackColor = &HE0E0E0
        End If
    'Thermal conductivity at condensing film temperature
        If Check_S_TC = 0 Then
            'Thermal conductivity, Kcal/hmºC
            XD79 = SHELL_OUT(1)
            'Thermal conductivity, Btu/hftºF
            XD78 = XD79 / 1.488
            SHELL_OUT(1).ForeColor = &HC0&
            SHELL_OUT(1).BackColor = &HE0E0E0
        ElseIf Check_CP = 1 Then
            'Thermal conductivity, Kcal/hmºC
            XD79 = HScroll_SHELL_TC / 1000
            SHELL_OUT(1) = Format(XD79 * 100, "0.000")
            'Thermal conductivity, Btu/hftºF
            XD78 = XD79 / 1.488
            SHELL_OUT(1).ForeColor = &HC0&
            SHELL_OUT(1).BackColor = &HE0E0E0
        End If
    'Heat capacity,J(kg-°K)
        If Check_S_SPH = 0 Then
            SPH_S = SHELL_OUT(3)
            SHELL_OUT(3).ForeColor = &HC0&
            SHELL_OUT(3).BackColor = &HE0E0E0
        ElseIf Check_S_SPH = 1 Then
            SPH_S = HScroll_SHELL_SPH / 1000
            SHELL_OUT(3) = Format(SPH_S, "0.000")
            SHELL_OUT(3).ForeColor = &HC0&
            SHELL_OUT(3).BackColor = &HE0E0E0
        End If
    'Density, kg/m3
        If Check_S_DENS = 0 Then
            XD75 = SHELL_OUT(4)
            SHELL_OUT(4).ForeColor = &HC0&
            SHELL_OUT(4).BackColor = &HE0E0E0
        ElseIf Check_S_DENS = 1 Then
            XD75 = HScroll_SHELL_DENS / 10
            SHELL_OUT(4) = Format(XD75, "0.0")
            SHELL_OUT(4).ForeColor = &HC0&
            SHELL_OUT(4).BackColor = &HE0E0E0
        End If
    'Density at condensing film temperature,lb/ft^3
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
    'Viscosity, cP
        If Check_S_VISC = 0 Then
            XD77 = SHELL_OUT(5)
            SHELL_OUT(5).ForeColor = &HC0&
            SHELL_OUT(5).BackColor = &HE0E0E0
        ElseIf Check_S_VISC = 1 Then
            XD77 = HScroll_SHELL_VISC / 1000
            SHELL_OUT(5) = Format(XD77, "0.000")
            SHELL_OUT(5).ForeColor = &HC0&
            SHELL_OUT(5).BackColor = &HE0E0E0
        End If
    'Latent heat Kcal/Kg
        If Check_LATENT = 0 Then
            XD24 = SHELL_OUT(10) / 4.1868
            SHELL_OUT(10).ForeColor = &HC0&
            SHELL_OUT(10).BackColor = &HE0E0E0
        ElseIf Check_LATENT = 1 Then
            XD24 = HScroll_LATENT / 4.1868
            SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
            SHELL_OUT(10).ForeColor = &HC0&
            SHELL_OUT(10).BackColor = &HE0E0E0
        End If
    'Condensation temperature
        If Check_CT = 0 Then
            XD73 = SHELL_OUT(11)
            SHELL_OUT(11).ForeColor = &HC0&
            SHELL_OUT(11).BackColor = &HE0E0E0
        ElseIf Check_CT = 1 Then
            XD73 = HScroll_C_TEMP / 100
            SHELL_OUT(11) = Format(XD73, "0.00")
            SHELL_OUT(11).ForeColor = &HC0&
            SHELL_OUT(11).BackColor = &HE0E0E0
        End If

End If
    'Shell side indiv. heat transfer coefficient (guess one till Z_0=0)
        XD80 = 1.5 * ((4 * XD69 / XD77) ^ -(1 / 3)) * (XD77 ^ 2 / (XD78 ^ 3 * XD76 ^ 2 * 9.81 * (3600 ^ 2 / 0.3048))) ^ -(1 / 3)
        XD81 = XD80 * 4.882
    'Heat transfer resistance due to process (shell side),[(hm^2ºC)/Kcal]*10^4
        XD115 = (1 / XD81) * 10000
    'Overall CLEAN heat transfer coefficient Kcal/(h m^2 ºC)
        XD117 = 10000 / (XD111 + XD114 + XD115)
    'Wall temperature, °C
        '(tw2-t2)/(1/alfa2)=(T1-t2)/(1/OHTC)
        'tw2 = (T1-t2)/(1/OHTC)*(1/alfa2)+ t2
        'tw2: wall temperature cold side
        't2: temperature cold side (here outlet)
        'alfa2: convective heat transfer cold side
        'T1: hot side temperature (inlet)
        'OHTC: overall heat transfer coefficient.
        tw2a = (XD19 - XD8) / (1 / XD117) * (1 / XD109) + XD8
        SKIN_TEMP = Format(tw2a, "0.00")
    X = (tw2a - xd72)
Return
500 'Condensation temperature, °C
    'Condensation temperature
        If Check_CT = 0 Then
            SHELL_OUT(11) = Format(XD73, "0.00")
'            SHELL_OUT(11).ForeColor = &HFF0000
'            SHELL_OUT(11).BackColor = &HE0E0E0
        ElseIf Check_CT = 1 Then
            XD73 = HScroll_C_TEMP / 100
            SHELL_OUT(11) = Format(XD73, "0.00")
            SHELL_OUT(11).ForeColor = &HC0&
            SHELL_OUT(11).BackColor = &HE0E0E0
        End If

PROP = "SHELL"
Call Properties

'SHELL thermal conductivity at condensing film temperature Kcal/h m ºC
    SHELL_OUT(1) = Format(XD79, "0.000")
    XD78 = XD79 / 1.488
'SHELL density at condensing film temperature,Kg/m^3
    SHELL_OUT(4) = Format(XD75, "0.0")
'SHELL viscosity at condensing film temperature,cp
    SHELL_OUT(5) = Format(XD77, "0.000")
'Wall temperature,ºC
    SKIN_TEMP.Text = Format(xd72, "0.00")
'SHELL Velocity
    L37 = XD61M   'SHELL_ID, mm
    O37 = XD52M   'SHELL_TUBES_PITCH, mm
    E37 = XD66M   'T_OD, mm
    N37 = XD64M   'SHELL_BAFFLES_SPACE, mm
'SHELL CLEARANCE
        V37 = O37 / 1000 - E37 / 1000
        Clearance.Text = Format(V37, "0.0000")
'SHELL CROSS FLOW AREA, m2
        U37 = L37 / 1000 * V37 * N37 / 1000 / (O37 / 1000)
        Flow_area.Text = Format(U37, "0.0000")
    k47 = (Vorg_IN_1 + Lorg_IN_1 + Vwat_IN_1 + Lwat_IN_1) / XD75 / U37 / 3600 / PARALLEL_N * XD59
    SHELL_OUT(2) = Format(k47, "0.00")
'SHELL Reynolds
    'SHELL_TUBES_PITCH / 1000
        EQ_D19 = XD52M / 1000
        EQ_PI = XPI
    'T_OD / 1000
        EQ_D14 = XD66
    'Equivalent diameter, m
        If SHELL_PITCH_CONF = "Triangular" Then
            EQ_E31 = 4 * (EQ_D19 ^ 2 - EQ_PI * EQ_D14 ^ 2 / 4) / (XPI * EQ_D14)
        Else
            EQ_E31 = (4 * (0.5 * EQ_D19 * 0.866 * EQ_D19 - 0.5 * XPI * EQ_D14 ^ 2 / 4) / (0.5 * XPI * EQ_D14))
        End If
    'Flow area section
        EQ_E29 = U37
    'Section SHELL_FLOW LIQUID, kg/m2
        EQ_E25 = CDbl(SHELL_LIQUID) + CDbl(SHELL_WATER)
        EQ_E30 = EQ_E25 / EQ_E29
    'Viscosity SHELL_OUT(5) * 3.6 =
        EQ_E8 = XD77 * 3.6
    ' Eq_dia,m * flow section,m2 / visc,?
        EQ_E32 = EQ_E31 * EQ_E30 / EQ_E8
    ' Eq_dia,mm
        Q_E22 = EQ_E31 * 1000
        Q_E22_1 = EQ_E32 * 1000
    'Density, kg/m3
        Q_E17 = XD75
    'Shell flow velocity, m/s
        Q_E27 = k47
    'SHELL viscosity, cP
        Q_E18 = XD77
        Q_E28 = Q_E22 * Q_E17 * Q_E27 / Q_E18
    SHELL_OUT(6) = Format(Q_E28, "#,##0")

'CALCULATING SHELL SIDE PRESSURE DROP
    If Check_P_DROP_S = Unchecked Then
'Pressure drop (tubes)
        P_E17 = XD75             'Density, kg/m3
        P_E22 = EQ_E31 * 1000    'Equivalent diameter, mm
        P_E27 = Q_E27            'Shell flow velocity, m/s
        P_E23 = XD55 * lungh     'T_len, mm
        P_E28 = Q_E28
        P_E29 = 0.44 * P_E28 ^ -0.19
        P_E30 = 4 * P_E29 * P_E23 * P_E27 ^ 2 / (P_E22 * 2 * 9.8) * P_E17 * 0.000096784 * 101.325
'Pressure drop(sheet), bar(a)
        P_E9 = XD59   'Shell passes
        P_E31 = 3 * P_E9 * P_E27 ^ 2 / 2 / 9.8 * P_E17 * 0.000096784 * 101.325
        P_E32 = P_E30 + P_E31
    'Pressure drop(sheet), KPa
        SHELL_OUT(9) = Format(P_E32 * 100, "0.00")
    Else
        P_E32 = (SHELL_P_IN - SHELL_P_OUT)
        SHELL_OUT(9) = Format(P_E32 * 100, "0.00")
    End If
'Water side fouling factor   [(hm^2ºC)/Kcal]*10^4
    'Overall heat transfer coefficient   Kcal/(h m^2 ºC)
        'Surface per linear ft, ft^2
            'Tube outlet diameter,inch
                XD50 = Format(XD66 * 1000 / 25.4, "0.000")
        'Surface per linear m, m^2
            XD90 = 3.14159 * D67 * 25.4 * 10 ^ -3 * D75 * D74 * SERIES_N * PARALLEL_N * lungh
        'Log Mean Temperature Difference CORRECTED, ºC
            AG6 = ((XD19 - XD8) - (XD20 - XD7)) / Log((XD19 - XD8) / (XD20 - XD7))
            RR = (XD19 - XD20) / (XD8 - XD7)
            ss = (XD8 - XD7) / (XD19 - XD7)
            If T_PASS > 1 And SERIES_N > 1 Then
                FT1 = (Sqr(RR ^ 2 + 1) / (2 * (RR - 1))) * Log((1 - ss) / (1 - RR * ss))
                FT2 = 2 / ss - 1 - RR + (2 / ss) * Sqr((1 - ss) * (1 - RR * ss)) + Sqr(RR ^ 2 + 1)
                FT3 = 2 / ss - 1 - RR + (2 / ss) * Sqr((1 - ss) * (1 - RR * ss)) - Sqr(RR ^ 2 + 1)
                FT4 = Log(FT2 / FT3)
                FT = FT1 / FT4
            ElseIf T_PASS > 1 And SHELL_PASS > 1 Then
                FT1 = (Sqr(RR ^ 2 + 1) / (2 * (RR - 1))) * Log((1 - ss) / (1 - RR * ss))
                FT2 = 2 / ss - 1 - RR + (2 / ss) * Sqr((1 - ss) * (1 - RR * ss)) + Sqr(RR ^ 2 + 1)
                FT3 = 2 / ss - 1 - RR + (2 / ss) * Sqr((1 - ss) * (1 - RR * ss)) - Sqr(RR ^ 2 + 1)
                FT4 = Log(FT2 / FT3)
                FT = FT1 / FT4
            Else
                FT1 = Sqr(RR ^ 2 + 1) * Log((1 - ss) / (1 - RR * ss))
                FT2 = 2 - ss * (RR + 1 - Sqr(RR ^ 2 + 1))
                FT3 = 2 - ss * (RR + 1 + Sqr(RR ^ 2 + 1))
                FT = FT1 / ((RR - 1) * Log(FT2 / FT3))
            End If
            AH6 = AG6 * FT
            XD31 = AH6
            XD38 = XD36 / XD90 / XD31

'Overall CLEAN heat transfer coefficient Kcal/(h m^2 ºC)
    If Check_U_CLEAN = Checked Then
        XD117 = HScroll_U_CLEAN
        U_COEFF_CLEAN = XD117
    ElseIf Check_U_CLEAN = Unchecked Then
        U_COEFF_CLEAN.Text = Format(XD117, "0.0")
    End If
'Water side fouling factor   [(hm^2ºC)/Kcal]*10^3
    xd118 = ((1 / XD38) - (1 / XD117) - (XD112 / 10000)) * 10000 * (XD85 / XD66)
'Total heat transfer resistance  [(h m^2 ºC)/Kcal]*10^4
    XD116 = 10000 / XD38
'Heat transfer resistance due to inside fouling factor,[(hm^2ºC)/Kcal]*10^4
    XD113 = xd118 * (XD66 / XD85)
'TUBES HEAT FLUX, kcal/m2
    Q6 = XD6 * SPH_T * (XD8 - XD7) / XD90
'TUBES HEAT FLUX, kW/m2
    TUBES_OUT(8).Text = Format(Q6 * 0.001163, "0.00")
'SHELL HEAT FLUX, kCal/m2
    'SHELL specific HEAT , kCal/kg
    If Check_S_SPH = Checked Then
        SPH_S = SHELL_OUT(3).Text
    ElseIf Check_S_SPH = Unchecked Then
        SPH_S = HScroll_SHELL_SPH / 1000
    End If
    Q6S = XD37 / XD90
    SHELL_OUT(8).Text = Format(Q6S * 0.001163, "0.00")
'Area, m^2
    Area.Text = Format(XD90, "0.00")
'Log Mean Temperature Difference, ºC
    LMTD.Text = Format(AG6, "0.00")
'Log Mean Temperature Difference corrected, ºC
    MTDc.Text = Format(AH6, "0.00")
'Condenser temperature approach, (T2-t2)  ºC
    XD32 = XD20 - XD8
'Terminal temperature difference T2-t2), °C
    TTD = Format(XD32, "0.00")
    Label22(15).Caption = "(T2 - t2)"
    TUBES_OUT(7).Text = Format(XD36 / 860.04, "#,##0") 'TUBES side duty,Kcal/h, KW
    SHELL_OUT(7).Text = Format(XD37 / 860.04, "#,##0") 'SHELL side duty, KW
'Overall CLEAN heat transfer coefficient Kcal/(h m^2 ºC)
'    If Check_U_CLEAN = Checked Then
'        XD117 = HScroll_U_CLEAN
'        U_COEFF_CLEAN = XD117
'    ElseIf Check_U_CLEAN = Unchecked Then
'        U_COEFF_CLEAN.Text = Format(XD117, "0.0")
'    End If
'Overall heat transfer coefficient   Kcal/(h m^2 ºC)
    U_COEFF_DIRTY.Text = Format(XD38, "0.0")
'Water side fouling factor   [(hm^2ºC)/Kcal]*10^3
    TUBES_FF_1 = xd118
    TUBES_FF.Text = Format(TUBES_FF_1, "0.000")
End Sub
Private Sub COOLERS()
On Error Resume Next
    Dim XPI, LN As Double
'Temperatures
    Spin_TUBES_T_IN.Max = 10000
    Spin_TUBES_T_OUT.Max = 10000
    Check_T_OUT.Visible = True
    lblLabels(25).Caption = "Approach temperature:"
'Labels
    lblLabels(3).Caption = "C Factor:"
    Label24(0).Caption = "m3/h/kPa^(1/2) - Tubes-side"
'Flow
    HScroll_SHELL_NON_COND.Visible = True
    Spin_WAT_VAP_IN.Visible = True
    Spin_WAT_LIQ_IN.Visible = True
    Spin_VAP_P.Visible = True
    SHELL_VAPOR.Visible = True
    SHELL_NON_COND.Visible = True
'SPECIFIC HEAT
    Check_S_SPH.Visible = True
    SHELL_OUT(3).Visible = True
    HScroll_SHELL_SPH.Visible = True
'Termal conductivity SHELL
    lbl_tubes(9).Visible = True
    SHELL_OUT(1).Visible = True
    Check_S_TC.Visible = True
    HScroll_SHELL_TC.Visible = True
    lbl_tubes(1).Visible = True
'Termal conductivity TUBES
    TUBES_OUT(1).Visible = True
    Check_T_TC.Visible = True
    HScroll_TUBES_TC.Visible = True
'Condensing steam  pressure
    lblLabels(10).Visible = False
    S_press_KP.Visible = False
    Check_CP.Visible = False
    Spin_S_PRESS.Visible = False
    Label1(10).Visible = False
'Condensing steam  temperature
    lblLabels(22).Visible = False
    SHELL_OUT(11).Visible = False
    Check_CT.Visible = False
    HScroll_C_TEMP.Visible = False
    Label22(11).Visible = False
'Latent heat
    lblLabels(1).Visible = False
    Check_LATENT.Visible = False
    SHELL_OUT(10).Visible = False
    HScroll_LATENT.Visible = False
    Label22(0).Visible = False
'Material factor
    lblLabels(34).Visible = False
    Mat_factor.Visible = False
    Spin_MAT_FACTOR.Visible = False
    Check_MAT_FACTOR.Visible = False
    lblLabels(16).Visible = False
'Wet steam
    Label22(17).Visible = False
    Wet_steam.Visible = False
'Vapor percent
    If Combo_S_FLUID = "Water" Then
        Frame_VAP.Visible = True
    Else
        Frame_VAP.Visible = True
    End If
'Skin temperature
    Label22(3).Visible = True
    SKIN_TEMP.Visible = True
    Label22(4).Visible = True
'Shell calculated flow
    SHELL_OUT(0).Visible = True
'Shell flow velocity
    SHELL_OUT(2).Visible = True
    Label22(26).Visible = True
'Shell Reynolds number
    SHELL_OUT(6).Visible = True
'Shell temperatures
    SHELL_TEMP_IN.Visible = True
    SHELL_TEMP_OUT.Visible = True
    HScroll_SHELL_T_IN.Visible = True
    HScroll_SHELL_T_OUT.Visible = True
'Shell pressures
    SHELL_P_IN.Visible = True
    SHELL_P_OUT.Visible = True
    HScroll_SHELL_P_IN.Visible = True
    HScroll_SHELL_P_OUT.Visible = True
    Check_P_DROP_S.Visible = True
'Thermal balance
    Thermal_bal_tubes.Visible = True
    Thermal_bal_shell.Visible = True
    Thermal_bal_shell_T.Visible = True

'Total shell flowrate,Kg/h
'IN
Ftot_IN_1 = CDbl(SHELL_FLOW)
Ftot_IN = Ftot_IN_1
XD18 = Ftot_IN
Vtot_IN_1 = 0      'Vtot_INP_1 / 100 * Ftot_IN_1
Vtot_IN = Format(Vtot_IN_1, "#,##0")
Vtot_INP_1 = 0      'Vtot_INP_1 / 100 * Ftot_IN_1
Vtot_INP = Format(Vtot_INP_1, "0.0")
    
    HScroll_SHELL_NON_COND.Max = Vtot_IN_1 / FACT_FLOW
    NC_IN_1 = HScroll_SHELL_NON_COND * FACT_FLOW
    NC_IN = Format(NC_IN_1, "#,##0")
        Vwat_INP_1 = 0
        Vwat_INP = Format(0, "0.0")
        Vwat_IN_1 = 0
        Vwat_IN = Format(0, "#,##0")
        Vorg_IN_1 = 0
        Vorg_IN = Format(0, "#,##0")
        Vorg_INP_1 = 0
        Vorg_INP = Format(Vorg_INP_1, "0.0")
    Ltot_IN_1 = Ftot_IN_1 - Vtot_IN_1
    Ltot_IN = Format(Ltot_IN_1, "#,##0")
    Ltot_INP_1 = 100 - Vtot_INP_1
    Ltot_INP = Format(Ltot_INP_1, "0.0")
        Spin_WAT_LIQ_IN.Max = 1000    'Ltot_INP_1 * 10
        Lwat_INP_1 = Spin_WAT_LIQ_IN / 10   'Lwat_IN_1 * 100 / Ltot_IN_1
        Lwat_INP = Format(Lwat_INP_1, "0.0")
        Lwat_IN_1 = Lwat_INP_1 / 100 * Ltot_IN_1
        Lwat_IN = Format(CDbl(Lwat_IN_1), "#,##0")
        Lorg_INP_1 = 100 - Lwat_INP_1
        Lorg_INP = Format(Lorg_INP_1, "0.0")
        Lorg_IN_1 = Ltot_IN_1 - Lwat_IN_1
        Lorg_IN = Format(Lorg_IN_1, "#,##0")
'OUT
Ltot_out = Vorg_IN_1 + Vwat_IN_1 + Lorg_IN_1 + Lwat_IN_1
    Lorg_OUT = Lorg_IN_1
    Lwat_OUT = Vwat_IN_1 + Lwat_IN_1
        SHELL_VAPOR = Format(Vorg_IN_1, "#,##0")
        SHELL_LIQUID = Format(Lorg_OUT, "#,##0")
        SHELL_WATER = Format(Lwat_OUT, "#,##0")
        SHELL_NON_COND = Format(NC_IN, "#,##0")
'Ftot
    Ftot_IN = Format(Vorg_IN_1 + Vwat_IN_1 + Lorg_IN_1 + Lwat_IN_1, "#,##0")
    Ftot_INP = Format(Vtot_INP_1 + Ltot_INP_1, "0.0")
'Percent of fractions OUT shell side
    Lorg_0UTP = Lorg_OUT / XD18 * 100
    Lwat_OUTP = Lwat_OUT / XD18 * 100
    NC_OUTP = NC_IN / XD18 * 100
    Ltot_OUTP = Lorg_0UTP + Lwat_OUTP + NC_OUTP

XPI = 3.141592654
LN = 2.302585093
XD6 = TUBES_FLOW                    'TUBES total flowrate,Kg/h
XD6L = TUBES_LIQUID                 'TUBES liquid flowrate,Kg/h
XD7 = TUBES_TEMP_IN                 'TUBES temperature in,ºC
XD8 = TUBES_TEMP_OUT                'TUBES temperature out,ºC
Range_T = (XD8 - XD7)
XD18 = Ftot_IN                     'Shell total fluid flowrate,Kg/h
XD19 = SHELL_TEMP_IN                'Shell fluid temperature in,ºC
XD20 = SHELL_TEMP_OUT               'Shell fluid temperature out,ºC
Range_S = (XD19 - XD20)
XD52M = SHELL_TUBES_PITCH           'Pitch, mm
XD52 = XD52M / 25.4                 'Pitch,inch
XD54 = T_NO                         'Number of tubes
XD55 = T_len                        'Tube lenght,m
XD56 = XD55 / 0.3048                'Tube lenght     ft
XD57 = T_PASS                       'Number of tube side passes
XD58 = Mat_cond                     'Thermal conductivity of tube material,Kcal/(h m^2 ºC/m)
XD59 = SHELL_PASS                   'Shell passes
XD61M = SHELL_ID                    'Shell ID, mm
XD61 = XD61M / 25.4                 'Shell ID, inch
XD63 = SHELL_BAFFLES_CUT            'Baffle cut, %
XD64M = SHELL_BAFFLES_SPACE         'Baffle spacing  mm
XD64 = XD64M / 25.4                 'Baffle spacing  inch
XD66M = T_OD                        'Tube Outlet diameter, mm
XD66 = XD66M / 1000                 'Tube Outlet diameter, m
XD50 = XD66 / 25.4 * 1000           'Tube outlet diameter, inch
XD85 = T_ID / 1000                  'Tube Inlet diameter, m
XD84 = XD85 / 0.3048                'Tube Inlet diameter,ft
XD83 = XD85 / 25.4 * 1000           'Tube Inlet diameter,inches
XD112 = SHELL_FF                    'Process side fouling factor [(hm^2ºC)/Kcal]*10^4
XD75 = SHELL_OUT(4)                 'Shell density
XD77 = SHELL_OUT(5)                 'SHELL VISCOSITY
S_press_KP_1 = 0
S_press_KP.Text = Format(S_press_KP_1, "0.00")

Call Mechanical
    
'Material conductivity of tubes
    Mat_cond.Text = D78
'Heat transfer surface,m^2
    D79 = 3.14159 * D67 * 25.4 * 10 ^ -3 * D75 * D74 * SERIES_N * PARALLEL_N * lungh
'    D79 = 3.14159 * D67 * 25.4 * 10 ^ -3 * D75 * D74 * lungh
'Heat transfer surface,inch^2
    D80 = D79 / (0.3048 ^ 2)
    Area.Text = Format(D79, "0.0")

Call FOULING

'PROP = "TUBES"
'Call Properties
        
'TUBES Caloric temperature,ºC
    D11 = XD7
    D13 = XD8
    D17 = D11 + (D13 - D11) / 2
    XD9 = D17
'TUBES Caloric temperature,ºF
    XD10 = XD9 * 1.8 + 32
'Water density at TUBES CALORIC TEMP,Kg/m3 ((t1+t2)/2)
    D19 = 0.0002 * D_17 ^ 3 - 0.028 * D17 ^ 2 + 0.0873 * D17 + 999.92
    If Check_T_DENS = 0 Then
        XD11 = D19
        TUBES_OUT(4).Text = Format(D19, "0.0")
        TUBES_OUT(4).ForeColor = &HFF0000
        TUBES_OUT(4).BackColor = &HE0E0E0
    ElseIf Check_T_DENS = 1 Then
        D19 = HScroll_TUBES_DENS / 10
        XD11 = D19
        TUBES_OUT(4) = Format(D19, "0.0")
        TUBES_OUT(4).ForeColor = &HC0&
        TUBES_OUT(4).BackColor = &HE0E0E0
    End If
'Water viscosity at tubes caloric temp. (t1+t2)/2,centipoise
    D20 = (100 / (2.1482 * ((273.16 + D17 - 281.435) + Sqr(8078.4 + (273.16 + D17 - 281.435) ^ 2)) - 120))
    If Check_T_VISC = 0 Then
        XD12 = D20
        TUBES_OUT(5).Text = Format(D20, "0.000")
        TUBES_OUT(5).ForeColor = &HFF0000
        TUBES_OUT(5).BackColor = &HE0E0E0
    ElseIf Check_T_VISC = 1 Then
        D20 = HScroll_TUBES_VISC / 1000
        XD12 = D20
        TUBES_OUT(5) = Format(D20, "0.000")
        TUBES_OUT(5).ForeColor = &HC0&
        TUBES_OUT(5).BackColor = &HE0E0E0
    End If
'Water specific heat at TUBES CALORIC TEMP,Kcal/KgºC
    h1 = 1.00691354509505
    h2 = -1.19506245657282E-03
    h3 = 5.57856020013537E-05
    h4 = -9.75376157602428E-07
    h5 = 6.26080712782905E-09
    SPH_T = h1 + h2 * D17 + h3 * D17 ^ 2 + h4 * D17 ^ 3 + h5 * D17 ^ 4
    If Check_T_SPH = 0 Then
        TUBES_OUT(3).Text = Format(SPH_T, "0.000")
        TUBES_OUT(3).ForeColor = &HFF0000
        TUBES_OUT(3).BackColor = &HE0E0E0
    ElseIf Check_T_SPH = 1 Then
        SPH_T = HScroll_TUBES_SPH / 1000
        TUBES_OUT(3) = Format(SPH_T, "0.000")
        TUBES_OUT(3).ForeColor = &HC0&
        TUBES_OUT(3).BackColor = &HE0E0E0
    End If
'Water fluid thermal conductivity at TUBES CALORIC TEMP, Kcal/h m ºC
    TH_C = 0.00000000592317 * D17 ^ 3 - 0.0000080425 * D17 ^ 2 + 0.0018262 * D17 + 0.478535
    If Check_T_TC = 0 Then
        TUBES_OUT(1).Text = Format(TH_C, "0.000")
        TUBES_OUT(1).ForeColor = &HFF0000
        TUBES_OUT(1).BackColor = &HE0E0E0
    ElseIf Check_T_TC = 1 Then
        TH_C = HScroll_TUBES_TC / 1000
        TUBES_OUT(1) = Format(TH_C, "0.000")
        TUBES_OUT(1).ForeColor = &HC0&
        TUBES_OUT(1).BackColor = &HE0E0E0
    End If
        
'PROP = "SHELL"
'Call Properties
    
'SHELL Caloric temperature,ºC
    D17 = XD20 + (XD19 - XD20) / 2
    XD9_S = D17
If Combo_S_FLUID = "Water" Then
    'Shell fluid density at film temperature,Kg/m^3
        D19 = 0.0002 * xd72 ^ 3 - 0.028 * xd72 ^ 2 + 0.0873 * xd72 + 999.92
        If Check_S_DENS = 0 Then
            XD75 = D19
            SHELL_OUT(4).Text = Format(D19, "0.0")
            SHELL_OUT(4).ForeColor = &HFF0000
            SHELL_OUT(4).BackColor = &HE0E0E0
        ElseIf Check_S_DENS = 1 Then
            D19 = HScroll_SHELL_DENS / 10
            XD75 = D19
            SHELL_OUT(4) = Format(D19, "0.0")
            SHELL_OUT(4).ForeColor = &HC0&
            SHELL_OUT(4).BackColor = &HE0E0E0
        End If
        'Shell fluid density at film temperature,lb/ft^3
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
    'Shell fluid viscosity at condensing film temperature,cp
        D20 = (100 / (2.1482 * ((273.16 + D17 - 281.435) + Sqr(8078.4 + (273.16 + D17 - 281.435) ^ 2)) - 120))
        If Check_S_VISC = 0 Then
            XD77 = D20
            SHELL_OUT(5).Text = Format(D20, "0.000")
            SHELL_OUT(5).ForeColor = &HFF0000
            SHELL_OUT(5).BackColor = &HE0E0E0
        ElseIf Check_S_VISC = 1 Then
            D20 = HScroll_SHELL_VISC / 1000
            XD77 = D20
            SHELL_OUT(5) = Format(D20, "0.000")
            SHELL_OUT(5).ForeColor = &HC0&
            SHELL_OUT(5).BackColor = &HE0E0E0
        End If
    'Water specific heat at SHELL CALORIC TEMP,Kcal/(Kg ºC)
        h1 = 1.00691354509505
        h2 = -1.19506245657282E-03
        h3 = 5.57856020013537E-05
        h4 = -9.75376157602428E-07
        h5 = 6.26080712782905E-09
        SPH_S = h1 + h2 * D17 + h3 * D17 ^ 2 + h4 * D17 ^ 3 + h5 * D17 ^ 4
        If Check_S_SPH = 0 Then
            SHELL_OUT(3).Text = Format(SPH_S, "0.000")
            SHELL_OUT(3).ForeColor = &HFF0000
            SHELL_OUT(3).BackColor = &HE0E0E0
        ElseIf Check_S_SPH = 1 Then
            SPH_S = HScroll_SHELL_SPH / 1000
            SHELL_OUT(3) = Format(SPH_S, "0.000")
            SHELL_OUT(3).ForeColor = &HC0&
            SHELL_OUT(3).BackColor = &HE0E0E0
        End If
    'Water fluid thermal conductivity at SHELL CALORIC TEMP, Kcal/h m ºC
        TH_C = 0.00000000592317 * xd72 ^ 3 - 0.0000080425 * xd72 ^ 2 + 0.0018262 * xd72 + 0.478535
        If Check_S_TC = 0 And Combo_S_FLUID = "Water" Then
            XD79 = TH_C
            'Shell fluid thermal conductivity at film temperature btu/h.ft.ºF
            XD78 = XD79 / 1.488
            SHELL_OUT(1).Text = Format(TH_C, "0.000")
            SHELL_OUT(1).ForeColor = &HFF0000
            SHELL_OUT(1).BackColor = &HE0E0E0
        ElseIf Check_S_TC = 1 Then
            TH_C = HScroll_SHELL_TC / 1000
            XD79 = TH_C
            'Shell fluid thermal conductivity at film temperature btu/h.ft.ºF
            XD78 = XD79 / 1.488
            SHELL_OUT(1) = Format(TH_C, "0.000")
            SHELL_OUT(1).ForeColor = &HC0&
            SHELL_OUT(1).BackColor = &HE0E0E0
        End If
Else
    'Shell fluid density at film temperature,Kg/m^3
        SHELL_OUT(4).ForeColor = &HC0&
        SHELL_OUT(4).BackColor = &HE0E0E0
        If Check_S_DENS = 1 Then
            D19 = HScroll_SHELL_DENS / 10
            XD75 = D19
            SHELL_OUT(4) = Format(D19, "0.0")
        End If
    'Shell fluid viscosity at condensing film temperature,cp
        SHELL_OUT(5).ForeColor = &HC0&
        SHELL_OUT(5).BackColor = &HE0E0E0
        If Check_S_VISC = 1 Then
            D20 = HScroll_SHELL_VISC / 1000
            XD77 = D20
            SHELL_OUT(5) = Format(D20, "0.000")
        End If
    'Water specific heat at SHELL CALORIC TEMP,Kcal/(Kg ºC)
        SHELL_OUT(3).ForeColor = &HC0&
        SHELL_OUT(3).BackColor = &HE0E0E0
        If Check_S_SPH = 1 Then
            SPH_S = HScroll_SHELL_SPH / 1000
            SHELL_OUT(3) = Format(SPH_S, "0.000")
        End If
    'Shell fluid thermal conductivity at SHELL CALORIC TEMP, Kcal/h m ºC
        SHELL_OUT(1).ForeColor = &HC0&
        SHELL_OUT(1).BackColor = &HE0E0E0
        If Check_S_TC = 1 Then
            XD79 = HScroll_SHELL_TC / 1000
            XD78 = XD79 / 1.488
            SHELL_OUT(1) = Format(XD79, "0.000")
        End If
End If
'Tubes flow rate,m3/h
    XD5 = XD6 / XD11
'DUTY
    TUBES_OUT(0) = Format(XD6, "#,##0")
    SHELL_OUT(0) = Format(XD18, "#,##0")
    COL = 0
'Shell side duty,Kcal/h
    Lorg_QS = Lorg_IN * SPH_S * Range_S
    Lwat_QS = Lwat_IN * SPH_T * Range_S
    XD37 = Lorg_QS + Lwat_QS
    xd37_1 = XD37 / 860.04
    Lorg_INP_Q = Lorg_QS / XD37 * 100
    Lwat_INP_Q = Lwat_QS / XD37 * 100
'TUBES side duty,Kcal/h
    XD36 = XD6 * SPH_T * Range_T
    If Thermal_bal_tubes = True Then
         COL = 1
         T_FLW = 1
         SHELL_FLOW.ForeColor = &HC0&
         SHELL_FLOW.BackColor = &H80000018
         TUBES_OUT(0).ForeColor = &HFFFFFF
         TUBES_OUT(0).BackColor = &HC0&
         SHELL_OUT(0).BackColor = &HE0E0E0
         SHELL_OUT(0).ForeColor = &HC0&
         If Check_T_OUT = 1 Then
            TUBES_FLOW.ForeColor = &HFFFFFF
            TUBES_FLOW.BackColor = &HC0&
            TUBES_TEMP_OUT.ForeColor = &HC0&
            TUBES_TEMP_OUT.BackColor = &H80000018
            XD6 = XD37 / (SPH_T * Range_T)
            XD6L = XD6
         ElseIf Check_T_OUT = 0 Then
            TUBES_FLOW.ForeColor = &HC0&
            TUBES_FLOW.BackColor = &H80000018
            TUBES_TEMP_OUT.ForeColor = &HFFFFFF
            TUBES_TEMP_OUT.BackColor = &HC0&
            XD6 = TUBES_FLOW
            Range_S = XD19 - XD20
            Range_T = XD8 - XD7
            Lorg_QS = Lorg_IN * SPH_S * Range_S
            Lwat_QS = Lwat_IN * SPH_T * Range_S
            XD37 = Lorg_QS + Lwat_QS
            XD8 = XD37 / XD6 / SPH_T + XD7
        End If
        TUBES_TEMP_IN = Format(XD7, "0.00")
        TUBES_TEMP_OUT = Format(XD8, "0.00")
        YXY = 1
        Spin_TUBES_T_OUT = XD8 * 100
        YXY = 0
    'Tube side duty, kcal/h
        Range_S = XD19 - XD20
        Range_T = XD8 - XD7
        XD36 = XD6 * SPH_T * Range_T
        TUBES_FLOW = Format(XD6, "#,##0")
        TUBES_LIQUID = Format(XD6, "#,##0")
        TUBES_OUT(0) = Format(XD6, "#,##0")
        HScroll_TUBES_FLOW = XD6 / FACT_FLOW
        TUBES_LIQUID.Text = Format(0, "#,##0")
        TUBES_WATER.Text = Format(XD6, "#,##0")
        HScroll_TUBES_WATER.Max = TUBES_WATER / FACT_FLOW
        XD5 = XD6 / D19
        TUBES_OUT(7).Text = Format(XD36 / 860.04, "#,##0") 'TUBES side duty,Kcal/h, KW
    ElseIf Thermal_bal_shell.Value = True Then
        COL = 1
        S_FLOW = 1
        SHELL_FLOW.BackColor = &HC0&
        SHELL_FLOW.ForeColor = &HFFFFFF
        TUBES_FLOW.BackColor = &H80000018
        TUBES_FLOW.ForeColor = &HC0&
        TUBES_TEMP_OUT.BackColor = &H80000018
        TUBES_TEMP_OUT.ForeColor = &HC0&
        SHELL_OUT(0).BackColor = &HC0&
        SHELL_OUT(0).ForeColor = &HFFFFFF
        TUBES_OUT(0).BackColor = &HE0E0E0
        TUBES_OUT(0).ForeColor = &HC0&
        Range_S = XD19 - XD20
        Range_T = XD8 - XD7
        Lorg_IN_Q = Lorg_IN * SPH_S * Range_S
        Lwat_IN_Q = Lwat_IN * SPH_T * Range_S
        XD37 = Lorg_IN_Q + Lwat_IN_Q
        xd37_1 = XD37 / 860.04
        Lorg_INP_Q = Lorg_IN_Q / XD37 * 100
        Lwat_INP_Q = Lwat_IN_Q / XD37 * 100
        xd36_1 = XD36 / 860.04
        Lorg_IN_1 = XD36 * Lorg_INP_Q / 100 / (SPH_S * Range_S)
        Lwat_IN_1 = XD36 * Lwat_INP_Q / 100 / (SPH_T * Range_S)
        Lwat_IN = Format(Lwat_IN_1, "#,##0")
        Lorg_IN = Format(Lorg_IN_1, "#,##0")
        Ltot_IN = Format(Lorg_IN_1 + Lwat_IN_1, "#,##0")
        Ftot_IN = Format(Ltot_IN, "#,##0")
        XD18 = Ftot_IN
        Lwat_INP = Format(Lwat_IN_1 / XD18 * 100, "0.0")
        Lorg_INP = Format(Lorg_IN_1 / XD18 * 100, "0.0")
    'Shell fluid side duty,Kcal/h
        Lorg_IN_Q = Lorg_IN_1 * SPH_S * Range_S
        Lwat_IN_Q = Lwat_IN_1 * SPH_T * Range_S
        XD37 = Lorg_IN_Q + Lwat_IN_Q
        xd37_1 = XD37 / 860.04
        SHELL_FLOW.Text = Format(XD18, "#,##0")
        SHELL_WATER = Lwat_IN
        SHELL_LIQUID = Lorg_IN
        SHELL_OUT(0) = Format(XD18, "#,##0")
        SHELL_OUT(7) = Format(XD37, "#,##0")
    ElseIf Thermal_bal_shell_T.Value = True Then
        COL = 1
        S_FLOW = 1
        If CSng(SHELL_TEMP_OUT) > CSng(PROCESS_TARGET_T_OUT) Then
            SHELL_TEMP_OUT.ForeColor = &H80000018
            SHELL_TEMP_OUT.BackColor = 255.255
        Else
            SHELL_TEMP_OUT.ForeColor = &H80000018
            SHELL_TEMP_OUT.BackColor = &HC0&
        End If

        Lorg_IN_Q = Lorg_IN * SPH_S * Range_S
        Lwat_IN_Q = Lwat_IN * SPH_T * Range_S
        XD37 = Lorg_IN_Q + Lwat_IN_Q
        xd37_1 = XD37 / 860.04
        Lorg_INP_Q = Lorg_IN_Q / XD37 * 100
        Lwat_INP_Q = Lwat_IN_Q / XD37 * 100
        Lorg_IN_Q1 = Lorg_INP_Q * XD36 / 100
        Lwat_IN_Q1 = Lwat_INP_Q * XD36 / 100
        Range_Sorg = Lorg_IN_Q1 / XD18 / SPH_S
        Range_Swat = Lwat_IN_Q1 / XD18 / SPH_T
        XD20 = XD19 - Range_Sorg - Range_Swat
        SHELL_TEMP_OUT = Format(XD20, "0.00")
        Range_S = XD19 - XD20
        Lorg_IN_Q = Lorg_IN * SPH_S * Range_S
        Lwat_IN_Q = Lwat_IN * SPH_T * Range_S
        XD37 = Lorg_IN_Q + Lwat_IN_Q
        xd37_1 = XD37 / 860.04
    Else
        If CSng(SHELL_TEMP_OUT) > CSng(PROCESS_TARGET_T_OUT) Then
            SHELL_TEMP_OUT.ForeColor = &H80000018
            SHELL_TEMP_OUT.BackColor = 255.255
        Else
            SHELL_TEMP_OUT.BackColor = &H80000018
            SHELL_TEMP_OUT.ForeColor = &HC0&
        End If
        SHELL_OUT(7).ForeColor = &HC0&
        SHELL_OUT(7).BackColor = &HE0E0E0
        SHELL_FLOW.ForeColor = &HC0&
        SHELL_FLOW.BackColor = &H80000018
        TUBES_FLOW.ForeColor = &HC0&
        TUBES_FLOW.BackColor = &H80000018
        TUBES_TEMP_OUT.ForeColor = &HC0&
        TUBES_TEMP_OUT.BackColor = &H80000018
        SHELL_OUT(0).ForeColor = &HC0&
        SHELL_OUT(0).BackColor = &HE0E0E0
        TUBES_OUT(0).ForeColor = &HC0&
        TUBES_OUT(0).BackColor = &HE0E0E0
    End If
    TUBES_OUT(7).Text = Format(XD36 / 860.04, "#,##0") 'Water side duty,Kcal/h, KW
    SHELL_OUT(7).Text = Format(XD37 / 860.04, "#,##0") 'Shell fluid side duty, KW
    XD5 = XD6 / XD11
    YXY = 1
    YXY = 0

'CALCULATING TUBE SIDE PRESSURE DROP
'Pressure drop in tubes,kPA
    'Tube side friction factor   ft^2/inch^2
        XD91 = (XD54 * XPI * XD85 ^ 2 / 4) / XD57  'Flow area tubes,m^2
        If lungh = 2 And XD57 > 1 Then
            XD91 = XD91 * lungh
        End If
    'Flow area tubes,ft^2
        XD92 = XD91 / (0.3048 ^ 2)
    'Water velocity through tubes    m/s
        XD95 = (XD5 / (XD91 * 3600)) / PARALLEL_N
    'Water velocity through tubes,ft/s
        XD96 = XD95 / 0.3048
    'Reynolds number
        XD97 = XD95 * XD85 * XD11 / (XD12 * 0.001)
        XD93 = 10 ^ ((-2.5165 - 0.263 * Log(XD97) / 2.30258))
    'Mass velocity,lb/h(ft^2)
        XD94 = XD5 * XD11 * 2.20462 / XD92
        XD98 = (XD93 * XD94 ^ 2 * XD56 * XD57) / (5.22 * 10 ^ 10 * XD84)
    'Pressure drop in tubes  bar
        XD99 = XD98 * 0.068947
    'Pressure drop in tubes  Kg/cm2
        XD100 = XD98 * 0.070307
    'Tube side pressure drop due to changes in direction,psi
        XD101 = (4 * XD57) * (XD96 ^ 2 / (2 * (9.81 / 0.3048))) * (62.5 / 144)
    'Tube side pressure drop due to changes in direction, bar
        XD102 = XD101 * 0.068947
    'Tube side pressure drop due to changes in direction Kg/cm2
        XD103 = XD101 * 0.070307
    'Total pressure drop for 100% clean tube side    psi
        XD104 = XD98 + XD101
    'Total pressure drop for 100% clean tube side    bar
        XD105 = XD99 + XD102
    'Total pressure drop for 100% clean tube side    Kg/cm2
        XD106 = XD100 + XD103
    'Total pressure drop for 100% clean tube side    kPa
        XD106_KPA = XD105 * 100
    'Tubes section
        TUBES_SECTION.Text = Format(XD91, "0.0000")
    'Water velocity through tubes    m/s
        TUBES_OUT(2) = Format(XD95, "0.00")
    'Reynolds number through tubes
        TUBES_OUT(6) = Format(XD97, "#,##0")
    'Total pressure drop
        If Check_P_DROP_T = Unchecked Then
            TUBES_OUT(9) = Format(XD106_KPA, "0.00")
        Else
            XD106_KPA = (TUBES_P_IN - TUBES_P_OUT) * 100
            TUBES_OUT(9).Text = Format(XD106_KPA, "0.00") ' KPa
        End If
    C_F = XD5 / (XD106_KPA) ^ (1 / 2)
    C_Factor.Text = Format(C_F, "0.0")

'CALCULATING HEAT TRANSFER
        
    'Water side individual heat transfer coeficient  Btu/(h ft^2 F)
        XD108 = 150 * (1 + 0.011 * XD10) * (XD96 ^ 0.8 / XD83 ^ 0.2)
    'Water side individual heat transfer coefficient  Kcal/(h m^2 C)
        XD109 = XD108 * 4.882
    'Water side indiv. heat transfer coefficient referred to ext. surface Kcal/(h m^2 C)
        XD110 = XD109 * (XD85 / XD66)
    'Heat transfer resistance due to the wall    [(hm^2ºC)/Kcal]*10^4
        XD111 = (XD66 * Log(XD66 / XD85) / (2 * XD58)) * 10000
        XD111B = 1 / (XD111 / 10000)
    'Heat transfer resistance due to outside fouling factor  [(h m^2 ºC)/Kcal]*10^4  1.00
        'XD112 = XD30
    'Heat transfer resistance due to water (tube side)   [(hm^2ºC)/Kcal]*10^4
        XD114 = 10 ^ 4 / XD110

'PROP = "SHELL"
'Call Properties

S_T_OUT = XD20

'CALCULATING h_o
    'Shell side crossflow area, ft^2
300     XD68 = XD61 * (XD52 - XD50) * XD64 / (XD52 * 144)
    'Shell fluid loading  lb/h ft
        XD69 = XD18 * 2.20462 / (XD56 * XD54 ^ (2 / 3))
    'Shell side indiv. heat transfer coefficient,Kcal/(h m^2 C)
        'Shell side indiv. heat transfer coefficient,Btu/(h ft^2 F)
    'Shell side indiv. heat transfer coefficient (guess one till Z_0=0),Kcal/(hm^2°C)

400     W = 0.1
410     j = 0.1
420     HE = W: GoSub 480
430     Y = X: HE = j + W
440     GoSub 480
450     g = W: W = g - j * Y / (X - Y)
460     If Abs(g - W) >= 0.00001 Then GoTo 420
470     W = HE: GoTo 500
480     xd72 = HE
        
    'Shell side film temperature ºC
        XD73 = (S_T_OUT + xd72) / 2
    'Shell side film temperature,K
        XD74 = 273.16 + XD73
    If Combo_S_FLUID = "Water" Then
        'Shell fluid density at film temperature,Kg/m^3
            D19 = 0.0002 * xd72 ^ 3 - 0.028 * xd72 ^ 2 + 0.0873 * xd72 + 999.92
            If Check_S_DENS = 0 Then
                XD75 = D19
                SHELL_OUT(4).Text = Format(D19, "0.0")
                SHELL_OUT(4).ForeColor = &HFF0000
                SHELL_OUT(4).BackColor = &HE0E0E0
            ElseIf Check_S_DENS = 1 Then
                D19 = HScroll_SHELL_DENS / 10
                XD75 = D19
                SHELL_OUT(4) = Format(D19, "0.0")
                SHELL_OUT(4).ForeColor = &HC0&
                SHELL_OUT(4).BackColor = &HE0E0E0
            End If
            'Shell fluid density at film temperature,lb/ft^3
            XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
        'Shell fluid viscosity at condensing film temperature,cp
            D20 = (100 / (2.1482 * ((273.16 + xd72 - 281.435) + Sqr(8078.4 + (273.16 + xd72 - 281.435) ^ 2)) - 120))
            If Check_S_VISC = 0 Then
                XD77 = D20
                SHELL_OUT(5).Text = Format(D20, "0.000")
                SHELL_OUT(5).ForeColor = &HFF0000
                SHELL_OUT(5).BackColor = &HE0E0E0
            ElseIf Check_S_VISC = 1 Then
                D20 = HScroll_SHELL_VISC / 1000
                XD77 = D20
                SHELL_OUT(5) = Format(D20, "0.000")
                SHELL_OUT(5).ForeColor = &HC0&
                SHELL_OUT(5).BackColor = &HE0E0E0
            End If
        'Water specific heat at SHELL CALORIC TEMP,Kcal/(Kg ºC)
            h1 = 1.00691354509505
            h2 = -1.19506245657282E-03
            h3 = 5.57856020013537E-05
            h4 = -9.75376157602428E-07
            h5 = 6.26080712782905E-09
            SPH_S = h1 + h2 * xd72 + h3 * xd72 ^ 2 + h4 * xd72 ^ 3 + h5 * xd72 ^ 4
            If Check_S_SPH = 0 Then
                SHELL_OUT(3).Text = Format(SPH_S, "0.000")
                SHELL_OUT(3).ForeColor = &HFF0000
                SHELL_OUT(3).BackColor = &HE0E0E0
            ElseIf Check_S_SPH = 1 Then
                SPH_S = HScroll_SHELL_SPH / 1000
                SHELL_OUT(3) = Format(SPH_S, "0.000")
                SHELL_OUT(3).ForeColor = &HC0&
                SHELL_OUT(3).BackColor = &HE0E0E0
            End If
        'Shell fluid thermal conductivity at SHELL CALORIC TEMP, Kcal/h m ºC
            XD79 = 0.00000000592317 * xd72 ^ 3 - 0.0000080425 * xd72 ^ 2 + 0.0018262 * xd72 + 0.478535
            If Check_S_TC = 0 And Combo_S_FLUID = "Water" Then
                XD79 = TH_C
                'Shell fluid thermal conductivity at film temperature btu/h.ft.ºF
                XD78 = XD79 / 1.488
                SHELL_OUT(1).Text = Format(TH_C, "0.000")
                SHELL_OUT(1).ForeColor = &HFF0000
                SHELL_OUT(1).BackColor = &HE0E0E0
            ElseIf Check_S_TC = 1 Then
                TH_C = HScroll_SHELL_TC / 1000
                XD79 = TH_C
                'Shell fluid thermal conductivity at film temperature btu/h.ft.ºF
                XD78 = XD79 / 1.488
                SHELL_OUT(1) = Format(TH_C, "0.000")
                SHELL_OUT(1).ForeColor = &HC0&
                SHELL_OUT(1).BackColor = &HE0E0E0
            End If
    Else
        'Shell fluid density at film temperature,Kg/m^3
            If Check_S_DENS = 1 Then
                D19 = HScroll_SHELL_DENS / 10
                XD75 = D19
                SHELL_OUT(4) = Format(D19, "0.0")
                SHELL_OUT(4).ForeColor = &HC0&
                SHELL_OUT(4).BackColor = &HE0E0E0
                'Shell fluid density at film temperature,lb/ft^3
                XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
            End If
        'Shell fluid viscosity at condensing film temperature,cp
            If Check_S_VISC = 1 Then
                D20 = HScroll_SHELL_VISC / 1000
                XD77 = D20
                SHELL_OUT(5) = Format(D20, "0.000")
                SHELL_OUT(5).ForeColor = &HC0&
                SHELL_OUT(5).BackColor = &HE0E0E0
            End If
        'Water specific heat at SHELL CALORIC TEMP,Kcal/(Kg ºC)
            If Check_S_SPH = 1 Then
                SPH_S = HScroll_SHELL_SPH / 1000
                SHELL_OUT(3) = Format(SPH_S, "0.000")
                SHELL_OUT(3).ForeColor = &HC0&
                SHELL_OUT(3).BackColor = &HE0E0E0
            End If
        'Shell fluid thermal conductivity at SHELL CALORIC TEMP, Kcal/h m ºC
            If Check_S_TC = 1 Then
                XD79 = HScroll_SHELL_TC / 1000
                XD78 = XD79 / 1.488
                SHELL_OUT(1) = Format(XD79, "0.000")
                SHELL_OUT(1).ForeColor = &HC0&
                SHELL_OUT(1).BackColor = &HE0E0E0
            End If
    End If
        'Shell side indiv. heat transfer coefficient,Btu/(h ft^2 F)
            XD80 = 1.5 * ((4 * XD69 / XD77) ^ -(1 / 3)) * (XD77 ^ 2 / (XD78 ^ 3 * XD76 ^ 2 * 9.81 * (3600 ^ 2 / 0.3048))) ^ -(1 / 3)
        'Shell side indiv. heat transfer coefficient,Kcal/(h m^2 C)
            XD81 = XD80 * 4.882
        'Heat transfer resistance due to process (shell side),[(hm^2ºC)/Kcal]*10^4
            XD115 = (1 / XD81) * 10000
        'Overall CLEAN heat transfer coefficient Kcal/(h m^2 ºC)
            XD117 = 10000 / (XD111 + XD114 + XD115)
            XD117B = 1 / (1 / XD111B + 1 / XD110 + 1 / XD81)
        'Wall temperature, °C
            '(tw2-t2)/(1/alfa2)=(T1-t2)/(1/OHTC)
            'tw2 = (T1-t2)/(1/OHTC)*(1/alfa2)+ t2
            'tw2: wall temperature cold side
            't2: temperature cold side (here outlet)
            'alfa2: convective heat transfer cold side
            'T1: hot side temperature (inlet)
            'OHTC: overall heat transfer coefficient.
            tw2a = (XD19 - XD8) / (1 / XD117) * (1 / XD109) + XD8
            SKIN_TEMP = Format(tw2a, "0.00")
        X = (tw2a - xd72)
    Return
500  If Combo_S_FLUID = "Water" Then
        SHELL_OUT(1) = Format(XD79, "0.000")       'Shell fluid thermal conductivity at condensing film temperature Kcal/h m ºC
        SHELL_OUT(3) = Format(SPH_S, "0.000")      'Shell specific heat
        SHELL_OUT(4) = Format(XD75, "0.0")         'Shell fluid density at condensing film temperature,Kg/m^3
        SHELL_OUT(5) = Format(XD77, "0.000")       'Shell fluid viscosity at condensing film temperature,cp
        SHELL_OUT(11) = Format(0, "0.00")          'Shell fluid film temperature ºC
        SKIN_TEMP = Format(xd72, "0.00")           'Wall temperature,ºC
    End If
'SHELL VELOCITY
    L37 = XD61M                     'SHELL_ID.Text, mm
    O37 = XD52M                     'SHELL_TUBES_PITCH, mm
    E37 = XD66M                     'T_OD, mm
    N37 = XD64M                     'SHELL_BAFFLES_SPACE, mm
'SHELL CLEARANCE, m
    V37 = O37 / 1000 - E37 / 1000
    Clearance.Text = Format(V37, "0.0000")
'SHELL FLOW AREA, m2
    U37 = L37 / 1000 * V37 * N37 / 1000 / (O37 / 1000)
    Flow_area.Text = Format(U37, "0.0000")
'SHELL VELOCITY
    k47 = XD18 / XD75 / U37 / 3600 / PARALLEL_N * XD59
    SHELL_OUT(2) = Format(k47, "0.00")
'SHELL Reynolds
    'SHELL_TUBES_PITCH / 1000
    EQ_D19 = XD52M / 1000
    EQ_PI = XPI
    EQ_D14 = XD66  'T_OD / 1000
    'Equivalent diameter, m
    If SHELL_PITCH_CONF = "Triangular" Then
        EQ_E31 = 4 * (EQ_D19 ^ 2 - EQ_PI * EQ_D14 ^ 2 / 4) / (XPI * EQ_D14)
    Else
        EQ_E31 = (4 * (0.5 * EQ_D19 * 0.866 * EQ_D19 - 0.5 * XPI * EQ_D14 ^ 2 / 4) / (0.5 * XPI * EQ_D14))
    End If
    EQ_E29 = U37                      'Flow_area
    EQ_E25 = SHELL_WATER + SHELL_LIQUID     'SHELL_FLOW WATER + SHELL_FLOW LIQUI
    EQ_E30 = EQ_E25 / EQ_E29
    EQ_E8 = XD77 * 3.6                'Shell fluid viscosity at condensing film temperature,cp
    EQ_E32 = EQ_E31 * EQ_E30 / EQ_E8
    Q_E22 = EQ_E31                    'Equivalent diameter
    Q_E17 = XD75                      'Shell fluid density at condensing film temperature,Kg/m^3
    Q_E27 = k47                       'Shell fluid velocity, m/s
    Q_E18 = XD77                      'Shell fluid viscosity at condensing film temperature,cp
    Q_E28 = Q_E22 * Q_E17 * Q_E27 / (Q_E18 * 0.001)
    SHELL_OUT(6) = Format(Q_E28, "#,##0")
'CALCULATING SHELL SIDE PRESSURE DROP
    If Check_P_DROP_S = Unchecked Then
    'Pressure drop (tubes)
        P_E17 = XD75                    'Shell density
        P_E22 = EQ_E31 * 1000           'Equivalent diameter, mm
        P_E27 = Q_E27                   'Flow velocity, m/s
        P_E23 = XD55 * 1000 * lungh     'Tubes lenght,mm
        P_E28 = Q_E28                   'Reynolds number
        P_E29 = 0.44 * P_E28 ^ -0.19    'Friction factor
        P_E30 = 4 * P_E29 * P_E23 * P_E27 ^ 2 / (P_E22 * 2 * 9.8) * P_E17 * 0.000096784 * 101.325
    'Pressure drop (sheet)
        P_E9 = XD59                     'Shell passes
        P_E31 = 3 * P_E9 * P_E27 ^ 2 / 2 / 9.8 * P_E17 * 0.000096784 * 101.325
        P_E32 = (P_E30 + P_E31)
    'Pressure drop (sheet), KPa
        SHELL_OUT(9).Text = Format(P_E32, "0.00")
    Else
    'Water_press_drop_bar.Text = Format(P_E32 / 100, "0.00"), bar
        P_E32 = (SHELL_P_IN - SHELL_P_OUT)
    'Water_press_drop_bar.Text = Format(P_E32 / 100, "0.00"), KPa
        SHELL_OUT(9).Text = Format(P_E32 * 100, "0.00")
    'Water_press_drop_bar.Text = Format(P_E32, "0.00"), bar
    End If
        
'Heat transfer resistance due to process (shell side),[(hm^2ºC)/Kcal]*10^4
        XD115 = (1 / XD81) * 10000
'Overall CLEAN heat transfer coefficient Kcal/(h m^2 ºC)
        XD117 = 10000 / (XD111 + XD114 + XD115)
'Water side fouling factor   [(hm^2ºC)/Kcal]*10^4
    'Overall heat transfer coefficient   Kcal/(h m^2 ºC)
    'Surface per linear ft, ft^2
        'Tube outlet diameter,inch
        XD50 = Format(XD66 * 1000 / 25.4, "0.000")
    'Surface per linear m, m^2
        XD90 = 3.14159 * D67 * 25.4 * 10 ^ -3 * D75 * D74 * SERIES_N * PARALLEL_N * lungh
    'Log Mean Temperature Difference CORRECTED, ºC
        If Combo_CURRENT = "Counter-flow" Then
            AG6 = ((XD19 - XD8) - (XD20 - XD7)) / Log((XD19 - XD8) / (XD20 - XD7))
        ElseIf Combo_CURRENT = "Cross-flow" Then
            AG6 = ((XD20 - XD7) - (XD19 - XD8)) / Log((XD20 - XD7) / (XD19 - XD8))
        ElseIf Combo_CURRENT = "Parallel-flow" Then
            AG6 = ((XD19 - XD7) - (XD20 - XD8)) / Log((XD19 - XD7) / (XD20 - XD8))
        End If
        AJ6 = (XD19 - XD20) / (XD8 - XD7)
        AK6 = (XD8 - XD7) / (XD19 - XD7)
        RR = AJ6
        ss = AK6
        If T_PASS > 1 And SHELL_PASS > 1 Then
            FT1 = (Sqr(RR ^ 2 + 1) / (2 * (RR - 1))) * Log((1 - ss) / (1 - RR * ss))
            FT2 = 2 / ss - 1 - RR + (2 / ss) * Sqr((1 - ss) * (1 - RR * ss)) + Sqr(RR ^ 2 + 1)
            FT3 = 2 / ss - 1 - RR + (2 / ss) * Sqr((1 - ss) * (1 - RR * ss)) - Sqr(RR ^ 2 + 1)
            FT4 = Log(FT2 / FT3)
            FT = FT1 / FT4
        ElseIf T_PASS > 1 And SERIES_N > 1 Then
            FT1 = (Sqr(RR ^ 2 + 1) / (2 * (RR - 1))) * Log((1 - ss) / (1 - RR * ss))
            FT2 = 2 / ss - 1 - RR + (2 / ss) * Sqr((1 - ss) * (1 - RR * ss)) + Sqr(RR ^ 2 + 1)
            FT3 = 2 / ss - 1 - RR + (2 / ss) * Sqr((1 - ss) * (1 - RR * ss)) - Sqr(RR ^ 2 + 1)
            FT4 = Log(FT2 / FT3)
            FT = FT1 / FT4
        Else
            AL6 = Sqr(AJ6 ^ 2 + 1) * Log((1 - AK6) / (1 - AJ6 * AK6))
            AM6 = 2 - AK6 * (AJ6 + 1 - Sqr(AJ6 ^ 2 + 1))
            AN6 = 2 - AK6 * (AJ6 + 1 + Sqr(AJ6 ^ 2 + 1))
            FT = AL6 / ((AJ6 - 1) * Log(AM6 / AN6))
        End If
        AH6 = AG6 * FT
        'Approach temperature  ºC
        If XD57 = 1 Then
            XD32 = XD19 - XD8
            Label22(15).Caption = "(T1 - t2)"
            If XD32 > XD20 - XD7 Then
                XD32 = XD20 - XD7
                Label22(15).Caption = "(T2 - t1)"
            End If
        ElseIf XD57 > 1 Then
            XD32 = XD20 - XD8
            Label22(15).Caption = "(T2 - t2)"
        End If
        XD31 = AH6
        XD38 = XD36 / (XD90 * XD31)
'Overall CLEAN heat transfer coefficient Kcal/(h m^2 ºC)
    If Check_U_CLEAN = Checked Then
        XD117 = HScroll_U_CLEAN
        U_COEFF_CLEAN = XD117
    ElseIf Check_U_CLEAN = Unchecked Then
'Overall CLEAN heat transfer coefficient Kcal/(h m^2 ºC)
        U_COEFF_CLEAN.Text = Format(XD117, "0.0")
    End If
'Water side fouling factor   [(hm^2ºC)/Kcal]*10^3
    xd118 = ((1 / XD38) - (1 / XD117) - (XD112 / 10000)) * 10000 * (XD85 / XD66)
'Total heat transfer resistance  [(h m^2 ºC)/Kcal]*10^4
    XD116 = 10000 / XD38
'Heat transfer resistance due to inside fouling factor,[(hm^2ºC)/Kcal]*10^4
    XD113 = xd118 * (XD66 / XD85)
'TUBES HEAT FLUX
    Q6 = XD36 / XD90
    TUBES_OUT(8).Text = Format(Q6 * 0.001163, "0.00")
'SHELL HEAT FLUX
    SPH_S = SHELL_OUT(3).Text
    Q6S = XD37 / XD90
    SHELL_OUT(8).Text = Format(Q6S * 0.001163, "0.00")
'Area, m^2
    Area.Text = Format(XD90, "0.00")
'Log Mean Temperature Difference, ºC
    LMTD.Text = Format(AG6, "0.00")
'Log Mean Temperature Difference corrected, ºC
    MTDc.Text = Format(AH6, "0.00")
    TTD = Format(XD32, "0.00")
    TUBES_OUT(7).Text = Format(XD36 / 860.04, "#,##0") 'Water side duty,Kcal/h, KW
    SHELL_OUT(7).Text = Format(XD37 / 860.04, "#,##0") 'Shell fluid side duty, KW
    U_COEFF_DIRTY.Text = Format(XD38, "0.0") 'Overall heat transfer coefficient   Kcal/(h m^2 ºC)
'Water side fouling factor, [(hm^2ºC)/Kcal]*10^3
    TUBES_FF.Text = Format(xd118, "0.000")   'Water side fouling factor   [(hm^2ºC)/Kcal]*10^3
    'TUBES_FF.Text = Format(XD118, "0.000")  'Water side fouling factor   [(hm^2ºC)/Kcal]*10^3
End Sub
Private Sub Steam()
On Error Resume Next
Combo_CURRENT = "Condensation"

'Vapor percent
    Frame_VAP.Visible = True
    SHELL_VAPOR.Visible = True
    SHELL_LIQUID.Visible = True
    SHELL_WATER.Visible = True
    SHELL_NON_COND.Visible = True
    HScroll_SHELL_NON_COND.Visible = False
    Spin_VAP_P.Visible = True
    Spin_WAT_VAP_IN.Visible = True
    Spin_WAT_LIQ_IN.Visible = True
'Wet steam
    Label22(17).Visible = True
    Wet_steam.Visible = True
'Temperatures
    SHELL_TEMP_IN.Visible = True
    SHELL_TEMP_OUT.Visible = True
    HScroll_SHELL_T_IN.Visible = True
    HScroll_SHELL_T_OUT.Visible = True
    Check_T_OUT.Visible = True
'Pressure
    SHELL_P_IN.Visible = True
    SHELL_P_OUT.Visible = True
    HScroll_SHELL_P_IN.Visible = True
    HScroll_SHELL_P_OUT.Visible = True
    Check_P_DROP_S.Visible = True
'SPECIFIC HEAT
    Check_S_SPH.Visible = True
    SHELL_OUT(3).Visible = True
    HScroll_SHELL_SPH.Visible = True
'Termal conductivity SHELL
    lbl_tubes(9).Visible = True
    SHELL_OUT(1).Visible = True
    Check_S_TC.Visible = True
    HScroll_SHELL_TC.Visible = True
    lbl_tubes(1).Visible = True
'Termal conductivity TUBES
    lbl_tubes(1).Visible = True
    lbl_tubes(9).Visible = True
    TUBES_OUT(1).Visible = True
    Check_T_TC.Visible = True
    HScroll_TUBES_TC.Visible = True
'Condensing steam  pressure
    lblLabels(10).Visible = True
    S_press_KP.Visible = True
    Check_CP.Visible = True
    Spin_S_PRESS.Visible = True
    Label1(10).Visible = True
'Condensing steam  temperature
    lblLabels(22).Visible = True
    SHELL_OUT(11).Visible = True
    Check_CT.Visible = True
    HScroll_C_TEMP.Visible = True
    Label22(11).Visible = True
'Latent heat
    lblLabels(1).Visible = True
    Check_LATENT.Visible = True
    SHELL_OUT(10).Visible = True
    HScroll_LATENT.Visible = True
    Label22(0).Visible = True
'Material factor
    lblLabels(34).Visible = True
    Mat_factor.Visible = True
    Spin_MAT_FACTOR.Visible = True
    Check_MAT_FACTOR.Visible = True
    lblLabels(16).Visible = False
'Skin temperature
    Label22(3).Visible = True
    SKIN_TEMP.Visible = True
    Label22(4).Visible = True
'Shell flow velocity
    Label22(26).Visible = True
    SHELL_OUT(2).Visible = True
'Shell Reynolds number
    SHELL_OUT(6).Visible = True
'Cleanliness factor
    lblLabels(1).Caption = "Latent heat:"
    lblLabels(3).Caption = "Cleanliness factor"
    Label24(0).Caption = "%  (Norm: 85 - 95 %)"
    lblLabels(25).Caption = "Terminal temperature:"
'Thermal balance
    Thermal_bal_tubes.Visible = True
    Thermal_bal_shell.Visible = True
    Thermal_bal_shell_T.Visible = False

    D33 = Wet_steam
'Total shell flowrate, Kg/h
    Ftot_IN_1 = CDbl(SHELL_FLOW)
    Ftot_IN = Ftot_IN_1
    XD18 = Ftot_IN
'TOTAL VAPOR IN, Kg/h
    Vtot_INP_1 = Vtot_INP
    Vtot_IN_1 = Vtot_INP_1 / 100 * Ftot_IN_1
    Vtot_IN = Format(Vtot_IN_1, "#,##0")
'VAPOR NOT CONDENSING IN = VAP_NON-CONDENSING OUT, Kg/h
    HScroll_SHELL_NON_COND.Max = Vtot_IN / FACT_FLOW
    NC_IN_1 = HScroll_SHELL_NON_COND * FACT_FLOW
    NC_IN = Format(NC_IN_1, "#,##0")
'CONDENSABLE WATER VAPOR IN, Kg/h
    Vwat_INP_1 = Vwat_INP
    Vwat_IN_1 = Vwat_INP_1 / 100 * Vtot_IN_1
    Vwat_IN = Format(CDbl(Vwat_IN_1), "#,##0")
    Vwat_INP_1 = Vwat_IN_1 * 100 / Vtot_IN_1
    Vwat_INP = Format(CDbl(Vwat_INP_1), "0.0")
'TOTAL LIQUID IN, Kg/h
    Ltot_IN_1 = Ftot_IN_1 - Vtot_IN_1
    Ltot_IN = Format(Ltot_IN_1, "#,##0")
    Ltot_INP_1 = 100 - Vtot_INP_1
    Ltot_INP = Format(Ltot_INP_1, "0.0")

'CONDENSABLE ORGANIC VAPOR IN (NOT WATER)
    Vorg_IN_1 = Vtot_IN - NC_IN - Vwat_IN_1
    Vorg_IN = Format(CDbl(Vorg_IN_1), "#,##0")
    Vorg_INP_1 = 100 - Vwat_INP_1
    Vorg_INP = Format(Vorg_INP_1, "0.0")
'WATER LIQUID IN, Kg/h
    Lwat_INP_1 = 100   'Lwat_IN_1 * 100 / Ltot_IN_1
    Lwat_INP = Format(Lwat_INP_1, "0.0")
    Lwat_IN_1 = Lwat_INP_1 / 100 * Ltot_IN_1
    Lwat_IN = Format(CDbl(Lwat_IN_1), "#,##0")
    Lwat_IN_1 = Lwat_INP_1 / 100 * Ltot_IN_1
    Lwat_IN = Format(CDbl(Lwat_IN_1), "#,##0")
    wet_steam_1 = Vtot_INP_1 / 100 '(Lwat_IN_1 / Ftot_IN * 100)
    Wet_steam = Format(wet_steam_1, "0.000")
'ORGANIC LIQUID IN, Kg/h
    Lorg_IN_1 = Ltot_IN_1 - Lwat_IN_1
    Lorg_IN = Format(Lorg_IN_1, "#,##0")
    Lorg_INP_1 = 100 - Lwat_INP_1
    Lorg_INP = Format(Lorg_INP_1, "0.0")
'TOTAL LIQUID OUT, Kg/h
    Ltot_out = Vorg_IN_1 + Vwat_IN_1 + Lorg_IN_1 + Lwat_IN_1
    Lorg_OUT = Lorg_IN_1
    Lwat_OUT = Vwat_IN_1 + Lwat_IN_1
        SHELL_VAPOR = Format(Vorg_IN_1, "#,##0")
        SHELL_LIQUID = Format(Lorg_OUT, "#,##0")
        SHELL_WATER = Format(Lwat_OUT, "#,##0")
        SHELL_NON_COND = Format(NC_IN, "#,##0")
'TOTAL FLOW
    Ftot_IN = Format(Vorg_IN_1 + Vwat_IN_1 + Lorg_IN_1 + Lwat_IN_1, "#,##0")
    Vorg_INP_1 = Vorg_IN_1 / Ftot * 100
    Ftot_INP = Format(Vtot_INP_1 + Ltot_INP_1, "0.0")
'Percent of fractions OUT shell side
    Lorg_0UTP = Lorg_OUT / XD18 * 100
    Lwat_OUTP = Lwat_OUT / XD18 * 100
    NC_OUTP = NC_IN / XD18 * 100
    Ltot_OUTP = Lorg_0UTP + Lwat_OUTP + NC_OUTP

Call Mechanical

'Surface per linear m, m^2
    D79 = 3.14159 * D67 * 25.4 * 10 ^ -3 * D75 * D74 * SERIES_N * PARALLEL_N * lungh
'Surface per linear inch, inch^2
    D80 = D79 / (0.3048 ^ 2)
'Surface, m^2
    Area.Text = Format(D79, "0.0")
    If Check_des = Checked Then
        'Water T_OUT, °C
        D13 = TUBES_TEMP_OUT.Text
    End If
'Water T_IN, °C
    D11 = TUBES_TEMP_IN.Text
'Water flow rate, kg/h
    D9 = CDbl(TUBES_FLOW.Text)
'Steam flow rate, kg/h
    D29 = CDbl(SHELL_FLOW.Text)
'Water T_IN, °F
    D12 = D11 * 1.8 + 32
'Steam Condensation pressure, bar
    D32 = S_press_KP_1 / 100
    If D32 = 0 Then D32 = 0.05
    S_press_KP.ForeColor = &HC0&
    S_press_KP.BackColor = &HE0E0E0

PROP = "TUBES"
Call Properties

'TUBES Caloric temperature,ºC
    D17 = XD7 + (XD8 - XD7) / 2
'TUBES Caloric temperature,ºF
    XD10 = XD9 * 1.8 + 32

'Water density at TUBES CALORIC TEMP,Kg/m3 ((t1+t2)/2)
    D19 = 0.0002 * D_17 ^ 3 - 0.028 * D17 ^ 2 + 0.0873 * D17 + 999.92
    If Check_T_DENS = 0 Then
        TUBES_OUT(4).Text = Format(D19, "0.0")
        TUBES_OUT(4).ForeColor = &HFF0000
        TUBES_OUT(4).BackColor = &HE0E0E0
    ElseIf Check_T_DENS = 1 Then
        D19 = HScroll_TUBES_DENS / 10
        TUBES_OUT(4) = Format(D19, "0.0")
        TUBES_OUT(4).ForeColor = &HC0&
        TUBES_OUT(4).BackColor = &HE0E0E0
    End If
'Water viscosity at tubes caloric temp. (t1+t2)/2,centipoise
    D20 = (100 / (2.1482 * ((273.16 + D17 - 281.435) + Sqr(8078.4 + (273.16 + D17 - 281.435) ^ 2)) - 120))
    If Check_T_VISC = 0 Then
        TUBES_OUT(5).Text = Format(D20, "0.000")
        TUBES_OUT(5).ForeColor = &HFF0000
        TUBES_OUT(5).BackColor = &HE0E0E0
    ElseIf Check_T_VISC = 1 Then
        D20 = HScroll_TUBES_VISC / 1000
        TUBES_OUT(5) = Format(D20, "0.000")
        TUBES_OUT(5).ForeColor = &HC0&
        TUBES_OUT(5).BackColor = &HE0E0E0
    End If
'Water specific heat at TUBES CALORIC TEMP,Kcal/KgºC
    h1 = 1.00691354509505
    h2 = -1.19506245657282E-03
    h3 = 5.57856020013537E-05
    h4 = -9.75376157602428E-07
    h5 = 6.26080712782905E-09
    SPH_T = h1 + h2 * D17 + h3 * D17 ^ 2 + h4 * D17 ^ 3 + h5 * D17 ^ 4
    If Check_T_SPH = 0 Then
        TUBES_OUT(3).Text = Format(SPH_T, "0.000")
        TUBES_OUT(3).ForeColor = &HFF0000
        TUBES_OUT(3).BackColor = &HE0E0E0
    ElseIf Check_T_SPH = 1 Then
        SPH_T = HScroll_TUBES_SPH / 1000
        TUBES_OUT(3) = Format(SPH_T, "0.000")
        TUBES_OUT(3).ForeColor = &HC0&
        TUBES_OUT(3).BackColor = &HE0E0E0
    End If
'Water fluid thermal conductivity at TUBES CALORIC TEMP, Kcal/h m ºC
    TH_C = 0.00000000592317 * D17 ^ 3 - 0.0000080425 * D17 ^ 2 + 0.0018262 * D17 + 0.478535
    If Check_T_TC = 0 Then
        TUBES_OUT(1).Text = Format(TH_C, "0.000")
        TUBES_OUT(1).ForeColor = &HFF0000
        TUBES_OUT(1).BackColor = &HE0E0E0
    ElseIf Check_T_TC = 1 Then
        TH_C = HScroll_TUBES_TC / 1000
        TUBES_OUT(1) = Format(TH_C, "0.000")
        TUBES_OUT(1).ForeColor = &HC0&
        TUBES_OUT(1).BackColor = &HE0E0E0
    End If

'VAPOR PRESSURE, kPa   (kayelaby.npl.co.uk/chemistry/3_4/3_4_2.html)
'ln(p/kPa) = 16.166 29 - 3736.276/((T/K) - 49.577) 95 - 105 °C
'ln(p/pc) = (a1*t + a2*t^1.5 + a3*t^3 + a4*t^3.5 + a5*t^4 + a6*t^7.5)*Tc/T (0 - 360°C)
'where p is the pressure, T = T90, and subscript c indicates the values at the critical point;
't = 1 - T/Tc. The values for substitution in the equation are:
'Tc = 647.096K     pc = 220 64 kPa     a1 = -7.85951783     a2 = 1.84408259
'a3 = -11.7866497;  a4 = 22.6807411;  a5 = -15.9618719;  a6 = 1.80122502
'Pressures are tabulated in kPa for Celsius temperatures t90.
'For example, at 34 °C the pressure is 5.3240 kPa and at 102°C is 108.87 kPa.
T_IN = 200
Tc = 647.096
Pc = 22064
t = 1 - (T_IN + 273.16) / Tc
A1 = -7.85951783
a2 = 1.84408259
a3 = -11.7866497
a4 = 22.6807411
a5 = -15.9618719
a6 = 1.80122502
P_1 = (A1 * t + a2 * t ^ 1.5 + a3 * t ^ 3 + a4 * t ^ 3.5 + a5 * t ^ 4 + a6 * t ^ 7.5) * Tc / T_IN
p = Exp(P_1) * Pc
'S_press_KP = Format(P, "0.00")
'Vapor pressure, bar
'D32 = P / 100

'Latent heat, Kcal/kg
    I9 = 0.168682569821809
    J9 = -1.80896828868017E-04
    J3 = -38.2917529410035
    D38_1 = (-I9 - Sqr(I9 ^ 2 - 4 * J9 * (J3 - Log(D32) / 2.3))) / (2 * J9)
    D38 = D38_1 * D33 + (1 - D33) * SPH_T
'Latent heat,kJ/kg
    D38_2 = D38 * 4.1868
    If Check_LATENT = 0 Then
        SHELL_OUT(10).Text = Format(D38_2, "0.0")
        SHELL_OUT(10).ForeColor = &HFF0000
        SHELL_OUT(10).BackColor = &HE0E0E0
        HScroll_LATENT = D38_2 * 4.1868
    ElseIf Check_LATENT = 1 Then
        D38_2 = HScroll_LATENT
        D38 = D38_2 / 4.1868
        SHELL_OUT(10) = Format(D38_2, "0.0")
        SHELL_OUT(10).ForeColor = &HC0&
        SHELL_OUT(10).BackColor = &HE0E0E0
    End If

'Condensed steam temperature if 0.023<P_c<0.067
    If D32 >= 0.0254 And D32 <= 0.0643 Then
        J3 = -4.15371836241458
        I9 = 1500.11867695013
        J9 = -25402.3185815265
        K9 = 259023.283197929
        L9 = -1093047.81628863
        D34 = J3 + I9 * D32 + J9 * D32 ^ 2 + K9 * D32 ^ 3 + L9 * D32 ^ 4
    Else
        D34 = 0
    End If
'Condensed steam temperature if 0.067<P_c<0.1
    If D32 > 0.0643 And D32 <= 0.1 Then
        J3 = 6.95024724711694
        I9 = 748.012008925124
        J9 = -5810.38128441877
        K9 = 27768.5506851489
        L9 = -55837.443148253
        D35 = J3 + I9 * D32 + J9 * D32 ^ 2 + K9 * D32 ^ 3 + L9 * D32 ^ 4
    Else
        D35 = 0
    End If
'Condensed steam temperature if P_c>0.1
    If D32 > 0.1 Then
        J3 = 17.4992653754406
        I9 = 402.36973631205
        J9 = -1491.36622741945
        K9 = 3336.98098975695
        L9 = -3080.83798502275
        D36 = J3 + I9 * D32 + J9 * D32 ^ 2 + K9 * D32 ^ 3 + L9 * D32 ^ 4
    Else
        D36 = 0
    End If
'Condensed steam  temperature
    D37 = D34 + D35 + D36
'Condensing temperature,°C
    If Check_CT = 0 Then
        SHELL_OUT(11).Text = Format(D37, "0.00")
        SHELL_OUT(11).ForeColor = &HFF0000
        SHELL_OUT(11).BackColor = &HE0E0E0
        SHELL_TEMP_IN.Text = Format(D37, "0.00")
        SHELL_TEMP_OUT.Text = Format(D37, "0.00")
    ElseIf Check_CT = 1 Then
        D37 = Format(HScroll_C_TEMP / 100, "0.00")
        SHELL_OUT(11).Text = Format(D37, "0.00")
        SHELL_OUT(11).ForeColor = &HC0&
        SHELL_OUT(11).BackColor = &HE0E0E0
        SHELL_TEMP_IN.Text = Format(D37, "0.00")
        SHELL_TEMP_OUT.Text = Format(D37, "0.00")
    End If

Call FOULING

'DELTA DUTY test
    If Check_des = Checked Then
        Check_T_OUT.Visible = False
        TUBES_FLOW.ForeColor = &HC0&
        TUBES_FLOW.BackColor = &H80000018
        TUBES_TEMP_OUT.ForeColor = &HC0&
        TUBES_TEMP_OUT.BackColor = &H80000018
    End If
'Water side duty
        D43 = D9 * SPH_T * (D13 - D11)
        DUTY_T = D43 / 860.04
'Steam side duty
'       Q_S_VAP = WATER_VAP_IN * D38
'       Q_S_LIQ = WATER_LIQ_IN * SPH_T * (XD19 - XD20)
'       D44 = Q_S_VAP + Q_S_LIQ
        D44 = D29 * D38
        DUTY_S = D44 / 860.04
'Water flow calculated
        flow_calc = D44 / ((D13 - D11) * SPH_T)
'Steam flow calculated
        SHELL_FLOW_CALC = D29
    If Thermal_bal_tubes = True Then
        Check_T_OUT.Visible = True
        TUBES_OUT(0).BackColor = &HC0&
        TUBES_OUT(0).ForeColor = &HFFFFFF
        SHELL_OUT(0).BackColor = &HE0E0E0
        SHELL_OUT(0).ForeColor = &HC0&
        Check_T_OUT.Visible = True
        If Check_T_OUT = 0 Then
            TUBES_TEMP_OUT.ForeColor = &HFFFFFF
            TUBES_TEMP_OUT.BackColor = &HC0&
            TUBES_FLOW.ForeColor = &HC0&
            TUBES_FLOW.BackColor = &H80000018
        'Steam side duty
'           Q_S_VAP = WATER_VAP_IN * D38
'           Q_S_LIQ = WATER_LIQ_IN * SPH_T * (XD19 - XD20)
'           D44 = Q_S_VAP + Q_S_LIQ
            D44 = D29 * D38
            DUTY_S = D44 / 860.04
        'Tubes flow calculated
            D13 = D44 / (D9 * SPH_T) + D11
        'Tubes side DUTY
            D43 = D9 * SPH_T * (D13 - D11)
            DUTY_T = D43 / 860.04
            flow_calc = D9 'D44 / ((D13 - D11) * SPH_T)
            SHELL_FLOW_CALC = D29
            TUBES_TEMP_OUT.Text = Format(D13, "0.00")
            YXY = 1
            Spin_TUBES_T_OUT = D13 * 100
        ElseIf Check_T_OUT = 1 Then
'            Check_T_Flow = Unchecked
            TUBES_TEMP_OUT.ForeColor = &HC0&
            TUBES_TEMP_OUT.BackColor = &H80000018
            TUBES_FLOW.ForeColor = &HFFFFFF
            TUBES_FLOW.BackColor = &HC0&
        'Steam side duty
'           Q_S_VAP = WATER_VAP_IN * D38
'           Q_S_LIQ = WATER_LIQ_IN * SPH_T * (XD19 - XD20)
'           D44 = Q_S_VAP + Q_S_LIQ
            D44 = D29 * D38
            DUTY_S = D44 / 860.04
            D9 = D44 / (SPH_T * (D13 - D11))
            D43 = D9 * SPH_T * (D13 - D11)
            DUTY_T = D43 / 860.04
            flow_calc = D9
            SHELL_FLOW_CALC = D29
            TUBES_TEMP_OUT.Text = Format(D13, "0.00")
            TUBES_FLOW = Format(flow_calc, "#,##0")
            TUBES_WATER = Format(flow_calc, "#,##0")
        End If
    ElseIf Thermal_bal_shell.Value = True Then
        TUBES_TEMP_OUT.ForeColor = &HC0&
        TUBES_TEMP_OUT.BackColor = &H80000018
        TUBES_FLOW.ForeColor = &HC0&
        TUBES_FLOW.BackColor = &H80000018
        
        SHELL_OUT(0).BackColor = &HC0&
        SHELL_OUT(0).ForeColor = &HFFFFFF
        TUBES_OUT(0).BackColor = &HE0E0E0
        TUBES_OUT(0).ForeColor = &HC0&
    'Steam side duty,
'       Q_S_VAP = WATER_VAP_IN * D38
'       Q_S_LIQ = WATER_LIQ_IN * SPH_T * (XD19 - XD20)
'       D44 = Q_S_VAP + Q_S_LIQ
        D44 = D29 * D38
        DUTY_S = D44 / 860.04
    'Tubes side duty
        D43 = D9 * SPH_T * (D13 - D11)
        DUTY_T = D43 / 860.04
    'Shell flow calculated
        D29 = D43 / D38
    'Steam side duty calculated
        D44 = D29 * D38
        DUTY_S = D44 / 860.04
        flow_calc = D44 / ((D13 - D11) * SPH_T)
        SHELL_FLOW_CALC = D29
    Else
        TUBES_OUT(0).ForeColor = &HC0&
        TUBES_OUT(0).BackColor = &HE0E0E0
        SHELL_FLOW.ForeColor = &HC0&
        SHELL_FLOW.BackColor = &H80000018
        SHELL_OUT(0).ForeColor = &HC0&
        SHELL_OUT(0).BackColor = &HE0E0E0
        TUBES_TEMP_OUT.ForeColor = &HC0&
        TUBES_TEMP_OUT.BackColor = &H80000018
        TUBES_FLOW.ForeColor = &HC0&
        TUBES_FLOW.BackColor = &H80000018
    'Shell side Duty, kcal/h
'       Q_S_VAP = WATER_VAP_IN * D38
'       Q_S_LIQ = WATER_LIQ_IN * SPH_T * (XD19 - XD20)
'       D44 = Q_S_VAP + Q_S_LIQ
        D44 = D29 * D38
        DUTY_S = D44 / 860.04
    'Tube side Duty, kcal/h
        D43 = D9 * SPH_T * (D13 - D11)
        DUTY_T = D43 / 860.04
    'Flow calculated, kg/h
        flow_calc = D9
        SHELL_FLOW_CALC = D29
    End If
YXY = 0
TUBES_FLOW = Format(D9, "#,##0")
TUBES_WATER = Format(flow_calc, "#,##0")
SHELL_FLOW = Format(D29, "#,##0")
SHELL_WATER = Format(D29, "#,##0")
Spin_SHELL_FLOW = D29 / FACT_FLOW
'Total shell flowrate, Kg/h
    Ftot_IN_1 = D29
    Ftot_IN = Ftot_IN_1
    XD18 = Ftot_IN
'TOTAL VAPOR IN, Kg/h
    Vtot_INP_1 = Vtot_INP
    Vtot_IN_1 = Vtot_INP_1 / 100 * Ftot_IN_1
    Vtot_IN = Format(Vtot_IN_1, "#,##0")
'VAPOR NOT CONDENSING IN = VAP_NON-CONDENSING OUT, Kg/h
    HScroll_SHELL_NON_COND.Max = Vtot_IN / FACT_FLOW
    NC_IN_1 = HScroll_SHELL_NON_COND * FACT_FLOW
    NC_IN = Format(NC_IN_1, "#,##0")
'CONDENSABLE WATER VAPOR IN, Kg/h
    Vwat_INP_1 = Vwat_INP
    Vwat_IN_1 = Vwat_INP_1 / 100 * Vtot_IN_1
    Vwat_IN = Format(CDbl(Vwat_IN_1), "#,##0")
    Vwat_INP_1 = Vwat_IN_1 * 100 / Vtot_IN_1
    Vwat_INP = Format(CDbl(Vwat_INP_1), "0.0")
'TOTAL LIQUID IN, Kg/h
    Ltot_IN_1 = Ftot_IN_1 - Vtot_IN_1
    Ltot_IN = Format(Ltot_IN_1, "#,##0")
    Ltot_INP_1 = 100 - Vtot_INP_1
    Ltot_INP = Format(Ltot_INP_1, "0.0")
'CONDENSABLE ORGANIC VAPOR IN (NOT WATER)
    Vorg_IN_1 = Vtot_IN - NC_IN - Vwat_IN_1
    Vorg_IN = Format(CDbl(Vorg_IN_1), "#,##0")
    Vorg_INP_1 = 100 - Vwat_INP_1
    Vorg_INP = Format(Vorg_INP_1, "0.0")
'WATER LIQUID IN, Kg/h
    Lwat_INP_1 = 100   'Lwat_IN_1 * 100 / Ltot_IN_1
    Lwat_INP = Format(Lwat_INP_1, "0.0")
    Lwat_IN_1 = Lwat_INP_1 / 100 * Ltot_IN_1
    Lwat_IN = Format(CDbl(Lwat_IN_1), "#,##0")
    Lwat_IN_1 = Lwat_INP_1 / 100 * Ltot_IN_1
    Lwat_IN = Format(CDbl(Lwat_IN_1), "#,##0")
    wet_steam_1 = Vtot_INP_1 / 100 '(Lwat_IN_1 / Ftot_IN * 100)
    Wet_steam = Format(wet_steam_1, "0.000")
'ORGANIC LIQUID IN, Kg/h
    Lorg_IN_1 = Ltot_IN_1 - Lwat_IN_1
    Lorg_IN = Format(Lorg_IN_1, "#,##0")
    Lorg_INP_1 = 100 - Lwat_INP_1
    Lorg_INP = Format(Lorg_INP_1, "0.0")
'TOTAL LIQUID OUT, Kg/h
    Ltot_out = Vorg_IN_1 + Vwat_IN_1 + Lorg_IN_1 + Lwat_IN_1
    Lorg_OUT = Lorg_IN_1
    Lwat_OUT = Vwat_IN_1 + Lwat_IN_1
        SHELL_VAPOR = Format(Vorg_IN_1, "#,##0")
        SHELL_LIQUID = Format(Lorg_OUT, "#,##0")
        SHELL_WATER = Format(Lwat_OUT, "#,##0")
        SHELL_NON_COND = Format(NC_IN, "#,##0")
'TOTAL FLOW
    Ftot_IN = Format(Vorg_IN_1 + Vwat_IN_1 + Lorg_IN_1 + Lwat_IN_1, "#,##0")
    Vorg_INP_1 = Vorg_IN_1 / Ftot * 100
    Ftot_INP = Format(Vtot_INP_1 + Ltot_INP_1, "0.0")
'Percent of fractions OUT shell side
    Lorg_0UTP = Lorg_OUT / XD18 * 100
    Lwat_OUTP = Lwat_OUT / XD18 * 100
    NC_OUTP = NC_IN / XD18 * 100
    Ltot_OUTP = Lorg_0UTP + Lwat_OUTP + NC_OUTP












'Tube side flow arte, kg/h
    TUBES_OUT(0).Text = Format(flow_calc, "#,##0")
'Shell side flow arte, kg/h
    SHELL_OUT(0) = Format(SHELL_FLOW_CALC, "#,##0")
'Tube side Duty, MW
    TUBES_OUT(7).Text = Format(D43 / 860.04, "#,##0")
'Shell side Duty, MW
    SHELL_OUT(7).Text = Format(D44 / 860.04, "#,##0")
'Mass of water in steam
    Hw = D29 * (1 - D33)
'Saturated steam flowrate INLET condenser, lb/h
    D30 = D29 * 2.20462
'Steam loading, lb/h ft2
    D31 = D30 / D80
'Water flow rate m3/h
    D15 = D9 / D19
'Flow area section, m2
    'd73= Inlet tube diameter
    'D74= Tubes number
    'D77= tubes passes
    'lungh =U pattern
    If lungh = 2 Then D77 = 2
    XD91 = 3.14159 / 4 * D73 ^ 2 * D74 / D77
    TUBES_SECTION.Text = Format(XD91, "0.0000")
'Water velocity through tubes, m/s
    D22 = D15 / XD91 / 3600 / PARALLEL_N
    TUBES_OUT(2).Text = Format(D22, "0.00")
'Water velocity through tubes, feet/s
    D23 = D22 / 0.3048                                                                                    'fps
'Reynolds through tubes
    D24 = D22 * D73 * D19 / (D20 / 1000)
    TUBES_OUT(6).Text = Format(D24, "#,##0")
'Pressure drop through tubes (Hazen-Williams with C=130) Related to CS
    D25 = ((6.05 * 10 ^ 5 * ((D15 * 1000 / 60) / (D74 / D77)) ^ 1.85) / (130 ^ 1.85 * (D73 * 1000) ^ 4.87)) * D75 * D77
'Pressure drop due to return (estimated four velocity heads)
    D26 = ((4 * (D22 ^ 2 / (2 * 9.81))) / 10) * 0.9807
'Total pressure drop for 100% clean tube side, kPa
    D27 = D25 + D26
    TUBES_OUT(9).Text = Format(D27 * 100, "0.00")
    If Check_P_DROP_T = Unchecked Then
'Tubes side total pressure drop, kPa
        TUBES_OUT(9) = Format(D27, "0.00")
    Else
'Tubes side total pressure drop, kPa
        D27 = (TUBES_P_IN - TUBES_P_OUT) * 100
        TUBES_OUT(9).Text = Format(D27, "0.00")
    End If
'C Factor
    C_Factor.Text = Format(D15 / (D27 * 100) ^ (1 / 2), "0.0")                      'm3/h/kPa
'Log Mean Temperature Difference, °C
    D41 = ((D37 - D11) - (D37 - D13)) / Log((D37 - D11) / (D37 - D13))
    LMTD.Text = Format(D41, "0.00")
    MTDc = Format(D41, "0.00")
'Terminal temperature difference, °C
    D42 = D37 - D13
    TTD.Text = Format(D42, "0.00")
    Label22(15).Caption = "(T2 - t2)"
'TUBES HEAT FLUX, kW/m2
    Q6 = D43 / D79 * 0.001163
    TUBES_OUT(8).Text = Format(Q6, "0.00")
'SHELL HEAT FLUX, kW/m2
    SPH_S = SHELL_OUT(3).Text
    Q6S = D44 / D79 * 0.001163
    SHELL_OUT(8).Text = Format(Q6S, "0.00")
'Water temperature correction factor
    If D12 <= 50 Then
        D45 = 0.1228463 + 1.483184 * 10 ^ -2 * D12 - 2.17211 * 10 ^ -5 * D12 ^ 2
    Else
        D45 = 0
    End If
    If D12 > 50 And D12 < 70 Then
        D46 = 4.269782 * 10 ^ -2 + 1.931213 * 10 ^ -2 * D12 - 8.034383 * 10 ^ -5 * D12 ^ 2
    Else
        D46 = 0
    End If
    If D12 >= 70 Then
        D47 = 0.4470833 + 1.121063 * 10 ^ -2 * D12 - 4.696334 * 10 ^ -5 * D12 ^ 2
    Else
        D47 = 0
    End If
    D48 = D45 + D46 + D47
'Constant for tube outside diameter
    If D67 <= 0.75 Then
        D52 = 267 * Sqr(D23)
    ElseIf D67 > 0.75 And D67 <= 1 Then
        D52 = 263 * Sqr(D23)
    ElseIf D67 > 1 Then
        D52 = 259 * Sqr(D23)
    Else
        D52 = 0
    End If
'Loading correction factor
    D53 = (D31 / 8) ^ 0.25
'Overall CLEAN heat transfer coefficient, kcal/(h m^2 ºC):
    D55 = D48 * D52 * D54 * 4.882 * D53
    If Check_U_CLEAN = Checked Then
        D55 = HScroll_U_CLEAN
        U_COEFF_CLEAN = D55
    ElseIf Check_U_CLEAN = Unchecked Then
        U_COEFF_CLEAN.Text = Format(D55, "0.0")
    End If
'Overall DIRTY heat transfer coefficient, kcal/(h m^2 ºC)
    D56 = D44 / D41 / D79
    U_COEFF_DIRTY.Text = Format(D56, "0.0")
'CLEANLINESS FACTOR
    D57 = D56 * 100 / D55
    C_Factor.Text = Format(D57, "0.0") 'CLEANLINESS FACTOR
    If D57 > 100 Or D57 <= 0 Then
        CF.BackColor = &HFF&
    Else
        CF.BackColor = &H8000000F
    End If

'D18 Tube side caloric temp., °F
'D23 Water velocity through tubes, feet/s
'D72 Tube inlet diameter (inches)
'D67 Tube Outlet Diameter (inches)
'D68 Tube Outlet Diameter (meters)
'D73 'Tube inlet diameter (meters)
'D78 Material conductivity
'D55 Overall CLEAN heat transfer coefficient, kcal/(h m^2 ºC)
'D56 Overall DIRTY heat transfer coefficient, kcal/(h m^2 ºC)

'Water side individual heat transfer coefficient referred to ext. surface, btu/h.ft2.°F
    D58_1 = (150 * (1 + 0.011 * D18) * (D23 ^ 0.8 / D72 ^ 0.2)) * (D72 / D67) * 4.882427685
    D58 = (150 * (1 + 0.011 * D18) * (D23 ^ 0.8 / D72 ^ 0.2))
    XD108 = D58
'Water side individual heat transfer coeficient, Kcal/(h m^2 C)
    XD109 = XD108 * 4.882427685
'Hio - Water side indiv. heat transfer coeficient referred to ext. surface Kcal/(h m^2 C)
    XD110 = XD109 * (D72 / D67)
'Heat transfer resistance due to water flowing inside the tubes, h.ft2.°F/btu
    D59 = 10000 / XD110 'D58
    XD114 = D59
'Heat transfer resistance due to  the wall, referred to ext. surface, (h m^2 ºC)/kcal
    D60 = (D68 * Log(D68 / D73) / (2 * D78)) * 10000
    XD111 = D60
'Ho - Heat transfer resistance due to the condensing steam film, (h m^2 ºC)/kcal
    D61 = (10000 / D55) - D59 - D60
'Heat transfer resistance due to outside fouling factor, (h m^2 ºC)/kcal
    D62 = D40_S * 10000
'Total heat transfer resistance, referred to external surface, kcal/(h m^2 ºC)
    D64 = 10000 / D56
'Heat transfer resistance due to inside fouling factor referred to ext surface, (h m^2 ºC)/kcal
    D63 = D64 - (D59 + D60 + D61 + D62)
'Water side fouling factor, [(m^2 ºC)/KW]*10^-4
    D65 = D63 * (D73 / D68)
    TUBES_FF_1 = D65
    TUBES_FF.Text = Format(TUBES_FF_1, "0.000")
    
PROP = "SHELL"
Call Properties
    
'Skin Temperature
'DT_1=[((KW/m2)/(mps)*3.739)+(tout)]

'The skin temperature will be equal to the cold side fluid plus the "film drop"
'temperature difference, which equals the value of film.
'DT_2 = (q/A) *(1/h + Rc)
'where h= convective heat transfer coef at cold side outlet and
'Rc= fouling coefficient on cold side outlet.

'I used the graphical approach from VDI Wärmeatlas (Cb1, Picture 3),
'which leads to formula without fouling for the wall temperature:
'(tw2-t2)/(1/alfa2)=(T1-t2)/(1/OHTC)
'tw2 = (T1-t2)/(1/OHTC)*(1/alfa2)+ t2
'tw2: wall temperature cold side
't2: temperature cold side (here outlet)
'alfa2: convective heat transfer cold side
'T1: hot side temperature (inlet)
'OHTC: overall heat transfer coefficient.
    
'XD9   TUBES Caloric temperature,ºC
'XD9_S SHELL Caloric temperature,ºC
'D79   Surface
'Ho = 1 / D61 * 10000
'Hio = 1 / XD110 * 10000
'rc = D65  'Water side fouling factor
'DT_1 = ((D43 / 860.04) / D79) / (D22 * 3.739) + XD8
'DT_2 = (D43 / D79) * (D60 / 10000 + rc / 10000) + XD8
'tw2 = (XD19 - XD8) / (1 / D55) * (D60 / 10000) + XD8
'SKIN_T = XD9 + Ho / (Ho + Hio) * (XD9_S - XD9)
    
tw2a = (XD19 - XD8) / (1 / D55) * (1 / XD109) + XD8
SKIN_T = tw2a
SKIN_TEMP = Format(tw2a, "0.00")
xd72 = tw2a
'Shell fluid density at film temperature,Kg/m^3
    D19 = 0.0002 * xd72 ^ 3 - 0.028 * xd72 ^ 2 + 0.0873 * xd72 + 999.92
    If Check_S_DENS = 0 Then
        SHELL_OUT(4).Text = Format(D19, "0.0")
        SHELL_OUT(4).ForeColor = &HFF0000
        SHELL_OUT(4).BackColor = &HE0E0E0
    ElseIf Check_S_DENS = 1 Then
        D19 = HScroll_SHELL_DENS / 10
        SHELL_OUT(4) = Format(D19, "0.0")
        SHELL_OUT(4).ForeColor = &HC0&
        SHELL_OUT(4).BackColor = &HE0E0E0
    End If
    'Shell fluid density at film temperature,lb/ft^3
    XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
'Shell fluid viscosity at condensing film temperature,cp
    D20 = (100 / (2.1482 * ((273.16 + xd72 - 281.435) + Sqr(8078.4 + (273.16 + xd72 - 281.435) ^ 2)) - 120))
    If Check_S_VISC = 0 Then
        SHELL_OUT(5).Text = Format(D20, "0.000")
        SHELL_OUT(5).ForeColor = &HFF0000
        SHELL_OUT(5).BackColor = &HE0E0E0
    ElseIf Check_S_VISC = 1 Then
        D20 = HScroll_SHELL_VISC / 1000
        SHELL_OUT(5) = Format(D20, "0.000")
        SHELL_OUT(5).ForeColor = &HC0&
        SHELL_OUT(5).BackColor = &HE0E0E0
    End If
'Water specific heat at SHELL CALORIC TEMP,Kcal/(Kg ºC)
    h1 = 1.00691354509505
    h2 = -1.19506245657282E-03
    h3 = 5.57856020013537E-05
    h4 = -9.75376157602428E-07
    h5 = 6.26080712782905E-09
    SPH_S = h1 + h2 * xd72 + h3 * xd72 ^ 2 + h4 * xd72 ^ 3 + h5 * xd72 ^ 4
    If Check_S_SPH = 0 Then
        SHELL_OUT(3).Text = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HFF0000
        SHELL_OUT(3).BackColor = &HE0E0E0
    ElseIf Check_S_SPH = 1 Then
        SPH_S = HScroll_SHELL_SPH / 1000
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HC0&
        SHELL_OUT(3).BackColor = &HE0E0E0
    End If
'Shell fluid thermal conductivity at SHELL CALORIC TEMP, Kcal/h m ºC
    TH_C = 0.00000000592317 * xd72 ^ 3 - 0.0000080425 * xd72 ^ 2 + 0.0018262 * xd72 + 0.478535
    If Check_S_TC = 0 Then
        SHELL_OUT(1).Text = Format(TH_C, "0.000")
        SHELL_OUT(1).ForeColor = &HFF0000
        SHELL_OUT(1).BackColor = &HE0E0E0
    ElseIf Check_S_TC = 1 Then
        TH_C = HScroll_SHELL_TC / 1000
        SHELL_OUT(1) = Format(TH_C, "0.000")
        SHELL_OUT(1).ForeColor = &HC0&
        SHELL_OUT(1).BackColor = &HE0E0E0
    End If
    
'SHELL VELOCITY
    L37 = XD61M                 'SHELL_ID, mm
    O37 = XD52M                 'SHELL_TUBES_PITCH, mm
    E37 = XD66M                 'T_OD, mm
    N37 = XD64M                 'SHELL_BAFFLES_SPACE, mm
'SHELL CLEARANCE
    V37 = O37 / 1000 - E37 / 1000
    Clearance.Text = Format(V37, "0.0000")
    'SHELL FLOW AREA
    U37 = L37 / 1000 * V37 * N37 / 1000 / (O37 / 1000)
    Flow_area.Text = Format(U37, "0.0000")
    k47 = XD18 / D19 / U37 / 3600 / PARALLEL_N * XD59
    SHELL_OUT(2) = Format(k47, "0.00")
'SHELL Reynolds
    EQ_D19 = XD52M / 1000           'SHELL_TUBES_PITCH / 1000
    EQ_PI = XPI
    EQ_D14 = XD66                   'T_OD / 1000
    If SHELL_PITCH_CONF = "Triangular" Then
        EQ_E31 = 4 * (EQ_D19 ^ 2 - EQ_PI * EQ_D14 ^ 2 / 4) / (XPI * EQ_D14)
    Else
        EQ_E31 = (4 * (0.5 * EQ_D19 * 0.866 * EQ_D19 - 0.5 * XPI * EQ_D14 ^ 2 / 4) / (0.5 * XPI * EQ_D14))
    End If
'Flow_area,m2
    EQ_E29 = U37
'Shell_Flow water
    EQ_E25 = XD18
    EQ_E30 = EQ_E25 / EQ_E29
'SHELL_OUT(5) * 3.6 = km/h
    EQ_E8 = D20 * 3.6
    EQ_E32 = EQ_E31 * EQ_E30 / EQ_E8
    Q_E22 = EQ_E31 * 1000
'Density, kg/m3
    Q_E17 = D19
'Shell flow velocity, m/s
    Q_E27 = k47
'SHELL viscosity, cP
    Q_E18 = D20
    Q_E28 = Q_E22 * Q_E17 * Q_E27 / D20
    SHELL_OUT(6) = Format(Q_E28, "#,##0")
'CALCULATING SHELL SIDE PRESSURE DROP
    If Check_P_DROP_S = Unchecked Then
'Pressure drop (tubes)
        P_E17 = D19                         'SHELL_OUT(4)
        P_E22 = EQ_E31 * 1000               'Equivalent diameter, mm
        P_E27 = Q_E27                       'Shell flow velocity, m/s
        P_E23 = XD55 * lungh                'T_len, mm
        P_E28 = Q_E28
        P_E29 = 0.44 * P_E28 ^ -0.19
        P_E30 = 4 * P_E29 * P_E23 * P_E27 ^ 2 / (P_E22 * 2 * 9.8) * P_E17 * 0.000096784 * 101.325
'Pressure drop (sheet)
        P_E9 = XD59                         'SHELL PASSES
        P_E31 = 3 * P_E9 * P_E27 ^ 2 / 2 / 9.8 * P_E17 * 0.000096784 * 101.325
        P_E32 = P_E30 + P_E31
'Pressure drop (tubes), kPa
        SHELL_OUT(9).Text = Format(P_E32, "0.00")
    Else
'Pressure drop (tubes), kPa
        P_E32 = (SHELL_P_IN - SHELL_P_OUT)
        SHELL_OUT(9).Text = Format(P_E32 * 100, "0.00")
    End If
End Sub
Private Sub Properties()
On Error Resume Next

XPI = 3.141592654
LN = 2.302585093
XD6 = TUBES_FLOW                'TUBES total flowrate,Kg/h
XD6L = TUBES_LIQUID             'TUBES liquid flowrate,Kg/h
XD7 = TUBES_TEMP_IN             'TUBES temperature in,ºC
XD8 = TUBES_TEMP_OUT            'TUBES temperature out,ºC
XD18 = SHELL_FLOW               'Shell total fluid flowrate,Kg/h
XD19 = SHELL_TEMP_IN            'Shell fluid temperature in,ºC
XD20 = SHELL_TEMP_OUT           'Shell fluid temperature out,ºC
XD52M = SHELL_TUBES_PITCH       'Pitch, mm
XD52 = XD52M / 25.4             'Pitch,inch
XD54 = T_NO                     'Number of tubes
XD55 = T_len                    'Tube lenght,m
XD56 = XD55 / 0.3048            'Tube lenght     ft
XD57 = T_PASS                   'Number of tube side passes
XD58 = Mat_cond                 'Thermal conductivity of tube material,Kcal/(h m^2 ºC/m)
XD59 = SHELL_PASS               'Shell passes
XD61M = SHELL_ID                'Shell ID, mm
XD61 = XD61M / 25.4             'Shell ID, inch
XD63 = SHELL_BAFFLES_CUT        'Baffle cut, %
XD64M = SHELL_BAFFLES_SPACE     'Baffle spacing  mm
XD64 = XD64M / 25.4             'Baffle spacing  inch
XD66M = T_OD                    'Tube Outlet diameter, mm
XD66 = XD66M / 1000             'Tube Outlet diameter, m
XD50 = XD66 / 25.4 * 1000       'Tube outlet diameter, inch
XD85 = T_ID / 1000              'Tube Inlet diameter, m
XD84 = XD85 / 0.3048            'Tube Inlet diameter,ft
XD83 = XD85 / 25.4 * 1000       'Tube Inlet diameter,inches
XD112 = SHELL_FF                'Process side fouling factor [(hm^2ºC)/Kcal]*10^4

'TUBES TEMP IN/OUT
    D11 = CDbl(XD7)                           'TUBES T_IN, °C
    D12 = D11 * 1.8 + 32                      'Water T_IN, °F
    D13 = CDbl(XD8)                           'TUBES T_OUT, °C
'TUBES RANGE
    Range_T = D13 - D11
'SHELL TEMP IN/OUT
    XD19 = CDbl(XD19)                         'SHELL T_IN, °C
    XD20 = CDbl(XD20)                         'SHELL T_OUT, °C
'SHELL RANGE
    Range_S = XD19 - XD20
'SHELL flowrate INLET
    D29 = CDbl(SHELL_FLOW.Text)         'kg/h
    D30 = D29 * 2.20462                 'lb/h
    D15S = D29 / D19
    If PROP = "TUBES" Then
        'TUBES Caloric temperature,ºC
            D17 = D11 + (D13 - D11) / 2
            XD9 = D17
    ElseIf PROP = "SHELL" Then
        'SHELL Caloric temperature,ºC
            D17 = XD20 + (XD19 - XD20) / 2
        'SHELL Caloric temperature
            XD9_S = D17
    End If

    'Thermal conductivity, Kcal/h m ºC
        TH_C = 0.00000000592317 * D17 ^ 3 - 0.0000080425 * D17 ^ 2 + 0.0018262 * D17 + 0.478535
    'Viscosity of water, cP
        D20 = (100 / (2.1482 * ((273.16 + D17 - 281.435) + Sqr(8078.4 + (273.16 + D17 - 281.435) ^ 2)) - 120))
    'Caloric temperature,°C
        D17_2 = Int(D17 / 2)
    'Caloric tubes-side temperature,°F
        D18 = D17 * 1.8 + 32
    'Density of water, kg/m3
        D19 = 1000
        D19 = 0.0002 * D_17_2 ^ 3 - 0.028 * D17_2 ^ 2 + 0.0873 * D17_2 + 999.92
'        Select Case D17_2
'            Case 1: D19 = 999.94
'            Case 2: D19 = 999.97
'            Case 3: D19 = 999.94
'            Case 4: D19 = 999.85
'            Case 5: D19 = 999.7
'            Case 6: D19 = 999.497
'            Case 7: D19 = 999.244
'            Case 8: D19 = 998.943
'            Case 9: D19 = 998.595
'            Case 10: D19 = 998.204
'            Case 11: D19 = 997.77
'            Case 12: D19 = 997.296
'            Case 13: D19 = 996.783
'            Case 14: D19 = 996.233
'            Case 15: D19 = 995.647
'            Case 16: D19 = 995.026
'            Case 17: D19 = 994.371
'            Case 18: D19 = 993.684
'            Case 19: D19 = 992.965
'            Case 20: D19 = 992.215
'            Case 21: D19 = 991.436
'            Case 22: D19 = 990.628
'            Case 23: D19 = 989.792
'            Case 24: D19 = 988.928
'            Case 25: D19 = 988.037
'            Case 26: D19 = 987.12
'            Case 27: D19 = 986.177
'            Case 28: D19 = 985.219
'            Case 29: D19 = 984.217
'        End Select
    'TUBES flowrate INLET
        D9 = TUBES_FLOW.Text                'Water flow rate, kg/h
        D15 = D9 / D19
    
    'Specific heat of water
        D_21 = 1
        h1 = 1.00691354509505
        h2 = -1.19506245657282E-03
        h3 = 5.57856020013537E-05
        h4 = -9.75376157602428E-07
        h5 = 6.26080712782905E-09
        D_21 = h1 + h2 * xd72 + h3 * xd72 ^ 2 + h4 * xd72 ^ 3 + h5 * xd72 ^ 4
'        Select Case D17_2
'            Case 1: D_21 = 1.00636
'            Case 2: D_21 = 1.00495
'            Case 3: D_21 = 1.00378
'            Case 4: D_21 = 1.00277
'            Case 5: D_21 = 1.00194
'            Case 6: D_21 = 1.00124
'            Case 7: D_21 = 1.00067
'            Case 8: D_21 = 1.00019
'            Case 9: D_21 = 0.999978
'            Case 10: D_21 = 0.99947
'            Case 11: D_21 = 0.99921
'            Case 12: D_21 = 0.99902
'            Case 13: D_21 = 0.99885
'            Case 14: D_21 = 0.99873
'            Case 15: D_21 = 0.99866
'            Case 16: D_21 = 0.99861
'            Case 17: D_21 = 0.99859
'            Case 18: D_21 = 0.99861
'            Case 19: D_21 = 0.99864
'            Case 20: D_21 = 0.99869
'            Case 21: D_21 = 0.99876
'            Case 22: D_21 = 0.99883
'            Case 23: D_21 = 0.99895
'            Case 24: D_21 = 0.99907
'            Case 25: D_21 = 0.99919
'            Case 26: D_21 = 0.99935
'            Case 27: D_21 = 0.9995
'            Case 28: D_21 = 0.99969
'            Case 29: D_21 = 0.99988
'        End Select
    If PROP = "TUBES" Then
        SPH_T = D_21
    ElseIf PROP = "SHELL" Then
        SPH_S = D_21
    End If
End Sub
Private Sub FOULING()
On Error Resume Next
Dim S_side(40), T_side(10), P_FF(40), T_FF(5)
'Process fouling
    D40_S = 0.0001
    D40_T = 0.0001
    If Check_PF = Unchecked Then
        S_side(0) = "Benzene"
        S_side(1) = "Demineralized water"
        S_side(2) = "Fuel oil"
        S_side(3) = "Gasoline"
        S_side(4) = "Heavy oil"
        S_side(5) = "Kerosene or Gas oil"
        S_side(6) = "Jacket water"
        S_side(7) = "Lube oil (low viscosity)"
        S_side(8) = "Lube oil (high viscosity)"
        S_side(9) = "Naphtha"
        S_side(10) = "Organic solvent"
        S_side(11) = "Wax distillate"
        S_side(12) = "Alchol vapor"
        S_side(13) = "Hydrocarbons hb"
        S_side(14) = "Hydrocarbons lb"
        S_side(15) = "Steam"
        S_side(16) = "Steam condensing"
        S_side(17) = "Butane"
        S_side(18) = "Propane"
        S_side(19) = "Water"
        S_side(20) = "Other"
        S_side(21) = "Jacket water"
        S_side(22) = "Feed Water"
        S_side(23) = "Cooling Water"
    
        P_FF(0) = 0.0001
        P_FF(1) = 0.0002049
        P_FF(2) = 0.001434
        P_FF(3) = 0.0006148
        P_FF(4) = 0.001025
        P_FF(5) = 0.001025
        P_FF(6) = 0.0004098
        P_FF(7) = 0.0004098
        P_FF(8) = 0.0006148
        P_FF(9) = 0.001025
        P_FF(10) = 0.0006148
        P_FF(11) = 0.001025
        P_FF(12) = 0.0004098
        P_FF(13) = 0.0006148
        P_FF(14) = 0.0006148
        P_FF(15) = 0.00008197
        P_FF(16) = 0.00008197
        P_FF(17) = 0.0003074
        P_FF(18) = 0.0003074
        P_FF(19) = 0.0006148
        P_FF(20) = 0.0001
        P_FF(21) = 0.0004098
        P_FF(22) = 0.0006148
        P_FF(23) = 0.0006148
        
        Shell_side = Combo_S_FLUID
        Tube_side = Combo_T_FLUID
        For i = 0 To 23
            If Shell_side = S_side(i) Then
                D40_S = P_FF(i)
                SHELL_FF.Text = D40_S * 10000
            End If
        Next i
        SHELL_FF.ForeColor = &HFF0000
        SHELL_FF.BackColor = &HE0E0E0
    ElseIf Check_PF = Checked Then
        D40_S = Spin_PF / 1000000     'SHELL_FF / 10000
        SHELL_FF.Text = Format(D40_S * 10000, "0.00")
        SHELL_FF.ForeColor = &HC0&
        SHELL_FF.BackColor = &HE0E0E0
    End If
    DES_1 = 0
End Sub
Private Sub Mechanical()
On Error Resume Next
'MECHANICAL DATA
    D68 = T_OD.Text / 1000           'Tube Outlet Diameter (m)
    D67 = D68 * 1000 / 25.4          'Tube Outlet Diameter (inches)
    D69 = Combo_BWG.Text             'BWG
'BWG / 'Wall Thickness (inches)
    Select Case D69
        Case 7:  D70 = 0.18
        Case 8:  D70 = 0.165
        Case 9:  D70 = 0.148
        Case 10: D70 = 0.134
        Case 11: D70 = 0.12
        Case 12: D70 = 0.109
        Case 13: D70 = 0.095
        Case 14: D70 = 0.083
        Case 15: D70 = 0.072
        Case 16: D70 = 0.065
        Case 17: D70 = 0.058
        Case 18: D70 = 0.049
        Case 19: D70 = 0.042
        Case 20: D70 = 0.035
        Case 22: D70 = 0.028
        Case 24: D70 = 0.022
        Case 26: D70 = 0.018
    End Select
'Tube inlet diameter
    D71 = D70 * 25.4 * 10 ^ -3              'Wall Thickness (m)
    D72 = D67 - 2 * D70                     'Tube inlet diameter (inches)
    D73 = D72 * 25.4 * 10 ^ -3              'Tube inlet diameter (meters)
    T_ID.Text = Format(D73 * 1000, "0.00")
    T_ID_E = Format(D73 * 1000, "0.00")

    D74 = T_NO.Text                         'Number of tubes
    D75 = T_len.Text                        'Tube lenght (m)
    D76 = D75 / 0.3048                      'Tube lenght (inches)
    D77 = T_PASS.Text                       'Number of tube side passes
    
'Tube Material Factor
    Dim mat(40), Fac(40, 7), Material_cond(40)
    mat(1) = "70-30 Cu-Ni"
    mat(2) = "90-10 Cu-Ni"
    mat(3) = "Admiralty brass"
    mat(4) = "Aluminum brass"
    mat(5) = "Aluminum bronze 612"
    mat(6) = "Arsenical Copper"
    mat(7) = "Carbon steel AISI 1020"
    mat(8) = "Cast gray iron"
    mat(9) = "Cold rolled carbon steel"
    mat(10) = "Copper iron 194"
    mat(11) = "Muntz metal 280"
    mat(12) = "Aluminium"
    mat(13) = "Stainless steel Type 304/316"
    mat(14) = "Stainless steel Type 410/430"
    mat(15) = "Olin 194"
    mat(16) = "Titanium"
    mat(17) = "Titanium 5Al-2,5Sn"
    mat(18) = "Titanium 6Al-4V"
    mat(19) = "Stainless steel Type 329"
    mat(20) = "Yellow brass (High brass 268)"
    mat(21) = "Zirconium (commercial)"
    mat(22) = "ASTM A106 Gr B"
    mat(23) = "ASTM A210 A1 79A"
    mat(24) = "ASTM A515 GR60"
    mat(25) = "ASTM A516 GR70"
    mat(26) = "ASTM B111 ADMIRALTY"
    mat(27) = "Fe 35.2 UNI 663"
    mat(28) = "Fe 410.1 UNI 5869"
    mat(29) = "Fe 42.1 UNI 5869"
    mat(30) = "ASTM A179"
    mat(31) = "ASTM A789 S"
    mat(32) = "ASTM A285 C"
    mat(33) = "ASTM A213 TP 304L"
    mat(34) = "AQ 92 UNI 3965"
    mat(35) = "Naval Brass"
    mat(36) = "AQ 42"
    mat(37) = "ASTM B111 69"
    mat(38) = "ASTM A213 TP 316L"
    mat(39) = "AQ 45"
    mat(40) = "AQ 48"

    Material_cond(1) = 24.8:    Material_cond(2) = 38.4
    Material_cond(3) = 95.4:    Material_cond(4) = 86.8
    Material_cond(5) = 60.7:    Material_cond(6) = 84
    Material_cond(7) = 44.6:    Material_cond(8) = 38.4
    Material_cond(9) = 34:      Material_cond(10) = 76
    Material_cond(11) = 108.7:  Material_cond(12) = 100.4
    Material_cond(13) = 14:     Material_cond(14) = 21.4
    Material_cond(15) = 21.9:   Material_cond(16) = 14.1
    Material_cond(17) = 6.7:    Material_cond(18) = 6.2
    Material_cond(19) = 172.3:  Material_cond(20) = 102.9
    Material_cond(21) = 11.8:   Material_cond(22) = 44.6
    Material_cond(23) = 44.6:   Material_cond(24) = 44.6
    Material_cond(25) = 44.6:   Material_cond(26) = 95.4
    Material_cond(27) = 44.6:   Material_cond(28) = 44.6
    Material_cond(29) = 44.6:   Material_cond(30) = 44.6
    Material_cond(31) = 44.6:   Material_cond(32) = 44.6
    Material_cond(33) = 14:     Material_cond(34) = 44.6
    Material_cond(35) = 95.4:   Material_cond(36) = 44.6
    Material_cond(37) = 95.4:   Material_cond(38) = 14
    Material_cond(39) = 44.6:   Material_cond(40) = 44.6

    Fac(1, 1) = 0.64: Fac(1, 2) = 0.71: Fac(1, 3) = 0.77
    Fac(1, 4) = 0.82: Fac(1, 5) = 0.87: Fac(1, 6) = 0.9
    Fac(1, 7) = 0.93
    Fac(2, 1) = 0.74: Fac(2, 2) = 0.8:  Fac(2, 3) = 0.85
    Fac(2, 4) = 0.9:  Fac(2, 5) = 0.94: Fac(2, 6) = 0.97
    Fac(2, 7) = 0.99
    Fac(3, 1) = 0.87: Fac(3, 2) = 0.92: Fac(3, 3) = 0.96
    Fac(3, 4) = 1:    Fac(3, 5) = 1.02: Fac(3, 6) = 1.04
    Fac(3, 7) = 1.06
    Fac(4, 1) = 0.84: Fac(4, 2) = 0.9: Fac(4, 3) = 0.94
    Fac(4, 4) = 0.97: Fac(4, 5) = 1:   Fac(4, 6) = 1.02
    Fac(4, 7) = 1.03
    Fac(5, 1) = 0.89: Fac(5, 2) = 0.9: Fac(5, 3) = 0.94
    Fac(5, 4) = 0.97: Fac(5, 5) = 1:   Fac(5, 6) = 1.02
    Fac(5, 7) = 1.03
    Fac(6, 1) = 0.87: Fac(6, 2) = 0.92: Fac(6, 3) = 0.96
    Fac(6, 4) = 1:    Fac(6, 5) = 1.02: Fac(6, 6) = 1.04
    Fac(6, 7) = 1.06
    Fac(7, 1) = 0.74: Fac(7, 2) = 0.8:  Fac(7, 3) = 0.86
    Fac(7, 4) = 0.91: Fac(7, 5) = 0.95: Fac(7, 6) = 0.98
    Fac(7, 7) = 1
    Fac(8, 1) = 0.74: Fac(8, 2) = 0.8:  Fac(8, 3) = 0.86
    Fac(8, 4) = 0.91: Fac(8, 5) = 0.95: Fac(8, 6) = 0.98
    Fac(8, 7) = 1
    Fac(9, 1) = 0.74: Fac(9, 2) = 0.8:  Fac(9, 3) = 0.86
    Fac(9, 4) = 0.91: Fac(9, 5) = 0.95: Fac(9, 6) = 0.98
    Fac(9, 7) = 1
    Fac(10, 1) = 0.87: Fac(10, 2) = 0.92: Fac(10, 3) = 0.96
    Fac(10, 4) = 1:    Fac(10, 5) = 1.02: Fac(10, 6) = 1.04
    Fac(10, 7) = 1.06
    Fac(11, 1) = 0:    Fac(11, 2) = 0:    Fac(11, 3) = 0
    Fac(11, 4) = 0:    Fac(11, 5) = 0:    Fac(11, 6) = 0
    Fac(11, 7) = 0
    Fac(12, 1) = 0.87: Fac(12, 2) = 0.92: Fac(12, 3) = 0.96
    Fac(12, 4) = 1:    Fac(12, 5) = 1.02: Fac(12, 6) = 1.04
    Fac(12, 7) = 1.06
    Fac(13, 1) = 0.49: Fac(13, 2) = 0.56: Fac(13, 3) = 0.63
    Fac(13, 4) = 0.69: Fac(13, 5) = 0.75: Fac(13, 6) = 0.79
    Fac(13, 7) = 0.83
    Fac(14, 1) = 0.59: Fac(14, 2) = 0.65: Fac(14, 3) = 0.7
    Fac(14, 4) = 0.76: Fac(14, 5) = 0.82: Fac(14, 6) = 0.85
    Fac(14, 7) = 0.88
    Fac(15, 1) = 0.87: Fac(15, 2) = 0.92: Fac(15, 3) = 0.96
    Fac(15, 4) = 1:    Fac(15, 5) = 1.02: Fac(15, 6) = 1.04
    Fac(15, 7) = 1.06
    Fac(16, 1) = 0.71: Fac(16, 2) = 0.71: Fac(16, 3) = 0.71
    Fac(16, 4) = 0.71: Fac(16, 5) = 0.77: Fac(16, 6) = 0.81
    Fac(16, 7) = 0.85
    Fac(17, 1) = 0.71: Fac(17, 2) = 0.71: Fac(17, 3) = 0.71
    Fac(17, 4) = 0.71: Fac(17, 5) = 0.77: Fac(17, 6) = 0.81
    Fac(17, 7) = 0.85
    Fac(18, 1) = 0.71: Fac(18, 2) = 0.71: Fac(18, 3) = 0.71
    Fac(18, 4) = 0.71: Fac(18, 5) = 0.77: Fac(18, 6) = 0.81
    Fac(18, 7) = 0.85
    Fac(19, 1) = 0.54: Fac(19, 2) = 0.6:  Fac(19, 3) = 0.65
    Fac(19, 4) = 0.69: Fac(19, 5) = 0.74: Fac(19, 6) = 0.76
    Fac(19, 7) = 0.78
    Fac(20, 1) = 0.87: Fac(20, 2) = 0.92: Fac(20, 3) = 0.96
    Fac(20, 4) = 1:    Fac(20, 5) = 1.02: Fac(20, 6) = 1.04
    Fac(20, 7) = 1.06
    Fac(21, 1) = 0:    Fac(21, 2) = 0:    Fac(21, 3) = 0
    Fac(21, 4) = 0:    Fac(21, 5) = 0:    Fac(21, 6) = 0
    Fac(21, 7) = 0
    Fac(22, 1) = 0.74: Fac(22, 2) = 0.8:  Fac(22, 3) = 0.86
    Fac(22, 4) = 0.91: Fac(22, 5) = 0.95: Fac(22, 6) = 0.98
    Fac(22, 7) = 1
    Fac(23, 1) = 0.74: Fac(23, 2) = 0.8:  Fac(23, 3) = 0.86
    Fac(23, 4) = 0.91: Fac(23, 5) = 0.95: Fac(23, 6) = 0.98
    Fac(23, 7) = 1
    Fac(24, 1) = 0.74: Fac(24, 2) = 0.8:  Fac(24, 3) = 0.86
    Fac(24, 4) = 0.91: Fac(24, 5) = 0.95: Fac(24, 6) = 0.98
    Fac(24, 7) = 1
    Fac(25, 1) = 0.74: Fac(25, 2) = 0.8:  Fac(25, 3) = 0.86
    Fac(25, 4) = 0.91: Fac(25, 5) = 0.95: Fac(25, 6) = 0.98
    Fac(25, 7) = 1
    Fac(26, 1) = 0.87: Fac(26, 2) = 0.92: Fac(26, 3) = 0.96
    Fac(26, 4) = 1:    Fac(26, 5) = 1.02: Fac(26, 6) = 1.04
    Fac(26, 7) = 1.06
    Fac(27, 1) = 0.74: Fac(27, 2) = 0.8:  Fac(27, 3) = 0.86
    Fac(27, 4) = 0.91: Fac(27, 5) = 0.95: Fac(27, 6) = 0.98
    Fac(27, 7) = 1
    Fac(28, 1) = 0.74: Fac(28, 2) = 0.8:  Fac(28, 3) = 0.86
    Fac(28, 4) = 0.91: Fac(28, 5) = 0.95: Fac(28, 6) = 0.98
    Fac(28, 7) = 1
    Fac(29, 1) = 0.74: Fac(29, 2) = 0.8:  Fac(29, 3) = 0.86
    Fac(29, 4) = 0.91: Fac(29, 5) = 0.95: Fac(29, 6) = 0.98
    Fac(29, 7) = 1
    Fac(30, 1) = 0.74: Fac(30, 2) = 0.8:  Fac(30, 3) = 0.86
    Fac(30, 4) = 0.91: Fac(30, 5) = 0.95: Fac(30, 6) = 0.98
    Fac(30, 7) = 1
    Fac(31, 1) = 0.74: Fac(31, 2) = 0.8:  Fac(31, 3) = 0.86
    Fac(31, 4) = 0.91: Fac(31, 5) = 0.95: Fac(31, 6) = 0.98
    Fac(31, 7) = 1
    Fac(32, 1) = 0.74: Fac(32, 2) = 0.8:  Fac(32, 3) = 0.86
    Fac(32, 4) = 0.91: Fac(32, 5) = 0.95: Fac(32, 6) = 0.98
    Fac(32, 7) = 1
    Fac(33, 1) = 0.49: Fac(33, 2) = 0.56: Fac(33, 3) = 0.63
    Fac(33, 4) = 0.69: Fac(33, 5) = 0.75: Fac(33, 6) = 0.79
    Fac(33, 7) = 0.83
    Fac(34, 1) = 0.74: Fac(34, 2) = 0.8:  Fac(34, 3) = 0.86
    Fac(34, 4) = 0.91: Fac(34, 5) = 0.95: Fac(34, 6) = 0.98
    Fac(34, 7) = 1
    Fac(35, 1) = 0.87: Fac(35, 2) = 0.92: Fac(35, 3) = 0.96
    Fac(35, 4) = 1:    Fac(35, 5) = 1.02: Fac(35, 6) = 1.04
    Fac(35, 7) = 1.06
    Fac(36, 1) = 0.74: Fac(36, 2) = 0.8:  Fac(36, 3) = 0.86
    Fac(36, 4) = 0.91: Fac(36, 5) = 0.95: Fac(36, 6) = 0.98
    Fac(36, 7) = 1
    Fac(37, 1) = 0.87: Fac(37, 2) = 0.92: Fac(37, 3) = 0.96
    Fac(37, 4) = 1:    Fac(37, 5) = 1.02: Fac(37, 6) = 1.04
    Fac(37, 7) = 1.06
    Fac(38, 1) = 0.49: Fac(38, 2) = 0.56: Fac(38, 3) = 0.63
    Fac(38, 4) = 0.69: Fac(38, 5) = 0.75: Fac(38, 6) = 0.79
    Fac(38, 7) = 0.83
    Fac(39, 1) = 0.74: Fac(39, 2) = 0.8:  Fac(39, 3) = 0.86
    Fac(39, 4) = 0.91: Fac(39, 5) = 0.95: Fac(39, 6) = 0.98
    Fac(39, 7) = 1
    Fac(40, 1) = 0.74: Fac(40, 2) = 0.8:  Fac(40, 3) = 0.86
    Fac(40, 4) = 0.91: Fac(40, 5) = 0.95: Fac(40, 6) = 0.98
    Fac(40, 7) = 1

    metal = Combo_TUBES_Mat.Text
    D54 = 0
    For i = 1 To 40
        If metal = mat(i) Then
            D78 = Material_cond(i)
            If D69 = 10 Then
                D54 = Fac(i, 1)
            ElseIf D69 = 12 Then
                D54 = Fac(i, 1)
            ElseIf D69 = 14 Then
                D54 = Fac(i, 2)
            ElseIf D69 = 16 Then
                D54 = Fac(i, 3)
            ElseIf D69 = 18 Then
                D54 = Fac(i, 4)
            ElseIf D69 = 20 Then
                D54 = Fac(i, 5)
            ElseIf D69 = 22 Then
                D54 = Fac(i, 6)
            ElseIf D69 = 24 Then
                D54 = Fac(i, 7)
            End If
        End If
    Next i
    If Check_MAT_FACTOR = Checked And Combo_S_FLUID = "Steam" Then
        D54 = Data1.Recordset.Mat_fact
        Mat_factor.ForeColor = &HFFFFFF
        Mat_factor.BackColor = &HC0&
    ElseIf Check_MAT_FACTOR = Unchecked And Combo_S_FLUID = "Steam" Then
        Mat_factor.Text = D54                 ' Material factor
        Mat_factor.ForeColor = &HC0&
        Mat_factor.BackColor = &HE0E0E0
    End If
End Sub
Private Sub Combo_Plant_1_LostFocus()
On Error Resume Next
    PPP1 = Combo_Plant_1.Text
10 End Sub
Private Sub Combo_Unit_1_GotFocus()
On Error Resume Next
    Dim Rs4 As Recordset
    Data4.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
    Data4.RecordSource = "Select * From [Query_Unit]"
    Data4.Refresh
    Set Rs4 = Data4.Recordset
    Combo_UNIT_1.Clear
    PPP1 = Combo_Plant_1
    If Rs4.RecordCount > 0 Then
       Do Until Rs4.EOF
            PPP2 = Data4.Recordset.Plant
            UUU1 = Data4.Recordset.Unit_name
            If PPP1 = PPP2 Then
                Combo_UNIT_1.AddItem UUU1
            End If
            Rs4.MoveNext
        Loop
    End If
End Sub
Private Sub Combo_UNIT_1_Lostfocus()
On Error Resume Next
    UUU1 = Combo_Plant_1.Text
End Sub
Private Sub Combo_Date_X_GotFocus()
On Error Resume Next
    Data7.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
    Data7.RecordSource = "Select * From [QUERY_Date]"
    Data7.Refresh
    Set Rs7 = Data7.Recordset
    Combo_Date_X.Clear
    PPP1 = Combo_Plant_1
    UUU1 = Combo_UNIT_1.Text
    If Rs7.RecordCount > 0 Then
        Do Until Rs7.EOF
            PPP2 = Data7.Recordset.Plant
            UUU2 = Data7.Recordset.Unit_name
            Date_X = Data7.Recordset.date_test
                If UUU1 = UUU2 And PPP1 = PPP2 Then
                    Combo_Date_X.AddItem Date_X
                End If
            Rs7.MoveNext
        Loop
    End If
End Sub
Private Sub Combo_Date_X_LostFocus()
On Error Resume Next
    Dim Date_1 As Date
    Dim Date_2 As Date
    Data1.Recordset.MoveLast
    n_record = Data1.Recordset.RecordCount
    Data1.Recordset.MoveFirst
    pos = Data1.Recordset.AbsolutePosition
    While pos < 0
        Data1.Recordset.MoveNext
    Wend
    Date_1 = Combo_Date_X.Text
    XXX = 1
    While Date_1 <> Date_2
        n_rec_a = Data1.Recordset.AbsolutePosition + 1
        Date_2 = Data1.Recordset.date_test
        If Date_1 = Date_2 Then
            XXX = 0
            If n_rec_a = 1 Then
                Data1.Recordset.MoveNext
                Data1.Recordset.MovePrevious
            Else
                Data1.Recordset.MovePrevious
                Data1.Recordset.MoveNext
            End If
            GoTo 12
        End If
        If n_rec_a = n_record Then
            MsgBox "Selected date not  found in the database"
            GoTo 12
        End If
        Data1.Recordset.MoveNext
    Wend
12  End Sub
Private Sub Spin_TARGET_T_Change()
    If YXY = 1 Then
        Exit Sub
    End If
    PROCESS_TARGET_T_OUT = Spin_TARGET_T / 10
    If CSng(SHELL_TEMP_OUT) > CSng(PROCESS_TARGET_T_OUT) Then
        SHELL_TEMP_OUT.ForeColor = &H80000018
        SHELL_TEMP_OUT.BackColor = 255.255
    Else
        SHELL_TEMP_OUT.BackColor = &H80000018
        SHELL_TEMP_OUT.ForeColor = &HC0&
    End If
End Sub
Private Sub TabStrip1_Click()
On Error Resume Next
    If TabStrip1.SelectedItem = "Update" Then
        Data1.UpdateRecord
        Data1.Recordset.Bookmark = Data1.Recordset.LastModified
    ElseIf TabStrip1.SelectedItem = "Add" Then
        Call CheckLockedStatus(temp)
        If temp = "locked" Then
            MsgBox "You cannot add new records in the trial version."
            Exit Sub
        End If
        XXX = 1
        XYX = 1
        Reply = MsgBox("Confirm to add one record?", vbYesNo, "Add record")
        If Reply = vbYes Then
            Data1.Recordset.AddNew
            Data1.UpdateRecord
            Data1.Recordset.Bookmark = Data1.Recordset.LastModified
            Data1.Recordset.MoveLast
            txt_num.Text = Data1.Recordset.AbsolutePosition + 1
            Call RESET
            Data1.UpdateRecord
            Data1.Recordset.Bookmark = Data1.Recordset.LastModified
        End If
        XXX = 0
    ElseIf TabStrip1.SelectedItem = "Close" Then
        Unload Me
    ElseIf TabStrip1.SelectedItem = "Delete" Then
        If Data1.Recordset.AbsolutePosition + 1 = 1 Then
            MsgBox (" You cannot delete this record")
            Exit Sub
        End If
        If txt_num = 1 Then
            MsgBox (" You cannot delete this record")
            Exit Sub
        End If
        If Check_des.Value = Unchecked Then
            Reply = MsgBox("Confirm to delete this record?", vbYesNo, "Delete record")
        Else
            Reply = MsgBox("This record contain the design data! Confirm to delete this record?", vbYesNo, "Delete record")
        End If
        If Reply = vbYes Then
            Data1.Recordset.Delete
            Data1.Recordset.MovePrevious
        End If
        XXX = 0
    ElseIf TabStrip1.SelectedItem = "Refresh" Then
        Data1.Refresh
    ElseIf TabStrip1.SelectedItem = "Print" Then
        frmHX.PrintForm
    ElseIf TabStrip1.SelectedItem = "Get design" Then
        Call CheckLockedStatus(temp)
        If temp = "locked" Then
            MsgBox "This feature is not allowed in the trial version."
            Exit Sub
        End If
        DES_1 = 1
        UNIT1 = Data1.Recordset.Unit_name
        PLANT1 = Data1.Recordset.Plant
        C_DES1 = Data1.Recordset.CHECK_DESIGN
        Data2.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
        Data2.RecordSource = "Select * From [Query_test]"
        Data2.Refresh
        Set Rs2 = Data2.Recordset
        Data2.Recordset.MoveFirst
        If Rs2.RecordCount > 0 Then
            Do Until Rs2.EOF
                UNIT2 = Data2.Recordset.Unit_name
                PLANT2 = Data2.Recordset.Plant
                C_DES2 = Data2.Recordset.CHECK_DESIGN
                If UNIT1 = UNIT2 And PLANT1 = PLANT2 And C_DES2 = -1 Then
                    Plant.Text = Data2.Recordset.Plant
                    Location.Text = Data2.Recordset.Location
                    Country.Text = Data2.Recordset.Country
                    Unit.Text = Data2.Recordset.Unit_name
                    
                    Combo_PLANT.Text = Plant.Text
                    Combo_LOC.Text = Location.Text
                    Combo_Country.Text = Country.Text
                    Combo_UNIT.Text = Unit.Text
                     
                    Combo_PLANT_UNIT = Data2.Recordset.PLANT_UNIT
                    Combo_PROCESS_DESCR = Data2.Recordset.PROCESS_DESCR
                    Combo_PROCESS_STREAM = Data2.Recordset.PROCESS_STREAM
                    Combo_COOL_TOWER = Data2.Recordset.COOL_TOWER
                    Combo_TEMA = Data2.Recordset.TEMA
                    Combo_POSITION = Data2.Recordset.Position
                    PARALLEL_N = Data2.Recordset.PARALLEL_N
                    SERIES_N = Data2.Recordset.SERIES_N
                    Combo_CURRENT = Data2.Recordset.Current
                    Combo_cooling_type = Data2.Recordset.COOLING_TYPE
                     
                    T_NO.Text = Data2.Recordset.TUBES_NO
                    T_len.Text = Data2.Recordset.TUBES_LE
                    T_PASS.Text = Data2.Recordset.TUBES_PASSES
                    T_OD.Text = Data2.Recordset.TUBES_OD
                    Combo_BWG.Text = Data2.Recordset.TUBES_BWG
                    Combo_TUBES_Mat.Text = Data2.Recordset.TUBES_MAT
                    
                    SHELL_PASS.Text = Data2.Recordset.SHELL_PASS
                    SHELL_BAFFLES_N.Text = Data2.Recordset.SHELL_BAFFLES_N
                    SHELL_BAFFLES_CUT.Text = Data2.Recordset.SHELL_BAFFLES_CUT
                    SHELL_BAFFLES_SPACE.Text = Data2.Recordset.SHELL_BAFFLES_SPACE
                    SHELL_ID.Text = Data2.Recordset.SHELL_ID
                    SHELL_TUBES_PITCH.Text = Data2.Recordset.SHELL_TUBES_PITCH
                    SHELL_PITCH_CONF.Text = Data2.Recordset.SHELL_PITCH_CONF
                    Combo_SHELL_MAT.Text = Data2.Recordset.SHELL_MAT
                    
                    FACT_FLOW = Data2.Recordset.FACT_FLOW
                    
                    Combo_T_FLUID.Text = Data2.Recordset.TUBES_FLUID
                    TUBES_FLOW.Text = Data2.Recordset.TUBES_FLOW
                    TUBES_VAPOR.Text = Data2.Recordset.TUBES_VAPOR
                    TUBES_LIQUID.Text = Data2.Recordset.TUBES_LIQUID
                    TUBES_WATER.Text = Data2.Recordset.TUBES_WATER
                    TUBES_NON_COND.Text = Data2.Recordset.TUBES_NON_COND
                    TUBES_TEMP_IN.Text = Data2.Recordset.TUBES_TEMP_IN
                    TUBES_TEMP_OUT.Text = Data2.Recordset.TUBES_TEMP_OUT
                    TUBES_P_IN.Text = Data2.Recordset.TUBES_P_IN
                    TUBES_P_OUT.Text = Data2.Recordset.TUBES_P_OUT
                     
                    Combo_S_FLUID.Text = Data2.Recordset.SHELL_FLUID
                    SHELL_FLOW.Text = Data2.Recordset.SHELL_FLOW
                    SHELL_VAPOR.Text = Data2.Recordset.SHELL_VAPOR
                    SHELL_LIQUID.Text = Data2.Recordset.SHELL_LIQUID
                    SHELL_WATER.Text = Data2.Recordset.SHELL_WATER
                    SHELL_NON_COND.Text = Data2.Recordset.SHELL_NON_COND
                    SHELL_TEMP_IN.Text = Data2.Recordset.SHELL_TEMP_IN
                    SHELL_TEMP_OUT.Text = Data2.Recordset.SHELL_TEMP_OUT
                    PROCESS_TARGET_T_OUT = Data2.Recordset.PROCESS_TARGET_TEMP
                    SHELL_P_IN.Text = Data2.Recordset.SHELL_P_IN
                    SHELL_P_OUT.Text = Data2.Recordset.SHELL_P_OUT
                    S_press_KP_1 = Data2.Recordset.Press_COND
                    S_press_KP.Text = S_press_KP_1
                    SHELL_OUT(10).Text = Data2.Recordset.SHELL_LATENT
                    PROCESS_TARGET_T_OUT = Data2.Recordset.PROCESS_TARGET_TEMP
                    
                    Vtot_INP_1 = Data2.Recordset.VAP_FRACTION
                    Vtot_INP.Text = Format(Vtot_INP_1, "0.0")
                    Vwat_INP_1 = Data2.Recordset.Vwat_perc
                    Vwat_INP.Text = Format(Vwat_INP_1, "0.0")
                    Ltot_INP_1 = Data2.Recordset.LIQ_FRACTION
                    Ltot_INP.Text = Format(Ltot_INP_1, "0.0")
                    Lwat_INP_1 = Data2.Recordset.Lwat_perc
                    Lwat_INP.Text = Format(Lwat_INP_1, "0.0")
                    
                    YXY = 1
                    Spin_VAP_P.Value = Vtot_INP_1 * 10
                    Spin_WAT_VAP_IN = Vwat_INP_1 * 10
                    Spin_WAT_LIQ_IN = Lwat_INP_1 * 10
                    YXY = 0
                    
                    U = Data2.Recordset.Check_X
                    If U = 0 Then
                        lungh = 1
                        Check_U = Unchecked
                    ElseIf U = -1 Then
                        lungh = 2
                        Check_U = Checked
                    End If
                    
                    Check_PF.Value = Data2.Recordset.Check_PF
                    Check_P_DROP_T.Value = Data2.Recordset.Check_P_DROP_T
                    Check_P_DROP_S.Value = Data2.Recordset.Check_P_DROP_S
                    Check_S_TC.Value = Data2.Recordset.Check_S_TC
                    Check_S_SPH.Value = Data2.Recordset.Check_S_SPH
                    Check_S_DENS.Value = Data2.Recordset.Check_S_DENS
                    Check_S_VISC.Value = Data2.Recordset.Check_S_VISC
                    Check_T_SPH.Value = Data2.Recordset.Check_T_SPH
                    Check_T_DENS.Value = Data2.Recordset.Check_T_DENS
                    Check_T_VISC.Value = Data2.Recordset.Check_T_VISC
                    Check_LATENT.Value = Data2.Recordset.Check_LATENT
                    Check_CT.Value = Data2.Recordset.Check_CT
                    Check_U_CLEAN.Value = Data2.Recordset.Check_U_CLEAN
                    
                    If Combo_S_FLUID.Text <> "Water" Then
                        SHELL_OUT(1) = Data2.Recordset.SHELL_T_COND
                        SHELL_OUT(3) = Data2.Recordset.SHELL_SPH
                        SHELL_OUT(4) = Data2.Recordset.SHELL_DENS
                        SHELL_OUT(5) = Data2.Recordset.SHELL_VISC
                    End If
                    
                    Spin_PARALLEL_N.Value = PARALLEL_N
                    Spin_SERIES_N.Value = SERIES_N
                    HScroll_T_NO.Value = T_NO.Text
                    Spin_TUBES_PITCH = SHELL_TUBES_PITCH * 10
                    Spin_S_PASS = SHELL_PASS
                    Spin_BAFFLES_N = SHELL_BAFFLES_N
                    Spin_BAFFLES_CUT = SHELL_BAFFLES_CUT
                    Spin_BAFFLES_SPACE = SHELL_BAFFLES_SPACE
                    Spin_SHELL_ID = SHELL_ID
                    Spin_TUBES_PITCH = SHELL_TUBES_PITCH * 10
                    Spin_T_LEN.Value = T_len.Text * 100
                    Spin_T_PAS.Value = T_PASS.Text
                    Spin_T_OD.Value = T_OD.Text * 100
                    YXY = 1
                    HScroll_TUBES_FLOW.Value = TUBES_FLOW / FACT_FLOW
                    HScroll_TUBES_VAPOR.Value = TUBES_VAPOR / FACT_FLOW
                    HScroll_TUBES_LIQUID.Value = TUBES_LIQUID / FACT_FLOW
                    HScroll_TUBES_WATER.Value = TUBES_WATER / FACT_FLOW
                    HScroll_TUBES_NON_COND.Value = TUBES_NON_COND / FACT_FLOW
    '                HScroll_TUBES_TC = TUBES_OUT(1) * 1000
                    HScroll_TUBES_SPH = SHELL_OUT(3) * 1000
                    HScroll_TUBES_DENS = SHELL_OUT(4) * 10
                    HScroll_TUBES_VISC = SHELL_OUT(5) * 1000
                    Spin_TUBES_T_IN.Value = TUBES_T_IN * 100
                    Spin_TUBES_T_OUT.Value = TUBES_T_OUT * 100
                    HScroll_TUBES_P_IN.Value = TUBES_P_IN * 100
                    HScroll_TUBES_P_OUT.Value = TUBES_P_OUT * 100
                    
                    Spin_SHELL_FLOW.Value = SHELL_FLOW / FACT_FLOW
                    HScroll_SHELL_NON_COND.Value = SHELL_NON_COND / FACT_FLOW
                    HScroll_SHELL_T_IN.Value = SHELL_TEMP_IN * 100
                    HScroll_SHELL_T_OUT.Value = SHELL_TEMP_OUT * 100
                    HScroll_SHELL_P_IN.Value = SHELL_P_IN * 100
                    HScroll_SHELL_P_OUT.Value = SHELL_P_OUT * 100
                    Spin_S_PRESS.Value = S_press_KP_1 * 100
                    HScroll_SHELL_SPH = SHELL_OUT(3) * 1000
                    HScroll_SHELL_DENS = SHELL_OUT(4) * 10
                    HScroll_SHELL_VISC = SHELL_OUT(5) * 1000
                    
                    LMTD = Data2.Recordset.LMTD
                    TTD = Data2.Recordset.TTD
                    MTDc = Data2.Recordset.MTDc
                    SKIN_TEMP = Data2.Recordset.SKIN_TEMP
                    C_Factor = Data2.Recordset.C_Factor
                    TUBES_FF_1 = Data2.Recordset.TUBES_FF
                    TUBES_FF.Text = TUBES_FF_1
                    U_COEFF_CLEAN = Data2.Recordset.Clean
                    HScroll_U_CLEAN = U_COEFF_CLEAN
                    WATER_FF = Data2.Recordset.WATER_FF
                    FACT_FLOW = Data2.Recordset.FACT_FLOW
                    Spin_FACT_FLOW.Value = FACT_FLOW
                    
                    If Check_PF = 1 Then
                        SHELL_FF = Format(Data2.Recordset.SHELL_FF, "0.00")
                        Spin_PF = SHELL_FF.Text ^ 100
                    End If
                    If Check_T_SPH = 1 Then
                        TUBES_OUT(3) = Data2.Recordset.TUBES_SPH
                    End If
                    If Check_T_DENS = 1 Then
                        TUBES_OUT(4) = Data2.Recordset.TUBES_DENS
                    End If
                    If Check_T_VISC = 1 Then
                        TUBES_OUT(5) = Data2.Recordset.TUBES_VISC
                    End If
                    If Check_S_TC = 1 Then
                        SHELL_OUT(1) = Data2.Recordset.SHELL_T_COND
                    End If
                    If Check_S_SPH = 1 Then
                        SHELL_OUT(3) = Data2.Recordset.SHELL_SPH
                    End If
                    If Check_S_DENS = 1 Then
                        SHELL_OUT(4) = Data2.Recordset.SHELL_DENS
                    End If
                    If Check_S_VISC = 1 Then
                        SHELL_OUT(5) = Data2.Recordset.SHELL_VISC
                    End If
                    If Check_LATENT = 1 Then
                        SHELL_OUT(10) = Data2.Recordset.SHELL_LATENT
                        HScroll_LATENT = SHELL_OUT(10)
                        S_10 = SHELL_OUT(10)
                    End If
                    If Check_CT = 1 Then
                        SHELL_OUT(11) = Data2.Recordset.Temp_COND
                        S_11 = SHELL_OUT(11)
                         HScroll_C_TEMP = S_11 * 100
                    End If
                    Exit Do
                End If
                RR = Data2.Recordset.AbsolutePosition + 1
                Data1.UpdateRecord
                Rs2.MoveNext
                If Data2.Recordset.EOF Then
                   MsgBox ("Not found the design data for this unit")
                   Exit Do
                End If
            Loop
        End If
        If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Or Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam" Or Combo_S_FLUID.Text = "Water" Then
            SHELL_OUT(1).ForeColor = &HFF0000
            SHELL_OUT(1).BackColor = &HE0E0E0
            SHELL_OUT(4).ForeColor = &HFF0000
            SHELL_OUT(4).BackColor = &HE0E0E0
            SHELL_OUT(5).ForeColor = &HFF0000
            SHELL_OUT(5).BackColor = &HE0E0E0
            SHELL_OUT(10).ForeColor = &HFF0000
            SHELL_OUT(10).BackColor = &HE0E0E0
            SHELL_OUT(11).ForeColor = &HFF0000
            SHELL_OUT(11).BackColor = &HE0E0E0
        Else
            SHELL_OUT(1).ForeColor = &HC0&
            SHELL_OUT(1).BackColor = &HE0E0E0
            SHELL_OUT(4).ForeColor = &HC0&
            SHELL_OUT(4).BackColor = &HE0E0E0
            SHELL_OUT(5).ForeColor = &HC0&
            SHELL_OUT(5).BackColor = &HE0E0E0
            SHELL_OUT(10).ForeColor = &HC0&
            SHELL_OUT(10).BackColor = &HE0E0E0
            SHELL_OUT(11).ForeColor = &HC0&
            SHELL_OUT(11).BackColor = &HE0E0E0
        End If
        If Combo_S_FLUID.Text = "Water" Then
            SHELL_OUT(3).ForeColor = &HFF0000
            SHELL_OUT(3).BackColor = &HE0E0E0
        Else
            SHELL_OUT(3).ForeColor = &HC0&
            SHELL_OUT(3).BackColor = &HE0E0E0
        End If
        Data1.UpdateRecord
        Data1.Recordset.Bookmark = Data1.Recordset.LastModified
        YXY = 0
Call Fluid_type
    ElseIf TabStrip1.SelectedItem = "Reset" Then
        Call RESET
        Call RESET
    End If
10 End Sub
Private Sub RESET()
On Error Resume Next
    Data1.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
    Data1.RecordSource = "Select * From [Query_test]"
    Set Rs1 = Data1.Recordset
        COL = 0
        DTPicker1.Value = Date
        Plant.Text = "Plant"
        Location.Text = "Location"
        Country.Text = "Country"
        Unit.Text = "UNIT"

        Combo_PLANT.Text = "Plant"
        Combo_LOC.Text = "Location"
        Combo_Country.Text = "Country"
        Combo_UNIT.Text = "UNIT"
        Combo_cooling_type = "Cooling"
        Combo_CURRENT = "Counter-flow"
        Check_des.Value = Unchecked
        Combo_PLANT_UNIT = "PLANT UNIT"
        Combo_PROCESS_DESCR = "PROCESS DESCRIPTION"
        Combo_PROCESS_STREAM = "PROCESS STREAM"
        Combo_COOL_TOWER = "TOWER"
        
        Combo_TEMA = "AES"
        Combo_POSITION = "Horizontal"
        PARALLEL_N.Text = 1
        SERIES_N.Text = 1
        
        Spin_PARALLEL_N.Value = PARALLEL_N
        Spin_SERIES_N.Value = SERIES_N
        
        T_NO.Text = 200
        T_len.Text = 6.09
        T_PASS.Text = 1
        T_OD.Text = 19.05
        Combo_BWG.Text = 14
        Combo_TUBES_Mat.Text = "Carbon steel AISI 1020"
        Combo_MAT_SHEET = "Carbon steel AISI 1020"
        lungh = 1
        HScroll_T_NO.Value = T_NO.Text
        Spin_T_LEN.Value = T_len.Text * 100
        Spin_T_PAS.Value = T_PASS.Text
        Spin_T_OD.Value = T_OD.Text * 100
        
        SHELL_PASS.Text = 1
        SHELL_BAFFLES_N.Text = 10
        SHELL_BAFFLES_CUT.Text = 25
        SHELL_BAFFLES_SPACE.Text = 200
        SHELL_ID.Text = 500
        SHELL_TUBES_PITCH.Text = 25.4
        SHELL_PITCH_CONF.Text = "Triangular"
        Combo_SHELL_MAT.Text = "Carbon steel AISI 1020"
                
        Spin_S_PASS.Value = SHELL_PASS
        Spin_BAFFLES_N.Value = SHELL_BAFFLES_N
        Spin_BAFFLES_CUT.Value = SHELL_BAFFLES_CUT
        Spin_BAFFLES_SPACE.Value = SHELL_BAFFLES_SPACE
        Spin_SHELL_ID.Value = SHELL_ID
        Spin_TUBES_PITCH.Value = SHELL_TUBES_PITCH * 10
        
        Combo_T_FLUID = "Water"
        Combo_S_FLUID = "Water"
        
        TUBES_FLOW.Text = Format(225000, "#,##0")
        TUBES_VAPOR.Text = Format(0, "#,##0")
        TUBES_LIQUID = Format(0, "#,##0")
        TUBES_WATER = Format(125000, "#,##0")
        TUBES_NON_COND = Format(0, "#,##0")
        TUBES_OUT(0) = Format(TUBES_FLOW.Text, "#,##0")
        TUBES_OUT(0).BackColor = &HE0E0E0
        TUBES_OUT(0).ForeColor = &HC0&
        
        Spin_FACT_FLOW.Value = 10
        FACT_FLOW = 10
        
        HScroll_TUBES_FLOW.Value = TUBES_FLOW / FACT_FLOW
        HScroll_TUBES_VAPOR = TUBES_VAPOR / FACT_FLOW
        HScroll_TUBES_LIQUID = TUBES_LIQUID / FACT_FLOW
        HScroll_TUBES_WATER = TUBES_WATER / FACT_FLOW
        HScroll_TUBES_NON_COND = TUBES_NON_COND / FACT_FLOW
        
        TUBES_TEMP_IN.Text = Format(25, "0.00")
        TUBES_TEMP_OUT.Text = Format(35, "0.00")
        TUBES_P_IN.Text = Format(4, "0.00")
        TUBES_P_OUT.Text = Format(3.5, "0.00")
        TUBES_OUT(1).Text = Format(0.5, "0.00")
        TUBES_OUT(3).Text = Format(0.999, "0.000")
        TUBES_OUT(4).Text = Format(1000, "0.0")
        TUBES_OUT(5).Text = Format(0.8, "0.00")
        
        TUBES_OUT(9).Text = Format(5, "0.00")
       
        Spin_TUBES_T_IN.Value = TUBES_TEMP_IN * 100
        Spin_TUBES_T_OUT.Value = TUBES_TEMP_OUT * 100
        HScroll_TUBES_P_IN.Value = TUBES_P_IN * 100
        HScroll_TUBES_P_OUT.Value = TUBES_P_OUT * 100
        
        SHELL_FLOW.Text = Format(50000, "#,##0")
        SHELL_VAPOR.Text = Format(0, "#,##0")
        SHELL_LIQUID = Format(45040, "#,##0")
        SHELL_WATER = Format(0, "#,##0")
        SHELL_NON_COND = Format(0, "#,##0")
        SHELL_OUT(0) = Format(SHELL_FLOW.Text, "#,##0")
        
        Spin_SHELL_FLOW.Value = SHELL_FLOW / FACT_FLOW
        HScroll_SHELL_NON_COND = SHELL_NON_COND / FACT_FLOW
        VAP_PERC = 100
        'V_P(4) = 100

        SHELL_TEMP_IN.Text = Format(85, "0.00")
        SHELL_TEMP_OUT.Text = Format(40, "0.00")
        PROCESS_TARGET_T_OUT = Format(45, "0.00")
        SHELL_P_IN.Text = Format(10, "0.00")
        SHELL_P_OUT.Text = Format(9.5, "0.00")
        S_press_KP_1 = 5
        S_press_KP.Text = Format(S_press_KP_1, "0.00")
        SHELL_FF = Format(1, "0.00")
        PROCESS_TARGET_T_OUT = 45
        Spin_TARGET_T.Value = 400
        
        HScroll_SHELL_T_IN.Value = SHELL_TEMP_IN * 100
        HScroll_SHELL_T_OUT.Value = SHELL_TEMP_OUT * 100
        HScroll_SHELL_P_IN.Value = SHELL_P_IN * 100
        HScroll_SHELL_P_OUT.Value = SHELL_P_OUT * 100
        Spin_S_PRESS.Value = S_press_KP_1 * 100
        Spin_PF = SHELL_FF * 100
        
        SHELL_OUT(1) = Format(0.5, "0.000")
        SHELL_OUT(3) = Format(0.999, "0.000")
        SHELL_OUT(4) = Format(1000, "0.0")
        SHELL_OUT(5) = Format(0.8, "0.000")
        SHELL_OUT(9).Text = Data1.Recordset.SHELL_PRESS_DROP
        SHELL_OUT(10) = Format(1000, "0.0")
        WATER_FF = Format(4, "0.00")
        
        Lwat_IN = 0
        Wet_steam = Format(1 - Lwat_IN, "0.000")
        U_COEFF_CLEAN = 1150
        
        HScroll_P_DROP_S = SHELL_OUT(9) * 100
    
        Check_U = Unchecked
        Check_T_TC = Unchecked
        Check_T_SPH = Unchecked
        Check_T_DENS = Unchecked
        Check_T_VISC = Unchecked
        Check_P_DROP_T = Unchecked
        Check_PF = Unchecked
        Check_CP = Unchecked
        Check_S_TC = Unchecked
        Check_S_SPH = Unchecked
        Check_S_DENS = Unchecked
        Check_S_VISC = Unchecked
        Check_LATENT = Unchecked
        Check_CT = Unchecked
        Check_P_DROP_S = Unchecked
        Check_U_CLEAN = Unchecked
Data1.UpdateRecord
Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Call Fluid_type
End Sub
Private Sub ACTUAL_CHECK_RESET()
On Error Resume Next
    Dim Rs2 As Recordset
    XXX = 1
    Data2.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
    Data2.RecordSource = "Select * From [Query_Test]"
    Set Rs2 = Data2.Recordset
    Rs2.MoveLast
    rec1 = Data1.Recordset.AbsolutePosition + 1
    Rs2.MoveFirst
    PPP1 = Data1.Recordset.Plant
    UUU1 = Data1.Recordset.Unit_name
    If Data2.Recordset.RecordCount > 0 Then
       Do Until Rs2.EOF
            PPP2 = Data2.Recordset.Plant
            UUU2 = Data2.Recordset.Unit_name
            rec2 = Data2.Recordset.AbsolutePosition + 1
            CHT2 = Data2.Recordset.CHECK_ACTUAL
            If rec1 <> rec2 Then
                If PPP1 = PPP2 And UUU1 = UUU2 And CHT2 = -1 Then
                    Data1.Recordset.AbsolutePosition = rec2 - 1
                    Check_ACT.Value = 0
                End If
            End If
            If Data2.Recordset.EOF Then
                  Exit Do
            End If
            Data2.Recordset.MoveNext
       Loop
    End If
    Data1.Recordset.AbsolutePosition = rec1 - 1
    XXX = 0
End Sub
Private Sub Comm_search_grid_Click()
On Error Resume Next
Call CheckLockedStatus(temp)
    If temp = "locked" Then
        MsgBox "This feature is not allowed in the trial version."
        Exit Sub
    End If
    Data12.Refresh
    Frame_Search.Visible = True
End Sub
Private Sub Comm_search_close_Click()
On Error Resume Next
    Frame_Search.Visible = False
End Sub
Private Sub No_balance_Click()
    If No_balance = True Then
        Check_T_OUT = Unchecked
        Thermal_bal_tubes.BackColor = &H8000000F
        Thermal_bal_tubes.ForeColor = &H80&
        Thermal_bal_shell.BackColor = &H8000000F
        Thermal_bal_shell.ForeColor = &H80&
        Thermal_bal_shell_T.BackColor = &H8000000F
        Thermal_bal_shell_T.ForeColor = &H80&
        If SHELL_TEMP_OUT > PROCESS_TARGET_T_OUT Then
            SHELL_TEMP_OUT.ForeColor = &H80000018
            SHELL_TEMP_OUT.BackColor = 255.255
        Else
            SHELL_TEMP_OUT.BackColor = &H80000018
            SHELL_TEMP_OUT.ForeColor = &HC0&
        End If
        SHELL_OUT(7).ForeColor = &HC0&
        SHELL_OUT(7).BackColor = &HE0E0E0
        SHELL_FLOW.ForeColor = &HC0&
        SHELL_FLOW.BackColor = &H80000018
        TUBES_FLOW.ForeColor = &HC0&
        TUBES_FLOW.BackColor = &H80000018
        TUBES_TEMP_OUT.ForeColor = &HC0&
        TUBES_TEMP_OUT.BackColor = &H80000018
        SHELL_OUT(0).ForeColor = &HC0&
        SHELL_OUT(0).BackColor = &HE0E0E0
        TUBES_OUT(0).ForeColor = &HC0&
        TUBES_OUT(0).BackColor = &HE0E0E0
    End If
End Sub
Private Sub Thermal_bal_tubes_Click()
On Error Resume Next
    If Thermal_bal_tubes = True Then
        Thermal_bal_tubes.BackColor = &H8000&
        Thermal_bal_tubes.ForeColor = &H80&
        Thermal_bal_shell.Value = False
        Thermal_bal_shell_T.Value = False
        Check_T_OUT = Unchecked
    Else
        Check_T_OUT = Unchecked
        Thermal_bal_tubes.BackColor = &H8000000F
        Thermal_bal_tubes.ForeColor = &H80&
    End If
Call Fluid_type
End Sub
Private Sub Thermal_bal_shell_Click()
    If Thermal_bal_shell = True Then
        Thermal_bal_shell.BackColor = &H8000&
        Thermal_bal_shell.ForeColor = &H80&
        Thermal_bal_tubes = False
        Thermal_bal_shell_T = False
'        Check_T_Flow = Unchecked
    Else
        Thermal_bal_shell.BackColor = &H8000000F
        Thermal_bal_shell.ForeColor = &H80&
    End If
Call Fluid_type
End Sub
Private Sub Thermal_bal_shell_T_Click()
    If Thermal_bal_shell = True Then
        Thermal_bal_shell_T.BackColor = &H8000&
        Thermal_bal_shell_T.ForeColor = &HFFFF&
        Thermal_bal_tubes = False
        Thermal_bal_shell = False
'        Check_T_Flow = Unchecked
    Else
        Thermal_bal_shell_T.BackColor = &H8000000F
        Thermal_bal_shell_T.ForeColor = &H80&
    End If
        Call COOLERS
End Sub
Private Sub Toggle_remarks_Click()
On Error Resume Next
Call CheckLockedStatus(temp)
    If temp = "locked" Then
        MsgBox "This feature is not allowed in the trial version."
        Exit Sub
    End If
    If Toggle_remarks = True Then
        Frame_remarks.Visible = True
    ElseIf Toggle_remarks = False Then
        Frame_remarks.Visible = False
    End If
Call Fluid_type
End Sub
Private Sub Check_T_OUT_Click()
'    If Check_T_OUT = 1 And Thermal_bal_tubes = True Then
''        Check_T_Flow = Unchecked
'    Else
''        Check_T_Flow = Unchecked
'        Check_T_OUT = Unchecked
'    End If
Call Fluid_type
End Sub
Private Sub Comm_property_Click()
On Error Resume Next
If Comm_property = 1 Then
 Dim Rs2 As Recordset
    Data2.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
    Data2.RecordSource = "Select * From [Query_Unit_sort]"
    Data2.Refresh
    Set Rs2 = Data2.Recordset
    If Rs2.RecordCount > 0 Then
        Do Until Rs2.EOF
            Fluid_1 = Combo_S_FLUID
            FLUID_2 = Data2.Recordset.SHELL_FLUID
            If Fluid_1 = FLUID_2 Then
                THC = Data2.Recordset.Check_S_TC
                SPH_1 = Data2.Recordset.Check_S_SPH
                DENS = Data2.Recordset.Check_S_DENS
                visc = Data2.Recordset.Check_S_VISC
                LAT = Data2.Recordset.Check_LATENT
                TEC = Data2.Recordset.Check_CT
                
                If THC = True Then
                    SHELL_OUT(1) = Data2.Recordset.SHELL_T_COND
                    SHELL_OUT(1).ForeColor = &HFF&
                    SHELL_OUT(1).BackColor = &HE0E0E0
                End If
                If SPH_1 = True Then
                    SHELL_OUT(3) = Data2.Recordset.SHELL_SPH
                    SHELL_OUT(3).ForeColor = &HFF&
                    SHELL_OUT(3).BackColor = &HE0E0E0
                End If
                If DENS = True Then
                    SHELL_OUT(4) = Data2.Recordset.SHELL_DENS
                    SHELL_OUT(4).ForeColor = &HFF&
                    SHELL_OUT(4).BackColor = &HE0E0E0
                End If
                If visc = True Then
                    visc = Data2.Recordset.SHELL_VISC
                    SHELL_OUT(5) = Format(visc, "0.000")
                    SHELL_OUT(5).ForeColor = &HFF&
                    SHELL_OUT(5).BackColor = &HE0E0E0
                End If
                If LAT = True Then
                    SHELL_OUT(10) = Data2.Recordset.SHELL_LATENT
                    SHELL_OUT(10).ForeColor = &HFF&
                    SHELL_OUT(10).BackColor = &HE0E0E0
                End If
                If TEC = True Then
                    SHELL_OUT(11) = Data2.Recordset.Temp_COND
                    SHELL_OUT(11).ForeColor = &HFF&
                    SHELL_OUT(11).BackColor = &HE0E0E0
                End If
            End If
            Rs2.MoveNext
        Loop
    Else
       MsgBox "Fluid not found"
    End If
End If
End Sub
Private Sub Ammonia()
MW = 17.001
TK = CDbl(XD20) + 273.16            'T, °K
Tc = 405.5                          'Critical temp., °K
Tr = TK / Tc                        'Reduced temperature
Tb = 239.8                          'Boiling point at 1atm °K

'Ammonia Condensation pressure, bar(a)
    ACP_1 = -11.2065110833853
    ACP_2 = 8.16294269621182E-02
    ACP_3 = 1.80542098712002E-04
    ACP_4 = 1.47860782222754E-07
    XD21 = 10 ^ (ACP_1 + ACP_2 * TK - ACP_3 * TK ^ 2 + ACP_4 * TK ^ 3)
    If Check_CP = 0 Then
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HFF0000
        S_press_KP.BackColor = &HE0E0E0
    ElseIf Check_CP = 1 Then
        XD21 = Spin_S_PRESS / 10000
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HC0&
        S_press_KP.BackColor = &HE0E0E0
    End If
'Ammonia vapor pressure, bar(A) (log10(P) = A - (B / (T + C)))
    If TK > 164 And TK < 239.6 Then
        A1 = 3.18757
        b1 = 506.713
        C1 = -80.78
    ElseIf TK > 239.6 And TK < 371.5 Then
        A1 = 4.86886
        b1 = 1113.928
        C1 = -10.409
    End If
    P1 = 10 ^ (A1 - (b1 / (TK + C1)))
'Ammonia IN enthalpy, Kcal/Kg
    A1 = -4.90421676857996E-04
    a2 = 4.90421676857996E-04
    a3 = 2.6065188815352
    a4 = -2.24502251374982E-06
    TK_IN = 273.16 + XD19
    XD22 = (A1 + Sqr((a2 ^ 2 - 4 * (a3 - Log(TK_IN) / LN) * a4))) / (2 * (a4))
    XD22 = (-4.90421676857996E-04 + Sqr((4.90421676857996E-04 ^ 2 - 4 * (2.6065188815352 - Log(TK_IN) / LN) * -2.24502251374982E-06))) / (2 * (-2.24502251374982E-06))
'Ammonia OUT enthalpy, Kcal/Kg
    A1 = -0.315317436306974
    a2 = 21.3312078166881
    a3 = 1.31615331719314E-03
    a4 = Log(TK) / LN
    XD23 = (-A1 + Sqr((A1 ^ 2 - 4 * (a2 - a4) * a3))) / (2 * (a3))
'Ammonia latent heat, Kcal/Kg
    XD24 = XD23 - XD22
    If Check_LATENT = 0 Then
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HFF0000
        SHELL_OUT(10).BackColor = &HE0E0E0
    ElseIf Check_LATENT = 1 Then
        XD24 = HScroll_LATENT / 4.1868
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HC0&
        SHELL_OUT(10).BackColor = &HE0E0E0
    End If
'Heat capacity,J(mol-°K) - Cp =CPVAP_A + CPVAP_B * K + CPVAP_C * K^2 + CPVAP_D * K^3
    H_cap = 27.32 + 0.02383 * TK + 0.000017075 * TK ^ 2 - 0.00000001185 * TK * 3
    'Heat capacity,J(kg-°K)
    SPH_S = H_cap * 1000 / MW * 0.0002388458966
    If Check_S_SPH = 0 Then
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HFF0000
        SHELL_OUT(3).BackColor = &HE0E0E0
    ElseIf Check_S_SPH = 1 Then
        SPH_S = HScroll_SHELL_SPH / 1000
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HC0&
        SHELL_OUT(3).BackColor = &HE0E0E0
    End If

'Ammonia density at condensing film temperature,Kg/m^3
    XD75 = (1.67201252301867 - 8.67573506044482E-03 * TK + 2.71463576525648E-05 * TK ^ 2 - 3.37884138338416E-08 * TK ^ 3) * 1000
    If Check_S_DENS = 0 Then
        SHELL_OUT(4).Text = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HFF0000
        SHELL_OUT(4).BackColor = &HE0E0E0
    ElseIf Check_S_DENS = 1 Then
        XD75 = HScroll_SHELL_DENS / 10
        SHELL_OUT(4).Text = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HC0&
        SHELL_OUT(4).BackColor = &HE0E0E0
    End If
'Ammonia density at condensing film temperature,lb/ft^3
    XD76 = XD75 * 2.20462 * (0.3048 ^ 3)

'Ammonia viscosity at condensing film temperature,cp
    XD77 = (22.8772727564961 - 0.154773809799488 * TK + 3.68290044153551E-04 * TK ^ 2 + -3.03030303928442E-07 * TK ^ 3) / 10
    If Check_S_VISC = 0 Then
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HFF0000
        SHELL_OUT(5).BackColor = &HE0E0E0
    ElseIf Check_S_VISC = 1 Then
        XD77 = HScroll_SHELL_VISC.Value / 1000
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HC0&
        SHELL_OUT(5).BackColor = &HE0E0E0
    End If
'Ammonia thermal conductivity at condensing film temperature,Btu/hftºF
    XD78 = (1.67572727536397 - 7.29141416629545E-03 * TK + 1.62878788658395E-05 * TK ^ 2 - 1.76767677578808E-08 * TK ^ 3) * 0.5778
    'Ammonia thermal conductivity at condensing film temperature Kcal/h m ºC
    XD79 = XD78 * 1.488
    If Check_S_TC = 0 Then
        SHELL_OUT(1) = Format(XD79, "0.000")
        SHELL_OUT(1).ForeColor = &HFF0000
        SHELL_OUT(1).BackColor = &HE0E0E0
    ElseIf Check_S_TC = 1 Then
        XD79 = HScroll_SHELL_TC / 1000
        XD78 = XD79 / 1.488
        SHELL_OUT(1) = Format(XD79, "0.000")
        SHELL_OUT(1).ForeColor = &HC0&
        SHELL_OUT(1).BackColor = &HE0E0E0
    End If
FLUID_VL = 1
End Sub
Private Sub BENZENE()
MW = 78.1118
TK = CDbl(XD20) + 273.16            'T, °K
Tc = 562.2                          'Critical temp., °K
Tr = TK / Tc                        'Reduced temperature
Tb = 353.2                          'Boiling point at 1atm °K

'Benzene Condensation pressure, bara
    XD21 = 10 ^ (-17.7266795554994 + 0.107558360498623 * TK - 2.20426777069489E-04 * TK ^ 2 + 1.64343153650998E-07 * TK ^ 3)
'Benzene Condensation pressure, bara
'Eq.1  -  lnPvp =(1-x)ì-1*((VPA*x+VPB*x^1.5+VPC*x^3+VPD*X^6)  -  X=1-K/Tc
'Eq.2  -  lnPvp =VPA - VPB/T+VPC*lnT+VPD*Pvp/T^2
'Eq.3  -  lnPvp = VPA - VPB / (T + VPC)
    If Check_CP = 0 Then
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HFF0000
        S_press_KP.BackColor = &HE0E0E0
    ElseIf Check_CP = 1 Then
        XD21 = Spin_S_PRESS / 10000
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HC0&
        S_press_KP.BackColor = &HE0E0E0
    End If
'Benzene vapor pressure, bar(a)   (log10(P) = A - (B / (T + C)))
    If TK > 287.7 And TK < 354.07 Then
        A1 = 4.01814
        b1 = 1203.835
        C1 = -53.226
    ElseIf TK > 354.06 And TK < 373.5 Then
        A1 = 4.72583
        b1 = 1660.652
        C1 = -1.461
    ElseIf TK > 373.4 And TK < 554.8 Then
        A1 = 4.60362
        b1 = 1701.073
        C1 = 20.806
    End If
    P1 = 10 ^ (A1 - (b1 / (TK + C1)))
'Benzene IN enthalpy, Kcal/Kg
    XD22 = (-5.84000817614495E-03 + Sqr((5.84000817614495E-03 ^ 2 - 4 * (2.04979316731536 - Log(273.16 + XD19) / LN) * -1.33496728809234E-05))) / (2 * (-1.33496728809234E-05))
'Benzene OUT enthalpy, Kcal/Kg
    XD23 = (-2.22078281567555E-02 + Sqr((2.22078281567555E-02 ^ 2 - 4 * (-0.26019046655962 - Log(TK) / LN) * -4.20919412093621E-05))) / (2 * (-4.20919412093621E-05))
'Benzene latent heat Kcal/Kg
    XD24 = XD23 - XD22
'Enthalpy, kJ/mole (EvapH = A exp(-aTr) (1 - Tr)^ß)
    A1 = 47.41                                      'kJ/mole
    a2 = 0.1231
    b1 = 0.3602
    TR1 = TK / 562.1                   'Reduced temperature T/Tc
    ET = A1 * Exp(-a2 * TR1) * (1 - TR1) ^ b1
'Enthalpy, kJ/kg (MW=78.1118)
    ET1 = ET * 1000 / MW
'Enthalpy, kcal/kg
    ET2 = ET1 / 4.18
    If Check_LATENT = 0 Then
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HFF0000
        SHELL_OUT(10).BackColor = &HE0E0E0
    ElseIf Check_LATENT = 1 Then
        XD24 = HScroll_LATENT / 4.1868
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HC0&
        SHELL_OUT(10).BackColor = &HE0E0E0
    End If

'Heat capacity,J(mol-°K) - Cp =CPVAP_A + CPVAP_B * K + CPVAP_C * K^2 + CPVAP_D * K^3
    H_cap = -33.92 + 0.4739 * TK - 0.0003017 * TK ^ 2 + 0.0000000713 * TK * 3
'Heat capacity,J(kg-°K)
    SPH_S = H_cap * 1000 / MW * 0.0002388458966
    If Check_S_SPH = 0 Then
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HFF0000
        SHELL_OUT(3).BackColor = &HE0E0E0
    ElseIf Check_S_SPH = 1 Then
        SPH_S = HScroll_SHELL_SPH / 1000
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HC0&
        SHELL_OUT(3).BackColor = &HE0E0E0
    End If

'Benzene density at condensing film temperature,Kg/m^3
    XD75 = (1.376997665 - 0.002829185 * TK + 5.6388778534534E-06 * TK ^ 2 - 6.07429148488935E-09 * TK ^ 3) * 1000
    If Check_S_DENS = 0 Then
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HFF0000
        SHELL_OUT(4).BackColor = &HE0E0E0
    ElseIf Check_S_DENS = 1 Then
        XD75 = HScroll_SHELL_DENS / 10
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HC0&
        SHELL_OUT(4).BackColor = &HE0E0E0
    End If
'Benzene density at condensing film temperature,lb/ft^3
    XD76 = XD75 * 2.20462 * (0.3048 ^ 3)

'Benzene viscosity at condensing film temperature,cp
    XD77 = (174.50837619225 - 1.30737638396497 * TK + 3.36425241742188E-03 * TK ^ 2 - 2.93447293753068E-06 * TK ^ 3) / 10
    If Check_S_VISC = 0 Then
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HFF0000
        SHELL_OUT(5).BackColor = &HE0E0E0
    ElseIf Check_S_VISC = 1 Then
        XD77 = HScroll_SHELL_VISC.Value / 1000
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HC0&
        SHELL_OUT(5).BackColor = &HE0E0E0
    End If

'Benzene thermal conductivity at condensing film temperature, Btu/hftºF
    XD78 = (0.234000000232994 - 3.00000002048061E-04 * TK + 5.96459460478361E-15 * TK ^ 2 - 5.75588815841478E-18 * TK ^ 3) * 0.5778
'Benzene thermal conductivity at condensing film temperature Kcal/h m ºC
    XD79 = XD78 * 1.488
'Benzene thermal conductivity at condensing film temperature W/mºK
    Landa = 0.1776 + 0.000004773 * TK - 0.0000003781 * TK ^ 2
'Benzene thermal conductivity at condensing film temperature Kcal/h m ºC
    LANDA_1 = Landa * 0.8604
    If Check_S_TC = 0 Then
        SHELL_OUT(1) = Format(XD79, "0.000")
        SHELL_OUT(1).ForeColor = &HFF0000
        SHELL_OUT(1).BackColor = &HE0E0E0
    ElseIf Check_S_TC = 1 Then
        XD79 = HScroll_SHELL_TC / 1000
        XD78 = XD79 / 1.488
        SHELL_OUT(1) = Format(XD79, "0.000")
        SHELL_OUT(1).ForeColor = &HC0&
        SHELL_OUT(1).BackColor = &HE0E0E0
    End If
FLUID_VL = 1
End Sub
Private Sub Butane()
MW = 58.1222
TK = CDbl(XD20) + 273.16            'T, °K
Tc = 425.2                          'Critical temp., °K
Tr = TK / Tc                        'Reduced temperature
Tb = 272.7                          'Boiling point at 1atm °K

'Butane vapor pressure, bar(a)   (log10(P) = A - (B / (T + C)))
    If TK > 272.66 And TK < 425 Then
        A1 = 4.35576
        b1 = 1175.581
        C1 = -2.071
    End If
    XD21 = 10 ^ (A1 - (b1 / (TK + C1)))
    If Check_CP = 0 Then
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HFF0000
        S_press_KP.BackColor = &HE0E0E0
    ElseIf Check_CP = 1 Then
        XD21 = Spin_S_PRESS / 10000
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HC0&
        S_press_KP.BackColor = &HE0E0E0
    End If

'Butane enthalpy, kJ/mol (EvapH = (J/mol*K))
    ET = 0.0823 * TK
    'Butane enthalpy, kJ/kg (MW=58.1222)
    ET1 = ET * 1000 / MW
    'Butane enthalpy, kcal/kg
    XD24 = ET1 / 4.18
    If Check_LATENT = 0 Then
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HFF0000
        SHELL_OUT(10).BackColor = &HE0E0E0
    ElseIf Check_LATENT = 1 Then
        XD24 = HScroll_LATENT / 4.1868
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HC0&
        SHELL_OUT(10).BackColor = &HE0E0E0
    End If

'Heat capacity,J(mol-°K) - Cp =CPVAP_A + CPVAP_B * K + CPVAP_C * K^2 + CPVAP_D * K^3
    H_cap = 9.487 + 0.3313 * TK - -0.00011 * TK ^ 2 + -0.000000002822 * TK * 3
    'Heat capacity,J(kg-°K)
    SPH_S = H_cap * 1000 / MW * 0.0002388458966
    If Check_S_SPH = 0 Then
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HFF0000
        SHELL_OUT(3).BackColor = &HE0E0E0
    ElseIf Check_S_SPH = 1 Then
        SPH_S = HScroll_SHELL_SPH / 1000
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HC0&
        SHELL_OUT(3).BackColor = &HE0E0E0
    End If

'Density at condensing film temperature,Kg/m^3
    a = -1.52816: b = 1.43907: c = -0.81446: d = 0.190454
    e = -0.296123: f = 0.386914: g = -0.0427258: h = -0.04480645
    Tc = 425.18
    Tr = TK / Tc
    Wsrk = 0.1825
    Vc = 0.2544
    Vro = 1 + a * (1 - Tr) ^ (1 / 3) + b * (1 - Tr) ^ (2 / 3) + c * (1 - Tr) + d * (1 - Tr) ^ (4 / 3) '0,25<Tr<0,95
    Vrd = (e + f * Tr + g * Tr ^ 2 + h * Tr ^ 3) / (Tr - 1.00001) '0,25<Tr<1,0
    Vs = Vro * (1 - Wsrk * Vrd) * Vc * 1000
    XD75 = 1 / (Vs * 1000 / MW) * 1000000
    If Check_S_DENS = 0 Then
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HFF0000
        SHELL_OUT(4).BackColor = &HE0E0E0
    ElseIf Check_S_DENS = 1 Then
        XD75 = HScroll_SHELL_DENS / 10
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HC0&
        SHELL_OUT(4).BackColor = &HE0E0E0
    End If
    'Density at condensing film temperature,lb/ft^3
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)

'Viscosity at condensing film temperature,cp
    'eq. 2: ln n = A + B / T
    'eq. 3: ln n = A + B / T + C *  T + D * T^2
    'T°C
    'T°K
    a = -3.82: b = 612.1
    XD77 = Exp(a + b / TK)
    If Check_S_VISC = 0 Then
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HFF0000
        SHELL_OUT(5).BackColor = &HE0E0E0
    ElseIf Check_S_VISC = 1 Then
        XD77 = HScroll_SHELL_VISC.Value / 1000
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HC0&
        SHELL_OUT(5).BackColor = &HE0E0E0
    End If
'Thermal conductivity at condensing film temperature W/mºK
    a = 0.2554
    b = -0.0003984
    c = -0.0000001135
    Landa = a + b * TK + c * TK ^ 2
    'Thermal conductivity at condensing film temperature Kcal/mºK
    XD79 = Landa * 0.8604
    'Thermal conductivity, Btu/hftºF
    XD78 = XD79 / 1.488
    If Check_S_TC = 0 Then
        SHELL_OUT(1) = Format(XD79, "0.000")
        SHELL_OUT(1).ForeColor = &HFF0000
        SHELL_OUT(1).BackColor = &HE0E0E0
    ElseIf Check_S_TC = 1 Then
        XD79 = HScroll_SHELL_TC / 1000
        'Thermal conductivity, Kcal/hmºC
        XD78 = XD79 / 1.488
        SHELL_OUT(1) = Format(XD79, "0.000")
        SHELL_OUT(1).ForeColor = &HC0&
        SHELL_OUT(1).BackColor = &HE0E0E0
    End If
FLUID_VL = 1
End Sub
Private Sub Butanol_1()
MW = 74.1216
TK = CDbl(XD20) + 273.16                  'T, °K
Tc = 563.1                          'Critical temp., °K
Tr = TK / Tc                        'Reduced temperature
Tb = 390.9                          'Boiling point at 1atm °K

'1-Butanol vapor pressure, bar(a)   (log10(P) = A - (B / (T + C)))
    If TK > 295.8 And TK < 391.1 Then
        A1 = 4.54607
        b1 = 1351.555
        C1 = -93.34
    ElseIf TK > 391 And TK < 479.1 Then
        A1 = 4.39031
        b1 = 1254.502
        C1 = -105.246
    ElseIf TK > 479 And TK < 562.98 Then
        A1 = 4.42921
        b1 = 1305.001
        C1 = -94.676
    End If
    XD21 = 10 ^ (A1 - (b1 / (TK + C1)))
    If Check_CP = 0 Then
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HFF0000
        S_press_KP.BackColor = &HE0E0E0
    ElseIf Check_CP = 1 Then
        XD21 = Spin_S_PRESS / 10000
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HC0&
        S_press_KP.BackColor = &HE0E0E0
    End If

'1-Butanol enthalpy, kJ/mol (EvapH = A exp(-aTr) (1 - Tr)^ß)(K = 298 - 410)
    A1 = 62.53                                  'kJ/mole
    a2 = -0.6584
    b1 = 0.696
    ET = A1 * Exp(-a2 * Tr) * (1 - Tr) ^ b1
'1-Butanol enthalpy, kJ/kg (MW=74.1216)
    ET1 = ET * 1000 / MW
'1-Butanol enthalpy, kcal/kg
    XD24 = ET1 / 4.18
    If Check_LATENT = 0 Then
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HFF0000
        SHELL_OUT(10).BackColor = &HE0E0E0
    ElseIf Check_LATENT = 1 Then
        XD24 = HScroll_LATENT / 4.1868
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HC0&
        SHELL_OUT(10).BackColor = &HE0E0E0
    End If

'Heat capacity,J(mol-°K) - Cp =CPVAP_A + CPVAP_B * K + CPVAP_C * K^2 + CPVAP_D * K^3
    H_cap = 3.266 + 0.418 * TK - 0.0002242 * TK ^ 2 + 0.00000004685 * TK * 3
'Heat capacity,J(kg-°K)
    SPH_S = H_cap * 1000 / MW * 0.0002388458966
    If Check_S_SPH = 0 Then
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HFF0000
        SHELL_OUT(3).BackColor = &HE0E0E0
    ElseIf Check_S_SPH = 1 Then
        SPH_S = HScroll_SHELL_SPH / 1000
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HC0&
        SHELL_OUT(3).BackColor = &HE0E0E0
    End If
'Density at condensing film temperature,Kg/m^3
    a = -1.52816: b = 1.43907: c = -0.81446: d = 0.190454
    e = -0.296123: f = 0.386914: g = -0.0427258: h = -0.04480645
    Tc = 562.93
    Tr = TK / Tc
    Wsrk = 0.5928
    Vc = 0.2841
    Vro = 1 + a * (1 - Tr) ^ (1 / 3) + b * (1 - Tr) ^ (2 / 3) + c * (1 - Tr) + d * (1 - Tr) ^ (4 / 3) '0,25<Tr<0,95
    Vrd = (e + f * Tr + g * Tr ^ 2 + h * Tr ^ 3) / (Tr - 1.00001) '0,25<Tr<1,0
    Vs = Vro * (1 - Wsrk * Vrd) * Vc * 1000
    XD75 = 1 / (Vs * 1000 / MW) * 1000000
    If Check_S_DENS = 0 Then
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HFF0000
        SHELL_OUT(4).BackColor = &HE0E0E0
    ElseIf Check_S_DENS = 1 Then
        XD75 = HScroll_SHELL_DENS / 10
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HC0&
        SHELL_OUT(4).BackColor = &HE0E0E0
    End If
    'Density at condensing film temperature,lb/ft^3
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)

'Viscosity at condensing film temperature,cp
    'eq. 2: ln n = A + B / T
    'eq. 3: ln n = A + B / T + C *  T + D * T^2
    'T°C
    'T°K
    a = -9.722: b = 2602: c = 0.00953: d = -0.000009966
    XD77 = Exp(a + b / TK + c * TK + d * TK ^ 2)
    If Check_S_VISC = 0 Then
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HFF0000
        SHELL_OUT(5).BackColor = &HE0E0E0
    ElseIf Check_S_VISC = 1 Then
        XD77 = HScroll_SHELL_VISC.Value / 1000
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HC0&
        SHELL_OUT(5).BackColor = &HE0E0E0
    End If

'Thermal conductivity at condensing film temperature W/mºK
    Landa = 0.2288 - 0.0002697 * TK + 0.00000001323 * TK ^ 2
'Thermal conductivity at condensing film temperature Kcal/mºK
    XD79 = Landa * 0.8604
'Thermal conductivity, Btu/hftºF
    XD78 = XD79 / 1.488
    If Check_S_TC = 0 Then
        SHELL_OUT(1) = Format(XD79, "0.000")
        SHELL_OUT(1).ForeColor = &HFF0000
        SHELL_OUT(1).BackColor = &HE0E0E0
    ElseIf Check_S_TC = 1 Then
        XD79 = HScroll_SHELL_TC / 1000
        XD78 = XD79 / 1.488
        SHELL_OUT(1) = Format(XD79, "0.000")
        SHELL_OUT(1).ForeColor = &HC0&
        SHELL_OUT(1).BackColor = &HE0E0E0
    End If
FLUID_VL = 1
End Sub
Private Sub Cyclohexane()
MW = 84.1595
TK = CDbl(XD20) + 273.16                  'T, °K
Tc = 553.5                          'Critical temp., °K
Tr = TK / Tc                        'Reduced temperature
Tb = 353.8                          'Boiling point at 1atm °K

'Cyclohexane vapor pressure, bar(a)   (log10(P) = A - (B / (T + C)))
    If TK > 293.06 And TK < 354.73 Then
        A1 = 3.96988
        b1 = 1203.526
        C1 = -50.287
    ElseIf TK > 354.7 And TK < 523 Then
        A1 = 4.13983
        b1 = 1316.554
        C1 = -35.581
    End If
    XD21 = 10 ^ (A1 - (b1 / (TK + C1)))
    If Check_CP = 0 Then
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HFF0000
        S_press_KP.BackColor = &HE0E0E0
    ElseIf Check_CP = 1 Then
        XD21 = Spin_S_PRESS / 10000
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HC0&
        S_press_KP.BackColor = &HE0E0E0
    End If

'Cyclohexane enthalpy, kJ/mol (EvapH = A exp(-aTr) (1 - Tr)^ß)
    A1 = 43.32                                    'kJ/mole
    a2 = -0.1437
    b1 = 0.4512
    ET = A1 * Exp(-a2 * Tr) * (1 - Tr) ^ b1
'Cyclohexane enthalpy, kJ/kg (MW=84.1595)
    ET1 = ET * 1000 / MW
'Cyclohexane enthalpy, kcal/kg
    XD24 = ET1 / 4.18
    If Check_LATENT = 0 Then
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HFF0000
        SHELL_OUT(10).BackColor = &HE0E0E0
    ElseIf Check_LATENT = 1 Then
        XD24 = HScroll_LATENT / 4.1868
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HC0&
        SHELL_OUT(10).BackColor = &HE0E0E0
    End If
    
'Heat capacity,J(mol-°K) - Cp =CPVAP_A + CPVAP_B * K + CPVAP_C * K^2 + CPVAP_D * K^3
    H_cap = -68.65 + 0.7252 * TK - 0.0005414 * TK ^ 2 + 0.0000001644 * TK * 3
'Heat capacity,J(kg-°K)
    SPH_S = H_cap * 1000 / MW * 0.0002388458966
    If Check_S_SPH = 0 Then
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HFF0000
        SHELL_OUT(3).BackColor = &HE0E0E0
    ElseIf Check_S_SPH = 1 Then
        SPH_S = HScroll_SHELL_SPH / 1000
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HC0&
        SHELL_OUT(3).BackColor = &HE0E0E0
    End If

'Density at condensing film temperature,Kg/m^3
    a = -1.52816: b = 1.43907: c = -0.81446: d = 0.190454
    e = -0.296123: f = 0.386914: g = -0.0427258: h = -0.04480645
    Tc = 553.54
    Tr = TK / Tc
    Wsrk = 0.2128
    Vc = 0.309
    Vro = 1 + a * (1 - Tr) ^ (1 / 3) + b * (1 - Tr) ^ (2 / 3) + c * (1 - Tr) + d * (1 - Tr) ^ (4 / 3) '0,25<Tr<0,95
    Vrd = (e + f * Tr + g * Tr ^ 2 + h * Tr ^ 3) / (Tr - 1.00001) '0,25<Tr<1,0
    Vs = Vro * (1 - Wsrk * Vrd) * Vc * 1000
    XD75 = 1 / (Vs * 1000 / MW) * 1000000
    If Check_S_DENS = 0 Then
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HFF0000
        SHELL_OUT(4).BackColor = &HE0E0E0
    ElseIf Check_S_DENS = 1 Then
        XD75 = HScroll_SHELL_DENS / 10
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HC0&
        SHELL_OUT(4).BackColor = &HE0E0E0
    End If
    'Density at condensing film temperature,lb/ft^3
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)

'Viscosity at condensing film temperature,cp
    'eq. 2: ln n = A + B / T
    'eq. 3: ln n = A + B / T + C *  T + D * T^2
    a = -4.398: b = 1380: c = -0.00155: d = 0.000001157
    XD77 = Exp(a + b / TK + c * TK + d * TK ^ 2)
    If Check_S_VISC = 0 Then
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HFF0000
        SHELL_OUT(5).BackColor = &HE0E0E0
    ElseIf Check_S_VISC = 1 Then
        XD77 = HScroll_SHELL_VISC.Value / 1000
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HC0&
        SHELL_OUT(5).BackColor = &HE0E0E0
    End If

'Thermal conductivity at condensing film temperature W/ m ºK
    Landa = 0.1626 - 0.00009513 * TK - 0.0000001382 * TK ^ 2
'Thermal conductivity at condensing film temperature Kcal/mºK
    XD79 = Landa * 0.8604
'Thermal conductivity, Btu/hftºF
    XD78 = XD79 / 1.488
    If Check_S_TC = 0 Then
        SHELL_OUT(1) = Format(XD79, "0.000")
        SHELL_OUT(1).ForeColor = &HFF0000
        SHELL_OUT(1).BackColor = &HE0E0E0
    ElseIf Check_S_TC = 1 Then
        XD79 = HScroll_SHELL_TC / 1000
        XD78 = XD79 / 1.488
        SHELL_OUT(1) = Format(XD79, "0.000")
        SHELL_OUT(1).ForeColor = &HC0&
        SHELL_OUT(1).BackColor = &HE0E0E0
    End If
FLUID_VL = 1
End Sub
Private Sub Ethanol()
MW = 46.0684
TK = CDbl(XD20) + 273.16                  'T, °K
Tc = 513.9                         'Critical temp., °K
Tr = TK / Tc                        'Reduced temperature
Tb = 351.4                          'Boiling point at 1atm °K

'Ethanol vapor pressure, bar(a)   (log10(P) = A - (B / (T + C)))
    If TK > 292.77 And TK < 366.63 Then
        A1 = 5.24677
        b1 = 1598.673
        C1 = -46.424
    ElseIf TK > 366.62 And TK < 513.91 Then
        A1 = 4.92531
        b1 = 1432.526
        C1 = -61.819
    End If
    XD21 = 10 ^ (A1 - (b1 / (TK + C1)))
    If Check_CP = 0 Then
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HFF0000
        S_press_KP.BackColor = &HE0E0E0
    ElseIf Check_CP = 1 Then
        XD21 = Spin_S_PRESS / 10000
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HC0&
        S_press_KP.BackColor = &HE0E0E0
    End If

'Ethanol enthalpy, kJ/mol (EvapH = A exp(-aTr) (1 - Tr)^ß)
    A1 = 50.43                                    'kJ/mole
    a2 = -0.4475
    b1 = 0.4989
    ET = A1 * Exp(-a2 * Tr) * (1 - Tr) ^ b1
'Ethanol enthalpy, kJ/kg (MW=46.0684)
    ET1 = ET * 1000 / MW
'Ethanol enthalpy, kcal/kg
    XD24 = ET1 / 4.18
    If Check_LATENT = 0 Then
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HFF0000
        SHELL_OUT(10).BackColor = &HE0E0E0
    ElseIf Check_LATENT = 1 Then
        XD24 = HScroll_LATENT / 4.1868
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HC0&
        SHELL_OUT(10).BackColor = &HE0E0E0
    End If

'Heat capacity,J(mol-°K) - Cp =CPVAP_A + CPVAP_B * K + CPVAP_C * K^2 + CPVAP_D * K^3
    H_cap = 9.014 + 0.2141 * TK - 0.0000839 * TK ^ 2 + 0.000000001373 * TK * 3
'Heat capacity,J(kg-°K)
    SPH_S = H_cap * 1000 / MW * 0.0002388458966
    If Check_S_SPH = 0 Then
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HFF0000
        SHELL_OUT(3).BackColor = &HE0E0E0
    ElseIf Check_S_SPH = 1 Then
        SPH_S = HScroll_SHELL_SPH / 1000
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HC0&
        SHELL_OUT(3).BackColor = &HE0E0E0
    End If

'Density at condensing film temperature,Kg/m^3
    a = -1.52816: b = 1.43907: c = -0.81446: d = 0.190454
    e = -0.296123: f = 0.386914: g = -0.0427258: h = -0.04480645
    Tc = 516.16
    Tr = TK / Tc
    Wsrk = 0.6378
    Vc = 0.1752
    Vro = 1 + a * (1 - Tr) ^ (1 / 3) + b * (1 - Tr) ^ (2 / 3) + c * (1 - Tr) + d * (1 - Tr) ^ (4 / 3) '0,25<Tr<0,95
    Vrd = (e + f * Tr + g * Tr ^ 2 + h * Tr ^ 3) / (Tr - 1.00001) '0,25<Tr<1,0
    Vs = Vro * (1 - Wsrk * Vrd) * Vc * 1000
    XD75 = 1 / (Vs * 1000 / MW) * 1000000
    If Check_S_DENS = 0 Then
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HFF0000
        SHELL_OUT(4).BackColor = &HE0E0E0
    ElseIf Check_S_DENS = 1 Then
        XD75 = HScroll_SHELL_DENS / 10
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HC0&
        SHELL_OUT(4).BackColor = &HE0E0E0
    End If
    'Density at condensing film temperature,lb/ft^3
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)

'Viscosity at condensing film temperature,cp
    'eq. 2: ln n = A + B / T
    'eq. 3: ln n = A + B / T + C *  T + D * T^2
    'T°C
    'T°K
    a = -6.21: b = 1614: c = 0.00618: d = -0.00001132
    XD77 = Exp(a + b / TK + c * TK + d * TK ^ 2)
    If Check_S_VISC = 0 Then
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HFF0000
        SHELL_OUT(5).BackColor = &HE0E0E0
    ElseIf Check_S_VISC = 1 Then
        XD77 = HScroll_SHELL_VISC.Value / 1000
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HC0&
        SHELL_OUT(5).BackColor = &HE0E0E0
    End If

'Thermal conductivity at condensing film temperature W/ m ºK
    Landa = 0.2629 - 0.00038475 * TK + 0.0000002211 * TK ^ 2
'Thermal conductivity at condensing film temperature Kcal/mºK
    XD79 = Landa * 0.8604
'Thermal conductivity, Btu/hftºF
    XD78 = XD79 / 1.488
    If Check_S_TC = 0 Then
        SHELL_OUT(1) = Format(XD79, "0.000")
        SHELL_OUT(1).ForeColor = &HFF0000
        SHELL_OUT(1).BackColor = &HE0E0E0
    ElseIf Check_S_TC = 1 Then
        XD79 = HScroll_SHELL_TC / 1000
        XD78 = XD79 / 1.488
        SHELL_OUT(1) = Format(XD79, "0.000")
        SHELL_OUT(1).ForeColor = &HC0&
        SHELL_OUT(1).BackColor = &HE0E0E0
    End If
FLUID_VL = 1
End Sub
Private Sub Heptane()
MW = 100.2019
TK = CDbl(XD20) + 273.16            'T, °K
Tc = 540.26                          'Critical temp., °K
Tr = TK / Tc                        'Reduced temperature
Tb = 371.6                          'Boiling point at 1atm °K

'Heptane vapor pressure, bar(a)   (log10(P) = A - (B / (T + C)))
    If TK > 185.29 And TK < 295.6 Then
        A1 = 4.81803
        b1 = 1635.409
        C1 = -27.338
    ElseIf TK > 295.5 And TK < 372.43 Then
        A1 = 4.02832
        b1 = 1268.636
        C1 = -56.199
    End If
    XD21 = 10 ^ (A1 - (b1 / (TK + C1)))
    If Check_CP = 0 Then
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HFF0000
        S_press_KP.BackColor = &HE0E0E0
    ElseIf Check_CP = 1 Then
        XD21 = Spin_S_PRESS / 10000
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HC0&
        S_press_KP.BackColor = &HE0E0E0
    End If

'Heptane enthalpy, kJ/mole   (EvapH = A exp(-ßTr) (1 - Tr)^ß)
    A1 = 53.66                                  'kJ/mole
    b1 = 0.2831
    ET = A1 * Exp(-b1 * Tr) * (1 - Tr) ^ b1
    'Heptane enthalpy, kJ/kg (MW = 100.2019)
    ET1 = ET * 1000 / MW
    'Heptane enthalpy, kcal/kg
    XD24 = ET1 / 4.18
    If Check_LATENT = 0 Then
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HFF0000
        SHELL_OUT(10).BackColor = &HE0E0E0
    ElseIf Check_LATENT = 1 Then
        XD24 = HScroll_LATENT / 4.1868
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HC0&
        SHELL_OUT(10).BackColor = &HE0E0E0
    End If

'Heat capacity,J(mol-°K) - Cp =CPVAP_A + CPVAP_B * K + CPVAP_C * K^2 + CPVAP_D * K^3
    H_cap = -5.146 + 0.6762 * TK - 0.0003661 * TK ^ 2 + 0.00000007658 * TK * 3
'Heat capacity,J(kg-°K)
    SPH_S = H_cap * 1000 / MW * 0.0002388458966
    If Check_S_SPH = 0 Then
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HFF0000
        SHELL_OUT(3).BackColor = &HE0E0E0
    ElseIf Check_S_SPH = 1 Then
        SPH_S = HScroll_SHELL_SPH / 1000
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HC0&
        SHELL_OUT(3).BackColor = &HE0E0E0
    End If

'Density at condensing film temperature,Kg/m^3
    a = -1.52816: b = 1.43907: c = -0.81446: d = 0.190454
    e = -0.296123: f = 0.386914: g = -0.0427258: h = -0.04480645
    Tc = 540.26
    Tr = TK / Tc
    Wsrk = 0.3507
    Vc = 0.4304
    Vro = 1 + a * (1 - Tr) ^ (1 / 3) + b * (1 - Tr) ^ (2 / 3) + c * (1 - Tr) + d * (1 - Tr) ^ (4 / 3) '0,25<Tr<0,95
    Vrd = (e + f * Tr + g * Tr ^ 2 + h * Tr ^ 3) / (Tr - 1.00001) '0,25<Tr<1,0
    Vs = Vro * (1 - Wsrk * Vrd) * Vc * 1000
    XD75 = 1 / (Vs * 1000 / MW) * 1000000
    If Check_S_DENS = 0 Then
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HFF0000
        SHELL_OUT(4).BackColor = &HE0E0E0
    ElseIf Check_S_DENS = 1 Then
        XD75 = HScroll_SHELL_DENS / 10
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HC0&
        SHELL_OUT(4).BackColor = &HE0E0E0
    End If
    'Density at condensing film temperature,lb/ft^3
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)

'Viscosity at condensing film temperature,cp
    'eq. 2: ln n = A + B / TK
    'eq. 3: ln n = A + B / TK + C *  TK + D * TK^2
    a = -4.325: b = 1006
    XD77 = Exp(a + b / TK)
    If Check_S_VISC = 0 Then
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HFF0000
        SHELL_OUT(5).BackColor = &HE0E0E0
    ElseIf Check_S_VISC = 1 Then
        XD77 = HScroll_SHELL_VISC.Value / 1000
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HC0&
        SHELL_OUT(5).BackColor = &HE0E0E0
    End If

'Thermal conductivity at condensing film temperature, W/(m.K)
    'Landa =  A1*(1-Tr)^0,38  / Tr^(1/6)
    Ax = 0.00335: a = 1.2: b = 0.5: c = 0.167
    A1 = (Ax * Tb ^ a) / (MW ^ b * Tc ^ c)
    XD79 = A1 * (1 - Tr) ^ 0.38 / Tr ^ (1 / 6)
    
    If Check_S_TC = 0 Then
        'Thermal conductivity, Kcal/hmºC
            SHELL_OUT(1) = Format(XD79, "0.000")
        'Thermal conductivity, Btu/hftºF
            XD78 = XD79 / 1.488
            SHELL_OUT(1).ForeColor = &HFF0000
            SHELL_OUT(1).BackColor = &HE0E0E0
    ElseIf Check_S_TC = 1 Then
        'Thermal conductivity, Kcal/hmºC
            XD79 = HScroll_SHELL_TC / 1000
        'Thermal conductivity, Btu/hftºF
            XD78 = XD79 / 1.488
            SHELL_OUT(1) = Format(XD79, "0.000")
            SHELL_OUT(1).ForeColor = &HC0&
            SHELL_OUT(1).BackColor = &HE0E0E0
    End If
FLUID_VL = 1
End Sub
Private Sub Hexane()

MW = 86.1754
TK = CDbl(XD20) + 273.16            'T, °K
Tc = 507.5                           'Critical temp., °K
Tr = TK / Tc                        'Reduced temperature
Tb = 341.9                          'Boiling point at 1atm °K

'Hexane vapor pressure, bar(a)   (log10(P) = A - (B / (T + C)))
    If TK > 185.29 And TK < 295.6 Then
        A1 = 3.45604
        b1 = 1044.038
        C1 = -53.893
    ElseIf TK > 295.5 And TK < 372.43 Then
        A1 = 4.00266
        b1 = 1171.53
        C1 = -48.784
    End If
    XD21 = 10 ^ (A1 - (b1 / (TK + C1)))
    If Check_CP = 0 Then
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HFF0000
        S_press_KP.BackColor = &HE0E0E0
    ElseIf Check_CP = 1 Then
        XD21 = Spin_S_PRESS / 10000
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HC0&
        S_press_KP.BackColor = &HE0E0E0
    End If

'Hexane enthalpy, kJ/mole   (EvapH = A exp(-aTr) (1 - Tr)^ß)
    A1 = 43.85                                  'kJ/mole
    a2 = -0.039
    b1 = 0.397
    ET = A1 * Exp(-a2 * Tr) * (1 - Tr) ^ b1
'Hexane enthalpy, kJ/kg (MW = 86.1754)
    ET1 = ET * 1000 / MW
'Hexane enthalpy, kcal/kg
    XD24 = ET1 / 4.18
    If Check_LATENT = 0 Then
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HFF0000
        SHELL_OUT(10).BackColor = &HE0E0E0
    ElseIf Check_LATENT = 1 Then
        XD24 = HScroll_LATENT / 4.1868
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HC0&
        SHELL_OUT(10).BackColor = &HE0E0E0
    End If

'Heat capacity,J(mol-°K) - Cp =CPVAP_A + CPVAP_B * K + CPVAP_C * K^2 + CPVAP_D * K^3
    H_cap = -1.746 + 0.5309 * TK - 0.0002903 * TK ^ 2 + 0.00000006054 * TK * 3
    'Heat capacity,J(kg-°K)
    SPH_S = H_cap * 1000 / MW * 0.0002388458966
    If Check_S_SPH = 0 Then
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HFF0000
        SHELL_OUT(3).BackColor = &HE0E0E0
    ElseIf Check_S_SPH = 1 Then
        SPH_S = HScroll_SHELL_SPH / 1000
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HC0&
        SHELL_OUT(3).BackColor = &HE0E0E0
    End If

'Density at condensing film temperature,Kg/m^3
    a = -1.52816: b = 1.43907: c = -0.81446: d = 0.190454
    e = -0.296123: f = 0.386914: g = -0.0427258: h = -0.04480645
    Tc = 507.43
    Tr = TK / Tc
    Wsrk = 0.3007
    Vc = 0.3682
    Vro = 1 + a * (1 - Tr) ^ (1 / 3) + b * (1 - Tr) ^ (2 / 3) + c * (1 - Tr) + d * (1 - Tr) ^ (4 / 3) '0,25<Tr<0,95
    Vrd = (e + f * Tr + g * Tr ^ 2 + h * Tr ^ 3) / (Tr - 1.00001) '0,25<Tr<1,0
    Vs = Vro * (1 - Wsrk * Vrd) * Vc * 1000
    XD75 = 1 / (Vs * 1000 / MW) * 1000000
    If Check_S_DENS = 0 Then
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HFF0000
        SHELL_OUT(4).BackColor = &HE0E0E0
    ElseIf Check_S_DENS = 1 Then
        XD75 = HScroll_SHELL_DENS / 10
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HC0&
        SHELL_OUT(4).BackColor = &HE0E0E0
    End If
    'Density at condensing film temperature,lb/ft^3
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)

'Viscosity at condensing film temperature,cp
    'eq. 2: ln n = A + B / T
    'eq. 3: ln n = A + B / T + C *  T + D * T^2
    'T°C
    'T°K
    a = -4.162: b = 823
    XD77 = Exp(a + b / TK)
    If Check_S_VISC = 0 Then
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HFF0000
        SHELL_OUT(5).BackColor = &HE0E0E0
    ElseIf Check_S_VISC = 1 Then
        XD77 = HScroll_SHELL_VISC.Value / 1000
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HC0&
        SHELL_OUT(5).BackColor = &HE0E0E0
    End If

'Thermal conductivity at condensing film temperature,Btu/hftºF
    'Landa =  A1*(1-Tr)^0,38  / Tr^(1/6)
    Ax = 0.00335: a = 1.2: b = 0.5: c = 0.167
    A1 = (Ax * Tb ^ a) / (MW ^ b * Tc ^ c)
    XD79 = A1 * (1 - Tr) ^ 0.38 / Tr ^ (1 / 6)
    
    If Check_S_TC = 0 Then
        'Thermal conductivity, Kcal/hmºC
            SHELL_OUT(1) = Format(XD79, "0.000")
        'Thermal conductivity, Btu/hftºF
            XD78 = XD79 / 1.488
            SHELL_OUT(1).ForeColor = &HFF0000
            SHELL_OUT(1).BackColor = &HE0E0E0
    ElseIf Check_S_TC = 1 Then
        'Thermal conductivity, Kcal/hmºC
            XD79 = HScroll_SHELL_TC / 1000
        'Thermal conductivity, Btu/hftºF
            XD78 = XD79 / 1.488
            SHELL_OUT(1) = Format(XD79, "0.000")
            SHELL_OUT(1).ForeColor = &HC0&
            SHELL_OUT(1).BackColor = &HE0E0E0
    End If
FLUID_VL = 1
End Sub
Private Sub Isobutane()
MW = 58.1222
TK = CDbl(XD20) + 273.16            'T, °K
Tc = 408.2                          'Critical temp., °K
Tr = TK / Tc                        'Reduced temperature
Tb = 261.4                          'Boiling point at 1atm °K

'Isobutane vapor pressure, bar(a)   (log10(P) = A - (B / (T + C)))
    If TK > 261.31 And TK < 408.12 Then
        A1 = 4.3281
        b1 = 1132.108
        C1 = 0.918
    ElseIf TK > 188 And TK < 261 Then
        A1 = 3.94417
        b1 = 912.141
        C1 = -29.808
    End If
    XD21 = 10 ^ (A1 - (b1 / (TK + C1)))
    If Check_CP = 0 Then
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HFF0000
        S_press_KP.BackColor = &HE0E0E0
    ElseIf Check_CP = 1 Then
        XD21 = Spin_S_PRESS / 10000
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HC0&
        S_press_KP.BackColor = &HE0E0E0
    End If

'Isobutane enthalpy, kJ/mole   (EvapH = (J/mol*K))
    ET = 81.46
    'Isobutane enthalpy, kJ/kg (MW = 58.1222)
    ET1 = ET * 1000 / MW
    'Isobutane enthalpy, kcal/kg
    XD24 = ET1 / 4.18
    If Check_LATENT = 0 Then
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HFF0000
        SHELL_OUT(10).BackColor = &HE0E0E0
    ElseIf Check_LATENT = 1 Then
        XD24 = HScroll_LATENT / 4.1868
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HC0&
        SHELL_OUT(10).BackColor = &HE0E0E0
    End If

'Heat capacity,J(mol-°K) - Cp =CPVAP_A + CPVAP_B * K + CPVAP_C * K^2 + CPVAP_D * K^3
    H_cap = -1.39 + 0.3847 * TK - 0.0001846 * TK ^ 2 + 0.00000002895 * TK * 3
'Heat capacity,J(kg-°K)
    SPH_S = H_cap * 1000 / MW * 0.0002388458966
    If Check_S_SPH = 0 Then
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HFF0000
        SHELL_OUT(3).BackColor = &HE0E0E0
    ElseIf Check_S_SPH = 1 Then
        SPH_S = HScroll_SHELL_SPH / 1000
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HC0&
        SHELL_OUT(3).BackColor = &HE0E0E0
    End If

'Density at condensing film temperature,Kg/m^3
    a = -1.52816: b = 1.43907: c = -0.81446: d = 0.190454
    e = -0.296123: f = 0.386914: g = -0.0427258: h = -0.04480645
    Tc = 408.14
    Tr = TK / Tc
    Wsrk = 0.1825
    Vc = 0.2568
    Vro = 1 + a * (1 - Tr) ^ (1 / 3) + b * (1 - Tr) ^ (2 / 3) + c * (1 - Tr) + d * (1 - Tr) ^ (4 / 3) '0,25<Tr<0,95
    Vrd = (e + f * Tr + g * Tr ^ 2 + h * Tr ^ 3) / (Tr - 1.00001) '0,25<Tr<1,0
    Vs = Vro * (1 - Wsrk * Vrd) * Vc * 1000
    XD75 = 1 / (Vs * 1000 / MW) * 1000000
    If Check_S_DENS = 0 Then
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HFF0000
        SHELL_OUT(4).BackColor = &HE0E0E0
    ElseIf Check_S_DENS = 1 Then
        XD75 = HScroll_SHELL_DENS / 10
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HC0&
        SHELL_OUT(4).BackColor = &HE0E0E0
    End If
    'Density at condensing film temperature,lb/ft^3
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)

'Viscosity at condensing film temperature,cp
    If Check_S_VISC = 0 Then
        XD77 = SHELL_OUT(5)
        SHELL_OUT(5).ForeColor = &HC0&
        SHELL_OUT(5).BackColor = &HE0E0E0
    ElseIf Check_S_VISC = 1 Then
        XD77 = HScroll_SHELL_VISC.Value / 1000
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HC0&
        SHELL_OUT(5).BackColor = &HE0E0E0
    End If

'Thermal conductivity at condensing film temperature,Btu/hftºF
    'Landa =  A1*(1-Tr)^0,38  / Tr^(1/6)
    Ax = 0.00335: a = 1.2: b = 0.5: c = 0.167
    A1 = (Ax * Tb ^ a) / (MW ^ b * Tc ^ c)
    XD79 = A1 * (1 - Tr) ^ 0.38 / Tr ^ (1 / 6)
    
    If Check_S_TC = 0 Then
        'Thermal conductivity, Kcal/hmºC
            SHELL_OUT(1) = Format(XD79, "0.000")
        'Thermal conductivity, Btu/hftºF
            XD78 = XD79 / 1.488
            SHELL_OUT(1).ForeColor = &HFF0000
            SHELL_OUT(1).BackColor = &HE0E0E0
    ElseIf Check_S_TC = 1 Then
        'Thermal conductivity, Kcal/hmºC
            XD79 = HScroll_SHELL_TC / 1000
        'Thermal conductivity, Btu/hftºF
            XD78 = XD79 / 1.488
            SHELL_OUT(1) = Format(XD79, "0.000")
            SHELL_OUT(1).ForeColor = &HC0&
            SHELL_OUT(1).BackColor = &HE0E0E0
    End If
FLUID_VL = 1
End Sub
Private Sub Isopropanol()
MW = 60.095
TK = CDbl(XD20) + 273.16                  'T, °K
Tc = 508.3                          'Critical temp., °K
Tr = TK / Tc                        'Reduced temperature
Tb = 355.4                          'Boiling point at 1atm °K

'Isopropanol vapor pressure, bar(a)   (log10(P) = A - (B / (T + C)))
    If TK > 309.92 And TK < 362.41 Then
        A1 = 4.861
        b1 = 1357.427
        C1 = -75.814
    ElseIf TK > 362.4 And TK < 508.24 Then
        A1 = 4.57795
        b1 = 1221.423
        C1 = -87.474
    End If
    XD21 = 10 ^ (A1 - (b1 / (TK + C1)))
    If Check_CP = 0 Then
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HFF0000
        S_press_KP.BackColor = &HE0E0E0
    ElseIf Check_CP = 1 Then
        XD21 = Spin_S_PRESS / 10000
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HC0&
        S_press_KP.BackColor = &HE0E0E0
    End If

'Isopropanol enthalpy, kJ/mole   (EvapH = A exp(-aTr) (1 - Tr)^ß)
    A1 = 53.38                                  'kJ/mole
    a2 = -0.708
    b1 = 0.6538
    ET = A1 * Exp(-a2 * Tr) * (1 - Tr) ^ b1
    'Isopropanol enthalpy, kJ/kg (MW =  60.0950)
    ET1 = ET * 1000 / MW
    'Isopropanol enthalpy, kcal/kg
    XD24 = ET1 / 4.18
    If Check_LATENT = 0 Then
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HFF0000
        SHELL_OUT(10).BackColor = &HE0E0E0
    ElseIf Check_LATENT = 1 Then
        XD24 = HScroll_LATENT / 4.1868
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HC0&
        SHELL_OUT(10).BackColor = &HE0E0E0
    End If

'Heat capacity,J(mol-°K) - Cp =CPVAP_A + CPVAP_B * K + CPVAP_C * K^2 + CPVAP_D * K^3
    H_cap = 32.43 + 0.1885 * TK - 0.00006405 * TK ^ 2 - 0.00000009261 * TK * 3
'Heat capacity,J(kg-°K)
    SPH_S = H_cap * 1000 / MW * 0.0002388458966
    If Check_S_SPH = 0 Then
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HFF0000
        SHELL_OUT(3).BackColor = &HE0E0E0
    ElseIf Check_S_SPH = 1 Then
        SPH_S = HScroll_SHELL_SPH / 1000
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HC0&
        SHELL_OUT(3).BackColor = &HE0E0E0
    End If

'Density at condensing film temperature,Kg/m^3
    a = -1.52816: b = 1.43907: c = -0.81446: d = 0.190454
    e = -0.296123: f = 0.386914: g = -0.0427258: h = -0.04480645
    Tc = 508.76
    Tr = TK / Tc
    Wsrk = 0.6637
    Vc = 0.2313
    Vro = 1 + a * (1 - Tr) ^ (1 / 3) + b * (1 - Tr) ^ (2 / 3) + c * (1 - Tr) + d * (1 - Tr) ^ (4 / 3) '0,25<Tr<0,95
    Vrd = (e + f * Tr + g * Tr ^ 2 + h * Tr ^ 3) / (Tr - 1.00001) '0,25<Tr<1,0
    Vs = Vro * (1 - Wsrk * Vrd) * Vc * 1000
    XD75 = 1 / (Vs * 1000 / MW) * 1000000
    
    If Check_S_DENS = 0 Then
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HFF0000
        SHELL_OUT(4).BackColor = &HE0E0E0
    ElseIf Check_S_DENS = 1 Then
        XD75 = HScroll_SHELL_DENS / 10
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HC0&
        SHELL_OUT(4).BackColor = &HE0E0E0
    End If
    'Density at condensing film temperature,lb/ft^3
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)

'Viscosity at condensing film temperature,cp
    'eq. 2: ln n = A + B / T
    'eq. 3: ln n = A + B / T + C *  T + D * T^2
    'T°C
    'T°K
    a = -12.28: b = 2666: c = 0.02008: d = -0.00002233
    XD77 = Exp(a + b / TK + c * TK + d * TK ^ 2)
    If Check_S_VISC = 0 Then
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HFF0000
        SHELL_OUT(5).BackColor = &HE0E0E0
    ElseIf Check_S_VISC = 1 Then
        XD77 = HScroll_SHELL_VISC.Value / 1000
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HC0&
        SHELL_OUT(5).BackColor = &HE0E0E0
    End If

'Thermal conductivity at condensing film temperature,Btu/hftºF
    'Landa =  A1*(1-Tr)^0,38  / Tr^(1/6)
    Ax = 0.00339: a = 1.2: b = 0.5: c = 0.167
    A1 = (Ax * Tb ^ a) / (MW ^ b * Tc ^ c)
    XD79 = A1 * (1 - Tr) ^ 0.38 / Tr ^ (1 / 6)
    
    If Check_S_TC = 0 Then
        'Thermal conductivity, Kcal/hmºC
            SHELL_OUT(1) = Format(XD79, "0.000")
        'Thermal conductivity, Btu/hftºF
            XD78 = XD79 / 1.488
            SHELL_OUT(1).ForeColor = &HFF0000
            SHELL_OUT(1).BackColor = &HE0E0E0
    ElseIf Check_S_TC = 1 Then
        'Thermal conductivity, Kcal/hmºC
            XD79 = HScroll_SHELL_TC / 1000
        'Thermal conductivity, Btu/hftºF
            XD78 = XD79 / 1.488
            SHELL_OUT(1) = Format(XD79, "0.000")
            SHELL_OUT(1).ForeColor = &HC0&
            SHELL_OUT(1).BackColor = &HE0E0E0
    End If
FLUID_VL = 1
End Sub
Private Sub Methanol()
MW = 32.0419
TK = CDbl(XD20) + 273.16                  'T, °K
Tc = 512.6                         'Critical temp., °K
Tr = TK / Tc                        'Reduced temperature
Tb = 337.7                          'Boiling point at 1atm °K

'Methanol vapor pressure, bar(a)   (log10(P) = A - (B / (T + C)))
    If TK > 288.1 And TK < 356.3 Then
        A1 = 5.20409
        b1 = 1581.341
        C1 = -33.5
    ElseIf TK > 356.2 And TK < 512.63 Then
        A1 = 5.15853
        b1 = 1569.613
        C1 = -34.846
    End If
    XD21 = 10 ^ (A1 - (b1 / (TK + C1)))
    If Check_CP = 0 Then
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HFF0000
        S_press_KP.BackColor = &HE0E0E0
    ElseIf Check_CP = 1 Then
        XD21 = Spin_S_PRESS / 10000
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HC0&
        S_press_KP.BackColor = &HE0E0E0
    End If

'Methanol enthalpy, kJ/mole   (EvapH = A exp(-aTr) (1 - Tr)^ß)
    A1 = 45.3                                   'kJ/mole
    a2 = -0.31
    b1 = 0.4241
    ET = A1 * Exp(-a2 * Tr) * (1 - Tr) ^ b1
'Methanol enthalpy, kJ/kg (MW =  32.0419)
    ET1 = ET * 1000 / MW
'Methanol enthalpy, kcal/kg
    XD24 = ET1 / 4.18
    If Check_LATENT = 0 Then
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HFF0000
        SHELL_OUT(10).BackColor = &HE0E0E0
    ElseIf Check_LATENT = 1 Then
        XD24 = HScroll_LATENT / 4.1868
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HC0&
        SHELL_OUT(10).BackColor = &HE0E0E0
    End If

'Heat capacity,J(mol-°K) - Cp =CPVAP_A + CPVAP_B * K + CPVAP_C * K^2 + CPVAP_D * K^3
    H_cap = 21.15 + 0.07091 * TK + 0.00002587 * TK ^ 2 - 0.00000002852 * TK * 3
'Heat capacity,J(kg-°K)
    SPH_S = H_cap * 1000 / MW * 0.0002388458966
    If Check_S_SPH = 0 Then
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HFF0000
        SHELL_OUT(3).BackColor = &HE0E0E0
    ElseIf Check_S_SPH = 1 Then
        SPH_S = HScroll_SHELL_SPH / 1000
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HC0&
        SHELL_OUT(3).BackColor = &HE0E0E0
    End If

'Density at condensing film temperature,Kg/m^3
    a = -1.52816: b = 1.43907: c = -0.81446: d = 0.190454
    e = -0.296123: f = 0.386914: g = -0.0427258: h = -0.04480645
    Tc = 513.15
    Tr = TK / Tc
    Wsrk = 0.5536
    Vc = 0.1198
    Vro = 1 + a * (1 - Tr) ^ (1 / 3) + b * (1 - Tr) ^ (2 / 3) + c * (1 - Tr) + d * (1 - Tr) ^ (4 / 3) '0,25<Tr<0,95
    Vrd = (e + f * Tr + g * Tr ^ 2 + h * Tr ^ 3) / (Tr - 1.00001) '0,25<Tr<1,0
    Vs = Vro * (1 - Wsrk * Vrd) * Vc * 1000
    XD75 = 1 / (Vs * 1000 / MW) * 1000000
    If Check_S_DENS = 0 Then
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HFF0000
        SHELL_OUT(4).BackColor = &HE0E0E0
    ElseIf Check_S_DENS = 1 Then
        XD75 = HScroll_SHELL_DENS / 10
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HC0&
        SHELL_OUT(4).BackColor = &HE0E0E0
    End If
    'Density at condensing film temperature,lb/ft^3
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)

'Viscosity at condensing film temperature,cp
    'eq. 2: ln n = A + B / T
    'eq. 3: ln n = A + B / T + C *  T + D * T^2
    'T°C
    'T°K
    a = -39.35: b = 4826: c = 0.1091: d = -0.0001127
    XD77 = Exp(a + b / TK + c * TK + d * TK ^ 2)
    If Check_S_VISC = 0 Then
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HFF0000
        SHELL_OUT(5).BackColor = &HE0E0E0
    ElseIf Check_S_VISC = 1 Then
        XD77 = HScroll_SHELL_VISC.Value / 1000
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HC0&
        SHELL_OUT(5).BackColor = &HE0E0E0
    End If

'Thermal conductivity at condensing film temperature W/ m ºK
    Landa = 0.3225 - 0.00047856 * TK + 0.0000001168 * TK ^ 2
    'Thermal conductivity at condensing film temperature Kcal/mºK
    XD79 = Landa * 0.8604
    'Thermal conductivity, Btu/hftºF
    XD78 = XD79 / 1.488
    If Check_S_TC = 0 Then
        SHELL_OUT(1) = Format(XD79, "0.000")
        SHELL_OUT(1).ForeColor = &HFF0000
        SHELL_OUT(1).BackColor = &HE0E0E0
    ElseIf Check_S_TC = 1 Then
        XD79 = HScroll_SHELL_TC / 1000
        XD78 = XD79 / 1.488
        SHELL_OUT(1) = Format(XD79, "0.000")
        SHELL_OUT(1).ForeColor = &HC0&
        SHELL_OUT(1).BackColor = &HE0E0E0
    End If
FLUID_VL = 1
End Sub
Private Sub Propylene()
MW = 42.081
TK = 273.16 + CDbl(XD20)            'T, °K
Tc = 364.9                          'Critical temp., °K
Tr = TK / Tc                        'Reduced temperature
Tb = 225.6                          'Boiling point at 1atm °K

'Propylene Condensation pressure, bara
    P_P1 = -11.0028519534878            'H5B5
    P_P2 = 8.88800106381893E-02         'H5C5
    P_P3 = -2.25058891968011E-04        'H5D5
    P_P4 = 2.10257407154681E-07         'H5E5
    P_P5 = TK
    XD21 = 10 ^ (P_P1 + P_P2 * P_P5 + P_P3 * P_P5 ^ 2 + P_P4 * P_P5 ^ 3)
    If Check_CP = 0 Then
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HFF0000
        S_press_KP.BackColor = &HE0E0E0
    ElseIf Check_CP = 1 Then
        XD21 = Spin_S_PRESS / 10000
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HC0&
        S_press_KP.BackColor = &HE0E0E0
    End If

'Propylene IN enthalpy, Kcal/Kg
    P_E1 = 6.12185604296598E-03         'H5C6
    P_E2 = 1.86493425411241             'H5B6
    P_E3 = -1.29346385178499E-05        'H5D6
    P_E4 = Log(273.16 + CDbl(XD19)) / LN
    XD22 = (-P_E1 + Sqr((P_E1 ^ 2 - 4 * (P_E2 - P_E4) * P_E3))) / (2 * P_E3)
'Propylene OUT enthalpy, Kcal/Kg
    P_A1 = -7.10616231537362E-02        'H5C7
    P_A2 = 9.1814480913749              'H5B7
    P_A3 = 1.84047520046369E-04         'H5D7
    P_A4 = Log(TK) / LN
    XD23 = (-P_A1 + Sqr((P_A1 ^ 2 - 4 * (P_A2 - P_A4) * P_A3))) / (2 * P_A3)
'Propylene latent heat, Kcal/Kg
    XD24 = XD23 - XD22
    If Check_LATENT = 0 Then
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HFF0000
        SHELL_OUT(10).BackColor = &HE0E0E0
    ElseIf Check_LATENT = 1 Then
        XD24 = HScroll_LATENT / 4.1868
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HC0&
        SHELL_OUT(10).BackColor = &HE0E0E0
    End If

'Heat capacity,J(mol-°K) - Cp =CPVAP_A + CPVAP_B * K + CPVAP_C * K^2 + CPVAP_D * K^3
    H_cap = 3.71 + 0.2345 * TK - 0.000116 * TK ^ 2 + 0.00000002205 * TK * 3
'Heat capacity,J(kg-°K)
    SPH_S = H_cap * 1000 / MW * 0.0002388458966
    If Check_S_SPH = 0 Then
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HFF0000
        SHELL_OUT(3).BackColor = &HE0E0E0
    ElseIf Check_S_SPH = 1 Then
        SPH_S = HScroll_SHELL_SPH / 1000
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HC0&
        SHELL_OUT(3).BackColor = &HE0E0E0
    End If

'Propylene density at condensing film temperature,Kg/m^3
    P_D1 = 1.36605329630155      'H5B11
    P_D2 = -7.21816263067366E-03 'H5C11
    P_D3 = 2.52654604770147E-05  'H5D11
    P_D4 = -3.59719473341389E-08 'H5E11
    XD75 = (P_D1 + P_D2 * TK + P_D3 * TK ^ 2 + P_D4 * TK ^ 3) * 1000
    If Check_S_DENS = 0 Then
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HFF0000
        SHELL_OUT(4).BackColor = &HE0E0E0
    ElseIf Check_S_DENS = 1 Then
        XD75 = HScroll_SHELL_DENS / 10
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HC0&
        SHELL_OUT(4).BackColor = &HE0E0E0
    End If
    'Propylene density at condensing film temperature,lb/ft^3
    XD76 = XD75 * 2.20462 * (0.3048 ^ 3)

'Propylene viscosity at condensing film temperature,cp
    P_V1 = 21.3988888935207         'H5B9
    P_V2 = -0.184389462440866       'H5C9
    P_V3 = 5.72566322754654E-04     'H5D9
    P_V4 = -6.20305620533474E-07    'H5E9
    XD77 = (P_V1 + P_V2 * TK + P_V3 * TK ^ 2 + P_V4 * TK ^ 3) / 10
    If Check_S_VISC = 0 Then
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HFF0000
        SHELL_OUT(5).BackColor = &HE0E0E0
    ElseIf Check_S_VISC = 1 Then
        XD77 = HScroll_SHELL_VISC.Value / 1000
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HC0&
        SHELL_OUT(5).BackColor = &HE0E0E0
    End If

'Propylene thermal conductivity at condensing film temperature,Btu/hftºF
    P_C1 = 0.183205128197412        'H5B10
    P_C2 = 3.90637140722696E-04     'H5C10
    P_C3 = -3.27505827537143E-06    'H5D10
    P_C4 = 3.88500388538241E-09     'H5E10
    XD78 = (P_C1 + P_C2 * TK + P_C3 * TK ^ 2 + P_C4 * TK ^ 3) * 0.5778
    'Propylene thermal conductivity at condensing film temperature, Kcal/h m ºC
    XD79 = XD78 * 1.488
    If Check_S_TC = 0 Then
        SHELL_OUT(1) = Format(XD79, "0.000")
        SHELL_OUT(1).ForeColor = &HFF0000
        SHELL_OUT(1).BackColor = &HE0E0E0
    ElseIf Check_S_TC = 1 Then
        XD79 = HScroll_SHELL_TC / 1000
        XD78 = XD79 / 1.488
        SHELL_OUT(1) = Format(XD79, "0.000")
        SHELL_OUT(1).ForeColor = &HC0&
        SHELL_OUT(1).BackColor = &HE0E0E0
    End If
FLUID_VL = 1
End Sub
Private Sub Propylene_glycol()
MW = 92.095
TK = 273.16 + CDbl(XD20)
Tc = 726
Tb = 563

'Propylene glycol vapor pressure, bar(a) (log10(P) = A - (B / (T + C)))
    A1 = 6.07936
    b1 = 2692.187
    C1 = -17.94
    XD21 = 10 ^ (A1 - (b1 / (TK + C1)))
    If Check_CP = 0 Then
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HFF0000
        S_press_KP.BackColor = &HE0E0E0
    ElseIf Check_CP = 1 Then
        XD21 = Spin_S_PRESS / 10000
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HC0&
        S_press_KP.BackColor = &HE0E0E0
    End If

'Propylene glycol ENTHALPY, kcal/kg
    If TK > 365 And TK < 496 Then
        XD24 = 52 * 1000 / 76.0944 / 4.18
    Else
        XD24 = 58.6 * 1000 / 76.0944 / 4.18
    End If
    If Check_LATENT = 0 Then
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HFF0000
        SHELL_OUT(10).BackColor = &HE0E0E0
    ElseIf Check_LATENT = 1 Then
        XD24 = HScroll_LATENT / 4.1868
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HC0&
        SHELL_OUT(10).BackColor = &HE0E0E0
    End If

'Heat capacity,J(mol-°K) - Cp =CPVAP_A + CPVAP_B * K + CPVAP_C * K^2 + CPVAP_D * K^3
    H_cap = 8.424 + 0.4442 * TK - 0.0003159 * TK ^ 2 + 0.00000009378 * TK * 3
    'Heat capacity,J(kg-°K)
    SPH_S = H_cap * 1000 / MW * 0.0002388458966
    If Check_S_SPH = 0 Then
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HFF0000
        SHELL_OUT(3).BackColor = &HE0E0E0
    ElseIf Check_S_SPH = 1 Then
        SPH_S = HScroll_SHELL_SPH / 1000
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HC0&
        SHELL_OUT(3).BackColor = &HE0E0E0
    End If

'Density at condensing film temperature,Kg/m^3
    a = -1.52816: b = 1.43907: c = -0.81446: d = 0.190454
    e = -0.296123: f = 0.386914: g = -0.0427258: h = -0.04480645
    Tc = 726
    Tr = TK / Tc
    Wsrk = 1.9845
    Vc = 0.4119
    Vro = 1 + a * (1 - Tr) ^ (1 / 3) + b * (1 - Tr) ^ (2 / 3) + c * (1 - Tr) + d * (1 - Tr) ^ (4 / 3) '0,25<Tr<0,95
    Vrd = (e + f * Tr + g * Tr ^ 2 + h * Tr ^ 3) / (Tr - 1.00001) '0,25<Tr<1,0
    Vs = Vro * (1 - Wsrk * Vrd) * Vc * 1000
    XD75 = 1 / (Vs * 1000 / MW) * 1000000
    If Check_S_DENS = 0 Then
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HFF0000
        SHELL_OUT(4).BackColor = &HE0E0E0
    ElseIf Check_S_DENS = 1 Then
        XD75 = HScroll_SHELL_DENS / 10
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HC0&
        SHELL_OUT(4).BackColor = &HE0E0E0
    End If
    'Density at condensing film temperature,lb/ft^3
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)

'Viscosity at condensing film temperature,cp
    'eq. 2: ln n = A + B / T
    'eq. 3: ln n = A + B / T + C *  T + D * T^2
    'T = °K
    a = -7.577: b = 3233
    XD77 = Exp(a + b / TK)
    If Check_S_VISC = 0 Then
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HFF0000
        SHELL_OUT(5).BackColor = &HE0E0E0
    ElseIf Check_S_VISC = 1 Then
        XD77 = HScroll_SHELL_VISC.Value / 1000
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HC0&
        SHELL_OUT(5).BackColor = &HE0E0E0
    End If
'Thermal conductivity at condensing film temperature,Btu/hftºF
    If Check_S_TC = 0 Then
        'Thermal conductivity, Kcal/hmºC
            XD79 = SHELL_OUT(1)
        'Thermal conductivity, Btu/hftºF
            XD78 = XD79 / 1.488
            SHELL_OUT(1).ForeColor = &HC0&
            SHELL_OUT(1).BackColor = &HE0E0E0
    ElseIf Check_S_TC = 1 Then
        'Thermal conductivity, Kcal/hmºC
            XD79 = HScroll_SHELL_TC / 1000
        'Thermal conductivity, Btu/hftºF
            XD78 = XD79 / 1.488
            SHELL_OUT(1) = Format(XD79, "0.000")
            SHELL_OUT(1).ForeColor = &HC0&
            SHELL_OUT(1).BackColor = &HE0E0E0
    End If
FLUID_VL = 1
End Sub
Private Sub Toluene()
MW = 92.1384
TK = CDbl(XD20) + 273.16            'T, °K
Tc = 591.8                          'Critical temp., °K
Tr = TK / Tc                        'Reduced temperature
Tb = 383.8                          'Boiling point at 1atm °K

'Toluene Condensation pressure, bara
    XD21 = 10 ^ (-20.4550478749456 + 0.122159504194137 * TK - 2.52019819590687E-04 * TK ^ 2 + 1.89239695329835E-07 * TK ^ 3)
    If Check_CP = 0 Then
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HFF0000
        S_press_KP.BackColor = &HE0E0E0
    ElseIf Check_CP = 1 Then
        XD21 = Spin_S_PRESS / 10000
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HC0&
        S_press_KP.BackColor = &HE0E0E0
    End If

'Toluene vapor pressure, bar(a)   (log10(P) = A - (B / (T + C)))
    If TK > 273.13 And TK < 297.89 Then
        A1 = 4.23679
        b1 = 1426.448
        C1 = -45.957
    ElseIf TK > 297.88 And TK < 384.66 Then
        A1 = 4.07827
        b1 = 1343.943
        C1 = -53.773
    ElseIf TK > 384.65 And TK < 580# Then
        A1 = 4.54436
        b1 = 1738.123
        C1 = 0.394
    End If
    P1 = 10 ^ (A1 - (b1 / (TK + C1)))

'Toluene IN enthalpy, Kcal/Kg
    XD22 = (-6.19081006602303E-03 + Sqr((6.19081006602303E-03 ^ 2 - 4 * (2.05395178519857 - Log(273.16 + XD19) / LN) * -1.55678980165631E-05))) / (2 * (-1.55678980165631E-05))
'Toluene OUT enthalpy, Kcal/Kg
    XD23 = (-1.87023824724839E-02 + Sqr((1.87023824724839E-02 ^ 2 - 4 * (0.253332590099177 - Log(TK) / LN) * -3.62854720193424E-05))) / (2 * (-3.62854720193424E-05))
    'Toluene latent heat Kcal/Kg
    XD24 = XD23 - XD22
    'Enthalpy, kJ/mole (EvapH = A exp(-ßTr) (1 - Tr)^ß)
    A1 = 53.09                       'kJ/mole
    b1 = 0.2774
    TR1 = TK / 591.7                 'Reduced temperature T/Tc
    ET = A1 * Exp(-b1 * TR1) * (1 - TR1) ^ b1
    'Enthalpy, kJ/kg (MW=92,1384)
    ET1 = ET * 1000 / MW
    'Enthalpy, kcal/kg
    ET2 = ET1 / 4.18
    If Check_LATENT = 0 Then
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HFF0000
        SHELL_OUT(10).BackColor = &HE0E0E0
    ElseIf Check_LATENT = 1 Then
        XD24 = HScroll_LATENT / 4.1868
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HC0&
        SHELL_OUT(10).BackColor = &HE0E0E0
    End If

'Heat capacity,J(mol-°K) - Cp =CPVAP_A + CPVAP_B * K + CPVAP_C * K^2 + CPVAP_D * K^3
    H_cap = -24.35 + 0.5125 * TK - 0.0002765 * TK ^ 2 + 0.00000004911 * TK * 3
    'Heat capacity,J(kg-°K)
    SPH_S = H_cap * 1000 / MW * 0.0002388458966
    If Check_S_SPH = 0 Then
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HFF0000
        SHELL_OUT(3).BackColor = &HE0E0E0
    ElseIf Check_S_SPH = 1 Then
        SPH_S = HScroll_SHELL_SPH / 1000
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HC0&
        SHELL_OUT(3).BackColor = &HE0E0E0
    End If

'Toluene density at condensing film temperature,Kg/m^3
    XD75 = (1.09657697034396 - 7.13586701948427E-04 * TK - 6.21961123052573E-08 * TK ^ 2 - 6.0875153993795E-10 * TK ^ 3) * 1000
    If Check_S_DENS = 0 Then
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HFF0000
        SHELL_OUT(4).BackColor = &HE0E0E0
    ElseIf Check_S_DENS = 1 Then
        XD75 = HScroll_SHELL_DENS / 10
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HC0&
        SHELL_OUT(4).BackColor = &HE0E0E0
    End If
    'Toluene density at condensing film temperature,lb/ft^3
    XD76 = XD75 * 2.20462 * (0.3048 ^ 3)

'Toluene viscosity at condensing film temperature,cp
    XD77 = (114.844320200991 - 0.826324102497882 * TK + 2.06924742120918E-03 * TK ^ 2 - 1.7689717708644E-06 * TK ^ 3) / 10
    If Check_S_VISC = 0 Then
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HFF0000
        SHELL_OUT(5).BackColor = &HE0E0E0
    ElseIf Check_S_VISC = 1 Then
        XD77 = HScroll_SHELL_VISC.Value / 1000
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HC0&
        SHELL_OUT(5).BackColor = &HE0E0E0
    End If

'Toluene thermal conductivity at condensing film temperature,Btu/hftºF
    XD78 = (0.266958042246326 - 7.4378954632313E-04 * TK + 1.34365635103354E-06 * TK ^ 2 - 1.1655011726189E-09 * TK ^ 3) * 0.5778
'Toluene thermal conductivity at condensing film temperature Kcal/h m ºC
    XD79 = XD78 * 1.488
    If Check_S_TC = 0 Then
        SHELL_OUT(1) = Format(XD79, "0.000")
        SHELL_OUT(1).ForeColor = &HFF0000
        SHELL_OUT(1).BackColor = &HE0E0E0
    ElseIf Check_S_TC = 1 Then
        XD79 = HScroll_SHELL_TC / 1000
        XD78 = XD79 / 1.488
        SHELL_OUT(1) = Format(XD79, "0.000")
        SHELL_OUT(1).ForeColor = &HC0&
        SHELL_OUT(1).BackColor = &HE0E0E0
    End If
FLUID_VL = 1
End Sub
Private Sub VCM()
MW = 62.499
TK = CDbl(XD20) + 273.16            'T, °K
Tc = 425                            'Critical temp., °K
Tr = TK / Tc                        'Reduced temperature
Tb = 259.8                          'Boiling point at 1atm °K

'VCM Condensation pressure, bara  (WAGNER EQUATION)
    V_B4 = TK
    V_B5 = 429.7
    V_B6 = V_B4 / V_B5
    V_B8 = 51.5
    V_B7 = 1 - V_B6
    V_B9 = -6.50008
    V_B10 = 1.21422
    V_B11 = -2.57876
    V_B12 = -2.00937
    XD21 = V_B8 * Exp((1 - V_B7) ^ -1 * (V_B9 * V_B7 + V_B10 * V_B7 ^ 1.5 + V_B11 * V_B7 ^ 3 + V_B12 * V_B7 ^ 6))
    If Check_CP = 0 Then
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HFF0000
        S_press_KP.BackColor = &HE0E0E0
    ElseIf Check_CP = 1 Then
        XD21 = Spin_S_PRESS / 10000
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HC0&
        S_press_KP.BackColor = &HE0E0E0
    End If
'VCM latent heat Kcal/Kg
    V_L5 = TK
    V_L6 = 8.314
    V_L7 = 429.7
    V_L8 = 0.122
    V_L9 = V_L5 / V_L7
    V_L10 = 62.499
    XD24 = V_L6 * V_L7 * (7.08 * (1 - V_L9) ^ (0.354) + 10.95 * V_L8 * (1 - V_L9) ^ 0.456) / (V_L10 * 4.18)
    If Check_LATENT = 0 Then
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HFF0000
        SHELL_OUT(10).BackColor = &HE0E0E0
    ElseIf Check_LATENT = 1 Then
        XD24 = HScroll_LATENT / 4.1868
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HC0&
        SHELL_OUT(10).BackColor = &HE0E0E0
    End If
'Heat capacity,J(kg-°K)
    'Cp = CPVAPA + CPVAPB * T + CPVAPC * T ^ 2 + CPVAPD * T ^ 3
    Cpvap_A = 5.949
    Cpvap_B = 0.2019
    Cpvap_C = -0.0001536
    Cpvap_D = 0.00000004773
    H_cap = Cpvap_A + Cpvap_B * TK + Cpvap_C * TK ^ 2 + Cpvap_D * TK * 3
    'Heat capacity,J(kg-°K)
    SPH_S = H_cap * 1000 / MW * 0.0002388458966
    If Check_S_SPH = 0 Then
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HFF0000
        SHELL_OUT(3).BackColor = &HE0E0E0
    ElseIf Check_S_SPH = 1 Then
        SPH_S = HScroll_SHELL_SPH / 1000
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HC0&
        SHELL_OUT(3).BackColor = &HE0E0E0
    End If

'VCM density at condensing film temperature,Kg/m^3
        V_D4 = TK             'T=
        V_D5 = 429.7            'Tc=
        V_D6 = 0.1293           'w_SRK=
        V_D7 = 0.1722           'V*=
        V_D8 = 62.499           'M=
        V_D9 = -1.52816         'a=
        V_D10 = 1.43907         'b=
        V_D11 = -0.81446        'c=
        V_D12 = 0.190454        'd=
        V_D13 = -0.296123       'e=
        V_D14 = 0.386914        'f=
        V_D15 = -0.0427258      'g=
        V_D16 = -0.0480645      'h=
        V_D17 = V_D4 / V_D5     'T_r=
        'V_R0=
        V_D18 = 1 + V_D9 * (1 - V_D17) ^ (1 / 3) + V_D10 * (1 - V_D17) ^ (2 / 3) + V_D11 * (1 - V_D17) + V_D12 * (1 - V_D17) ^ (4 / 3)
        'V_Rd=
        V_D19 = (V_D13 + V_D14 * V_D17 + V_D15 * V_D17 ^ 2 + V_D16 * V_D17 ^ 3) / (V_D17 - 1.00001)
        'V_s=
        V_D20 = V_D7 * V_D18 * (1 - V_D6 * V_D19)
        'Density
        XD75 = 1 / (V_D20 / V_D8)
    If Check_S_DENS = 0 Then
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HFF0000
        SHELL_OUT(4).BackColor = &HE0E0E0
    ElseIf Check_S_DENS = 1 Then
        XD75 = HScroll_SHELL_DENS / 10
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HC0&
        SHELL_OUT(4).BackColor = &HE0E0E0
    End If
    'VCM density at condensing film temperature,lb/ft^3
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)

'VCM viscosity at condensing film temperature,cp
    V_V3 = TK
    V_V4 = 2
    V_V5 = -0.152 - 0.042 * B4
    V_V6 = 1.91 - 1.459
    V_V7 = V_V4 + V_V5 + V_V6
    V_V9 = 28.86 + 37.439 * V_V7 - 1.3547 * V_V7 ^ 2 + 0.0276 * V_V7 ^ 3
    V_V10 = 24.79 + 66.885 * V_V7 - 1.3173 * V_V7 ^ 2 - 0.00377 * V_V7 ^ 3 - 44.94 + 5.41 * V_V7 - 26.38
    XD77 = 10 ^ (V_V10 * (V_V3 ^ -1 - V_V9 ^ -1))
    If Check_S_VISC = 0 Then
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HFF0000
        SHELL_OUT(5).BackColor = &HE0E0E0
    ElseIf Check_S_VISC = 1 Then
        XD77 = HScroll_SHELL_VISC.Value / 1000
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HC0&
        SHELL_OUT(5).BackColor = &HE0E0E0
    End If

'VCM thermal conductivity at condensing film temperature,Kcal/h m ºC
    V_T5 = TK
    V_T6 = 259.8
    V_T7 = 62.499
    V_T8 = 429.7
    V_T9 = V_T5 / V_T8
    V_T10 = V_T6 / V_T8
    XD79 = ((1.11 / (V_T7 ^ 0.5)) * (3 + 20 * (1 - V_T9) ^ (2 / 3)) / (3 + 20 * (1 - V_T10) ^ (2 / 3))) * (3.6 / 4.18)
'Thermal conductivity, Btu/hftºF
    XD78 = XD79 / 1.488
    If Check_S_TC = 0 Then
        SHELL_OUT(1) = Format(XD79, "0.000")
        SHELL_OUT(1).ForeColor = &HFF0000
        SHELL_OUT(1).BackColor = &HE0E0E0
    ElseIf Check_S_TC = 1 Then
        XD79 = HScroll_SHELL_TC / 1000
        XD78 = XD79 / 1.488
        SHELL_OUT(1) = Format(XD79, "0.000")
        SHELL_OUT(1).ForeColor = &HC0&
        SHELL_OUT(1).BackColor = &HE0E0E0
    End If
FLUID_VL = 1
End Sub
Private Sub Naphtalene()
MW = 128.174
TK = CDbl(XD20) + 273.16            'T, °K
Tc = 748.4                          'Critical temp., °K
Tr = TK / Tc                        'Reduced temperature
Tb = 490.1                          'Boiling point at 1atm °K
Pc = 40.5                           'Critical pression, bar
Tbr = Tb / Tc                       'Reduced boiling point
RR = 8.314                          'gas costant, J/(mol.°K)
FLUID_VL = 1
'Naphtalene vapor pressure, bar(a)
    Vp_A = -7.85178
    Vp_B = 2.17172
    Vp_C = -3.70504
    Vp_D = -4.81238
    X = 1 - Tr
    XD21 = Exp((1 - X) ^ -1 * (Vp_A * X + Vp_B * X ^ 1.5 + Vp_C * X ^ 3 + Vp_D * X ^ 6)) * Pc
    If Check_CP = 0 Then
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HFF0000
        S_press_KP.BackColor = &HE0E0E0
    ElseIf Check_CP = 1 Then
        XD21 = Spin_S_PRESS / 10000
        S_press_KP = Format(XD21 * 100, "0.00")
        S_press_KP.ForeColor = &HC0&
        S_press_KP.BackColor = &HE0E0E0
    End If

'Naphtalene enthalpy, J/mole
    Hv1 = 1.093 * RR * Tc * ((Tbr * (Log(Pc) - 1.013)) / (0.93 - Tbr))
    X = (Tbr / Tr) * ((1 - Tr) / (1 - Tbr))
    q = 0.35298
    p = 0.13856
    Hv2 = Hv1 * (Tr / Tbr) * ((X + X ^ q) / (1 + X ^ p))
    'Naphtalene enthalpy, kJ/kg
    Hv3 = Hv2 / MW
    'Naphtalene enthalpy, kcal/kg
    XD24 = Hv3 * 0.2388458966
    If Check_LATENT = 0 Then
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HFF0000
        SHELL_OUT(10).BackColor = &HE0E0E0
    ElseIf Check_LATENT = 1 Then
        XD24 = HScroll_LATENT / 4.1868
        SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0")
        SHELL_OUT(10).ForeColor = &HC0&
        SHELL_OUT(10).BackColor = &HE0E0E0
    End If

'Naphtalene Heat capacity,J(mol-°K) - Cp =CPVAP_A + CPVAP_B * K + CPVAP_C * K^2 + CPVAP_D * K^3
    Cpvap_A = -68.8
    Cpvap_B = 0.8499
    Cpvap_C = -0.0006506
    Cpvap_D = 0.000000198
    H_cap = Cpvap_A + Cpvap_B * TK + Cpvap_C * TK ^ 2 + Cpvap_D * TK * 3
    'Heat capacity,J(kg-°K)
    SPH_S = H_cap * 1000 / MW * 0.0002388458966
    'Heat capacity,J(kg-°K)
    SPH_S = H_cap * 1000 / MW * 0.0002388458966
    
    If Check_S_SPH = 0 Then
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HFF0000
        SHELL_OUT(3).BackColor = &HE0E0E0
    ElseIf Check_S_SPH = 1 Then
        SPH_S = HScroll_SHELL_SPH / 1000
        SHELL_OUT(3) = Format(SPH_S, "0.000")
        SHELL_OUT(3).ForeColor = &HC0&
        SHELL_OUT(3).BackColor = &HE0E0E0
    End If

'Naphtalene Density at condensing film temperature,Kg/m^3
    a = -1.52816: b = 1.43907: c = -0.81446: d = 0.190454
    e = -0.296123: f = 0.386914: g = -0.0427258: h = -0.04480645
    Wsrk = 0.3
    Vc = 0.383
    Vro = 1 + a * (1 - Tr) ^ (1 / 3) + b * (1 - Tr) ^ (2 / 3) + c * (1 - Tr) + d * (1 - Tr) ^ (4 / 3) '0,25<Tr<0,95
    Vrd = (e + f * Tr + g * Tr ^ 2 + h * Tr ^ 3) / (Tr - 1.00001) '0,25<Tr<1,0
    Vs = Vro * (1 - Wsrk * Vrd) * Vc * 1000
    XD75 = 1 / (Vs * 1000 / MW) * 1000000
    If Check_S_DENS = 0 Then
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HFF0000
        SHELL_OUT(4).BackColor = &HE0E0E0
    ElseIf Check_S_DENS = 1 Then
        XD75 = HScroll_SHELL_DENS / 10
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
        SHELL_OUT(4) = Format(XD75, "0.0")
        SHELL_OUT(4).ForeColor = &HC0&
        SHELL_OUT(4).BackColor = &HE0E0E0
    End If
    'Density at condensing film temperature,lb/ft^3
        XD76 = XD75 * 2.20462 * (0.3048 ^ 3)

'Naphtalene Viscosity at condensing film temperature,cp
    'eq. 2: ln n = A + B / T
    'eq. 3: ln n = A + B / T + C *  T + D * T^2
    'T°C
    'T°K
    a = -10.27: b = 2517: c = 0.01098: d = -0.000005867
    XD77 = Exp(a + b / TK + c * TK + d * TK ^ 2)
    If Check_S_VISC = 0 Then
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HFF0000
        SHELL_OUT(5).BackColor = &HE0E0E0
    ElseIf Check_S_VISC = 1 Then
        XD77 = HScroll_SHELL_VISC.Value / 1000
        SHELL_OUT(5) = Format(XD77, "0.000")
        SHELL_OUT(5).ForeColor = &HC0&
        SHELL_OUT(5).BackColor = &HE0E0E0
    End If

'Thermal conductivity at condensing film temperature W/ m ºK
    A1 = 0.0346
    a = 1.2
    b = 1
    c = 0.167
    Y = (A1 * Tb ^ a) / (MW ^ b * Tc ^ c)
    Landa = Y * (1 - Tr) ^ 0.38 / Tr ^ (1 / 6)
    'Thermal conductivity at condensing film temperature Kcal/mºK
    XD79 = Landa * 0.8604
    'Thermal conductivity, Btu/hftºF
    XD78 = XD79 / 1.488
    If Check_S_TC = 0 Then
        SHELL_OUT(1) = Format(XD79, "0.000")
        SHELL_OUT(1).ForeColor = &HFF0000
        SHELL_OUT(1).BackColor = &HE0E0E0
    ElseIf Check_S_TC = 1 Then
        XD79 = HScroll_SHELL_TC / 1000
        XD78 = XD79 / 1.488
        SHELL_OUT(1) = Format(XD79, "0.000")
        SHELL_OUT(1).ForeColor = &HC0&
        SHELL_OUT(1).BackColor = &HE0E0E0
    End If
FLUID_VL = 1
End Sub
