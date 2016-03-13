VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHX 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WaterProof HX - Heat Exchangers Performance Monitoring - Data Input"
   ClientHeight    =   10110
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   15225
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9888.021
   ScaleMode       =   0  'User
   ScaleWidth      =   14310.17
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
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
      Height          =   4875
      Left            =   300
      TabIndex        =   296
      ToolTipText     =   "Double click the unit name or date."
      Top             =   4260
      Visible         =   0   'False
      Width           =   3975
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   195
         Left            =   120
         TabIndex        =   321
         Top             =   4560
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         Height          =   270
         Left            =   3000
         TabIndex        =   299
         Top             =   4500
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Bindings        =   "frmHX.frx":0000
         Height          =   3615
         Left            =   240
         TabIndex        =   297
         Top             =   300
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   7
         Cols            =   3
         FixedCols       =   0
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Double click a UNIT or a DATE or TEST_NO. The first item found for the unit or for the date will be set."
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   60
         TabIndex        =   300
         Top             =   3900
         Width           =   3855
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
      Height          =   6075
      Left            =   5400
      TabIndex        =   309
      Top             =   3060
      Visible         =   0   'False
      Width           =   5655
      Begin RichTextLib.RichTextBox RichTextBox_REMARKS 
         DataField       =   "REMARKS"
         DataSource      =   "Data1"
         Height          =   5595
         Left            =   180
         TabIndex        =   310
         Top             =   360
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   9869
         _Version        =   393217
         TextRTF         =   $"frmHX.frx":0015
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "    Searching   all tests"
      Height          =   555
      Left            =   60
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   298
      Top             =   6660
      Width           =   1275
   End
   Begin VB.TextBox Date_Y 
      Height          =   255
      Left            =   180
      TabIndex        =   295
      Text            =   "Date_Y"
      Top             =   9600
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
      Left            =   9960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_Unit_sort"
      Top             =   10440
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
      Height          =   5715
      Left            =   1440
      TabIndex        =   147
      Top             =   3420
      Width           =   6675
      Begin VB.ComboBox Combo_CURRENT 
         BackColor       =   &H80000018&
         DataField       =   "CURRENT"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   330
         ItemData        =   "frmHX.frx":0091
         Left            =   1620
         List            =   "frmHX.frx":00A1
         TabIndex        =   278
         Text            =   "Combo_CURRENT"
         ToolTipText     =   "Chose if the exchanger is horizontal or vertical."
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
         Height          =   4395
         Left            =   0
         TabIndex        =   225
         Top             =   1320
         Width           =   3495
         Begin VB.CheckBox Check_U 
            Alignment       =   1  'Right Justify
            Caption         =   """U"""
            DataField       =   "Check_X"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   3060
            TabIndex        =   318
            ToolTipText     =   "Check if ""U"" tubes type"
            Top             =   480
            Width           =   195
         End
         Begin VB.CheckBox Check_MAT_FACTOR 
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
            Left            =   3300
            MaskColor       =   &H0000FF00&
            Style           =   1  'Graphical
            TabIndex        =   307
            ToolTipText     =   "Check to activate the cursor and enter the value."
            Top             =   3780
            UseMaskColor    =   -1  'True
            Width           =   135
         End
         Begin VB.HScrollBar Spin_T_LEN 
            Height          =   255
            LargeChange     =   100
            Left            =   2220
            Max             =   2000
            Min             =   1
            TabIndex        =   292
            Top             =   720
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
            TabIndex        =   23
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            Height          =   225
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   243
            Top             =   3540
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
            Height          =   225
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   239
            Top             =   3300
            Width           =   795
         End
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
            Height          =   225
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   238
            Top             =   3780
            Width           =   795
         End
         Begin VB.ComboBox Combo_MAT_SHEET 
            BackColor       =   &H80000018&
            DataField       =   "TUBES_SHEET_MAT"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   330
            ItemData        =   "frmHX.frx":00DC
            Left            =   1080
            List            =   "frmHX.frx":0158
            Sorted          =   -1  'True
            TabIndex        =   28
            ToolTipText     =   "Select from the list the material of tube-sheet. Use TAB Key to save the new enter"
            Top             =   2880
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
            TabIndex        =   18
            ToolTipText     =   "Use the scrollbar to enter the NUMBER of TUBES."
            Top             =   480
            Width           =   915
         End
         Begin VB.HScrollBar HScroll_T_NO 
            Height          =   255
            LargeChange     =   100
            Left            =   2220
            Max             =   20000
            Min             =   10
            TabIndex        =   19
            Top             =   480
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
            Height          =   225
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   226
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
            ItemData        =   "frmHX.frx":040B
            Left            =   1320
            List            =   "frmHX.frx":044B
            TabIndex        =   26
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
            ItemData        =   "frmHX.frx":049C
            Left            =   1080
            List            =   "frmHX.frx":0537
            Sorted          =   -1  'True
            TabIndex        =   27
            ToolTipText     =   "Select from the list the material of tubes.Use TAB Key to save the new enter."
            Top             =   2520
            Width           =   2355
         End
         Begin VB.ComboBox ComboT_OD 
            BackColor       =   &H80000018&
            ForeColor       =   &H000000C0&
            Height          =   330
            ItemData        =   "frmHX.frx":07EA
            Left            =   1320
            List            =   "frmHX.frx":080F
            TabIndex        =   25
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
            TabIndex        =   319
            Top             =   480
            Width           =   225
         End
         Begin VB.Label lblLabels 
            Caption         =   "Enter the value (suggested: 0.5 - 1.0)"
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
            Height          =   255
            Index           =   16
            Left            =   840
            TabIndex        =   308
            Top             =   4020
            Width           =   2685
         End
         Begin MSForms.SpinButton Spin_MAT_FACTOR 
            Height          =   195
            Left            =   2220
            TabIndex        =   306
            Top             =   3780
            Width           =   435
            Size            =   "767;344"
            Min             =   50
            Max             =   150
            Position        =   50
            Orientation     =   1
         End
         Begin VB.Label Label48 
            Caption         =   "Flow area section, m2:"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   120
            TabIndex        =   242
            Top             =   3540
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Thermal cond., Kcal/(h m^2 ºC/m):"
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
            Index           =   33
            Left            =   120
            TabIndex        =   241
            Top             =   3300
            Width           =   2475
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
            Height          =   255
            Index           =   34
            Left            =   120
            TabIndex        =   240
            Top             =   3780
            Width           =   1245
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
            TabIndex        =   237
            Top             =   180
            Width           =   765
         End
         Begin VB.Label lblLabels 
            Caption         =   "Tubes:"
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   56
            Left            =   540
            TabIndex        =   236
            Top             =   2640
            Width           =   465
         End
         Begin VB.Label lblLabels 
            Caption         =   "Tube sheet:"
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   55
            Left            =   180
            TabIndex        =   235
            Top             =   2940
            Width           =   885
         End
         Begin VB.Label lblLabels 
            Caption         =   "Tubes Number:"
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   234
            Top             =   480
            Width           =   2205
         End
         Begin VB.Label lblLabels 
            Caption         =   "Tubes lenght, m:"
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   233
            Top             =   735
            Width           =   1245
         End
         Begin VB.Label lblLabels 
            Caption         =   "O.D., mm:"
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   232
            Top             =   1260
            Width           =   1245
         End
         Begin VB.Label lblLabels 
            Caption         =   "BWG:"
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   231
            Top             =   1935
            Width           =   765
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
            Height          =   255
            Index           =   32
            Left            =   120
            TabIndex        =   230
            Top             =   2220
            Width           =   825
         End
         Begin VB.Label lblLabels 
            Caption         =   "Passes, n°:"
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   229
            Top             =   1005
            Width           =   1005
         End
         Begin VB.Label lblLabels 
            Caption         =   "Material:"
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   35
            Left            =   120
            TabIndex        =   228
            Top             =   2460
            Width           =   645
         End
         Begin MSForms.SpinButton Spin_T_PAS 
            Height          =   255
            Left            =   2220
            TabIndex        =   22
            Top             =   960
            Width           =   855
            Size            =   "1508;450"
            Min             =   1
            Max             =   8
            Position        =   1
            Orientation     =   1
         End
         Begin MSForms.SpinButton Spin_T_OD 
            Height          =   255
            Left            =   2220
            TabIndex        =   24
            Top             =   1200
            Width           =   855
            Size            =   "1508;450"
            Max             =   10000
            Position        =   100
            Orientation     =   1
         End
         Begin VB.Label Label47 
            Caption         =   "O.D, inches:"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   120
            TabIndex        =   227
            Top             =   1560
            Width           =   1395
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
         TabIndex        =   16
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
         TabIndex        =   14
         ToolTipText     =   "Use the cursors to enter the number of units in parallel."
         Top             =   600
         Width           =   555
      End
      Begin VB.TextBox ELEVATION 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         DataField       =   "ELEVATION"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "Use the cursors to enter the elevation of the exchanger."
         Top             =   240
         Width           =   735
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
         Height          =   4395
         Left            =   3420
         TabIndex        =   203
         Top             =   1320
         Width           =   3255
         Begin VB.HScrollBar Spin_BAFFLES_SPACE 
            Height          =   255
            LargeChange     =   50
            Left            =   2280
            Max             =   2000
            Min             =   1
            TabIndex        =   294
            Top             =   1380
            Value           =   1000
            Width           =   855
         End
         Begin VB.HScrollBar Spin_SHELL_ID 
            Height          =   255
            LargeChange     =   100
            Left            =   2280
            Max             =   3000
            Min             =   1
            TabIndex        =   293
            Top             =   1680
            Value           =   1000
            Width           =   855
         End
         Begin VB.CheckBox CHECK_BAFFLES_N 
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
            Height          =   120
            Left            =   1380
            MaskColor       =   &H0000FF00&
            Style           =   1  'Graphical
            TabIndex        =   280
            ToolTipText     =   "Check to enter the value"
            Top             =   780
            UseMaskColor    =   -1  'True
            Width           =   135
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
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   246
            Top             =   3120
            Width           =   975
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
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   244
            Top             =   3420
            Width           =   975
         End
         Begin VB.ComboBox SHELL_PITCH_CONF 
            BackColor       =   &H80000018&
            DataField       =   "SHELL_PITCH_CONF"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   330
            ItemData        =   "frmHX.frx":084E
            Left            =   1380
            List            =   "frmHX.frx":0858
            Sorted          =   -1  'True
            TabIndex        =   39
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
            TabIndex        =   31
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
            TabIndex        =   37
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
            TabIndex        =   35
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
            TabIndex        =   33
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
            TabIndex        =   36
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
            TabIndex        =   29
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
            ItemData        =   "frmHX.frx":0870
            Left            =   1080
            List            =   "frmHX.frx":08E9
            Sorted          =   -1  'True
            TabIndex        =   40
            ToolTipText     =   "Select from the list the material of the shell.Use TAB Key to save the new enter"
            Top             =   2640
            Width           =   2055
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
            TabIndex        =   247
            Top             =   3180
            Width           =   1125
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
            TabIndex        =   245
            Top             =   3480
            Width           =   1785
         End
         Begin VB.Label lblLabels 
            Caption         =   "Pitch pattern.:"
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   69
            Left            =   120
            TabIndex        =   222
            Top             =   2400
            Width           =   1005
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
            TabIndex        =   221
            Top             =   840
            Width           =   795
         End
         Begin VB.Label lblLabels 
            Caption         =   " I.D., mm:"
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   59
            Left            =   120
            TabIndex        =   220
            Top             =   1740
            Width           =   765
         End
         Begin VB.Label lblLabels 
            Caption         =   "Passes:"
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   58
            Left            =   120
            TabIndex        =   219
            Top             =   540
            Width           =   585
         End
         Begin VB.Label lblLabels 
            Caption         =   "Material shell:"
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   57
            Left            =   120
            TabIndex        =   218
            Top             =   2760
            Width           =   1005
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
            TabIndex        =   217
            Top             =   1140
            Width           =   1035
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
            TabIndex        =   216
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label30 
            Caption         =   "Baffle space, mm:"
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
            Left            =   120
            TabIndex        =   215
            Top             =   1440
            Width           =   1275
         End
         Begin MSForms.SpinButton Spin_BAFFLES_N 
            Height          =   255
            Left            =   2280
            TabIndex        =   32
            Top             =   780
            Width           =   855
            Size            =   "1508;450"
            Position        =   10
            Orientation     =   1
         End
         Begin MSForms.SpinButton Spin_TUBES_PITCH 
            Height          =   255
            Left            =   2280
            TabIndex        =   38
            Top             =   1980
            Width           =   855
            Size            =   "1508;450"
            Max             =   1000
            Position        =   25
            Orientation     =   1
         End
         Begin MSForms.SpinButton Spin_BAFFLES_CUT 
            Height          =   255
            Left            =   2280
            TabIndex        =   34
            Top             =   1080
            Width           =   855
            Size            =   "1508;450"
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
            TabIndex        =   204
            Top             =   180
            Width           =   825
         End
         Begin MSForms.SpinButton Spin_S_PASS 
            Height          =   255
            Left            =   2280
            TabIndex        =   30
            Top             =   480
            Width           =   855
            Size            =   "1508;450"
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
         ItemData        =   "frmHX.frx":0B8E
         Left            =   3060
         List            =   "frmHX.frx":0B98
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
         ItemData        =   "frmHX.frx":0BB2
         Left            =   1080
         List            =   "frmHX.frx":0BD7
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
         TabIndex        =   149
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label Lab_COND 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "COOLING TYPE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   4140
         TabIndex        =   288
         Top             =   960
         Width           =   2475
      End
      Begin VB.Label lblLabels 
         Caption         =   "Flow arrangement:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   279
         Top             =   1020
         Width           =   1425
      End
      Begin VB.Label lblLabels 
         Caption         =   "Series, n°:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   71
         Left            =   2280
         TabIndex        =   224
         Top             =   645
         Width           =   765
      End
      Begin MSForms.SpinButton Spin_SERIES_N 
         Height          =   255
         Left            =   3600
         TabIndex        =   17
         Top             =   600
         Width           =   615
         Size            =   "1085;450"
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
         TabIndex        =   223
         Top             =   645
         Width           =   825
      End
      Begin MSForms.SpinButton Spin_PARALLEL_N 
         Height          =   255
         Left            =   1620
         TabIndex        =   15
         Top             =   600
         Width           =   555
         Size            =   "979;450"
         Min             =   1
         Max             =   8
         Position        =   1
         Orientation     =   1
      End
      Begin VB.Label lblLabels 
         Caption         =   "Elevation, m:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   62
         Left            =   4380
         TabIndex        =   205
         Top             =   285
         Width           =   1005
      End
      Begin MSForms.SpinButton Spin_ELEVATION 
         Height          =   255
         Left            =   6000
         TabIndex        =   13
         Top             =   240
         Width           =   615
         Size            =   "1085;450"
         Min             =   1
         Position        =   1
         Orientation     =   1
      End
      Begin VB.Label Label39 
         Caption         =   "Position:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2280
         TabIndex        =   176
         Top             =   300
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "TEMA type:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   52
         Left            =   180
         TabIndex        =   167
         Top             =   300
         Width           =   825
      End
      Begin VB.Label lblLabels 
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
         Left            =   4380
         TabIndex        =   150
         Top             =   600
         Width           =   1005
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
      Height          =   1275
      Left            =   1440
      TabIndex        =   200
      Top             =   2160
      Width           =   6675
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
         ItemData        =   "frmHX.frx":0C12
         Left            =   1740
         List            =   "frmHX.frx":0C14
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
         TabIndex        =   277
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
         TabIndex        =   206
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
         TabIndex        =   202
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
         TabIndex        =   201
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
      TabIndex        =   138
      Top             =   0
      Width           =   6675
      Begin VB.CommandButton Command1 
         Caption         =   "Reset check"
         Height          =   270
         Left            =   5460
         TabIndex        =   286
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
         TabIndex        =   285
         ToolTipText     =   "Check to compare this record with design"
         Top             =   1860
         UseMaskColor    =   -1  'True
         Width           =   2355
      End
      Begin VB.TextBox CHECK_ACTUAL 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   4140
         TabIndex        =   284
         Top             =   1860
         Width           =   150
      End
      Begin VB.TextBox CHECK_DESIGN 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   276
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
         ItemData        =   "frmHX.frx":0C16
         Left            =   3000
         List            =   "frmHX.frx":0C18
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
         CalendarTitleBackColor=   -2147483624
         CalendarTitleForeColor=   16512
         CalendarTrailingForeColor=   4210816
         CustomFormat    =   "01/01/01"
         Format          =   51642369
         CurrentDate     =   37920
         MinDate         =   36526
      End
      Begin VB.TextBox txt_num 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         DataField       =   "Test_NO"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   5940
         Locked          =   -1  'True
         TabIndex        =   166
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
         TabIndex        =   168
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
         TabIndex        =   169
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
         TabIndex        =   170
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
         TabIndex        =   171
         Text            =   "Unit"
         Top             =   1440
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "Actual:"
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   3600
         TabIndex        =   287
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
         TabIndex        =   180
         Top             =   1860
         Width           =   555
      End
      Begin VB.Label lblLabels 
         Caption         =   "Test_n°:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   21
         Left            =   5340
         TabIndex        =   148
         Top             =   180
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "Date  (must be unique for same unit!):"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   0
         Left            =   225
         TabIndex        =   143
         Top             =   285
         Width           =   2805
      End
      Begin VB.Label lblLabels 
         Caption         =   "Plant name:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   20
         Left            =   225
         TabIndex        =   142
         Top             =   615
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Location:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   19
         Left            =   225
         TabIndex        =   141
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Country:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   18
         Left            =   225
         TabIndex        =   140
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Heat transfer Unit I.D.:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   17
         Left            =   225
         TabIndex        =   139
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
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_TOWER"
      Top             =   10440
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
      Left            =   3780
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_PROCESS_STREAM"
      Top             =   10440
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
      Left            =   1980
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_PROCESS_DESCR"
      Top             =   10440
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
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_Plant_UNIT"
      Top             =   10440
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
      Height          =   5655
      Left            =   8160
      TabIndex        =   146
      Top             =   3480
      Width           =   6780
      Begin VB.HScrollBar HScroll_WATER_FF 
         Height          =   195
         LargeChange     =   100
         Left            =   3660
         Max             =   2000
         TabIndex        =   315
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
         TabIndex        =   312
         Top             =   5220
         Width           =   1215
      End
      Begin VB.CheckBox Check_WET_STEAM 
         DataField       =   "Check_WET_STEAM"
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
         Left            =   5280
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   305
         ToolTipText     =   "Check to enter your wet steam value"
         Top             =   3540
         UseMaskColor    =   -1  'True
         Width           =   135
      End
      Begin VB.HScrollBar HScroll_WET_STEAM 
         Height          =   195
         LargeChange     =   100
         Left            =   5880
         Max             =   1000
         TabIndex        =   304
         Top             =   3600
         Value           =   100
         Width           =   795
      End
      Begin VB.TextBox Wet_steam 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "Wet_steam"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   302
         Top             =   3540
         Width           =   555
      End
      Begin VB.CheckBox Check_water_steam 
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
         TabIndex        =   301
         ToolTipText     =   "Check to link to the total inlet tubes flow."
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   135
      End
      Begin VB.CheckBox Check_LATENT 
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
         TabIndex        =   290
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   2640
         UseMaskColor    =   -1  'True
         Width           =   135
      End
      Begin VB.HScrollBar HScroll_LATENT 
         Height          =   195
         LargeChange     =   100
         Left            =   4800
         Max             =   3000
         TabIndex        =   289
         Top             =   2700
         Value           =   1000
         Width           =   855
      End
      Begin VB.CheckBox Check_T_TC 
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
         TabIndex        =   281
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   135
      End
      Begin VB.CheckBox Check_S_TC 
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
         TabIndex        =   107
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   135
      End
      Begin VB.HScrollBar HScroll_SHELL_TC 
         Height          =   195
         LargeChange     =   100
         Left            =   4800
         Max             =   1500
         TabIndex        =   108
         Top             =   540
         Value           =   150
         Width           =   795
      End
      Begin VB.CheckBox Check_CT 
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
         TabIndex        =   126
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   2880
         UseMaskColor    =   -1  'True
         Width           =   135
      End
      Begin VB.HScrollBar HScroll_C_TEMP 
         Height          =   195
         LargeChange     =   100
         Left            =   4800
         Max             =   30000
         TabIndex        =   127
         Top             =   2940
         Value           =   2500
         Width           =   855
      End
      Begin VB.HScrollBar HScroll_TUBES_SPH 
         Height          =   195
         LargeChange     =   100
         Left            =   2880
         Max             =   1500
         TabIndex        =   91
         Top             =   780
         Value           =   150
         Width           =   795
      End
      Begin VB.HScrollBar HScroll_TUBES_DENS 
         Height          =   195
         LargeChange     =   100
         Left            =   2880
         Max             =   15000
         TabIndex        =   94
         Top             =   1020
         Value           =   150
         Width           =   795
      End
      Begin VB.HScrollBar HScroll_TUBES_VISC 
         Height          =   195
         LargeChange     =   100
         Left            =   2880
         Max             =   5000
         TabIndex        =   97
         Top             =   1260
         Value           =   150
         Width           =   795
      End
      Begin VB.CheckBox Check_T_DENS 
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
         TabIndex        =   93
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   135
      End
      Begin VB.CheckBox Check_T_VISC 
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
         TabIndex        =   96
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   135
      End
      Begin VB.CheckBox Check_T_SPH 
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
         TabIndex        =   90
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   135
      End
      Begin VB.HScrollBar HScroll_TUBES_TC 
         Height          =   195
         LargeChange     =   100
         Left            =   2880
         Max             =   1500
         TabIndex        =   104
         Top             =   540
         Value           =   150
         Width           =   795
      End
      Begin VB.CheckBox Check_P_DROP_S 
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
         TabIndex        =   123
         ToolTipText     =   "Check to see the allowed pressure drop."
         Top             =   2400
         UseMaskColor    =   -1  'True
         Width           =   135
      End
      Begin VB.CheckBox Check_P_DROP_T 
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
         TabIndex        =   103
         ToolTipText     =   "Check to see the allowed pressure drop."
         Top             =   2400
         UseMaskColor    =   -1  'True
         Width           =   135
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
         TabIndex        =   125
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
         Height          =   255
         Index           =   10
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   124
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
         Height          =   255
         Index           =   9
         Left            =   3900
         TabIndex        =   122
         Top             =   2400
         Width           =   915
      End
      Begin VB.TextBox SHELL_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   8
         Left            =   3900
         TabIndex        =   121
         Top             =   2160
         Width           =   915
      End
      Begin VB.TextBox SHELL_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "SHELL_DUTY"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   7
         Left            =   3900
         TabIndex        =   120
         Top             =   1920
         Width           =   915
      End
      Begin VB.TextBox SHELL_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "SHELL_REYNOLDS"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   6
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   119
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
         TabIndex        =   102
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
         TabIndex        =   101
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
         TabIndex        =   100
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
         TabIndex        =   99
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
         TabIndex        =   98
         Top             =   1440
         Width           =   915
      End
      Begin VB.TextBox SHELL_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "SHELL_VEL"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   2
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   118
         Top             =   1440
         Width           =   915
      End
      Begin VB.CheckBox Check_U_CLEAN 
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
         TabIndex        =   133
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   4140
         UseMaskColor    =   -1  'True
         Width           =   135
      End
      Begin VB.HScrollBar HScroll_U_CLEAN 
         Height          =   195
         LargeChange     =   100
         Left            =   3660
         Max             =   5000
         TabIndex        =   134
         Top             =   4200
         Value           =   1000
         Width           =   1035
      End
      Begin VB.HScrollBar HScroll_SHELL_VISC 
         Height          =   195
         LargeChange     =   100
         Left            =   4800
         Max             =   10000
         TabIndex        =   117
         Top             =   1260
         Value           =   150
         Width           =   795
      End
      Begin VB.HScrollBar HScroll_SHELL_DENS 
         Height          =   195
         LargeChange     =   100
         Left            =   4800
         Max             =   15000
         TabIndex        =   114
         Top             =   1020
         Value           =   150
         Width           =   795
      End
      Begin VB.HScrollBar HScroll_SHELL_SPH 
         Height          =   195
         LargeChange     =   100
         Left            =   4800
         Max             =   1500
         TabIndex        =   111
         Top             =   780
         Value           =   150
         Width           =   795
      End
      Begin VB.CheckBox Check_S_SPH 
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
         TabIndex        =   110
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   135
      End
      Begin VB.CheckBox Check_S_VISC 
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
         TabIndex        =   116
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   135
      End
      Begin VB.CheckBox Check_S_DENS 
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
         TabIndex        =   113
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   135
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
         TabIndex        =   130
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
         TabIndex        =   132
         Top             =   4140
         Width           =   1215
      End
      Begin VB.TextBox SHELL_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "SHELL_VISC"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   5
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   115
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
         TabIndex        =   95
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
         Height          =   255
         Index           =   4
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   112
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
         TabIndex        =   92
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
         Height          =   255
         Index           =   3
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   109
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
         TabIndex        =   89
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
         TabIndex        =   88
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
         TabIndex        =   87
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
         TabIndex        =   128
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
         TabIndex        =   129
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
         TabIndex        =   136
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
         TabIndex        =   135
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
         TabIndex        =   131
         Top             =   3900
         Width           =   1215
      End
      Begin VB.TextBox TUBES_FF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   137
         Top             =   4920
         Width           =   1215
      End
      Begin VB.TextBox SHELL_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "SHELL_T_COND"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   1
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   106
         ToolTipText     =   "If blue and not checked,, the value is calculated."
         Top             =   480
         Width           =   915
      End
      Begin VB.TextBox SHELL_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   105
         Top             =   240
         Width           =   915
      End
      Begin MSForms.ToggleButton Thermal_bal_tubes 
         Height          =   600
         Left            =   2880
         TabIndex        =   330
         ToolTipText     =   "Check to balance duty"
         Top             =   1860
         Width           =   825
         BackColor       =   -2147483633
         ForeColor       =   192
         DisplayStyle    =   6
         Size            =   "1455;1058"
         Value           =   "0"
         Caption         =   "Thermal balance"
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.ToggleButton Thermal_bal_shell 
         Height          =   345
         Left            =   5160
         TabIndex        =   329
         ToolTipText     =   "Check to balance duty"
         Top             =   1680
         Width           =   1530
         BackColor       =   -2147483633
         ForeColor       =   192
         DisplayStyle    =   6
         Size            =   "2708;609"
         Value           =   "0"
         Caption         =   "Thermal balance F"
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Allowed:  "
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
         TabIndex        =   314
         Top             =   5280
         Width           =   2235
      End
      Begin VB.Label Label12 
         Caption         =   "[(hm^2 ºC)/kcal]*10^-4"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   4965
         TabIndex        =   313
         Top             =   5280
         Width           =   1635
      End
      Begin VB.Label Label22 
         Caption         =   "Wet steam, %"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   17
         Left            =   4200
         TabIndex        =   303
         Top             =   3600
         Width           =   1035
      End
      Begin VB.Label Label22 
         Caption         =   "Approach"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   15
         Left            =   4260
         TabIndex        =   291
         Top             =   3180
         Width           =   915
      End
      Begin VB.Label Label22 
         Caption         =   "(Allowable)"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   14
         Left            =   5220
         TabIndex        =   283
         Top             =   2460
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label22 
         Caption         =   "(Allowable)"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   12
         Left            =   2940
         TabIndex        =   282
         Top             =   2460
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label22 
         Caption         =   "m/s (liquid fraction)"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   26
         Left            =   4920
         TabIndex        =   253
         Top             =   1500
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "[(hm^2 ºC)/kcal]*10^-4"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   3780
         TabIndex        =   266
         Top             =   4980
         Width           =   1635
      End
      Begin VB.Label Label22 
         Caption         =   "kg/h"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   13
         Left            =   4860
         TabIndex        =   263
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label22 
         Caption         =   "°C"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   29
         Left            =   3780
         TabIndex        =   262
         Top             =   3660
         Width           =   435
      End
      Begin VB.Label Label22 
         Caption         =   "kJ/Kg"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   5760
         TabIndex        =   261
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
         TabIndex        =   257
         Top             =   1260
         Width           =   930
      End
      Begin VB.Label lblLabels 
         Caption         =   "Kg/m3"
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
         TabIndex        =   255
         Top             =   1020
         Width           =   930
      End
      Begin VB.Label lblLabels 
         Caption         =   "Kcal/(Kg ºC)"
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
         TabIndex        =   252
         Top             =   780
         Width           =   930
      End
      Begin VB.Label lbl_tubes 
         Caption         =   "Kcal/h m ºC"
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
         TabIndex        =   250
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
         Left            =   4800
         TabIndex        =   190
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label lblLabels 
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   186
         Top             =   4920
         Width           =   2235
      End
      Begin VB.Label Label40 
         Caption         =   "KW"
         ForeColor       =   &H00404080&
         Height          =   195
         Left            =   4800
         TabIndex        =   177
         Top             =   1980
         Width           =   375
      End
      Begin VB.Label Label22 
         Caption         =   "*C"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   11
         Left            =   5760
         TabIndex        =   164
         Top             =   2940
         Width           =   255
      End
      Begin VB.Label Label22 
         Caption         =   "kcal/(h m^2 ºC)"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   10
         Left            =   3780
         TabIndex        =   163
         Top             =   4440
         Width           =   1155
      End
      Begin VB.Label Label22 
         Caption         =   "kcal/(h m^2 ºC)"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   9
         Left            =   4740
         TabIndex        =   162
         Top             =   4200
         Width           =   1395
      End
      Begin VB.Label Label22 
         Caption         =   "°C"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   6
         Left            =   3780
         TabIndex        =   161
         Top             =   3180
         Width           =   555
      End
      Begin VB.Label Label22 
         Caption         =   "°C"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   5
         Left            =   3780
         TabIndex        =   160
         Top             =   3420
         Width           =   435
      End
      Begin VB.Label Label22 
         Caption         =   "°C"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   4
         Left            =   3780
         TabIndex        =   159
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label Label22 
         Caption         =   "kPa"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   4920
         TabIndex        =   158
         Top             =   2460
         Width           =   435
      End
      Begin VB.Label Label24 
         Caption         =   "m3/h/kPa^(1/2) - Tubes-side"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   3780
         TabIndex        =   157
         Top             =   4680
         Width           =   2235
      End
      Begin VB.Label lbl_tubes 
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   248
         Top             =   240
         Width           =   1755
      End
      Begin VB.Label lbl_tubes 
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   264
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lbl_tubes 
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   254
         Top             =   720
         Width           =   1755
      End
      Begin VB.Label lbl_tubes 
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   256
         Top             =   960
         Width           =   1755
      End
      Begin VB.Label lbl_tubes 
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   258
         Top             =   1200
         Width           =   1755
      End
      Begin VB.Label lbl_tubes 
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   183
         Top             =   1440
         Width           =   1755
      End
      Begin VB.Label lbl_tubes 
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   181
         Top             =   1680
         Width           =   1755
      End
      Begin VB.Label lbl_tubes 
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   189
         Top             =   1920
         Width           =   1755
      End
      Begin VB.Label lbl_tubes 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Heat flux:"
         ForeColor       =   &H00404080&
         Height          =   255
         Index           =   8
         Left            =   180
         TabIndex        =   178
         Top             =   2160
         Width           =   1755
      End
      Begin VB.Label lblLabels 
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   184
         Top             =   2400
         Width           =   1755
      End
      Begin VB.Label lblLabels 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Latent Heat:"
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
         TabIndex        =   260
         Top             =   2640
         Width           =   3735
      End
      Begin VB.Label lblLabels 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Condensing  temperature:"
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
         TabIndex        =   185
         Top             =   2880
         Width           =   3735
      End
      Begin VB.Label lblLabels 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   180
         TabIndex        =   188
         Top             =   3180
         Width           =   2235
      End
      Begin VB.Label lblLabels 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   180
         TabIndex        =   187
         Top             =   3420
         Width           =   2235
      End
      Begin VB.Label Label22 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MTDc"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   16
         Left            =   180
         TabIndex        =   251
         Top             =   3660
         Width           =   2235
      End
      Begin VB.Label Label22 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Skin temp."
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   249
         Top             =   3900
         Width           =   2235
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   180
         TabIndex        =   191
         Top             =   4140
         Width           =   2235
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Overall DIRTY heat transfer coefficient:Overall DIRTY heat transfer coefficient"
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
         Left            =   180
         TabIndex        =   192
         Top             =   4380
         Width           =   2235
      End
      Begin VB.Label lblLabels 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   180
         TabIndex        =   182
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
      Left            =   9960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_Date"
      Top             =   10080
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
      Left            =   7500
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_Country"
      Top             =   10080
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
      Left            =   5580
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_LOC"
      Top             =   10080
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
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_Unit"
      Top             =   10080
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
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_Plant"
      Top             =   10080
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
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_test"
      Top             =   10080
      Visible         =   0   'False
      Width           =   1695
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
      Top             =   9180
      Width           =   4695
   End
   Begin VB.TextBox PLANT_X 
      DataField       =   "PLANT_Z"
      DataSource      =   "Data2"
      Height          =   255
      Left            =   8760
      TabIndex        =   175
      Text            =   "PLANT_X"
      Top             =   10440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame7 
      Caption         =   "Search"
      ForeColor       =   &H00FF0000&
      Height          =   3435
      Left            =   0
      TabIndex        =   172
      Top             =   3060
      Width           =   1455
      Begin VB.ComboBox Combo_Date_X 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00004080&
         Height          =   330
         Left            =   60
         TabIndex        =   199
         Text            =   "Search Date"
         ToolTipText     =   "Chose the test-date to see."
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton Com_Go 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Go to date"
         Height          =   255
         Left            =   120
         TabIndex        =   198
         Top             =   3060
         Width           =   1095
      End
      Begin VB.ComboBox Combo_UNIT_1 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00004080&
         Height          =   330
         Left            =   60
         TabIndex        =   194
         Text            =   "Search unit"
         ToolTipText     =   "Chose the unit of the selected plant"
         Top             =   1740
         Width           =   1335
      End
      Begin VB.ComboBox Combo_Plant_1 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00004080&
         Height          =   330
         Left            =   60
         TabIndex        =   193
         Text            =   "Search plant"
         ToolTipText     =   "Chose the plant in the database"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         DrawMode        =   14  'Copy Pen
         X1              =   0
         X2              =   1380
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         Caption         =   "Search by date"
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   120
         TabIndex        =   197
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Label Label45 
         Caption         =   "Unit"
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   120
         TabIndex        =   196
         Top             =   1500
         Width           =   795
      End
      Begin VB.Label Label44 
         Caption         =   "Plant:"
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   120
         TabIndex        =   195
         Top             =   840
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
         TabIndex        =   174
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Caption         =   "Search by Plant and Unit selection"
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   60
         TabIndex        =   173
         Top             =   300
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
      Left            =   7500
      TabIndex        =   165
      Text            =   "UNIT_X"
      Top             =   10440
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2475
      Left            =   120
      TabIndex        =   156
      Top             =   120
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
      Height          =   3555
      Left            =   8100
      TabIndex        =   144
      Top             =   0
      Width           =   6795
      Begin VB.CheckBox Check_T_OUT 
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
         TabIndex        =   331
         ToolTipText     =   "Check to enter outlet temp. value and calculate flow rate"
         Top             =   2220
         UseMaskColor    =   -1  'True
         Width           =   135
      End
      Begin VB.Frame Frame_VAP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Inlet %"
         ForeColor       =   &H00FF0000&
         Height          =   1395
         Left            =   5940
         TabIndex        =   323
         Top             =   600
         Width           =   855
         Begin VB.TextBox LIQ_PERC 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   326
            Text            =   "LIQ_PERC"
            ToolTipText     =   "Fraction of liquid + non condensable"
            Top             =   960
            Width           =   555
         End
         Begin VB.CheckBox Check_VAP_P 
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
            Left            =   60
            MaskColor       =   &H0000FF00&
            Style           =   1  'Graphical
            TabIndex        =   324
            ToolTipText     =   "Check to link to vapor shell flow"
            Top             =   420
            UseMaskColor    =   -1  'True
            Width           =   135
         End
         Begin VB.TextBox VAP_PERC 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            DataField       =   "VAP_FRACTION"
            DataSource      =   "Data1"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   60
            TabIndex        =   325
            Text            =   "VAP_PERC"
            ToolTipText     =   "Fraction of condensing vapor"
            Top             =   420
            Width           =   555
         End
         Begin MSForms.SpinButton Spin_VAP_P 
            Height          =   255
            Left            =   600
            TabIndex        =   328
            ToolTipText     =   "Fraction of condensing vapor"
            Top             =   420
            Width           =   195
            Size            =   "344;450"
            Max             =   1000
            Position        =   1
         End
         Begin VB.Line Line3 
            BorderColor     =   &H000000FF&
            X1              =   0
            X2              =   420
            Y1              =   840
            Y2              =   960
         End
         Begin VB.Line Line4 
            BorderColor     =   &H000000FF&
            X1              =   0
            X2              =   420
            Y1              =   1320
            Y2              =   1200
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   " fractions   "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   0
            TabIndex        =   327
            Top             =   180
            Width           =   795
         End
      End
      Begin VB.TextBox PROCESS_TARGET_T_OUT 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         DataField       =   "PROCESS_TARGET_TEMP"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   6060
         TabIndex        =   316
         ToolTipText     =   "Process target outlet temp."
         Top             =   2220
         Width           =   495
      End
      Begin VB.CheckBox Check_CP 
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
         Left            =   3840
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Check to activate the cursor and enter the value."
         Top             =   2940
         UseMaskColor    =   -1  'True
         Width           =   135
      End
      Begin VB.TextBox FACT_FLOW 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         DataField       =   "FACT_FLOW"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   6000
         TabIndex        =   41
         Text            =   "FACT_FLOW"
         ToolTipText     =   "Span factor of flow"
         Top             =   360
         Width           =   495
      End
      Begin VB.HScrollBar Spin_PF 
         Height          =   195
         LargeChange     =   100
         Left            =   4740
         Max             =   2000
         TabIndex        =   86
         Top             =   3240
         Value           =   400
         Width           =   855
      End
      Begin VB.CheckBox Check_PF 
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
         Left            =   3840
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Uncheck to to have process fouling calculated"
         Top             =   3180
         UseMaskColor    =   -1  'True
         Width           =   135
      End
      Begin VB.TextBox SHELL_FF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "SHELL_FF"
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
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   84
         Top             =   3180
         Width           =   915
      End
      Begin VB.HScrollBar Spin_S_PRESS 
         Height          =   195
         LargeChange     =   100
         Left            =   4740
         Max             =   3000
         TabIndex        =   83
         Top             =   3000
         Value           =   100
         Width           =   855
      End
      Begin VB.TextBox S_press_KP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "Press_COND"
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
         Height          =   225
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   81
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   2940
         Width           =   915
      End
      Begin VB.TextBox SHELL_P_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "SHELL_P_OUT"
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
         Height          =   225
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   79
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   2700
         Width           =   915
      End
      Begin VB.HScrollBar HScroll_SHELL_P_OUT 
         Height          =   195
         LargeChange     =   100
         Left            =   4740
         Max             =   30000
         TabIndex        =   80
         Top             =   2760
         Value           =   5000
         Width           =   855
      End
      Begin VB.TextBox TUBES_P_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "TUBES_P_OUT"
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
         Height          =   225
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   60
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   2700
         Width           =   915
      End
      Begin VB.HScrollBar HScroll_TUBES_P_OUT 
         Height          =   195
         LargeChange     =   100
         Left            =   2880
         Max             =   30000
         TabIndex        =   61
         Top             =   2760
         Value           =   5000
         Width           =   855
      End
      Begin VB.TextBox SHELL_P_IN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "SHELL_P_IN"
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
         Height          =   225
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   77
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   2460
         Width           =   915
      End
      Begin VB.HScrollBar HScroll_SHELL_P_IN 
         Height          =   195
         LargeChange     =   100
         Left            =   4740
         Max             =   30000
         TabIndex        =   78
         Top             =   2520
         Value           =   5000
         Width           =   855
      End
      Begin VB.TextBox TUBES_P_IN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "TUBES_P_IN"
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
         Height          =   225
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   58
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   2460
         Width           =   915
      End
      Begin VB.HScrollBar HScroll_TUBES_P_IN 
         Height          =   195
         LargeChange     =   100
         Left            =   2880
         Max             =   30000
         TabIndex        =   59
         Top             =   2520
         Value           =   5000
         Width           =   855
      End
      Begin VB.TextBox SHELL_TEMP_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "SHELL_TEMP_OUT"
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
         Height          =   240
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   75
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   2205
         Width           =   915
      End
      Begin VB.HScrollBar HScroll_SHELL_T_OUT 
         Height          =   195
         LargeChange     =   100
         Left            =   4740
         Max             =   30000
         Min             =   5
         TabIndex        =   76
         Top             =   2250
         Value           =   3500
         Width           =   855
      End
      Begin VB.HScrollBar Spin_TUBES_T_OUT 
         Height          =   195
         LargeChange     =   100
         Left            =   2880
         Max             =   30000
         Min             =   5
         TabIndex        =   57
         Top             =   2250
         Value           =   3500
         Width           =   855
      End
      Begin VB.TextBox TUBES_TEMP_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "TUBES_TEMP_OUT"
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
         Height          =   225
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   56
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   2205
         Width           =   915
      End
      Begin VB.TextBox SHELL_TEMP_IN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "SHELL_TEMP_IN"
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
         Height          =   225
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   73
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   1980
         Width           =   915
      End
      Begin VB.HScrollBar HScroll_SHELL_T_IN 
         Height          =   195
         LargeChange     =   100
         Left            =   4740
         Max             =   30000
         Min             =   5
         TabIndex        =   74
         Top             =   2025
         Value           =   2500
         Width           =   855
      End
      Begin VB.HScrollBar Spin_TUBES_T_IN 
         Height          =   195
         LargeChange     =   100
         Left            =   2880
         Max             =   30000
         Min             =   5
         TabIndex        =   55
         Top             =   2025
         Value           =   2500
         Width           =   855
      End
      Begin VB.TextBox TUBES_TEMP_IN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "TUBES_TEMP_IN"
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
         Height          =   225
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   54
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   1980
         Width           =   915
      End
      Begin VB.TextBox TUBES_NON_COND 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "TUBES_NON_COND"
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
         Height          =   225
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   52
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   1740
         Width           =   915
      End
      Begin VB.HScrollBar HScroll_TUBES_NON_COND 
         Height          =   195
         LargeChange     =   10
         Left            =   2880
         Max             =   30000
         TabIndex        =   53
         Top             =   1800
         Width           =   855
      End
      Begin VB.HScrollBar HScroll_SHELL_NON_COND 
         Height          =   195
         LargeChange     =   10
         Left            =   4740
         Max             =   30000
         TabIndex        =   72
         Top             =   1815
         Width           =   855
      End
      Begin VB.TextBox SHELL_NON_COND 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "SHELL_NON_COND"
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
         Height          =   225
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   71
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox TUBES_WATER 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "TUBES_WATER"
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
         Height          =   225
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   50
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   1500
         Width           =   915
      End
      Begin VB.HScrollBar HScroll_TUBES_WATER 
         Height          =   195
         LargeChange     =   10
         Left            =   2880
         Max             =   30000
         TabIndex        =   51
         Top             =   1560
         Width           =   855
      End
      Begin VB.HScrollBar HScroll_SHELL_WATER 
         Height          =   195
         LargeChange     =   10
         Left            =   4740
         Max             =   30000
         TabIndex        =   70
         Top             =   1575
         Width           =   855
      End
      Begin VB.TextBox SHELL_WATER 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "SHELL_WATER"
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
         Height          =   225
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   69
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   1500
         Width           =   915
      End
      Begin VB.TextBox TUBES_LIQUID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "TUBES_LIQUID"
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
         Height          =   225
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   48
         ToolTipText     =   "Includes water."
         Top             =   1260
         Width           =   915
      End
      Begin VB.HScrollBar HScroll_TUBES_LIQUID 
         Height          =   195
         LargeChange     =   10
         Left            =   2880
         Max             =   30000
         TabIndex        =   49
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.HScrollBar HScroll_SHELL_LIQUID 
         Height          =   195
         LargeChange     =   10
         Left            =   4740
         Max             =   30000
         TabIndex        =   68
         Top             =   1335
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox SHELL_LIQUID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "SHELL_LIQUID"
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
         Height          =   225
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   67
         ToolTipText     =   "Includes water."
         Top             =   1260
         Width           =   915
      End
      Begin VB.TextBox TUBES_VAPOR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "TUBES_VAPOR"
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
         Height          =   225
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   46
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   1020
         Width           =   915
      End
      Begin VB.HScrollBar HScroll_TUBES_VAPOR 
         Height          =   195
         LargeChange     =   10
         Left            =   2880
         Max             =   30000
         TabIndex        =   47
         Top             =   1080
         Width           =   855
      End
      Begin VB.HScrollBar HScroll_SHELL_VAPOR 
         Height          =   195
         LargeChange     =   10
         Left            =   4740
         Max             =   30000
         TabIndex        =   66
         Top             =   1095
         Width           =   855
      End
      Begin VB.TextBox SHELL_VAPOR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "SHELL_VAPOR"
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
         Height          =   225
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   65
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   1020
         Width           =   915
      End
      Begin VB.TextBox SHELL_FLOW 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "SHELL_FLOW"
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
         Height          =   225
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   63
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   780
         Width           =   915
      End
      Begin VB.HScrollBar Spin_SHELL_FLOW 
         Height          =   195
         LargeChange     =   10
         Left            =   4740
         Max             =   30000
         TabIndex        =   64
         Top             =   840
         Width           =   855
      End
      Begin VB.HScrollBar HScroll_TUBES_FLOW 
         Height          =   195
         LargeChange     =   10
         Left            =   2880
         Max             =   30000
         TabIndex        =   45
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox TUBES_FLOW 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "TUBES_FLOW"
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
         Height          =   225
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   44
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   780
         Width           =   915
      End
      Begin VB.ComboBox Combo_T_FLUID 
         BackColor       =   &H80000018&
         DataField       =   "TUBES_FLUID"
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
         ItemData        =   "frmHX.frx":0C1A
         Left            =   1740
         List            =   "frmHX.frx":0C5D
         TabIndex        =   43
         ToolTipText     =   "Select from the list.Use TAB Key to save the new enter"
         Top             =   420
         Width           =   1995
      End
      Begin VB.ComboBox Combo_S_FLUID 
         BackColor       =   &H80000018&
         DataField       =   "SHELL_FLUID"
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
         ItemData        =   "frmHX.frx":0DC1
         Left            =   3840
         List            =   "frmHX.frx":0E16
         TabIndex        =   62
         ToolTipText     =   "Select from the list.Use TAB Key to save the new enter"
         Top             =   420
         Width           =   2055
      End
      Begin VB.Label Label17 
         Caption         =   "Span flow"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5940
         TabIndex        =   322
         Top             =   180
         Width           =   795
      End
      Begin MSForms.SpinButton Spin_TARGET_T 
         Height          =   255
         Left            =   6540
         TabIndex        =   320
         ToolTipText     =   "Target T out"
         Top             =   2220
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
         Left            =   6000
         TabIndex        =   317
         Top             =   2040
         Width           =   735
      End
      Begin MSForms.SpinButton Spin_FACT_FLOW 
         Height          =   255
         Left            =   6480
         TabIndex        =   42
         ToolTipText     =   "Span factor of flow"
         Top             =   360
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
         Left            =   180
         TabIndex        =   275
         Top             =   3240
         Width           =   2010
      End
      Begin VB.Label lblLabels 
         Caption         =   "Condensing pressure:"
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
         Index           =   10
         Left            =   180
         TabIndex        =   274
         Top             =   2937
         Width           =   1545
      End
      Begin VB.Label lblLabels 
         Caption         =   "Temperature OUT:"
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
         Index           =   8
         Left            =   180
         TabIndex        =   273
         Top             =   2223
         Width           =   1365
      End
      Begin VB.Label lblLabels 
         Caption         =   "Temperature IN:"
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
         Index           =   7
         Left            =   180
         TabIndex        =   272
         Top             =   1985
         Width           =   1185
      End
      Begin VB.Label lblLabels 
         Caption         =   "Total INLET flow rate:"
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
         Index           =   6
         Left            =   180
         TabIndex        =   271
         Top             =   795
         Width           =   1665
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
         Height          =   195
         Index           =   38
         Left            =   180
         TabIndex        =   270
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label lblLabels 
         Caption         =   "Non-cond. (out):"
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
         Index           =   67
         Left            =   600
         TabIndex        =   269
         Top             =   1740
         Width           =   1230
      End
      Begin VB.Label Label22 
         Caption         =   "Pressure IN (design):"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   268
         Top             =   2460
         Width           =   1635
      End
      Begin VB.Label Label22 
         Caption         =   "Pressure OUT (design):"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   28
         Left            =   180
         TabIndex        =   267
         Top             =   2700
         Width           =   1755
      End
      Begin VB.Label Label12 
         Caption         =   "[(hm^2 ºC)/kcal]*10^-4"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   2160
         TabIndex        =   265
         Top             =   3240
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "kPa(a)"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   10
         Left            =   5640
         TabIndex        =   155
         Top             =   3000
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "bar"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5640
         TabIndex        =   259
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "bar"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5640
         TabIndex        =   154
         Top             =   2520
         Width           =   315
      End
      Begin VB.Label Label9 
         Caption         =   "°C"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5640
         TabIndex        =   153
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "°C"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5640
         TabIndex        =   152
         Top             =   2040
         Width           =   315
      End
      Begin VB.Label Label51 
         Caption         =   "kg/h"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5640
         TabIndex        =   212
         Top             =   1800
         Width           =   315
      End
      Begin VB.Label Label52 
         Caption         =   "kg/h"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5640
         TabIndex        =   214
         Top             =   1560
         Width           =   315
      End
      Begin VB.Label lblLabels 
         Caption         =   "Water (out):"
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
         Index           =   68
         Left            =   900
         TabIndex        =   213
         Top             =   1515
         Width           =   870
      End
      Begin VB.Label Label50 
         Caption         =   "kg/h"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5640
         TabIndex        =   211
         Top             =   1320
         Width           =   315
      End
      Begin VB.Label lblLabels 
         Caption         =   "Liquid (out):"
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
         Index           =   66
         Left            =   900
         TabIndex        =   210
         Top             =   1275
         Width           =   870
      End
      Begin VB.Label Label49 
         Caption         =   "kg/h"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5640
         TabIndex        =   209
         Top             =   1080
         Width           =   315
      End
      Begin VB.Label lblLabels 
         Caption         =   "Vapor (out):"
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
         Index           =   9
         Left            =   900
         TabIndex        =   145
         Top             =   1035
         Width           =   870
      End
      Begin VB.Label Label6 
         Caption         =   "kg/h"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5640
         TabIndex        =   151
         Top             =   840
         Width           =   315
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
         Left            =   4320
         TabIndex        =   208
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
         Left            =   2400
         TabIndex        =   207
         Top             =   180
         Width           =   765
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Height          =   330
      Left            =   120
      TabIndex        =   179
      Top             =   9240
      Visible         =   0   'False
      Width           =   14715
      _ExtentX        =   25956
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20770
            Text            =   "Stato"
            TextSave        =   "Stato"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "02/01/2016"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "22.19"
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
   Begin MSForms.ToggleButton Toggle_remarks 
      Height          =   375
      Left            =   60
      TabIndex        =   311
      Top             =   7440
      Width           =   1215
      BackColor       =   12648384
      ForeColor       =   128
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
Public A40, D11, D13, D15, D17, D18, D19, D20, D21, D29, D38, D40, D43, D44, D54, D67, D68, D69, D70, D71, D72, D73, D74, D75, D76, D77, D78, D80 As Double
Public T_FLOW, S_FLOWAs As Double
Public K_PF, K_P_DROP_T, K_P_DROP_S, K_S_SPH, K_S_DENS, K_S_VISC, K_T_SPH, K_T_DENS, K_T_VISC, K_U_CLEAN
Public XD21, XD22, XD23, XD24, TH_C As Double
Public XPI, LN, XD6, XD7, XD8, XD9, XD10, XD18, XD18L, XD19, XD20, XD9_S
Public XD37, XD52M, XD52, XD54, XD55, XD56, XD57, XD58, XD59, XD61M
Public XD61, XD63, XD64M, XD64, XD66M, XD66, XD50, XD85, XD84, XD83, XD112
Private Sub Form_Load()
On Error Resume Next
    Width = frmMain.Width * 0.98 ' Imposta la larghezza del form.
    Height = frmMain.Height * 0.9      ' Imposta l'altezza del form.
    Left = 50 '(frmMain.Width - Width) / 2 ' Centra il form orizzontalmente.
    Top = 0 '(frmMain.Height - Height) / 2 ' Centra il form verticalmente.

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
    Data4.RecordSource = "Select * From [Query_Unit]"
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
    
    Data7.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
    Data7.RecordSource = "Select * From [QUERY_Date]"
    Data7.Refresh
    Set Rs7 = Data7.Recordset
    If Rs7.RecordCount > 0 Then
       Do Until Rs7.EOF
          Date_X = Data7.Recordset.date_test
          Combo_Date_X.AddItem Date_X
          Rs7.MoveNext
       Loop
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
    Data8.RecordSource = "Select * From [Query_Plant_UNIT]"
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
    Data11.RecordSource = "Select * From [Query_TOWER]"
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
    Check_WET_STEAM.Value = Data1.Recordset.Check_WET_STEAM
    Check_U = Data1.Recordset.Check_X
    foul = 1
    
With Grid1
    .COL = 2
    .Row = 7
    .ColWidth(0) = 1000
    .ColWidth(1) = 1100
    .DataSource = Rs12
    .AddItem (Unit_name), (date_test)
    End With
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
        TEST_1 = Val(TEST_Y)
    End If
    Do Until Data1.Recordset.EOF
        n_rec_a = Data1.Recordset.AbsolutePosition + 1
        Date_2 = Data1.Recordset.date_test
        Unit_2 = Data1.Recordset.Unit_name
        TEST_2 = Data1.Recordset.TEST_NO
        If Grid1.ColSel = 0 Then
            If Unit_1 = Unit_2 Then
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
            Bar1 = 10
        End If
    Loop
End If
10  XXX = 0
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
            Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
    Bar1 = Bar1 + Bar1
    
    Data1.UpdateRecord
    Data1.Recordset.Bookmark = Data1.Recordset.LastModified
    Bar1.Visible = False
End Sub
Private Sub Data1_Error(DataErr As Integer, Response As Integer)
On Error Resume Next
  ER = DataErr
  MsgBox "Intercettato errore dei dati:" & Error$(DataErr)
  Response = 0  'Ignora l'errore
End Sub
Private Sub Data1_Reposition()
On Error Resume Next
  Screen.MousePointer = vbDefault
  Checkrec = Data1.Recordset.AbsolutePosition + 1
  Data1.Caption = Data1.Recordset("Unit_Name")
'    Dim cn As New ADODB.Connection
'    cn.Open "Provider=Microsoft Jet 6.0 OLE DB Provider;Data source=C:\Program Files\WaterProof HX\HX.mdb;Jet & OLEDB:Database Password=gifra"
'    cn.Provider = "Microsoft Jet 4.0 OLE DB Provider"
'    cn.ConnectionString = "Data source=C:\Program Files\WaterProof HX\HX.mdb"
'    cn.Properties("Jet OLEDB: :Database Password") = "gifra"
'    cn.Open
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
    ELEVATION = Data1.Recordset.ELEVATION
    PARALLEL_N = Data1.Recordset.PARALLEL_N
    SERIES_N = Data1.Recordset.SERIES_N
    Combo_CURRENT.Text = Data1.Recordset.CURRENT
    
    T_NO.Text = Val(Data1.Recordset.TUBES_NO)
    T_len.Text = Val(Data1.Recordset.TUBES_LE)
    T_PASS.Text = Val(Data1.Recordset.TUBES_PASSES)
    T_OD.Text = Val(Data1.Recordset.TUBES_OD)
    Combo_BWG.Text = Val(Data1.Recordset.TUBES_BWG)
    Combo_TUBES_Mat.Text = Data1.Recordset.TUBES_MAT
    D54 = Data1.Recordset.TUBES_Mat_fact
    Mat_factor.Text = D54
    U = Data1.Recordset.Check_X
    If U = 0 Then
        lungh = 1
        Check_U = Unchecked
    ElseIf U = -1 Then
        lungh = 2
        Check_U = Checked
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
    TUBES_FLOW.Text = Data1.Recordset.TUBES_FLOW
    TUBES_VAPOR.Text = Data1.Recordset.TUBES_VAPOR
    TUBES_LIQUID.Text = Data1.Recordset.TUBES_LIQUID
    TUBES_WATER.Text = Data1.Recordset.TUBES_WATER
    TUBES_NON_COND.Text = Data1.Recordset.TUBES_NON_COND
    
    TUBES_TEMP_IN.Text = Data1.Recordset.TUBES_TEMP_IN
    TUBES_TEMP_OUT.Text = Data1.Recordset.TUBES_TEMP_OUT
    TUBES_P_IN.Text = Data1.Recordset.TUBES_P_IN
    TUBES_P_OUT.Text = Data1.Recordset.TUBES_P_OUT
    
    TUBES_OUT(1).Text = Data1.Recordset.TUBES_T_COND
    TUBES_OUT(3).Text = Data1.Recordset.TUBES_SPH
    TUBES_OUT(4).Text = Data1.Recordset.TUBES_DENS
    TUBES_OUT(5).Text = Data1.Recordset.TUBES_VISC
    
    Combo_S_FLUID.Text = Data1.Recordset.SHELL_FLUID
    SHELL_FLOW.Text = Data1.Recordset.SHELL_FLOW
    SHELL_VAPOR.Text = Data1.Recordset.SHELL_VAPOR
    SHELL_LIQUID.Text = Data1.Recordset.SHELL_LIQUID
    SHELL_WATER.Text = Data1.Recordset.SHELL_WATER
    SHELL_NON_COND.Text = Data1.Recordset.SHELL_NON_COND
    
    SHELL_TEMP_IN.Text = Data1.Recordset.SHELL_TEMP_IN
    SHELL_TEMP_OUT.Text = Data1.Recordset.SHELL_TEMP_OUT
    SHELL_P_IN.Text = Data1.Recordset.SHELL_P_IN
    SHELL_P_OUT.Text = Data1.Recordset.SHELL_P_OUT
    S_press_KP.Text = Data1.Recordset.Press_COND
    
    SHELL_OUT(1) = Data1.Recordset.SHELL_T_COND
    SHELL_OUT(3) = Data1.Recordset.SHELL_SPH
    SHELL_OUT(4) = Data1.Recordset.SHELL_DENS
    SHELL_OUT(5) = Data1.Recordset.SHELL_VISC
    SHELL_OUT(11) = Data1.Recordset.Temp_COND
    If Combo_S_FLUID.Text = "Steam" Then
        SHELL_OUT(10) = Data1.Recordset.SHELL_LATENT
    End If
YXY = 1
    Spin_BAFFLES_N.Value = SHELL_BAFFLES_N
    
    TUBES_T_IN = Val(TUBES_TEMP_IN)
    TUBES_T_OUT = Val(TUBES_TEMP_OUT)
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
    
    TUBES_TC = Val(TUBES_OUT(1))
    HScroll_TUBES_TC = TUBES_TC * 1000
    
    HScroll_SHELL_VAPOR.Max = SHELL_FLOW / FACTOR
    HScroll_SHELL_LIQUID.Max = SHELL_FLOW / FACTOR
    HScroll_SHELL_WATER.Max = SHELL_FLOW / FACTOR
    HScroll_SHELL_NON_COND.Max = SHELL_FLOW / FACTOR
    
    Spin_SHELL_FLOW.Value = SHELL_FLOW / FACTOR
    HScroll_SHELL_VAPOR.Value = SHELL_VAPOR / FACTOR
    HScroll_SHELL_LIQUID.Value = SHELL_LIQUID / FACTOR
    HScroll_SHELL_WATER.Value = SHELL_WATER / FACTOR
    HScroll_SHELL_NON_COND.Value = SHELL_NON_COND / FACTOR
    
    VAP_P = Data1.Recordset.VAP_FRACTION
    VAP_PERC = Format(VAP_P, "0.0")
    LIQ_PERC = Format(100 - VAP_PERC, "0.0")
    Spin_VAP_P.Value = VAP_PERC * 10
    
    Spin_ELEVATION.Value = ELEVATION
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
    
    HScroll_TUBES_P_IN.Value = TUBES_P_IN * 1000
    HScroll_TUBES_P_OUT.Value = TUBES_P_OUT * 1000

    HScroll_SHELL_SPH = SHELL_OUT(3) * 1000
    HScroll_SHELL_DENS = SHELL_OUT(4) * 10
    HScroll_SHELL_VISC = SHELL_OUT(5) * 1000
    HScroll_SHELL_T_IN.Value = SHELL_TEMP_IN * 100
    HScroll_SHELL_T_OUT.Value = SHELL_TEMP_OUT * 100
    HScroll_SHELL_P_IN.Value = SHELL_P_IN * 1000
    HScroll_SHELL_P_OUT.Value = SHELL_P_OUT * 1000
    S_press = Val(S_press_KP) * 10
    Spin_S_PRESS.Value = S_press '* 10
    
    HScroll_SHELL_TC = SHELL_OUT(1) * 1000
    SHELL_OUT(9).Text = Data1.Recordset.SHELL_PRESS_DROP
    HScroll_P_DROP_S = SHELL_OUT(9) * 10
    HScroll_C_TEMP = SHELL_OUT(11) * 100
    HScroll_LATENT.Value = Val(SHELL_OUT(10))
    
    PROCESS_TARGET_T_OUT = Data1.Recordset.PROCESS_TARGET_TEMP
    Spin_TARGET_T = PROCESS_TARGET_T_OUT * 10
    
    Wet_steam = Data1.Recordset.Wet_steam
    HScroll_WET_STEAM = Wet_steam * 10
    Check_WET_STEAM.Value = Data1.Recordset.Check_WET_STEAM


    XXX = 1
    
    FFX = Data1.Recordset.SHELL_FF
    SHELL_FF = Format(FFX, "0.00")
    D40 = FFX / 10000
    W_FF = Data1.Recordset.WATER_FF
    WATER_FF = Format(W_FF, "0.00")
    For i = 1 To 5
        If Combo_T_FLUID = "Water" Then
            If i = 2 Then i = 3
            TUBES_OUT(i).ForeColor = &HFF0000
            TUBES_OUT(i).BackColor = &HE0E0E0
        Else
            TUBES_OUT(i).ForeColor = &HC0&
            TUBES_OUT(i).BackColor = &HE0E0E0
        End If
    Next i
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Or Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam" Or Combo_S_FLUID.Text = "Water" Or Combo_S_FLUID = "Steam condensing" Then
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
    If Combo_S_FLUID.Text = "Water" Or Combo_S_FLUID = "Steam condensing" Then
        SHELL_OUT(3).ForeColor = &HFF0000
        SHELL_OUT(3).BackColor = &HE0E0E0
    Else
        SHELL_OUT(3).ForeColor = &HC0&
        SHELL_OUT(3).BackColor = &HE0E0E0
    End If
    
    If Combo_T_FLUID <> "Water" Then
        'Termal conductivity SHELL
            TUBES_OUT(1).Visible = True
            Check_T_TC.Visible = True
            HScroll_TUBES_TC.Visible = True
            lbl_tubes(1).Visible = True
    ElseIf Combo_T_FLUID = "Water" Then
        'Termal conductivity SHELL
            TUBES_OUT(1).Visible = False
            Check_T_TC.Visible = False
            HScroll_TUBES_TC.Visible = False
            lbl_tubes(1).Visible = False
        'SPECIFIC HEAT
            Check_S_SPH.Visible = False
            SHELL_OUT(3).Visible = False
            HScroll_SHELL_SPH.Visible = False
    End If
    
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Lab_COND = "CONDENSATION"
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Lab_COND = "CONDENSATION"
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Lab_COND = "Steam Exhaust Condenser"
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Lab_COND = "COOLER"
        Call COOLERS
    End If
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
  XXX = 1
End Sub
Private Sub Check_des_Click()
On Error Resume Next
    If Check_des = Checked Then
        CHECK_DESIGN.BackColor = &HFF&
        Label4.Visible = False
        CHECK_ACTUAL.Visible = False
        Check_ACT.Visible = False
        Line2.Visible = False
        Command1.Visible = False
        Check_T_OUT.Visible = False
    
    ElseIf Check_des = Unchecked Then
        CHECK_DESIGN.BackColor = &HE0E0E0
        Label4.Visible = True
        CHECK_ACTUAL.Visible = True
        Check_ACT.Visible = True
        Line2.Visible = True
        Command1.Visible = True
    End If
End Sub
Private Sub Check_act_Click()
On Error Resume Next
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
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Command1_Click()
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
    If Combo_CURRENT.Text = "Condensation" Then
        Lab_COND.Caption = "CONDENSATION"
        If Combo_S_FLUID = "Water" Then
            MsgBox ("Tubes-side condensation cannot be set")
            Combo_CURRENT = "Counter-flow"
            Exit Sub
            Lab_COND.Caption = "COOLER"
        End If
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub HScroll_LATENT_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If

    If Check_LATENT = Checked Then
        SHELL_OUT(10) = HScroll_LATENT
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Check_LATENT_Click()
On Error Resume Next
If XXX = 1 Then
    Exit Sub
End If
    If Check_LATENT = Checked Then
        SHELL_OUT(10) = HScroll_LATENT
    End If

    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub HScroll_WATER_FF_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If

    W_FF = HScroll_WATER_FF / 100
    WATER_FF = Format(W_FF, "0.00")
End Sub
Private Sub HScroll_WET_STEAM_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If

    If Check_WET_STEAM = Checked Then
        Wet_steam = HScroll_WET_STEAM / 10
    Else
        D33 = Data1.Recordset.Wet_steam 'Exit Sub
        HScroll_WET_STEAM = Wet_steam * 10
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Check_WET_STEAM_Click()
On Error Resume Next
    If Check_WET_STEAM = Checked Then
        Wet_steam = HScroll_WET_STEAM / 10
    Else
        D33 = Data1.Recordset.Wet_steam
        HScroll_WET_STEAM = Wet_steam * 10
    End If
    
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
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
    TUBES_FLOW = T_FLOW * FACT_FLOW
    TUBES_OUT(0) = Format(TUBES_FLOW, "##,##0")
    
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
    
    TUBES_LIQUID = TUBES_FLOW - TUBES_VAPOR - TUBES_NON_COND
    T_LIQUID = TUBES_LIQUID / FACT_FLOW
    HScroll_TUBES_LIQUID.Value = T_LIQUID
    
    HScroll_TUBES_WATER.Max = T_LIQUID
    If Combo_T_FLUID = "Water" Then
        TUBES_WATER = TUBES_LIQUID
        HScroll_TUBES_WATER = T_LIQUID
    End If
    
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
'    If COL = 0 Then
'        T_FLW = 0
'        TUBES_OUT(0).BackColor = &HE0E0E0
'        TUBES_OUT(0).ForeColor = &HC0&
'    End If
End Sub
Private Sub HScroll_TUBES_VAPOR_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If

    T_VAPOR = HScroll_TUBES_VAPOR
    TUBES_VAPOR = T_VAPOR * FACT_FLOW
    
    TUBES_LIQUID = TUBES_FLOW - TUBES_VAPOR - TUBES_NON_COND
    T_LIQUID = TUBES_LIQUID / FACT_FLOW
    HScroll_TUBES_LIQUID.Value = T_LIQUID
    If Combo_T_FLUID = "Water" Then
        TUBES_WATER = TUBES_LIQUID
        HScroll_TUBES_WATER = TUBES_WATER / FACT_FLOW
    End If
    
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

    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub HScroll_TUBES_NON_COND_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If

    T_NON_COND = HScroll_TUBES_NON_COND
    TUBES_NON_COND = T_NON_COND * FACT_FLOW
    
    TUBES_LIQUID = TUBES_FLOW - TUBES_VAPOR - TUBES_NON_COND
    T_LIQUID = TUBES_LIQUID / FACT_FLOW
    HScroll_TUBES_LIQUID.Value = T_LIQUID
    HScroll_TUBES_WATER.Max = T_LIQUID
    If Combo_T_FLUID = "Water" Then
        TUBES_WATER = TUBES_LIQUID
        HScroll_TUBES_WATER = TUBES_WATER / FACT_FLOW
    End If
    If COL = 0 Then
        T_FLOW = 0
        TUBES_OUT(0).BackColor = &HE0E0E0
        TUBES_OUT(0).ForeColor = &HC0&
    End If

    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
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
    TUBES_WATER = T_WATER * FACT_FLOW
End Sub
Private Sub Spin_SHELL_FLOW_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If

    S_FLOW = Spin_SHELL_FLOW
    SHELL_FLOW = S_FLOW * FACT_FLOW
    SHELL_OUT(0) = Format(SHELL_FLOW, "##,##0")
    
    HScroll_SHELL_VAPOR.Max = SHELL_FLOW / FACT_FLOW
    HScroll_SHELL_LIQUID.Max = SHELL_FLOW / FACT_FLOW
    HScroll_SHELL_WATER.Max = SHELL_FLOW / FACT_FLOW
    HScroll_SHELL_NON_COND.Max = SHELL_FLOW / FACT_FLOW
    
    SHELL_VAPOR = Data1.Recordset.SHELL_VAPOR
    SHELL_LIQUID = Data1.Recordset.SHELL_LIQUID
    SHELL_WATER = Data1.Recordset.SHELL_WATER
    SHELL_NON_COND = Data1.Recordset.SHELL_NON_COND
    
    HScroll_SHELL_VAPOR.Value = SHELL_VAPOR / FACT_FLOW
    HScroll_SHELL_LIQUID.Value = SHELL_LIQUID / FACT_FLOW
    HScroll_SHELL_WATER.Value = SHELL_WATER / FACT_FLOW
    HScroll_SHELL_NON_COND.Value = SHELL_NON_COND / FACT_FLOW
    
    SHELL_LIQUID = SHELL_FLOW - SHELL_VAPOR - SHELL_NON_COND
    S_LIQUID = SHELL_LIQUID / FACT_FLOW
    HScroll_SHELL_LIQUID.Value = S_LIQUID
    
    If Combo_S_FLUID = "Water" Then
        SHELL_WATER = SHELL_LIQUID
        HScroll_SHELL_WATER = S_LIQUID
    End If
    
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
    AA = XD37 / 860
    If COL = 0 Then
        S_FLOW = 0
        SHELL_OUT(0).BackColor = &HE0E0E0
        SHELL_OUT(0).ForeColor = &HC0&
    End If
End Sub
Private Sub HScroll_SHELL_VAPOR_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If

    S_VAPOR = HScroll_SHELL_VAPOR.Value
    SHELL_VAPOR = S_VAPOR * FACT_FLOW
    
    SHELL_LIQUID = SHELL_FLOW - SHELL_VAPOR - SHELL_NON_COND
    HScroll_SHELL_LIQUID = SHELL_LIQUID / FACT_FLOW
    S_LIQUID = SHELL_LIQUID / FACT_FLOW
    HScroll_SHELL_LIQUID.Value = S_LIQUID
        
    If Check_VAP_P = Checked Then
        VAP_P = Spin_VAP_P / 10
        VAP_OUT = SHELL_VAPOR / SHELL_FLOW * 100
        VAP_IN = VAP_P - VAP_OUT
        VAP_PERC = Format(VAP_IN, "0.0")
        LIQ_PERC = Format(100 - VAP_P, "0.0")
    ElseIf Check_VAP_P = Unchecked Then
        VAP_P = Spin_VAP_P / 10
        VAP_PERC = Format(VAP_P, "0.0")
        LIQ_PERC = Format(100 - VAP_P, "0.0")
    End If


    If Combo_S_FLUID = "Water" Then
        SHELL_WATER = SHELL_LIQUID
        HScroll_SHELL_WATER = SHELL_WATER / FACT_FLOW
    End If
    
    If COL = 0 Then
        S_FLOW = 0
        SHELL_OUT(0).BackColor = &HE0E0E0
        SHELL_OUT(0).ForeColor = &HC0&
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
    If COL = 0 Then
        S_FLOW = 0
        SHELL_OUT(0).BackColor = &HE0E0E0
        SHELL_OUT(0).ForeColor = &HC0&
    End If
End Sub
Private Sub Spin_VAP_P_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    If Check_VAP_P = Checked Then
        VAP_P = Spin_VAP_P / 10
        VAP_OUT = SHELL_VAPOR / SHELL_FLOW * 100
        VAP_IN = VAP_P - VAP_OUT
        VAP_PERC = Format(VAP_IN, "0.0")
        LIQ_PERC = Format(100 - VAP_P, "0.0")
    ElseIf Check_VAP_P = Unchecked Then
        VAP_P = Spin_VAP_P / 10
        VAP_PERC = Format(VAP_P, "0.0")
        LIQ_PERC = Format(100 - VAP_P, "0.0")
    End If

    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Check_VAP_P_Click()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    If Check_VAP_P = Checked Then
        VAP_P = Spin_VAP_P / 10
        VAP_OUT = SHELL_VAPOR / SHELL_FLOW * 100
        VAP_IN = VAP_P - VAP_OUT
        VAP_PERC = Format(VAP_IN, "0.0")
        LIQ_PERC = Format(100 - VAP_P, "0.0")
    ElseIf Check_VAP_P = Unchecked Then
        VAP_P = Spin_VAP_P / 10
        VAP_PERC = Format(VAP_P, "0.0")
        LIQ_PERC = Format(100 - VAP_P, "0.0")
    End If
    
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
    If COL = 0 Then
        S_FLOW = 0
        SHELL_OUT(0).BackColor = &HE0E0E0
        SHELL_OUT(0).ForeColor = &HC0&
    End If
End Sub
Private Sub HScroll_SHELL_NON_COND_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    S_NON_COND = HScroll_SHELL_NON_COND
    SHELL_NON_COND = S_NON_COND * FACT_FLOW
    SHELL_LIQUID = SHELL_FLOW - SHELL_VAPOR - SHELL_NON_COND
    S_LIQUID = SHELL_LIQUID / FACT_FLOW
    HScroll_SHELL_LIQUID.Value = S_LIQUID
    
    If Combo_S_FLUID = "Water" Then
        SHELL_WATER = SHELL_LIQUID
        HScroll_SHELL_WATER = SHELL_WATER / FACT_FLOW
    End If
    If COL = 0 Then
        S_FLOW = 0
        SHELL_OUT(0).BackColor = &HE0E0E0
        SHELL_OUT(0).ForeColor = &HC0&
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
    If COL = 0 Then
        S_FLOW = 0
        SHELL_OUT(0).BackColor = &HE0E0E0
        SHELL_OUT(0).ForeColor = &HC0&
    End If
End Sub
Private Sub HScroll_SHELL_WATER_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    S_WATER = HScroll_SHELL_WATER
    SHELL_WATER = S_WATER * FACT_FLOW
End Sub
Private Sub SHELL_PITCH_CONF_LostFocus()
On Error Resume Next
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Spin_TUBES_T_IN_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    TUBES_TEMP_IN = Spin_TUBES_T_IN / 100
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Spin_TUBES_T_OUT_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    TUBES_TEMP_OUT = Spin_TUBES_T_OUT / 100
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub HScroll_TUBES_P_IN_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    TUBES_P_IN = HScroll_TUBES_P_IN / 1000
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub HScroll_TUBES_P_OUT_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    TUBES_P_OUT = HScroll_TUBES_P_OUT / 1000
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub HScroll_SHELL_T_IN_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    SHELL_TEMP_IN = HScroll_SHELL_T_IN / 100
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub HScroll_SHELL_T_OUT_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    SHELL_TEMP_OUT = HScroll_SHELL_T_OUT / 100
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub HScroll_SHELL_P_IN_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    SHELL_P_IN = HScroll_SHELL_P_IN / 1000
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub HScroll_SHELL_P_OUT_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    SHELL_P_OUT = HScroll_SHELL_P_OUT / 1000
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Spin_TUBES_PITCH_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    SHELL_TUBES_PITCH = Spin_TUBES_PITCH / 10
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Spin_BAFFLES_CUT_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    SHELL_BAFFLES_CUT = Spin_BAFFLES_CUT
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Spin_BAFFLES_N_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    If CHECK_BAFFLES_N = Unchecked Then
        SHELL_BAFFLES_N = Int(XD55 / (XD64M / 1000))
        Spin_BAFFLES_N = SHELL_BAFFLES_N
    Else
        SHELL_BAFFLES_N = Spin_BAFFLES_N
    End If
    
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub CHECK_BAFFLES_N_Click()
On Error Resume Next
If XXX = 1 Then
    Exit Sub
End If
    If CHECK_BAFFLES_N = Unchecked Then
        SHELL_BAFFLES_N = Int(XD55 / (XD64M / 1000))
        Spin_BAFFLES_N = SHELL_BAFFLES_N
    Else
        SHELL_BAFFLES_N = Spin_BAFFLES_N
    End If
    
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
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

    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Spin_ELEVATION_Change()
If YXY = 1 Then
    Exit Sub
End If
    ELEVATION.Text = Spin_ELEVATION
End Sub
Private Sub Spin_PARALLEL_N_Change()
On Error Resume Next
    PARALLEL_N.Text = Spin_PARALLEL_N
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Spin_S_PASS_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    SHELL_PASS = Spin_S_PASS
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Spin_SERIES_N_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    SERIES_N.Text = Spin_SERIES_N
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub HScroll_TUBES_SPH_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    If Check_T_SPH = Checked Then
        TUBES_OUT(3) = HScroll_TUBES_SPH / 1000
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub HScroll_TUBES_TC_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    If Check_T_TC = Checked Then
        TUBES_OUT(1) = HScroll_TUBES_TC / 1000
        If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
            Call CONDENSER
        ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
            Call CONDENSER
        ElseIf Combo_S_FLUID = "Steam" Then
            Call Steam
        ElseIf Combo_CURRENT = "Condensation" Then
            Call CONDENSER
        Else
            Call COOLERS
        End If
    End If
End Sub
Private Sub Check_T_TC_Click()
On Error Resume Next
    If Check_T_TC = Checked Then
        TUBES_OUT(1) = HScroll_TUBES_TC / 1000
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub HScroll_SHELL_TC_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    If Check_S_TC = Checked Then
        SHELL_OUT(1) = HScroll_SHELL_TC / 1000
        If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
            Call CONDENSER
        ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
            Call CONDENSER
        ElseIf Combo_S_FLUID = "Steam" Then
            Call Steam
        ElseIf Combo_CURRENT = "Condensation" Then
            Call CONDENSER
        Else
            Call COOLERS
        End If
    End If
End Sub
Private Sub Check_S_TC_Click()
On Error Resume Next
If XXX = 1 Then
    Exit Sub
End If
    If Check_S_TC = Checked Then
        SHELL_OUT(1) = HScroll_SHELL_TC / 1000
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Check_T_SPH_Click()
On Error Resume Next
    If Check_T_SPH = Checked Then
        TUBES_OUT(3) = HScroll_TUBES_SPH / 1000
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub HScroll_SHELL_SPH_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    If Check_S_SPH = Checked Then
        SHELL_OUT(3) = HScroll_SHELL_SPH / 1000
        If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
            Call CONDENSER
        ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
            Call CONDENSER
        ElseIf Combo_S_FLUID = "Steam" Then
            Call Steam
        ElseIf Combo_CURRENT = "Condensation" Then
            Call CONDENSER
        Else
            Call COOLERS
        End If
    End If
End Sub
Private Sub Check_S_SPH_Click()
On Error Resume Next
If XXX = 1 Then
    Exit Sub
End If
    If Check_S_SPH = Checked Then
        SHELL_OUT(3) = HScroll_SHELL_SPH / 1000
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub HScroll_TUBES_DENS_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    If Check_T_DENS = Checked Then
        TUBES_OUT(4) = Format(HScroll_TUBES_DENS / 10, "0.0")
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Check_T_DENS_Click()
On Error Resume Next
    If Check_T_DENS = Checked Then
        TUBES_OUT(4) = Format(HScroll_TUBES_DENS / 10, "0.0")
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub HScroll_SHELL_DENS_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    If Check_S_DENS = Checked Then
        SHELL_OUT(4) = HScroll_SHELL_DENS / 10
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Check_S_DENS_Click()
On Error Resume Next
If XXX = 1 Then
    Exit Sub
End If
    If Check_S_DENS = Checked Then
        SHELL_OUT(4) = HScroll_SHELL_DENS / 10
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub HScroll_TUBES_VISC_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    If Check_T_VISC = Checked Then
        TUBES_OUT(5) = HScroll_TUBES_VISC / 1000
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Check_T_VISC_Click()
On Error Resume Next
    If Check_T_VISC = Checked Then
        TUBES_OUT(5) = HScroll_TUBES_VISC / 1000
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub HScroll_SHELL_VISC_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    If Check_S_VISC = Checked Then
        SHELL_OUT(5) = HScroll_SHELL_VISC / 1000
        If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
            Call CONDENSER
        ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
            Call CONDENSER
        ElseIf Combo_S_FLUID = "Steam" Then
            Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
        Else
            Call COOLERS
        End If
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
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub HScroll_U_CLEAN_Change()
On Error Resume Next
If XXX = 1 Then
    Exit Sub
End If
    If Check_U_CLEAN = Checked Then
        U_COEFF_CLEAN = HScroll_U_CLEAN
    Else
        U_COEFF_CLEAN = Data1.Recordset.Clean
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Check_U_CLEAN_Click()
On Error Resume Next
If XXX = 1 Then
    Exit Sub
End If
    If Check_U_CLEAN = Checked Then
        U_COEFF_CLEAN = HScroll_U_CLEAN
    Else
        U_COEFF_CLEAN = Data1.Recordset.Clean
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
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
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
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
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Combo_S_FLUID_lostfocus()
On Error Resume Next
    If Combo_S_FLUID.Text = "Demineralized water" Or Combo_S_FLUID.Text = "Jacket water" Then
        Combo_S_FLUID.Text = "Water"
    End If
    If Combo_S_FLUID.Text <> "Water" Then
        Combo_T_FLUID.Text = "Water"
    End If
    If Combo_S_FLUID = "Water" And Combo_CURRENT = "Condensation" Then
        MsgBox ("    Tubes-side condensation cannot be set ")
        Combo_CURRENT = "Counter-flow"
        Lab_COND.Caption = "COOLER"
        Exit Sub
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Or Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam" Or Combo_S_FLUID.Text = "Water" Or Combo_S_FLUID = "Steam condensing" Then
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
    If Combo_S_FLUID.Text = "Water" Or Combo_S_FLUID = "Steam condensing" Then
        SHELL_OUT(3).ForeColor = &HFF0000
        SHELL_OUT(3).BackColor = &HE0E0E0
    Else
        SHELL_OUT(3).ForeColor = &HC0&
        SHELL_OUT(3).BackColor = &HE0E0E0
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Combo_T_FLUID_lostfocus()
On Error Resume Next
    If Combo_T_FLUID.Text <> "Water" Then
        Combo_S_FLUID.Text = "Water"
        Combo_CURRENT = "Counter-flow"
        Lab_COND = "COOLER"
    End If
    For i = 1 To 5
        If Combo_T_FLUID = "Water" Then
            If i = 2 Then i = 3
            TUBES_OUT(i).ForeColor = &HFF0000
            TUBES_OUT(i).BackColor = &HE0E0E0
        Else
            TUBES_OUT(i).ForeColor = &HC0&
            TUBES_OUT(i).BackColor = &HE0E0E0
        End If
    Next i
    If Combo_T_FLUID <> "Water" Then
        'Termal conductivity SHELL
            TUBES_OUT(1).Visible = True
            Check_T_TC.Visible = True
            HScroll_TUBES_TC.Visible = True
            lbl_tubes(1).Visible = True
    ElseIf Combo_T_FLUID = "Water" Then
        'Termal conductivity SHELL
            TUBES_OUT(1).Visible = False
            Check_T_TC.Visible = False
            HScroll_TUBES_TC.Visible = False
            lbl_tubes(1).Visible = False
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Combo_TUBES_Mat_LostFocus()
On Error Resume Next
    metal = Combo_TUBES_Mat.Text
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Check_PF_Click()
On Error Resume Next
If XXX = 1 Then
    Exit Sub
End If
    If Check_PF.Value = Checked Then
        FFX = Spin_PF.Value / 100
        SHELL_FF = Format(FFX, "0.00")
'        Spin_PF.Value = FFX * 100
    ElseIf Check_PF = Unchecked Then
        FFX = Data1.Recordset.SHELL_FF
        SHELL_FF = FFX
'        Spin_PF.Value = FFX * 100
    End If
    D40 = FFX / 10000
    
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Spin_PF_Change()
On Error Resume Next
    If Check_PF.Value = Checked Then
        FFX = Spin_PF.Value / 100
        SHELL_FF = Format(FFX, "0.00")
    ElseIf Check_PF = Unchecked Then
        FFX = Data1.Recordset.SHELL_FF
        SHELL_FF = FFX
        Spin_PF.Value = FFX * 100
    End If
    D40 = FFX / 10000
    
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Spin_S_FLOW_IN_Change()
On Error Resume Next
    SFIN_D = Spin_S_FLOW_IN
    S_flow_IN.Text = Format(SFIN_D * 100, "0,00")
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Spin_S_PRESS_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    If Check_CP = Checked Then
        S_press_KP = Format(Spin_S_PRESS / 10, "0.00")
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Check_CP_Click()
On Error Resume Next
If XXX = 1 Then
    Exit Sub
End If
    If Check_CP = Checked Then
        S_press_KP = Format(Spin_S_PRESS / 10, "0.00")
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub HScroll_C_TEMP_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    If Check_CT = Checked Then
        SHELL_OUT(11) = Format(HScroll_C_TEMP / 100, "0.00")
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Check_CT_Click()
On Error Resume Next
    If Check_CT = Checked Then
        SHELL_OUT(11) = Format(HScroll_C_TEMP / 100, "0.00")
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub HScroll_T_NO_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    T_NO.Text = Val(HScroll_T_NO.Value)
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Spin_SHELL_ID_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    SHELL_ID = Spin_SHELL_ID
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Spin_T_LEN_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    T_len.Text = Spin_T_LEN.Value / 100
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Spin_T_PAS_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    T_PASS.Text = Spin_T_PAS.Value
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Spin_T_OD_Change()
On Error Resume Next
If YXY = 1 Then
    Exit Sub
End If
    T_OD = Spin_T_OD / 100
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Spin_MAT_FACTOR_Change()
On Error Resume Next
If XXX = 1 Then
    Exit Sub
End If
    If Check_MAT_FACTOR = Checked Then
        Mat_factor = Spin_MAT_FACTOR / 100
    ElseIf Check_MAT_FACTOR = Unchecked Then
        Mat_factor = Data1.Recordset.TUBES_Mat_fact
    End If
    Call Steam
End Sub
Private Sub Check_MAT_FACTOR_Click()
On Error Resume Next
If XXX = 1 Then
    Exit Sub
End If
    If Check_MAT_FACTOR = Checked Then
        Mat_factor = Spin_MAT_FACTOR / 100
    ElseIf Check_MAT_FACTOR = Unchecked Then
        Mat_factor = Data1.Recordset.TUBES_Mat_fact
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
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Combo_BWG_LostFocus()
On Error Resume Next
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub HScroll_W_FLOW_Change()
On Error Resume Next
    WFIN = HScroll_W_FLOW
    W_flow_IN.Text = Format(Val(WFIN) * 10000, "0,00")
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Spin_W_T_IN_Change()
On Error Resume Next
    W_T_IN = Format(Spin_W_T_IN / 100, "0.00")
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Spin_W_T_OUT_Change()
On Error Resume Next
    W_T_OUT = Format(Spin_W_T_OUT / 100, "0.00")
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub CONDENSER()
On Error Resume Next

Dim XPI, LN As Double

If Combo_T_FLUID = "Water" Then
    Spin_TUBES_T_IN.Max = 10000
    Spin_TUBES_T_OUT.Max = 10000
Else
    Spin_TUBES_T_IN.Max = 30000
    Spin_TUBES_T_OUT.Max = 30000
End If

Lab_COND.Caption = "CONDENSATION"
Combo_CURRENT.Text = "Condensation"
lblLabels(3).Caption = "C Factor:"
Label24(0).Caption = "m3/h/kPa^(1/2) - Tubes-side"

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
    TUBES_OUT(1).Visible = False
    Check_T_TC.Visible = False
    HScroll_TUBES_TC.Visible = False
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
'Wet steam
    Label22(17).Visible = False
    Wet_steam.Visible = False
    HScroll_WET_STEAM.Visible = False
    Check_WET_STEAM.Visible = False
    Check_water_steam.Visible = False

Check_T_OUT.Visible = True

'Skin temperature
    Label22(3).Visible = True
    SKIN_TEMP.Visible = True
    Label22(4).Visible = True
' Flow calculated
    SHELL_OUT(0).Visible = True
' Flow velocity
    Label22(26).Visible = True
    SHELL_OUT(2).Visible = True
' Reynolds number
    SHELL_OUT(6).Visible = True
' Shell temperatures
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
lblLabels(25).Caption = "Terminal temperature:"
Thermal_bal_tubes.Visible = True
Thermal_bal_shell.Visible = True


Call Mechanical
    
    Mat_cond.Text = Val(D78)                                 'Thermal conductivity of tube material
    D79 = 3.14159 * D67 * 25.4 * 10 ^ -3 * D75 * D74 * lungh 'Heat transfer surface,m^2
    D80 = D79 / (0.3048 ^ 2)                                 'Heat transfer surface,inch^2
    Area.Text = Format(Val(D79), "0.0")

Call FOULING

XPI = 3.141592654
LN = 2.302585093
XD5 = Val(TUBES_LIQUID) / 1000   'TUBES TOTAL flow rate,m3/h
XD6 = Val(TUBES_FLOW)            'TUBES TOTAL flowrate,Kg/h
XD6L = Val(TUBES_LIQUID)         'TUBES LIQUID flowrate,Kg/h
XD7 = Val(TUBES_TEMP_IN)         'Water temperature in,ºC
XD8 = Val(TUBES_TEMP_OUT)        'Water temperature out,ºC
RANGE_T = Abs(XD8 - XD7)         'Tubes side delta T
XD9 = XD7 + (XD8 - XD7) / 2      'Caloric water temperature,ºC
XD10 = XD9 * 1.8 + 32            'Average water temperature,ºF

PROP = "TUBES"
Call Properties

XD5 = Val(TUBES_LIQUID) / D19    'Water flow rate,m3/h
XD11 = D19                       'Water density at ((t1+t2)/2),Kg/m3
XD12 = D20                       'Water viscosity at (t1+t2)/2,centipoise
XD13 = D21                       'Specific heat of water,Kcal/(Kg ºC)
XD13_S = SHELL_OUT(3)            'SHELL SIDE SPECIFIC HEAT
XD18 = Val(SHELL_FLOW)           'Total shell flowrate,Kg/h
XD19 = Val(SHELL_TEMP_IN)        'Shell temperature in,ºC
XD20 = Val(SHELL_TEMP_OUT)       'Shell temperature out,ºC
RANGE_S = Abs(XD19 - XD20)       'Shell Delta T
XD9_S = XD20 + (XD19 - XD20) / 2 'Caloric shell temperature
XD52M = Val(SHELL_TUBES_PITCH)   'Pitch, mm
XD52 = XD52M / 25.4              'Pitch,inch
XD54 = Val(T_NO)                 'Number of tubes
XD55 = Val(T_len)                'Tube lenght,m
XD56 = XD55 / 0.3048             'Tube lenght     ft
XD57 = Val(T_PASS)               'Number of tube side passes
XD58 = Val(Mat_cond)             'Thermal conductivity of tube material,Kcal/(h m^2 ºC/m)
XD59 = Val(SHELL_PASS)           'Shell passes
XD61M = Val(SHELL_ID)            'Shell ID, mm
XD61 = XD61M / 25.4              'Shell ID, inch
XD63 = Val(SHELL_BAFFLES_CUT)    'Baffle cut, %
XD64M = Val(SHELL_BAFFLES_SPACE) 'Baffle spacing  mm
XD64 = XD64M / 25.4              'Baffle spacing  inch
XD66M = Val(T_OD)                'Tube Outlet diameter, mm
XD66 = XD66M / 1000              'Tube Outlet diameter, m
XD50 = XD66 / 25.4 * 1000        'Tube outlet diameter, inch
XD85 = Val(T_ID) / 1000          'Tube Inlet diameter, m
XD84 = XD85 / 0.3048             'Tube Inlet diameter,ft
XD83 = XD85 / 25.4 * 1000        'Tube Inlet diameter,inches
XD112 = Val(SHELL_FF)            'Process side fouling factor [(hm^2ºC)/Kcal]*10^4
XD30 = XD112
XD75 = SHELL_OUT(4)              'Shell density
XD77 = SHELL_OUT(5)              'SHELL VISCOSITY

'SHELL Caloric temperature,ºC
    If Combo_S_FLUID.Text = "Water" Then
        XD9_S = XD20 + (XD19 - XD20) / 2
    Else
        XD9_S = XD19 + (XD20 - XD19) / 2
    End If

If Combo_T_FLUID = "Water" Then
    If Check_T_DENS = Unchecked Then
        XD11 = D19                       'Water density at ((t1+t2)/2),Kg/m3
    ElseIf Check_T_DENS = Checked Then
        XD11 = TUBES_OUT(4)              'Water density at ((t1+t2)/2),Kg/m3
    End If
    If Check_T_VISC = Unchecked Then
        XD12 = D20                       'Water viscosity at (t1+t2)/2,centipoise
    ElseIf Check_T_VISC = Checked Then
        XD12 = TUBES_OUT(5)              'Water density at ((t1+t2)/2),Kg/m3
    End If
    If Check_T_SPH = Unchecked Then
        XD13 = D21                       'Water specific heat,Kcal/(Kg ºC)
    ElseIf Check_T_SPH = Checked Then
        XD13 = TUBES_OUT(3)              'Water specific heat,Kcal/(Kg ºC)
    End If
Else
        XD11 = TUBES_OUT(4)
        XD12 = TUBES_OUT(5)
        XD13 = TUBES_OUT(3)
End If
XD5 = Val(TUBES_LIQUID) / XD11    'Water flow rate,m3/h

TUBES_OUT(3).Text = Format(XD13, "0.000")
TUBES_OUT(4).Text = Format(XD11, "0.0")
TUBES_OUT(5).Text = Format(XD12, "0.000")

If Combo_S_FLUID = "Benzene" Then
'BENZENE
    'Benzene condensation pressure, bara
        XD21 = 10 ^ (-17.7266795554994 + 0.107558360498623 * (273.16 + XD20) - 2.20426777069489E-04 * (273.16 + XD20) ^ 2 + 1.64343153650998E-07 * (273.16 + XD20) ^ 3)
    'Benzene IN enthalpy, Kcal/Kg
        XD22 = (-5.84000817614495E-03 + Sqr((5.84000817614495E-03 ^ 2 - 4 * (2.04979316731536 - Log(273.16 + XD19) / LN) * -1.33496728809234E-05))) / (2 * (-1.33496728809234E-05))
    'Benzene OUT enthalpy, Kcal/Kg
        XD23 = (-2.22078281567555E-02 + Sqr((2.22078281567555E-02 ^ 2 - 4 * (-0.26019046655962 - Log(273.16 + XD20) / LN) * -4.20919412093621E-05))) / (2 * (-4.20919412093621E-05))
    'Benzene latent heat Kcal/Kg
        XD24 = XD23 - XD22
    If Check_VAP_P = Unchecked Then
        VAP_PERC = 100
        Spin_VAP_P = VAP_PERC * 10
    End If
    
ElseIf Combo_S_FLUID = "Toluene" Then
'TOLUENE
    'Toluene condensation pressure, bara
        XD21 = 10 ^ (-20.4550478749456 + 0.122159504194137 * (273.16 + XD20) - 2.52019819590687E-04 * (273.16 + XD20) ^ 2 + 1.89239695329835E-07 * (273.16 + XD20) ^ 3)
    'Toluene IN enthalpy, Kcal/Kg
        XD22 = (-6.19081006602303E-03 + Sqr((6.19081006602303E-03 ^ 2 - 4 * (2.05395178519857 - Log(273.16 + XD19) / LN) * -1.55678980165631E-05))) / (2 * (-1.55678980165631E-05))
    'Toluene OUT enthalpy, Kcal/Kg
        XD23 = (-1.87023824724839E-02 + Sqr((1.87023824724839E-02 ^ 2 - 4 * (0.253332590099177 - Log(273.16 + XD20) / LN) * -3.62854720193424E-05))) / (2 * (-3.62854720193424E-05))
    'Toluene latent heat Kcal/Kg
        XD24 = XD23 - XD22
    If Check_VAP_P = Unchecked Then
        VAP_PERC = 100
        Spin_VAP_P = VAP_PERC * 10
    End If

ElseIf Combo_S_FLUID = "Ammonia" Then
'AMMONIA
    'Ammonia condensation pressure, bara
        XD21 = 10 ^ (-11.2065110833853 + 8.16294269621182E-02 * (273.16 + XD20) - 1.80542098712002E-04 * (273.16 + XD20) ^ 2 + 1.47860782222754E-07 * (273.16 + XD20) ^ 3)
    'Ammonia IN enthalpy, Kcal/Kg
        XD22 = (-4.90421676857996E-04 + Sqr((4.90421676857996E-04 ^ 2 - 4 * (2.6065188815352 - Log(273.16 + XD19) / LN) * -2.24502251374982E-06))) / (2 * (-2.24502251374982E-06))
    'Ammonia OUT enthalpy, Kcal/Kg
        A1 = -0.315317436306974
        A2 = 21.3312078166881
        A3 = 1.31615331719314E-03
        A4 = Log(273.16 + XD20) / LN
        XD23 = (-A1 + Sqr((A1 ^ 2 - 4 * (A2 - A4) * A3))) / (2 * (A3))
    'Ammonia latent heat Kcal/Kg
        XD24 = XD23 - XD22
    If Check_VAP_P = Unchecked Then
        VAP_PERC = 100
        Spin_VAP_P = VAP_PERC * 10
    End If

ElseIf Combo_S_FLUID = "Propylene" Then
'PROPYLENE
    'Propylene condensation pressure, bara
        P_P1 = -11.0028519534878            'H5B5
        P_P2 = 8.88800106381893E-02         'H5C5
        P_P3 = -2.25058891968011E-04        'H5D5
        P_P4 = 2.10257407154681E-07         'H5E5
        P_P5 = 273.16 + XD20
        XD21 = 10 ^ (P_P1 + P_P2 * P_P5 + P_P3 * P_P5 ^ 2 + P_P4 * P_P5 ^ 3)
    'Propylene IN enthalpy, Kcal/Kg
        P_E1 = 6.12185604296598E-03         'H5C6
        P_E2 = 1.86493425411241             'H5B6
        P_E3 = -1.29346385178499E-05        'H5D6
        P_E4 = Log(273.16 + XD19) / LN
        XD22 = (-P_E1 + Sqr((P_E1 ^ 2 - 4 * (P_E2 - P_E4) * P_E3))) / (2 * P_E3)
    'Propylene OUT enthalpy, Kcal/Kg
        P_A1 = -7.10616231537362E-02        'H5C7
        P_A2 = 9.1814480913749              'H5B7
        P_A3 = 1.84047520046369E-04         'H5D7
        P_A4 = Log(273.16 + XD20) / LN
        XD23 = (-P_A1 + Sqr((P_A1 ^ 2 - 4 * (P_A2 - P_A4) * P_A3))) / (2 * P_A3)
    'Propylene latent heat Kcal/Kg
        XD24 = XD23 - XD22
    If Check_VAP_P = Unchecked Then
        VAP_PERC = 100
        Spin_VAP_P = VAP_PERC * 10
    End If

ElseIf Combo_S_FLUID = "VCM" Then
'VCM
    'VCM condensation pressure, bara  (WAGNER EQUATION)
        V_B4 = XD20 + 273.16
        V_B5 = 429.7
        V_B6 = V_B4 / V_B5
        V_B8 = 51.5
        V_B7 = 1 - V_B6
        V_B9 = -6.50008
        V_B10 = 1.21422
        V_B11 = -2.57876
        V_B12 = -2.00937
        XD21 = V_B8 * Exp((1 - V_B7) ^ -1 * (V_B9 * V_B7 + V_B10 * V_B7 ^ 1.5 + V_B11 * V_B7 ^ 3 + V_B12 * V_B7 ^ 6))
    'VCM latent heat Kcal/Kg
        V_L5 = XD20 + 273.16
        V_L6 = 8.314
        V_L7 = 429.7
        V_L8 = 0.122
        V_L9 = V_L5 / V_L7
        V_L10 = 62.499
        XD24 = V_L6 * V_L7 * (7.08 * (1 - V_L9) ^ (0.354) + 10.95 * V_L8 * (1 - V_L9) ^ 0.456) / (V_L10 * 4.18)
    If Check_VAP_P = Unchecked Then
        VAP_PERC = 100
        Spin_VAP_P = VAP_PERC * 10
    End If

ElseIf Combo_S_FLUID = "Steam condensing" Then
'STEAM
    'Steam condensation pressure
        XD21 = Val(S_press_KP.Text) / 100
        If XD21 = 0 Then XD21 = 0.05
    'Latent heat, Kcal/kg
        I9 = 0.168682569821809
        J9 = -1.80896828868017E-04
        J3 = -38.2917529410035
        XD24 = (-I9 - Sqr(I9 ^ 2 - 4 * J9 * (J3 - Log(XD21) / 2.3))) / (2 * J9)
    'Steam specific heat at condensing film temperature Kcal/h m ºC
        If Check_S_SPH = Unchecked Then
            XD13_S = 0.00000000124678 * XD9_S ^ 4 - 0.00000023989201 * XD9_S ^ 3 + 0.00001910636228 * XD9_S ^ 2 - 0.00066516557128 * XD9_S + 1.00688381947618
            SHELL_OUT(3) = Format(XD13_S, "0.000")
        ElseIf Check_S_SPH = Checked Then
            XD13_S = HScroll_SHELL_SPH / 1000
        End If
    If Check_VAP_P = Unchecked Then
        VAP_PERC = 100
        Spin_VAP_P = VAP_PERC * 10
    End If

ElseIf Combo_CURRENT = "Condensation" Then
'CONDENSATION
    XD21 = Val(S_press_KP.Text) / 100
    'Latent heat Kcal/Kg
        XD24 = Val(SHELL_OUT(10).Text) / 4.1868
End If

If Check_LATENT = Unchecked Then
    SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0") 'Shell fluid latent heat KJ/Kg
ElseIf Check_LATENT = Checked Then
    XD24 = HScroll_LATENT / 4.1868
    SHELL_OUT(10) = Format(XD24 * 4.1868, "0.0") 'Shell fluid latent heat KJ/Kg
End If
'DUTY
    TUBES_OUT(0) = Format(XD6, "##,##0")
    SHELL_OUT(0) = Format(XD18, "##,##0")
    COL = 0
    XD18V = XD18 * VAP_PERC / 100
    XD18L = XD18 * (1 - VAP_PERC / 100)
    XD37 = XD18V * XD24 + XD18L * XD13_S * RANGE_S 'Shell side duty,Kcal/h
    XD36 = XD6 * XD13 * RANGE_T                    'Tubes side duty,Kcal/h
    DUTY_S = XD37 / 860

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
            XD6 = XD37 / (XD13 * RANGE_T)
            XD6L = XD6
        ElseIf Check_T_OUT = 0 Then
            TUBES_FLOW.ForeColor = &HC0&
            TUBES_FLOW.BackColor = &H80000018
            TUBES_TEMP_OUT.ForeColor = &HFFFFFF
            TUBES_TEMP_OUT.BackColor = &HC0&
            XD6 = TUBES_FLOW
            XD8 = XD37 / XD6 / XD13 + XD7
            RANGE_T = XD8 - XD7
            Spin_TUBES_T_OUT = XD8 * 100
            XD36 = XD6 * XD13 * RANGE_T    'Tubes side duty,Kcal/h
            
       End If
     'Tube side duty, kcal/h
        XD36 = XD6 * XD13 * RANGE_T
        TUBES_FLOW = Format(XD6, "0")
        TUBES_LIQUID = Format(XD6, "0")
        TUBES_OUT(0) = Format(XD6, "##,##0")
        HScroll_TUBES_FLOW = XD6 / FACT_FLOW
        If Combo_T_FLUID = "Water" Then
            TUBES_LIQUID.Text = Format(XD6, "0")
            TUBES_WATER.Text = Format(XD6, "0")
            HScroll_TUBES_WATER.Max = Val(TUBES_WATER) / FACT_FLOW
            XD5 = XD6 / D19
        End If
        TUBES_OUT(7).Text = Format(XD36 * 0.001163, "##,##0") 'TUBES side duty,Kcal/h, KW
    ElseIf Thermal_bal_shell = True Then
        COL = 1
        S_FLOW = 1
        SHELL_FLOW.ForeColor = &HFFFFFF
        SHELL_FLOW.BackColor = &HC0&
        TUBES_FLOW.ForeColor = &HC0&
        TUBES_FLOW.BackColor = &H80000018
        TUBES_TEMP_OUT.ForeColor = &HC0&
        TUBES_TEMP_OUT.BackColor = &H80000018
        SHELL_OUT(0).BackColor = &HC0&
        SHELL_OUT(0).ForeColor = &HFFFFFF
        TUBES_OUT(0).BackColor = &HE0E0E0
        TUBES_OUT(0).ForeColor = &HC0&
        WAT = Data1.Recordset.SHELL_WATER
        N_COND = Data1.Recordset.SHELL_NON_COND
        SHELL_VAP = Data1.Recordset.SHELL_VAPOR

10      W = 0.1
11      j = 0.1
12      HE = W: GoSub 18
13      Y = X: HE = j + W
14      GoSub 18
15      G = W: W = G - j * Y / (X - Y)
16      If Abs(G - W) >= 0.00001 Then GoTo 12
17      W = HE: GoTo 19
18      FT = HE
        'Shell duty, kcal/h
        FX = (FT * VAP_PERC * XD24 / 100 + FT * LIQ_PERC / 100 * XD13_S * RANGE_S)
        DUTY_S = FX / 860
        XD_37 = FX
        X = XD36 - FX
        SHELL_OUT(7).Text = Format(DUTY_S, "##,##0") 'SHELL side duty, KW
Return

19      XD18_V = FT * VAP_PERC / 100
        XD18_L = FT * LIQ_PERC / 100 - N_COND
        XD18 = XD18_V + XD18_L + N_COND
        LIQ_P = Format(XD18_L / XD18 * 100, "0.0")
        N_COND_PERC = Format(N_COND / XD18 * 100, "0.0")
        
        SHELL_FLOW.Text = Format(XD18, "0")
        SHELL_VAPOR.Text = Format(SHELL_VAP, "0")
        SHELL_LIQUID.Text = Format(XD18 - SHELL_VAPOR - SHELL_NON_COND, "0")
        SHELL_WATER.Text = WAT
        SHELL_NON_COND.Text = N_COND

        HScroll_SHELL_VAPOR.Max = SHELL_FLOW / FACT_FLOW
        HScroll_SHELL_LIQUID.Max = SHELL_FLOW / FACT_FLOW
        HScroll_SHELL_WATER.Max = SHELL_FLOW / FACT_FLOW
        HScroll_SHELL_NON_COND.Max = SHELL_FLOW / FACT_FLOW
        SHELL_OUT(0) = Format(XD18, "##,##0")
        Spin_SHELL_FLOW = XD18 / FACT_FLOW
        AA1 = XD37 / 860
        If Combo_S_FLUID = "Water" Then
            SHELL_LIQUID.Text = Format(XD18, "0")
            SHELL_WATER.Text = Format(XD18, "0")
            HScroll_SHELL_WATER.Max = Val(SHELL_WATER) / FACT_FLOW
        End If
        SHELL_OUT(7).Text = Format(XD37 * 0.001163, "##,##0") 'SHELL side duty, KW
    Else
        SHELL_FLOW.ForeColor = &HC0&
        SHELL_FLOW.BackColor = &H80000018
        TUBES_FLOW.ForeColor = &HC0&
        TUBES_FLOW.BackColor = &H80000018
        TUBES_TEMP_OUT.ForeColor = &HC0&
        TUBES_TEMP_OUT.BackColor = &H80000018
        TUBES_OUT(0).BackColor = &HE0E0E0
        TUBES_OUT(0).ForeColor = &HC0&
        SHELL_OUT(0).BackColor = &HE0E0E0
        SHELL_OUT(0).ForeColor = &HC0&
    End If

'CALCULATING TUBE SIDE PRESSURE DROP

'Pressure drop in tubes,psi
    'Flow area per tube  m^2
        XD91 = (XD54 * XPI * XD85 ^ 2 / 4) / XD57  'Flow area tubes,m^2
        If lungh = 2 And XD57 > 1 Then
            XD91 = XD91 * lungh
        End If
    'Flow area per tube,ft^2
        XD92 = XD91 / (0.3048 ^ 2)
    'Water velocity through tubes    m/s
        XD95 = (XD5 / (XD91 * 3600))
    'Water velocity through tubes,ft/s
        XD96 = XD95 / 0.3048
    'Reynolds number
        XD97 = XD95 * XD85 * XD11 / (XD12 * 0.001)
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
    'Total pressure drop for 100% clean tube side    psi
        XD104 = XD98 + XD101
    'Total pressure drop for 100% clean tube side    bar
        XD105 = XD99 + XD102
    'Total pressure drop for 100% clean tube side    Kg/cm2
        XD106 = XD100 + XD103
    'Total pressure drop for 100% clean tube side    kPa
        XD106_KPA = XD105 * 100

    TUBES_SECTION.Text = Format(XD91, "0.0000")
    TUBES_OUT(2) = Format(XD95, "0.00")      'Water velocity through tubes    m/s
    TUBES_OUT(6) = Format(XD97, "##,##0")    'Reynolds number through tubes
    If Check_P_DROP_T = Unchecked Then
        TUBES_OUT(9) = Format(XD106_KPA, "0.00") 'Total pressure drop
    Else
        XD106_KPA = (Val(TUBES_P_IN) - Val(TUBES_P_OUT)) * 100
        TUBES_OUT(9).Text = Format(XD106_KPA, "0.00")        ' KPa
    End If
    C_F = XD5 / (XD106_KPA) ^ (1 / 2)
    C_Factor.Text = Format(C_F, "0")

'CALCULATING HEAT TRANSFER
    
    'Water side individual heat transfer coeficient  Btu/(h ft^2 F)
        XD108 = 150 * (1 + 0.011 * XD10) * (XD96 ^ 0.8 / XD83 ^ 0.2)
    'Water side individual heat transfer coeficient  Kcal/(h m^2 C)
        XD109 = XD108 * 4.882
    'Water side indiv. heat transfer coeficient referred to ext. surface Kcal/(h m^2 C)
        XD110 = XD109 * (XD85 / XD66)
    'Heat transfer resistance due to the wall    [(hm^2ºC)/Kcal]*10^4
        XD111 = (XD66 * Log(XD66 / XD85) / (2 * XD58)) * 10000
    'Heat transfer resistance due to outside fouling factor  [(h m^2 ºC)/Kcal]*10^4  1.00
        'XD112 = XD30
    'Heat transfer resistance due to water (tube side)   [(hm^2ºC)/Kcal]*10^4
        XD114 = 10 ^ 4 / XD110

'CALCULATING h_o
    
'Shell side crossflow area, ft^2
    XD68 = XD61 * (XD52 - XD50) * XD64 / (XD52 * 144) * lungh
'Condensate loading  lb/h ft
    XD69 = XD18 * 2.20462 / (XD56 * XD54 ^ (2 / 3))
'Shell side indiv. heat transfer coefficient,Kcal/(h m^2 C)
    'Shell side indiv. heat transfer coefficient,Btu/(h ft^2 F)
'Shell side indiv. heat transfer coefficient (guess one till Z_0=0),Kcal/(hm^2°C)
100      W = 0.1
110      j = 0.1
120      HE = W: GoSub 180
130      Y = X: HE = j + W
140      GoSub 180
150      G = W: W = G - j * Y / (X - Y)
160      If Abs(G - W) >= 0.00001 Then GoTo 120
170      W = HE: GoTo 190
180      XD70 = HE
        'Wall temperature,ºC
             XD72 = XD9 + (XD70 / (XD110 + XD70)) * (XD9_S - XD9)
        'Condensing film temperature ºC
            XD73 = (XD20 + XD72) / 2
        'Condensing film temperature ºC
            If Check_CT = Unchecked Then
                SHELL_OUT(11) = Format(XD73, "0.00")
            ElseIf Check_CT = Checked Then
            XD73 = HScroll_C_TEMP / 100
        End If
        'Condensing film temperature,K
            XD74 = 273.16 + XD73
'BENZENE
    If Combo_S_FLUID = "Benzene" Then
        'Benzene density at condensing film temperature,Kg/m^3
            If Check_S_DENS = Unchecked Then
                XD75 = (1.376997665 - 0.002829185 * XD74 + 5.6388778534534E-06 * XD74 ^ 2 - 6.07429148488935E-09 * XD74 ^ 3) * 1000
            ElseIf Check_S_DENS = Checked Then
                XD75 = HScroll_SHELL_DENS / 10
            End If
        'Benzene density at condensing film temperature,lb/ft^3
            XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
        'Benzene viscosity at condensing film temperature,cp
            If Check_S_VISC = Unchecked Then
                XD77 = (174.50837619225 - 1.30737638396497 * XD74 + 3.36425241742188E-03 * XD74 ^ 2 - 2.93447293753068E-06 * XD74 ^ 3) / 10
            ElseIf Check_S_VISC = Checked Then
                XD77 = HScroll_SHELL_VISC.Value / 1000
            End If
        'Benzene thermal conductivity at condensing film temperature,Btu/hftºF
            XD78 = (0.234000000232994 - 3.00000002048061E-04 * XD74 + 5.96459460478361E-15 * XD74 ^ 2 - 5.75588815841478E-18 * XD74 ^ 3) * 0.5778
        'Benzene thermal conductivity at condensing film temperature Kcal/h m ºC
            If Check_S_TC = Unchecked Then
                XD79 = XD78 * 1.488
                SHELL_OUT(1) = XD79
            ElseIf Check_S_TC = Checked Then
                XD79 = SHELL_OUT(1)
                XD78 = XD79 / 1.488
            End If
'TOLUENE
    ElseIf Combo_S_FLUID = "Toluene" Then
        'Toluene density at condensing film temperature,Kg/m^3
            If Check_S_DENS = Unchecked Then
                XD75 = (1.09657697034396 - 7.13586701948427E-04 * XD74 - 6.21961123052573E-08 * XD74 ^ 2 - 6.0875153993795E-10 * XD74 ^ 3) * 1000
            ElseIf Check_S_DENS = Checked Then
                XD75 = HScroll_SHELL_DENS / 10
            End If
        'Toluene density at condensing film temperature,lb/ft^3
            XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
        'Toluene viscosity at condensing film temperature,cp
            If Check_S_VISC = Unchecked Then
                XD77 = (114.844320200991 - 0.826324102497882 * XD74 + 2.06924742120918E-03 * XD74 ^ 2 - 1.7689717708644E-06 * XD74 ^ 3) / 10
            ElseIf Check_S_VISC = Checked Then
                XD77 = HScroll_SHELL_VISC.Value / 1000
            End If
        'Toluene thermal conductivity at condensing film temperature,Btu/hftºF
            XD78 = (0.266958042246326 - 7.4378954632313E-04 * XD74 + 1.34365635103354E-06 * XD74 ^ 2 - 1.1655011726189E-09 * XD74 ^ 3) * 0.5778
        'Toluene thermal conductivity at condensing film temperature Kcal/h m ºC
            If Check_S_TC = Unchecked Then
                XD79 = XD78 * 1.488
                SHELL_OUT(1) = XD79
            ElseIf Check_S_TC = Checked Then
                XD79 = SHELL_OUT(1)
                XD78 = XD79 / 1.488
            End If
'AMMONIA
    ElseIf Combo_S_FLUID = "Ammonia" Then
        'Ammonia density at condensing film temperature,Kg/m^3
            If Check_S_DENS = Unchecked Then
                XD75 = (1.67201252301867 - 8.67573506044482E-03 * XD74 + 2.71463576525648E-05 * XD74 ^ 2 - 3.37884138338416E-08 * XD74 ^ 3) * 1000
            ElseIf Check_S_DENS = Checked Then
                XD75 = HScroll_SHELL_DENS / 10
            End If
        'Ammonia density at condensing film temperature,lb/ft^3
            XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
        'Ammonia viscosity at condensing film temperature,cp
            If Check_S_VISC = Unchecked Then
                XD77 = (22.8772727564961 - 0.154773809799488 * XD74 + 3.68290044153551E-04 * XD74 ^ 2 + -3.03030303928442E-07 * XD74 ^ 3) / 10
            ElseIf Check_S_VISC = Checked Then
                XD77 = HScroll_SHELL_VISC.Value / 1000
            End If
        'Ammonia thermal conductivity at condensing film temperature,Btu/hftºF
            XD78 = (1.67572727536397 - 7.29141416629545E-03 * XD74 + 1.62878788658395E-05 * XD74 ^ 2 - 1.76767677578808E-08 * XD74 ^ 3) * 0.5778
        'Ammonia thermal conductivity at condensing film temperature Kcal/h m ºC
            XD79 = XD78 * 1.488
            If Check_S_TC = Unchecked Then
                SHELL_OUT(1) = Format(XD79, "0.000")    'Ammonia thermal conductivity
            ElseIf Check_S_TC = Checked Then
                XD79 = SHELL_OUT(1)
                XD78 = XD79 / 1.488
            End If
'PROPYLENE
    ElseIf Combo_S_FLUID = "Propylene" Then
        'Propylene density at condensing film temperature,Kg/m^3
            If Check_S_DENS = Unchecked Then
                P_D1 = 1.36605329630155      'H5B11
                P_D2 = -7.21816263067366E-03 'H5C11
                P_D3 = 2.52654604770147E-05  'H5D11
                P_D4 = -3.59719473341389E-08 'H5E11
                XD75 = (P_D1 + P_D2 * XD74 + P_D3 * XD74 ^ 2 + P_D4 * XD74 ^ 3) * 1000
            ElseIf Check_S_DENS = Checked Then
                XD75 = HScroll_SHELL_DENS / 10
            End If
        'Propylene density at condensing film temperature,lb/ft^3
            XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
        'Propylene viscosity at condensing film temperature,cp
            If Check_S_VISC = Unchecked Then
                P_V1 = 21.3988888935207         'H5B9
                P_V2 = -0.184389462440866       'H5C9
                P_V3 = 5.72566322754654E-04     'H5D9
                P_V4 = -6.20305620533474E-07    'H5E9
                XD77 = (P_V1 + P_V2 * XD74 + P_V3 * XD74 ^ 2 + P_V4 * XD74 ^ 3) / 10
            ElseIf Check_S_VISC = Checked Then
                XD77 = HScroll_SHELL_VISC.Value / 1000
            End If
        'Propylene thermal conductivity at condensing film temperature,Btu/hftºF
            P_C1 = 0.183205128197412        'H5B10
            P_C2 = 3.90637140722696E-04     'H5C10
            P_C3 = -3.27505827537143E-06    'H5D10
            P_C4 = 3.88500388538241E-09     'H5E10
            XD78 = (P_C1 + P_C2 * XD74 + P_C3 * XD74 ^ 2 + P_C4 * XD74 ^ 3) * 0.5778
        'Propylene thermal conductivity at condensing film temperature Kcal/h m ºC
            XD79 = XD78 * 1.488
            If Check_S_TC = Unchecked Then
                SHELL_OUT(1) = Format(XD79, "0.000")    'Propylene thermal conductivity
            ElseIf Check_S_TC = Checked Then
                XD79 = SHELL_OUT(1)
                XD78 = XD79 / 1.488
            End If
'VCM
    ElseIf Combo_S_FLUID = "VCM" Then
        'VCM density at condensing film temperature,Kg/m^3
            If Check_S_DENS = Unchecked Then
                V_D4 = XD74             'T=
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
            ElseIf Check_S_DENS = Checked Then
                XD75 = HScroll_SHELL_DENS / 10
            End If
        'VCM density at condensing film temperature,lb/ft^3
            XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
        'VCM viscosity at condensing film temperature,cp
            If Check_S_VISC = Unchecked Then
                V_V3 = XD74
                V_V4 = 2
                V_V5 = -0.152 - 0.042 * B4
                V_V6 = 1.91 - 1.459
                V_V7 = V_V4 + V_V5 + V_V6
                V_V9 = 28.86 + 37.439 * V_V7 - 1.3547 * V_V7 ^ 2 + 0.0276 * V_V7 ^ 3
                V_V10 = 24.79 + 66.885 * V_V7 - 1.3173 * V_V7 ^ 2 - 0.00377 * V_V7 ^ 3 - 44.94 + 5.41 * V_V7 - 26.38
                XD77 = 10 ^ (V_V10 * (V_V3 ^ -1 - V_V9 ^ -1))
            ElseIf Check_S_VISC = Checked Then
                XD77 = HScroll_SHELL_VISC.Value / 1000
            End If
        'VCM thermal conductivity at condensing film temperature,Kcal/h m ºC
            V_T5 = XD74
            V_T6 = 259.8
            V_T7 = 62.499
            V_T8 = 429.7
            V_T9 = V_T5 / V_T8
            V_T10 = V_T6 / V_T8
            XD79 = ((1.11 / (V_T7 ^ 0.5)) * (3 + 20 * (1 - V_T9) ^ (2 / 3)) / (3 + 20 * (1 - V_T10) ^ (2 / 3))) * (3.6 / 4.18)
            XD78 = XD79 / 1.488
            If Check_S_TC = Unchecked Then
                SHELL_OUT(1) = Format(XD79, "0.000")    'VCM thermal conductivity
            ElseIf Check_S_TC = Checked Then
                XD79 = SHELL_OUT(1)
                XD78 = XD79 / 1.488
            End If
        
'STEAM
    ElseIf Combo_S_FLUID = "Steam condensing" Then
        'Steam density at condensing film temperature,Kg/m^3
            If Check_S_DENS = Unchecked Then
                XD75 = -0.4351444 * XD73 + 1009.209598
            ElseIf Check_S_DENS = Checked Then
                XD75 = HScroll_SHELL_DENS / 10
            End If
        'Steam density at condensing film temperature,lb/ft^3
            XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
        'Steam viscosity at condensing film temperature,cp
            If Check_S_VISC = Unchecked Then
                XD77 = (100 / (2.1482 * ((XD74 - 281.435) + Sqr(8078.4 + (XD74 - 281.435) ^ 2)) - 120))
            ElseIf Check_S_VISC = Checked Then
                XD77 = HScroll_SHELL_VISC.Value / 1000
            End If
        'Steam thermal conductivity at condensing film temperature,Btu/hftºF
            XD78 = (0.00000000592317 * XD73 ^ 3 - 0.0000080425 * XD73 ^ 2 + 0.0018262 * XD73 + 0.478535) * 1.488
        'Steam thermal conductivity at condensing film temperature Kcal/h m ºC
            If Check_S_TC = Unchecked Then
                XD79 = XD78 / 1.488
                SHELL_OUT(1) = Format(XD79, "0.000")
            ElseIf Check_S_TC = Checked Then
                XD79 = SHELL_OUT(1)
                XD78 = XD79 / 1.488
            End If

'CONDENSATION
    ElseIf Combo_CURRENT = "Condensation" Then
        'Density at condensing film temperature,Kg/m^3
                XD75 = SHELL_OUT(4)  'HScroll_SHELL_DENS / 10
        'Density at condensing film temperature,lb/ft^3
            XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
        'Viscosity at condensing film temperature,cp
                XD77 = SHELL_OUT(5) 'HScroll_SHELL_VISC.Value / 1000
        'Thermal conductivity at condensing film temperature Kcal/h m ºC
                XD79 = SHELL_OUT(1)
                XD78 = XD79 / 1.488
    End If
    XD80 = 1.5 * ((4 * XD69 / XD77) ^ -(1 / 3)) * (XD77 ^ 2 / (XD78 ^ 3 * XD76 ^ 2 * 9.81 * (3600 ^ 2 / 0.3048))) ^ -(1 / 3)
    XD81 = XD80 * 4.882
    X = (XD70 - XD81)
Return
190 'XD70 = Shell side indiv. heat transfer coefficient (guess one till Z_0=0)

'Heat transfer resistance due to process (shell side),[(hm^2ºC)/Kcal]*10^4
    XD115 = (1 / XD81) * 10000
'Overall CLEAN heat transfer coefficient Kcal/(h m^2 ºC)
    XD117 = 10000 / (XD111 + XD114 + XD115)

PROP = "SHELL"
Call Properties

SHELL_OUT(1) = Format(XD79, "0.000")  'SHELL thermal conductivity at condensing film temperature Kcal/h m ºC
SHELL_OUT(4) = Format(XD75, "0.0")    'SHELL density at condensing film temperature,Kg/m^3
SHELL_OUT(5) = Format(XD77, "0.000")  'SHELL viscosity at condensing film temperature,cp
SHELL_OUT(11) = Format(XD73, "0.00")  'SHELL condensing film temperature ºC
SKIN_TEMP.Text = Format(XD72, "0.00") 'Wall temperature,ºC

'SHELL Velocity
    L37 = XD61M   'Val(SHELL_ID.Text)
    O37 = XD52M   'Val(SHELL_TUBES_PITCH.Text)
    E37 = XD66M   'Val(T_OD.Text)
    N37 = XD64M   'Val(SHELL_BAFFLES_SPACE.Text)
    'SHELL CLEARANCE
    V37 = O37 / 1000 - E37 / 1000
    Clearance.Text = Format(V37, "0.0000")
    'SHELL FLOW AREA
    U37 = L37 / 1000 * V37 * N37 / 1000 / (O37 / 1000)
    Flow_area.Text = Format(U37, "0.0000")
    k47 = XD18L / XD75 / U37 / 3600
    SHELL_OUT(2) = Format(k47, "0.00")

'SHELL Reynolds
    EQ_D19 = XD52M / 1000 'SHELL_TUBES_PITCH / 1000
    EQ_PI = XPI
    EQ_D14 = XD66  'T_OD / 1000
    If SHELL_PITCH_CONF = "Triangular" Then
        EQ_E31 = 4 * (EQ_D19 ^ 2 - EQ_PI * EQ_D14 ^ 2 / 4) / (XPI * EQ_D14)
    Else
        EQ_E31 = (4 * (0.5 * EQ_D19 * 0.866 * EQ_D19 - 0.5 * XPI * EQ_D14 ^ 2 / 4) / (0.5 * XPI * EQ_D14))
    End If
    EQ_E29 = U37       'Flow_area
    EQ_E25 = Val(SHELL_LIQUID)      'SHELL_FLOW LIQUID
    EQ_E30 = EQ_E25 / EQ_E29
    EQ_E8 = XD77 * 3.6   'SHELL_OUT(5) * 3.6
    EQ_E32 = EQ_E31 * EQ_E30 / EQ_E8
    Q_E22 = EQ_E31 * 1000
    Q_E17 = XD75            'Density, kg/m3
    Q_E27 = k47             'Shell flow velocity, m/s
    Q_E18 = XD77            'SHELL viscosity, cP
    Q_E28 = Q_E22 * Q_E17 * Q_E27 / Q_E18
    SHELL_OUT(6) = Format(Q_E28, "##,##0")

'CALCULATING SHELL SIDE PRESSURE DROP
    If Check_P_DROP_S = Unchecked Then
        'Pressure drop (tubes)
        P_E17 = XD75             'SHELL_OUT(4)
        P_E22 = EQ_E31 * 1000    ' Equivalent diameter, mm
        P_E27 = Q_E27            'Shell flow velocity, m/s
        P_E23 = XD55 * lungh     'T_len, mm
        P_E28 = Q_E28
        P_E29 = 0.44 * P_E28 ^ -0.19
        P_E30 = 4 * P_E29 * P_E23 * P_E27 ^ 2 / (P_E22 * 2 * 9.8) * P_E17 * 0.000096784 * 101.325
        
        'Pressure drop (sheet)
        P_E9 = XD59               'Val(SHELL_PASS)
        P_E31 = 3 * P_E9 * P_E27 ^ 2 / 2 / 9.8 * P_E17 * 0.000096784 * 101.325
        P_E32 = P_E30 + P_E31
        
        SHELL_OUT(9).Text = Format(P_E32, "0.00")      ' KPa
    Else
        P_E32 = (Val(SHELL_P_IN) - Val(SHELL_P_OUT))
        SHELL_OUT(9).Text = Format(P_E32 * 100, "0.00")      ' KPa
    End If

'Water side fouling factor   [(hm^2ºC)/Kcal]*10^4
    'Overall heat transfer coefficient   Kcal/(h m^2 ºC)
        'Surface per linear ft, ft^2
            XD50 = Format(XD66 * 1000 / 25.4, "0.000")          'Tube outlet diameter,inch
        'Surface per linear m, m^2
            XD90 = 3.14159 * D67 * 25.4 * 10 ^ -3 * D75 * D74 * SERIES_N * PARALLEL_N * lungh
        'Log Mean Temperature Difference CORRECTED, ºC
            AG6 = ((XD19 - XD8) - (XD20 - XD7)) / Log((XD19 - XD8) / (XD20 - XD7))
            RR = (XD19 - XD20) / (XD8 - XD7)
            SS = (XD8 - XD7) / (XD19 - XD7)
    If T_PASS > 1 And SERIES_N > 1 Then
            FT1 = (Sqr(RR ^ 2 + 1) / (2 * (RR - 1))) * Log((1 - SS) / (1 - RR * SS))
            FT2 = 2 / SS - 1 - RR + (2 / SS) * Sqr((1 - SS) * (1 - RR * SS)) + Sqr(RR ^ 2 + 1)
            FT3 = 2 / SS - 1 - RR + (2 / SS) * Sqr((1 - SS) * (1 - RR * SS)) - Sqr(RR ^ 2 + 1)
            FT4 = Log(FT2 / FT3)
            FT = FT1 / FT4
    ElseIf T_PASS > 1 And SHELL_PASS > 1 Then
            FT1 = (Sqr(RR ^ 2 + 1) / (2 * (RR - 1))) * Log((1 - SS) / (1 - RR * SS))
            FT2 = 2 / SS - 1 - RR + (2 / SS) * Sqr((1 - SS) * (1 - RR * SS)) + Sqr(RR ^ 2 + 1)
            FT3 = 2 / SS - 1 - RR + (2 / SS) * Sqr((1 - SS) * (1 - RR * SS)) - Sqr(RR ^ 2 + 1)
            FT4 = Log(FT2 / FT3)
            FT = FT1 / FT4
    Else
            FT1 = Sqr(RR ^ 2 + 1) * Log((1 - SS) / (1 - RR * SS))
            FT2 = 2 - SS * (RR + 1 - Sqr(RR ^ 2 + 1))
            FT3 = 2 - SS * (RR + 1 + Sqr(RR ^ 2 + 1))
            FT = FT1 / ((RR - 1) * Log(FT2 / FT3))
    End If
    AH6 = AG6 * FT
    XD31 = AH6
    XD38 = XD36 / (XD90 * XD31)
    xd118 = ((1 / XD38) - (1 / XD117) - (XD112 / 10000)) * 10000 * (XD85 / XD66)
'Total heat transfer resistance  [(h m^2 ºC)/Kcal]*10^4
    XD116 = 10000 / XD38

'Heat transfer resistance due to inside fouling factor,[(hm^2ºC)/Kcal]*10^4
    XD113 = xd118 * (XD66 / XD85)

'TUBES HEAT FLUX
    Q6 = XD6 * XD13 * RANGE_T / XD90
    TUBES_OUT(8).Text = Format(Q6 * 0.001163, "0.00")
'SHELL HEAT FLUX
    If Check_S_SPH = Checked Then
        S_SPH = SHELL_OUT(3).Text
    ElseIf Check_S_SPH = Unchecked Then
        S_SPH = HScroll_SHELL_SPH
    End If
    Q6S = XD37 / XD90
    SHELL_OUT(8).Text = Format(Q6S * 0.001163, "0.00")
    Area.Text = Format(XD90, "0.00")  'Area, m^2
    LMTD.Text = Format(AG6, "0.00")  'Log Mean Temperature Difference, ºC
    MTDc.Text = Format(AH6, "0.00")   'Log Mean Temperature Difference corrected, ºC
    XD32 = XD20 - XD8                 'Condenser temperature approach  ºC
    TTD = Format(XD32, "0.00")
    Label22(15).Caption = "(T2 - t2)"
    
    TUBES_OUT(7).Text = Format(XD36 * 0.001163, "##,##0") 'TUBES side duty,Kcal/h, KW
    SHELL_OUT(7).Text = Format(XD37 * 0.001163, "##,##0") 'SHELL side duty, KW
    If Check_U_CLEAN = Checked Then
        U_COEFF_CLEAN = HScroll_U_CLEAN
    ElseIf Check_U_CLEAN = Unchecked Then
        U_COEFF_CLEAN.Text = Format(XD117, "0.0")      'Overall CLEAN heat transfer coefficient Kcal/(h m^2 ºC)
    End If
    U_COEFF_DIRTY.Text = Format(XD38, "0.0")       'Overall heat transfer coefficient   Kcal/(h m^2 ºC)
    TUBES_FF.Text = Format(xd118, "0.000")         'Water side fouling factor   [(hm^2ºC)/Kcal]*10^3
End Sub
Private Sub COOLERS()
On Error Resume Next
Dim XPI, LN As Double

If Combo_T_FLUID = "Water" Then
    Spin_TUBES_T_IN.Max = 10000
    Spin_TUBES_T_OUT.Max = 10000
Else
    Spin_TUBES_T_IN.Max = 30000
    Spin_TUBES_T_OUT.Max = 30000
End If

Lab_COND.Caption = "COOLER"
lblLabels(3).Caption = "C Factor:"
Label24(0).Caption = "m3/h/kPa^(1/2) - Tubes-side"

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
    TUBES_OUT(1).Visible = False
    Check_T_TC.Visible = False
    HScroll_TUBES_TC.Visible = False
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
    HScroll_WET_STEAM.Visible = False
    Check_WET_STEAM.Visible = False
    Check_water_steam.Visible = False
'Vapor percent
    Frame_VAP.Visible = False

Check_T_OUT.Visible = True

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
lblLabels(25).Caption = "Approach temperature:"
Thermal_bal_tubes.Visible = True
Thermal_bal_shell.Visible = True



XPI = 3.141592654
LN = 2.302585093
XD6 = TUBES_FLOW                 'TUBES total flowrate,Kg/h
XD6L = TUBES_LIQUID              'TUBES liquid flowrate,Kg/h
XD7 = TUBES_TEMP_IN              'TUBES temperature in,ºC
XD8 = TUBES_TEMP_OUT             'TUBES temperature out,ºC
XD18 = SHELL_FLOW                'Shell total fluid flowrate,Kg/h
XD18L = SHELL_LIQUID             'Shell liquid fluid flowrate,Kg/h
XD19 = SHELL_TEMP_IN             'Shell fluid temperature in,ºC
XD20 = SHELL_TEMP_OUT            'Shell fluid temperature out,ºC
XD52M = SHELL_TUBES_PITCH        'Pitch, mm
XD52 = XD52M / 25.4              'Pitch,inch
XD54 = T_NO                      'Number of tubes
XD55 = T_len                     'Tube lenght,m
XD56 = XD55 / 0.3048             'Tube lenght     ft
XD57 = T_PASS                    'Number of tube side passes
XD58 = Mat_cond                  'Thermal conductivity of tube material,Kcal/(h m^2 ºC/m)
XD59 = SHELL_PASS                'Shell passes
XD61M = SHELL_ID                 'Shell ID, mm
XD61 = XD61M / 25.4              'Shell ID, inch
XD63 = SHELL_BAFFLES_CUT         'Baffle cut, %
XD64M = SHELL_BAFFLES_SPACE      'Baffle spacing  mm
XD64 = XD64M / 25.4              'Baffle spacing  inch
XD66M = T_OD                     'Tube Outlet diameter, mm
XD66 = XD66M / 1000              'Tube Outlet diameter, m
XD50 = XD66 / 25.4 * 1000        'Tube outlet diameter, inch
XD85 = T_ID / 1000               'Tube Inlet diameter, m
XD84 = XD85 / 0.3048             'Tube Inlet diameter,ft
XD83 = XD85 / 25.4 * 1000        'Tube Inlet diameter,inches
XD112 = SHELL_FF                 'Process side fouling factor [(hm^2ºC)/Kcal]*10^4
XD75 = SHELL_OUT(4)              'Shell density
XD77 = SHELL_OUT(5)              'SHELL VISCOSITY
S_press_KP = 0

Call Mechanical
    
Mat_cond.Text = Val(D78)                                     'Thermal conductivity of tube material
'Heat transfer surface
    D79 = 3.14159 * D67 * 25.4 * 10 ^ -3 * D75 * D74 * lungh 'Heat transfer surface,m^2
    D80 = D79 / (0.3048 ^ 2)                                 'Heat transfer surface,inch^2
    Area.Text = Format(Val(D79), "0.0")

Call FOULING

RANGE_T = Abs(XD8 - XD7)
RANGE_S = Abs(XD19 - XD20)

PROP = "TUBES"
Call Properties

XD9 = D17                        'TUBES Caloric temperature,ºC
XD10 = XD9 * 1.8 + 32            'TUBES Caloric temperature,ºF

If Combo_T_FLUID = "Water" Then
    If Check_T_DENS = Unchecked Then
        XD11 = D19                       'Water density at ((t1+t2)/2),Kg/m3
    ElseIf Check_T_DENS = Checked Then
        XD11 = TUBES_OUT(4)              'Water density at ((t1+t2)/2),Kg/m3
    End If
    If Check_T_VISC = Unchecked Then
        XD12 = D20                       'Water viscosity at (t1+t2)/2,centipoise
    ElseIf Check_T_VISC = Checked Then
        XD12 = TUBES_OUT(5)              'Water viscosity at (t1+t2)/2,centipoise
    End If
    If Check_T_SPH = Unchecked Then
        XD13 = D21                       'Water specific heat,Kcal/(Kg ºC)
    ElseIf Check_T_SPH = Checked Then
        XD13 = TUBES_OUT(3)              'Water specific heat,Kcal/(Kg ºC)
    End If
Else
    XD11 = TUBES_OUT(4)
    XD12 = TUBES_OUT(5)
    XD13 = TUBES_OUT(3)
End If
XD5 = XD6 / XD11    'Water flow rate,m3/h

TUBES_OUT(3).Text = Format(XD13, "0.000")
TUBES_OUT(4).Text = Format(XD11, "0.0")
TUBES_OUT(5).Text = Format(XD12, "0.000")

'Specific heat
    XD13S = SHELL_OUT(3) 'Shell specific heat Kcal/Kg

'DUTY
    TUBES_OUT(0) = Format(XD6, "##,##0")
    SHELL_OUT(0) = Format(XD18, "##,##0")
    COL = 0
    XD37 = XD18 * XD13S * RANGE_S   'Shell fluid side duty,Kcal/h
    XD36 = XD6 * XD13 * RANGE_T     'TUBES fluid side duty,Kcal/h
    
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
            XD6 = XD37 / (XD13 * RANGE_T)
            XD6L = XD6
        ElseIf Check_T_OUT = 0 Then
            TUBES_FLOW.ForeColor = &HC0&
            TUBES_FLOW.BackColor = &H80000018
            TUBES_TEMP_OUT.ForeColor = &HFFFFFF
            TUBES_TEMP_OUT.BackColor = &HC0&
            XD6 = TUBES_FLOW
            XD8 = XD37 / XD6 / XD13 + XD7
            RANGE_T = XD8 - XD7
            Spin_TUBES_T_OUT = XD8 * 100
       End If
        
    'Tube side duty, kcal/h
        XD36 = XD6 * XD13 * RANGE_T    'Tubes side duty,Kcal/h
        TUBES_FLOW = Format(XD6, "0")
        TUBES_LIQUID = Format(XD6, "0")
        TUBES_OUT(0) = Format(XD6, "##,##0")
        HScroll_TUBES_FLOW = XD6 / FACT_FLOW
        If Combo_T_FLUID = "Water" Then
            TUBES_LIQUID.Text = Format(XD6, "0")
            TUBES_WATER.Text = Format(XD6, "0")
            HScroll_TUBES_WATER.Max = Val(TUBES_WATER) / FACT_FLOW
            XD5 = XD6 / D19
        End If
        TUBES_OUT(7).Text = Format(XD36 * 0.001163, "##,##0") 'TUBES side duty,Kcal/h, KW
        
'        TUBES_FLOW.ForeColor = &HFFFFFF
'        TUBES_FLOW.BackColor = &HC0&
'        SHELL_FLOW.ForeColor = &HC0&
'        SHELL_FLOW.BackColor = &H80000018
'        TUBES_OUT(0).BackColor = &HC0&
'        TUBES_OUT(0).ForeColor = &HFFFFFF
'        SHELL_OUT(0).BackColor = &HE0E0E0
'        SHELL_OUT(0).ForeColor = &HC0&
'
'        XD6 = XD37 / (XD13 * RANGE_T)
'        XD36 = XD6 * XD13 * RANGE_T
'        TUBES_FLOW = Format(XD6, "0")
'        TUBES_LIQUID = Format(XD6, "0")
'        XD6L = XD6
'        TUBES_OUT(0) = Format(XD6, "##,##0")
'        HScroll_TUBES_FLOW = XD6 / FACT_FLOW
'        If Combo_T_FLUID = "Water" Then
'            TUBES_LIQUID.Text = Format(XD6, "0")
'            TUBES_WATER.Text = Format(XD6, "0")
'            HScroll_TUBES_WATER.Max = Val(TUBES_WATER) / FACT_FLOW
'        End If
    ElseIf Thermal_bal_shell = True Then
        COL = 1
        S_FLOW = 1
        SHELL_FLOW.ForeColor = &HFFFFFF
        SHELL_FLOW.BackColor = &HC0&
        TUBES_FLOW.ForeColor = &HC0&
        TUBES_FLOW.BackColor = &H80000018
        TUBES_TEMP_OUT.ForeColor = &HC0&
        TUBES_TEMP_OUT.BackColor = &H80000018
        SHELL_OUT(0).BackColor = &HC0&
        SHELL_OUT(0).ForeColor = &HFFFFFF
        TUBES_OUT(0).BackColor = &HE0E0E0
        TUBES_OUT(0).ForeColor = &HC0&
        
        XD18 = XD36 / (XD13S * RANGE_S)
        SHELL_FLOW.Text = Format(XD18, "0")
        SHELL_LIQUID.Text = Format(XD18, "0")
        XD18L = XD18
        SHELL_OUT(0) = Format(XD18, "##,##0")
        Spin_SHELL_FLOW = XD18 / FACT_FLOW
        If Combo_S_FLUID = "Water" Then
            SHELL_LIQUID.Text = Format(XD18, "0")
            HScroll_SHELL_LIQUID.Max = SHELL_FLOW / FACT_FLOW
            SHELL_WATER.Text = Format(XD18, "0")
            HScroll_SHELL_WATER.Max = Val(SHELL_WATER) / FACT_FLOW
        End If
        XD37 = XD18 * XD13S * RANGE_S      'Shell fluid side duty,Kcal/h
    Else
        SHELL_FLOW.ForeColor = &HC0&
        SHELL_FLOW.BackColor = &H80000018
        TUBES_FLOW.ForeColor = &HC0&
        TUBES_FLOW.BackColor = &H80000018
        TUBES_TEMP_OUT.ForeColor = &HC0&
        TUBES_TEMP_OUT.BackColor = &H80000018
        SHELL_OUT(0).BackColor = &HE0E0E0
        SHELL_OUT(0).ForeColor = &HC0&
        TUBES_OUT(0).BackColor = &HE0E0E0
        TUBES_OUT(0).ForeColor = &HC0&
    End If
    XD5 = XD6 / XD11

'CALCULATING TUBE SIDE PRESSURE DROP

'Pressure drop in tubes,kPA
    'Tube side friction factor   ft^2/inch^2
        XD91 = (XD54 * XPI * XD85 ^ 2 / 4) / XD57  'Flow area tubes,m^2
        If lungh = 2 And XD57 > 1 Then
            XD91 = XD91 * lungh
        End If
        XD92 = XD91 / (0.3048 ^ 2)                 'Flow area tubes,ft^2
        'Water velocity through tubes    m/s
        XD95 = (XD5 / (XD91 * 3600))               'Water velocity through tubes    m/s
        'Water velocity through tubes,ft/s
        XD96 = XD95 / 0.3048                       'Water velocity through tubes,ft/s
        'Reynolds number
        XD97 = XD95 * XD85 * XD11 / (XD12 * 0.001) 'Reynolds number
        XD93 = 10 ^ ((-2.5165 - 0.263 * Log(XD97) / 2.30258))
        XD94 = XD5 * XD11 * 2.20462 / XD92             'Mass velocity,lb/h(ft^2)
        XD98 = (XD93 * XD94 ^ 2 * XD56 * XD57) / (5.22 * 10 ^ 10 * XD84)
        XD99 = XD98 * 0.068947           'Pressure drop in tubes  bar
        XD100 = XD98 * 0.070307          'Pressure drop in tubes  Kg/cm2
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
    
    TUBES_SECTION.Text = Format(XD91, "0.0000")

    'Water velocity through tubes    m/s
        TUBES_OUT(2) = Format(XD95, "0.00")
    'Reynolds number through tubes
        TUBES_OUT(6) = Format(XD97, "##,##0")
    'Total pressure drop
        If Check_P_DROP_T = Unchecked Then
            TUBES_OUT(9) = Format(XD106_KPA, "0.00")
        Else
            XD106_KPA = (Val(TUBES_P_IN) - Val(TUBES_P_OUT)) * 100
            TUBES_OUT(9).Text = Format(XD106_KPA, "0.00") ' KPa
        End If
    C_F = XD5 / (XD106_KPA) ^ (1 / 2)
    C_Factor.Text = Format(C_F, "0")

'CALCULATING HEAT TRANSFER
        
    'Water side individual heat transfer coeficient  Btu/(h ft^2 F)
        XD108 = 150 * (1 + 0.011 * XD10) * (XD96 ^ 0.8 / XD83 ^ 0.2)
    'Water side individual heat transfer coefficient  Kcal/(h m^2 C)
        XD109 = XD108 * 4.882
    'Water side indiv. heat transfer coefficient referred to ext. surface Kcal/(h m^2 C)
        XD110 = XD109 * (XD85 / XD66)
    'Heat transfer resistance due to the wall    [(hm^2ºC)/Kcal]*10^4
        XD111 = (XD66 * Log(XD66 / XD85) / (2 * XD58)) * 10000
    'Heat transfer resistance due to outside fouling factor  [(h m^2 ºC)/Kcal]*10^4  1.00
        'XD112 = XD30
    'Heat transfer resistance due to water (tube side)   [(hm^2ºC)/Kcal]*10^4
        XD114 = 10 ^ 4 / XD110

PROP = "SHELL"
Call Properties
    
'CALCULATING h_o
    'Shell side crossflow area, ft^2
        XD68 = XD61 * (XD52 - XD50) * XD64 / (XD52 * 144)
    'Shell fluid loading  lb/h ft
        XD69 = XD18 * 2.20462 / (XD56 * XD54 ^ (2 / 3))
    'Shell side indiv. heat transfer coefficient,Kcal/(h m^2 C)
        'Shell side indiv. heat transfer coefficient,Btu/(h ft^2 F)
            'Wall temperature,ºC
    'Shell side indiv. heat transfer coefficient (guess one till Z_0=0),Kcal/(hm^2°C)
100              W = 0.1
110              j = 0.1
120              HE = W: GoSub 180
130              Y = X: HE = j + W
140              GoSub 180
150              G = W: W = G - j * Y / (X - Y)
160              If Abs(G - W) >= 0.00001 Then GoTo 120
170              W = HE: GoTo 190
180              XD70 = HE
            XD72 = XD9 + (XD70 / (XD110 + XD70)) * (XD9_S - XD9)
            'Shell side film temperature ºC
            XD73 = (XD20 + XD72) / 2
            'Shell side film temperature,K
            XD74 = 273.16 + XD73
    'Shell fluid
            'Shell fluid density at film temperature,Kg/m^3
                If Check_S_DENS = Unchecked Then
                    XD75 = SHELL_OUT(4)
                ElseIf Check_S_DENS = Checked Then
                    XD75 = HScroll_SHELL_DENS / 10
                End If
            'Shell fluid density at film temperature,lb/ft^3
                XD76 = XD75 * 2.20462 * (0.3048 ^ 3)
            'Shell fluid viscosity at condensing film temperature,cp
                If Check_S_VISC = Unchecked Then
                    XD77 = SHELL_OUT(5)
                ElseIf Check_S_VISC = Checked Then
                    XD77 = HScroll_SHELL_VISC / 1000
                End If
            'Shell fluid thermal conductivity at film temperature Kcal/h m ºC
                If Check_S_TC = Unchecked And Combo_S_FLUID = "Water" Then
                    XD79 = 0.00000000592317 * XD72 ^ 3 - 0.0000080425 * XD72 ^ 2 + 0.0018262 * XD72 + 0.478535
                    XD78 = XD79 / 1.488
                    SHELL_OUT(1) = XD79
                ElseIf Check_S_TC = Checked Then
                    XD79 = HScroll_SHELL_TC / 1000
                    XD78 = XD79 / 1.488
                Else
                    XD79 = SHELL_OUT(1)
                    XD78 = XD79 / 1.488
                End If
        If Combo_T_FLUID = "Water" Then
            'Shell side indiv. heat transfer coefficient,Btu/(h ft^2 F)
            XD80 = 1.5 * ((4 * XD69 / XD77) ^ -(1 / 3)) * (XD77 ^ 2 / (XD78 ^ 3 * XD76 ^ 2 * 9.81 * (3600 ^ 2 / 0.3048))) ^ -(1 / 3)
            'Shell side indiv. heat transfer coefficient,Kcal/(h m^2 C)
            XD81 = XD80 * 4.882
        ElseIf Combo_T_FLUID <> "Water" Then
            XD77E = XD77 * 2.42
            XD80E = 0.36 * XD78 / EQ_E31E * (EQ_E31E * EQ_E30E / XD77E) ^ 0.55 * (XD13S * XD77E / XD78) ^ (1 / 3) '  * (XD77E / XD77E) ^ 0.14
            XD81 = XD80E * 4.882
        End If
        X = (XD70 - XD81)
    Return

190 'XD70 = Shell side indiv. heat transfer coefficient (guess one till Z_0=0)
    If Combo_S_FLUID = "Water" Then
        If Check_S_DENS = Unchecked Then
            XD75 = D19                       'Water density at CALORIC TEMP,Kg/m3
        ElseIf Check_S_DENS = Checked Then
            XD75 = SHELL_OUT(4)              'Water density at CALORIC TEMP,Kg/m3
        End If
        If Check_S_VISC = Unchecked Then
            XD77 = D20                       'Water viscosity at CALORIC TEMP,centipoise
        ElseIf Check_S_VISC = Checked Then
            XD77 = SHELL_OUT(5)              'Water density at CALORIC TEMP,Kg/m3
        End If
        If Check_S_SPH = Unchecked Then
            XD13S = D21                      'Water specific heat at CALORIC TEMP,Kcal/(Kg ºC)
        ElseIf Check_S_SPH = Checked Then
            XD13S = SHELL_OUT(3)             'Water specific heat at CALORIC TEMP,Kcal/(Kg ºC)
        End If
        If Check_S_TC = Unchecked Then
            XD79 = TH_C                      'Shell fluid thermal conductivity at CALORIC TEMP, Kcal/h m ºC
        ElseIf Check_S_TC = Checked Then
            XD79 = SHELL_OUT(1)              'Shell fluid thermal conductivity at CALORIC TEMP, Kcal/h m ºC
            XD78 = XD79 / 1.488
        End If
    Else
            XD75 = SHELL_OUT(4)
            XD77 = SHELL_OUT(5)
            XD13S = SHELL_OUT(3)
            XD79 = SHELL_OUT(1)
            XD78 = XD79 / 1.488
    End If
    
    SHELL_OUT(1) = Format(XD79, "0.000")       'Shell fluid thermal conductivity at condensing film temperature Kcal/h m ºC
    SHELL_OUT(3) = Format(XD13S, "0.000")      'Shell specific heat
    SHELL_OUT(4) = Format(XD75, "0.0")         'Shell fluid density at condensing film temperature,Kg/m^3
    SHELL_OUT(5) = Format(XD77, "0.000")       'Shell fluid viscosity at condensing film temperature,cp
    SHELL_OUT(11) = Format(0, "0.00")          'Shell fluid film temperature ºC
    SKIN_TEMP = Format(XD72, "0.00")           'Wall temperature,ºC
    
    'SHELL VELOCITY
        L37 = XD61M                     'Val(SHELL_ID.Text)
        O37 = XD52M                     'Val(SHELL_TUBES_PITCH.Text)
        E37 = XD66M                     'Val(T_OD.Text)
        N37 = XD64M                     'Val(SHELL_BAFFLES_SPACE.Text)
        'SHELL CLEARANCE
        V37 = O37 / 1000 - E37 / 1000
        'SHELL FLOW AREA
        U37 = L37 / 1000 * V37 * N37 / 1000 / (O37 / 1000)
        k47 = XD18L / XD75 / U37 / 3600
        
        Clearance.Text = Format(V37, "0.000")
        Flow_area.Text = Format(U37, "0.0000")
        SHELL_OUT(2) = Format(k47, "0.00")
        
    'SHELL Reynolds
        EQ_D19 = XD52M / 1000 'SHELL_TUBES_PITCH / 1000
        EQ_PI = XPI
        EQ_D14 = XD66  'T_OD / 1000
        'Equivalent diameter, m
        If SHELL_PITCH_CONF = "Triangular" Then
            EQ_E31 = 4 * (EQ_D19 ^ 2 - EQ_PI * EQ_D14 ^ 2 / 4) / (XPI * EQ_D14)
        Else
            EQ_E31 = (4 * (0.5 * EQ_D19 * 0.866 * EQ_D19 - 0.5 * XPI * EQ_D14 ^ 2 / 4) / (0.5 * XPI * EQ_D14))
        End If
        EQ_E29 = U37                      'Flow_area
        EQ_E25 = Val(SHELL_LIQUID)        'SHELL_FLOW LIQUID
        EQ_E30 = EQ_E25 / EQ_E29
        EQ_E8 = XD77 * 3.6                'Shell fluid viscosity at condensing film temperature,cp
        EQ_E32 = EQ_E31 * EQ_E30 / EQ_E8
        Q_E22 = EQ_E31                    'Equivalent diameter
        Q_E17 = XD75                      'Shell fluid density at condensing film temperature,Kg/m^3
        Q_E27 = k47                       'Shell fluid velocity, m/s
        Q_E18 = XD77                      'Shell fluid viscosity at condensing film temperature,cp
        Q_E28 = Q_E22 * Q_E17 * Q_E27 / (Q_E18 * 0.001)
        SHELL_OUT(6) = Format(Q_E28, "##,##0")
    
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
        
        SHELL_OUT(9).Text = Format(P_E32, "0.00")      ' KPa
    '    Water_press_drop_bar.Text = Format(P_E32 / 100, "0.00")  ' bar
    Else
        P_E32 = (Val(SHELL_P_IN) - Val(SHELL_P_OUT))
        SHELL_OUT(9).Text = Format(P_E32 * 100, "0.00")      ' KPa
    '    Water_press_drop_bar.Text = Format(P_E32, "0.00")    ' bar
    End If
    
    If Combo_T_FLUID <> "Water" Then
        'Termal conductivity SHELL
            TUBES_OUT(1).Visible = True
            Check_T_TC.Visible = True
            HScroll_TUBES_TC.Visible = True
            lbl_tubes(1).Visible = True
            If Check_T_TC = Unchecked Then
                TH_C_TUBES = Val(TUBES_OUT(1))        'Shell fluid thermal conductivity at CALORIC TEMP, Kcal/h m ºC
            ElseIf Check_T_TC = Checked Then
                TH_C_TUBES = HScroll_TUBES_TC / 1000  'Shell fluid thermal conductivity at CALORIC TEMP, Kcal/h m ºC
            End If
        'Prandtl number
            PRANDTL = XD13 * XD12 * 3.6 / TH_C_TUBES
        'Film coefficient,Kcal/m2.h.C
            XD114A = 0.027 * TH_C_TUBES / XD85 * XD97 ^ 0.8 * PRANDTL ^ (1 / 3)
            XD114 = 10000 / (XD114A * XD85 / XD66)
    End If
    
    'Heat transfer resistance due to process (shell side),[(hm^2ºC)/Kcal]*10^4
        XD115 = (1 / XD81) * 10000
    'Overall CLEAN heat transfer coefficient Kcal/(h m^2 ºC)
        XD117 = 10000 / (XD111 + XD114 + XD115)
    'Water side fouling factor   [(hm^2ºC)/Kcal]*10^4
        'Overall heat transfer coefficient   Kcal/(h m^2 ºC)
            'Surface per linear ft, ft^2
            XD50 = Format(XD66 * 1000 / 25.4, "0.000")          'Tube outlet diameter,inch
            'Surface per linear m, m^2
            XD90 = 3.14159 * D67 * 25.4 * 10 ^ -3 * D75 * D74 * SERIES_N * PARALLEL_N * lungh
            'Log Mean Temperature Difference CORRECTED, ºC
            If Combo_T_FLUID = "Water" And XD7 < XD8 Then
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
                SS = AK6
                If T_PASS > 1 And SHELL_PASS > 1 Then
                    FT1 = (Sqr(RR ^ 2 + 1) / (2 * (RR - 1))) * Log((1 - SS) / (1 - RR * SS))
                    FT2 = 2 / SS - 1 - RR + (2 / SS) * Sqr((1 - SS) * (1 - RR * SS)) + Sqr(RR ^ 2 + 1)
                    FT3 = 2 / SS - 1 - RR + (2 / SS) * Sqr((1 - SS) * (1 - RR * SS)) - Sqr(RR ^ 2 + 1)
                    FT4 = Log(FT2 / FT3)
                    FT = FT1 / FT4
                ElseIf T_PASS > 1 And SERIES_N > 1 Then
                    FT1 = (Sqr(RR ^ 2 + 1) / (2 * (RR - 1))) * Log((1 - SS) / (1 - RR * SS))
                    FT2 = 2 / SS - 1 - RR + (2 / SS) * Sqr((1 - SS) * (1 - RR * SS)) + Sqr(RR ^ 2 + 1)
                    FT3 = 2 / SS - 1 - RR + (2 / SS) * Sqr((1 - SS) * (1 - RR * SS)) - Sqr(RR ^ 2 + 1)
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
            ElseIf Combo_S_FLUID = "Water" And XD19 < XD20 Then
                If Combo_CURRENT = "Counter-flow" Then
                    AG6 = ((XD7 - XD20) - (XD8 - XD19)) / Log((XD7 - XD20) / (XD8 - XD19))
                ElseIf Combo_CURRENT = "Cross-flow" Then
                    AG6 = ((XD8 - XD19) - (XD7 - XD20)) / Log((XD8 - XD19) / (XD7 - XD20))
                ElseIf Combo_CURRENT = "Parallel-flow" Then
                    AG6 = ((XD7 - XD19) - (XD8 - XD20)) / Log((XD7 - XD19) / (XD8 - XD20))
                End If
                AJ6 = (XD7 - XD8) / (XD20 - XD19)
                AK6 = (XD20 - XD19) / (XD7 - XD19)
                RR = AJ6
                SS = AK6
                If T_PASS > 1 And SHELL_PASS > 1 Then
                    FT1 = (Sqr(RR ^ 2 + 1) / (2 * (RR - 1))) * Log((1 - SS) / (1 - RR * SS))
                    FT2 = 2 / SS - 1 - RR + (2 / SS) * Sqr((1 - SS) * (1 - RR * SS)) + Sqr(RR ^ 2 + 1)
                    FT3 = 2 / SS - 1 - RR + (2 / SS) * Sqr((1 - SS) * (1 - RR * SS)) - Sqr(RR ^ 2 + 1)
                    FT4 = Log(FT2 / FT3)
                    FT = FT1 / FT4
                ElseIf T_PASS > 1 And SERIES_N > 1 Then
                    FT1 = (Sqr(RR ^ 2 + 1) / (2 * (RR - 1))) * Log((1 - SS) / (1 - RR * SS))
                    FT2 = 2 / SS - 1 - RR + (2 / SS) * Sqr((1 - SS) * (1 - RR * SS)) + Sqr(RR ^ 2 + 1)
                    FT3 = 2 / SS - 1 - RR + (2 / SS) * Sqr((1 - SS) * (1 - RR * SS)) - Sqr(RR ^ 2 + 1)
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
                    XD32 = XD7 - XD20
                    Label22(15).Caption = "(T1 - t2)"
                    If XD32 > (XD8 - XD19) Then
                        XD32 = XD8 - XD19
                        Label22(15).Caption = "(T2 - t1)"
                    End If
                ElseIf XD57 > 1 Then
                    XD32 = XD8 - XD20
                    Label22(15).Caption = "(T2 - t2)"
                End If
            End If
            XD31 = AH6
            XD38 = XD36 / (XD90 * XD31)
    xd118 = ((1 / XD38) - (1 / XD117) - (XD112 / 10000)) * 10000 * (XD85 / XD66)
    'Total heat transfer resistance  [(h m^2 ºC)/Kcal]*10^4
    XD116 = 10000 / XD38
    
    'Heat transfer resistance due to inside fouling factor,[(hm^2ºC)/Kcal]*10^4
    XD113 = xd118 * (XD66 / XD85)
    
    'TUBES HEAT FLUX
        Q6 = XD36 / XD90
        TUBES_OUT(8).Text = Format(Q6 * 0.001163, "0.00")
    'SHELL HEAT FLUX
    '    Check_S_SPH = Checked
        S_SPH = SHELL_OUT(3).Text
        Q6S = XD37 / XD90
        SHELL_OUT(8).Text = Format(Q6S * 0.001163, "0.00")
    
    'Area, m^2
        Area.Text = Format(XD90, "0.00")
    'Log Mean Temperature Difference, ºC
        LMTD.Text = Format(AG6, "0.00")
    'Log Mean Temperature Difference corrected, ºC
        MTDc.Text = Format(AH6, "0.00")
        
    TTD = Format(XD32, "0.00")
    TUBES_OUT(7).Text = Format(XD36 * 0.001163, "0") 'Water side duty,Kcal/h, KW
    SHELL_OUT(7).Text = Format(XD37 * 0.001163, "0") 'Shell fluid side duty, KW
    If Check_U_CLEAN = Checked Then
        U_COEFF_CLEAN = HScroll_U_CLEAN
    ElseIf Check_U_CLEAN = Unchecked Then
        U_COEFF_CLEAN.Text = Format(XD117, "0.0")      'Overall CLEAN heat transfer coefficient Kcal/(h m^2 ºC)
    End If
    U_COEFF_DIRTY.Text = Format(XD38, "0.0")       'Overall heat transfer coefficient   Kcal/(h m^2 ºC)
    
    'TUBES_FF = xd118         'Water side fouling factor   [(hm^2ºC)/Kcal]*10^3
    TUBES_FF.Text = Format(xd118, "0.000")         'Water side fouling factor   [(hm^2ºC)/Kcal]*10^3
    'TUBES_FF.Text = Format(Val(XD118), "0.000")         'Water side fouling factor   [(hm^2ºC)/Kcal]*10^3

End Sub
Private Sub Steam()
On Error Resume Next
    Lab_COND.Caption = "Steam Exhaust Condenser"
    Combo_CURRENT.Text = "Condensation"
    SHELL_TEMP_IN.Visible = False
    SHELL_TEMP_OUT.Visible = False
    HScroll_SHELL_T_IN.Visible = False
    HScroll_SHELL_T_OUT.Visible = False
    SHELL_P_IN.Visible = False
    SHELL_P_OUT.Visible = False
    HScroll_SHELL_P_IN.Visible = False
    HScroll_SHELL_P_OUT.Visible = False
    Check_P_DROP_S.Visible = False

'SPECIFIC HEAT
    Check_S_SPH.Visible = False
    SHELL_OUT(3).Visible = False
    HScroll_SHELL_SPH.Visible = False
'Termal conductivity SHELL
    lbl_tubes(9).Visible = False
    SHELL_OUT(1).Visible = False
    Check_S_TC.Visible = False
    HScroll_SHELL_TC.Visible = False
    lbl_tubes(1).Visible = False
'Termal conductivity TUBES
    TUBES_OUT(1).Visible = False
    Check_T_TC.Visible = False
    HScroll_TUBES_TC.Visible = False
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
'Wet steam
    Label22(17).Visible = True
    Wet_steam.Visible = True
    HScroll_WET_STEAM.Visible = True
    Check_WET_STEAM.Visible = True
    Check_water_steam.Visible = True
'Vapor percent
    Frame_VAP.Visible = True
'Cleanliness factor
    lblLabels(3).Caption = "Cleanliness factor"
    Label24(0).Caption = "%  (Norm: 85 - 95 %)"
lblLabels(25).Caption = "Terminal temperature:"
Thermal_bal_tubes.Visible = True
Thermal_bal_shell.Visible = True
Check_T_OUT.Visible = True


    VAP_PERC = 100
    LIQ_PERC = 0
    Spin_VAP_P = 1000

Call Mechanical
    Mat_cond.Text = D78                                                    'Thermal conductivity of tube material
    If Check_MAT_FACTOR = Checked Then
        D54 = Spin_MAT_FACTOR / 100
        Mat_factor = Format(D54, "0.00")
    ElseIf Check_MAT_FACTOR = Unchecked Then
        Mat_factor = Format(D54, "0.00")
    End If

'Surface per linear m, m^2
    D79 = 3.14159 * D67 * 25.4 * 10 ^ -3 * D75 * D74 * SERIES_N * PARALLEL_N * lungh
    D80 = D79 / (0.3048 ^ 2)                'inch^2
    Area.Text = Format(D79, "0.0")          'Surface, m^2

    If Check_des = Checked Then
        D13 = Val(TUBES_TEMP_OUT.Text)          'Water T_OUT, °C
    End If
    D11 = Val(TUBES_TEMP_IN.Text)           'Water T_IN, °C
    D9 = TUBES_FLOW.Text                    'Water flow rate, kg/h
    D12 = D11 * 1.8 + 32                    'Water T_IN, °F
''Saturated steam flowrate INLET condenser
'    D29 = Val(SHELL_FLOW.Text)              'kg/h                                                               'kg/h
'    D30 = D29 * 2.20462                     'lb/h
''Steam loading
'    D31 = D30 / D80                         'lb/h ft2
'Steam condensation pressure
    D32 = Val(S_press_KP.Text) / 100        'bar
    If D32 = 0 Then D32 = 0.05              'bar

PROP = "TUBES"
Call Properties

    If Check_T_DENS = Checked Then
        D19 = TUBES_OUT(4)              'Water density at CALORIC TEMP,Kg/m3
    End If
    If Check_T_VISC = Checked Then
        D20 = TUBES_OUT(5)              'Water density at CALORIC TEMP,Kg/m3
    End If
    If Check_T_SPH = Checked Then
        D21 = TUBES_OUT(3)              'Water specific heat at CALORIC TEMP,Kcal/(Kg ºC)
    End If
    If Check_T_TC = Checked Then
        TH_C = TUBES_OUT(1)             'Shell fluid thermal conductivity at CALORIC TEMP, Kcal/h m ºC
        XD78 = XD79 / 1.488
    End If
    TUBES_OUT(4).Text = Format(Val(D19), "0.0")       'Steam density
    TUBES_OUT(5).Text = Format(Val(D20), "0.000")     'Steam viscosity
    TUBES_OUT(3).Text = Format(Val(D21), "0.000")     'Steam specific heat
    TUBES_OUT(1).Text = Format(Val(TH_C), "0.000")    'Thermal conductivity
'Latent heat, Kcal/kg
    If Check_LATENT = Unchecked Then
        I9 = 0.168682569821809
        J9 = -1.80896828868017E-04
        J3 = -38.2917529410035
        D38 = (-I9 - Sqr(I9 ^ 2 - 4 * J9 * (J3 - Log(D32) / 2.3))) / (2 * J9)
        SHELL_OUT(10).Text = Format(D38 * 4.1868, "0.00")  'Latent heat, KJ/kg
    ElseIf Check_LATENT = Checked Then
        D38 = SHELL_OUT(10) / 4.1868
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
    SHELL_OUT(10).Text = Format(D38 * 4.1868, "0.00") 'Steam latent heat ,Kcal/(Kg ºC)
    SHELL_OUT(11).Text = Format(D37, "0.00")          'Condensing temperature,°C
    SHELL_TEMP_IN.Text = Format(D37, "0.00")          'Condensing temperature,°C
    SHELL_TEMP_OUT.Text = Format(D37, "0.00")         'Condensing temperature,°C

Call FOULING

'DELTA DUTY test
If Check_des = Checked Then
    Check_WET_STEAM = Unchecked
    Check_T_OUT.Visible = False
'    Thermal_bal_tubes.Visible = False
'    Thermal_bal_shell.Visible = False
    TUBES_FLOW.ForeColor = &HC0&
    TUBES_FLOW.BackColor = &H80000018
    TUBES_TEMP_OUT.ForeColor = &HC0&
    TUBES_TEMP_OUT.BackColor = &H80000018
    'Water side duty
        D43 = D9 * D21 * (D13 - D11)
    '% of wet steam
        D33 = 100 - ((D43 * 100) / (D29 * D38))
    'Steam side duty
        D44 = D29 * D38 * (100 - D33) / 100
    'Water flow calculated
        flow_calc = D44 / ((D13 - D11) * D21)
    'Steam flow calculated
        SHELL_FLOW_CALC = D29
Else
    Check_T_OUT.Visible = True
    If Thermal_bal_tubes = True Then
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
            D33 = HScroll_WET_STEAM.Value / 10
            D44 = D29 * D38 * (100 - D33) / 100
            D13 = D44 / (D9 * D21) + D11
            D43 = D9 * D21 * (D13 - D11)
            flow_calc = D9 'D44 / ((D13 - D11) * D21)
            SHELL_FLOW_CALC = D29
            TUBES_TEMP_OUT.Text = Format(D13, "0.00")
            YXY = 1
            Spin_TUBES_T_OUT = D13 * 100
        ElseIf Check_T_OUT = 1 Then
            Check_T_FLOW = Unchecked
            TUBES_TEMP_OUT.ForeColor = &HC0&
            TUBES_TEMP_OUT.BackColor = &H80000018
            TUBES_FLOW.ForeColor = &HFFFFFF
            TUBES_FLOW.BackColor = &HC0&
            D33 = HScroll_WET_STEAM.Value / 10
            D44 = D29 * D38 * (100 - D33) / 100
            D9 = D44 / (D21 * (D13 - D11))
            D43 = D9 * D21 * (D13 - D11)
            flow_calc = D9
            SHELL_FLOW_CALC = D29
            TUBES_TEMP_OUT.Text = Format(D13, "0.00")
            TUBES_FLOW = Format(flow_calc, "0")
            TUBES_LIQUID = Format(flow_calc, "0")
            TUBES_WATER = Format(flow_calc, "0")
        End If
    ElseIf Thermal_bal_shell = True Then
        TUBES_TEMP_OUT.ForeColor = &HC0&
        TUBES_TEMP_OUT.BackColor = &H80000018
        TUBES_FLOW.ForeColor = &HC0&
        TUBES_FLOW.BackColor = &H80000018
        
        SHELL_OUT(0).BackColor = &HC0&
        SHELL_OUT(0).ForeColor = &HFFFFFF
        TUBES_OUT(0).BackColor = &HE0E0E0
        TUBES_OUT(0).ForeColor = &HC0&
    
        D33 = HScroll_WET_STEAM.Value / 10
        D43 = D9 * D21 * (D13 - D11)
        D29 = D43 / (D38 * (100 - D33) / 100)
        D44 = D29 * D38 * (100 - D33) / 100
        flow_calc = D44 / ((D13 - D11) * D21)
        SHELL_FLOW_CALC = D29
    Else
        TUBES_OUT(0).ForeColor = &HC0&
        TUBES_OUT(0).BackColor = &HE0E0E0
        SHELL_OUT(0).ForeColor = &HC0&
        SHELL_OUT(0).BackColor = &HE0E0E0
        
        TUBES_TEMP_OUT.ForeColor = &HC0&
        TUBES_TEMP_OUT.BackColor = &H80000018
        TUBES_FLOW.ForeColor = &HC0&
        TUBES_FLOW.BackColor = &H80000018
        
        D33 = HScroll_WET_STEAM.Value / 10
        D44 = D29 * D38 * (100 - D33) / 100
        D43 = D9 * D21 * (D13 - D11)
        flow_calc = D9
        SHELL_FLOW_CALC = D29
    End If
End If

    YXY = 0
    TUBES_FLOW = Format(D9, "0")
    TUBES_LIQUID = Format(flow_calc, "0")
    TUBES_WATER = Format(flow_calc, "0")
    SHELL_FLOW = Format(D29, "0")
    SHELL_LIQUID = Format(D29, "0")
    SHELL_WATER = Format(D29, "0")
    
    TUBES_OUT(0).Text = Format(flow_calc, "##,##0")
    SHELL_OUT(0) = Format(SHELL_FLOW_CALC, "##,##0")
    TUBES_OUT(7).Text = Format(D43 / 859.845, "##,##0")                         'MW
    SHELL_OUT(7).Text = Format(D44 / 859.845, "##,##0")
    Wet_steam.Text = Format(D33, "0.0")

'Saturated steam flowrate INLET condenser
'    D29 = Val(SHELL_FLOW.Text)              'kg/h                                                               'kg/h
    D30 = D29 * 2.20462                     'lb/h
'Steam loading
    D31 = D30 / D80                         'lb/h ft2
'Water flow rate m3/h
    D15 = D9 / D19      'm3/h
'Water velocity through tubes
    XD91 = 3.14159 / 4 * D73 ^ 2 * D74 / D77
    If lungh = 2 And D77 > 1 Then
        XD91 = XD91 * lungh
    End If
    TUBES_SECTION.Text = Format(XD91, "0.0000")
    D22 = D15 / XD91 / 3600
    If lungh = 2 And D77 > 1 Then
        D22 = D22 / lungh
    End If
    D23 = D22 / 0.3048                                                                                    'fps
    TUBES_OUT(2).Text = Format(D22, "0.00")
'Reynolds through tubes
    D24 = D22 * D73 * D19 / (D20 / 1000)
    TUBES_OUT(6).Text = Format(D24, "0,00")
'Pressure drop through tubes (Hazen-Williams with C=130) Related to CS
    D25 = ((6.05 * 10 ^ 5 * ((D15 * 1000 / 60) / (D74 / D77)) ^ 1.85) / (130 ^ 1.85 * (D73 * 1000) ^ 4.87)) * D75 * D77
'Pressure drop due to return (estimated four velocity heads)
    D26 = ((4 * (D22 ^ 2 / (2 * 9.81))) / 10) * 0.9807
'Total pressure drop for 100% clean tube side
    D27 = D25 + D26
    TUBES_OUT(9).Text = Format(D27 * 100, "0.00")                            ' KPa
    If Check_P_DROP_T = Unchecked Then
        TUBES_OUT(9) = Format(D27, "0.00") 'Total pressure drop
    Else
        D27 = (Val(TUBES_P_IN) - Val(TUBES_P_OUT)) * 100
        TUBES_OUT(9).Text = Format(D27, "0.00")        ' KPa
    End If
'C Factor
    C_Factor.Text = Format(D15 / (D27 * 100) ^ (1 / 2), "0.0")                      'm3/h/kPa

'Log Mean Temperature Difference
    D41 = ((D37 - D11) - (D37 - D13)) / Log((D37 - D11) / (D37 - D13))
    LMTD.Text = Format(D41, "0.00")
    MTDc = Format(D41, "0.00")
'Terminal temperature difference
    D42 = D37 - D13
    TTD.Text = Format(D42, "0.00")
    Label22(15).Caption = "(T2 - t2)"
'TUBES HEAT FLUX
    Q6 = D43 / D79 * 0.001163
    TUBES_OUT(8).Text = Format(Q6, "0.00")
'SHELL HEAT FLUX
    S_SPH = SHELL_OUT(3).Text
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
    U_COEFF_CLEAN.Text = Format(Val(D55) * 1, "0.0")
'Overall DIRTY heat transfer coefficient, kcal/(h m^2 ºC)
    D56 = D44 / (D41 * D79)
    U_COEFF_DIRTY.Text = Format(Val(D56), "0.0")
'CLEANLINESS FACTOR
    D57 = D56 * 100 / D55
    C_Factor.Text = Format(Val(D57), "0.0") 'CLEANLINESS FACTOR
    If D57 > 100 Or D57 <= 0 Then
        CF.BackColor = &HFF&
    Else
        CF.BackColor = &H8000000F
    End If
'Water side individual heat transfer coefficient referred to ext. surface
    D58 = (150 * (1 + 0.011 * D18) * (D23 ^ 0.8 / D72 ^ 0.2)) * 4.882 * (D72 / D67)             '[(h m^2 ºC)/Kcal]*10^-4
'Heat transfer resistance due to water flowing inside the tubes
    D59 = 10000 / D58
'Heat transfer resistance due to  the wall, referred to ext. surface
    D60 = (D68 * Log(D68 / D73) / (2 * D78)) * 10000
'Heat transfer resistance due to the condensing steam film
    D61 = (10000 / D55) - D59 - D60
'Heat transfer resistance due to outside fouling factor
    D62 = D40 * 10000
'Total heat transfer resistance, referred to external surface
    D64 = 10000 / D56
'Heat transfer resistance due to inside fouling factor referred to ext surface
    D63 = D64 - (D59 + D60 + D61 + D62)
'Water side fouling factor
    D65 = D63 * (D73 / D68)
    TUBES_FF.Text = Format(Val(D65), "0.000")                     '[(m^2 ºC)/KW]*10^-4

PROP = "SHELL"
Call Properties

'Skin Temperature - [((KW/m2)/(mps)*3.739)+(tout)]
    HO = 1 / D61 * 10000
    HIO = 1 / D59 * 10000
    SKIN_T = XD9 + HO / (HO + HIO) * (XD9_S - XD9)
    SKIN_TEMP = Format(SKIN_T, "0.00")

    If Check_S_DENS = Checked Then
        D19 = SHELL_OUT(4)              'Water density at CALORIC TEMP,Kg/m3
    End If
    If Check_S_VISC = Checked Then
        D20 = SHELL_OUT(5)              'Water density at CALORIC TEMP,Kg/m3
    End If
    If Check_S_SPH = Checked Then
        D21 = SHELL_OUT(3)              'Water specific heat at CALORIC TEMP,Kcal/(Kg ºC)
    End If
    If Check_S_TC = Checked Then
        TH_C = SHELL_OUT(1)             'Shell fluid thermal conductivity at CALORIC TEMP, Kcal/h m ºC
        XD78 = XD79 / 1.488
    End If
    If Check_LATENT = Checked Then
        D38 = SHELL_OUT(10) / 4.1868    'Steam latent heat ,Kcal/Kg
    End If
    If Check_CT = Checked Then
        D37 = SHELL_OUT(11)             'Condensing temperature,°C
    End If
    SHELL_OUT(4).Text = Format(Val(D19), "0.0")       'Steam density
    SHELL_OUT(5).Text = Format(Val(D20), "0.000")     'Steam viscosity
    SHELL_OUT(3).Text = Format(Val(D21), "0.000")     'Steam specific heat
    SHELL_OUT(1).Text = Format(Val(TH_C), "0.000")    'Thermal conductivity

'SHELL VELOCITY
    L37 = XD61M                 'SHELL_ID
    O37 = XD52M                 'SHELL_TUBES_PITCH
    E37 = XD66M                 'T_OD
    N37 = XD64M                 'SHELL_BAFFLES_SPACE
'SHELL CLEARANCE
    V37 = O37 / 1000 - E37 / 1000
    Clearance.Text = V37
    'SHELL FLOW AREA
    U37 = L37 / 1000 * V37 * N37 / 1000 / (O37 / 1000)
    Flow_area.Text = Format(U37, "0.0000")
    k47 = XD18L / D19 / U37 / 3600
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
    EQ_E29 = U37                    'Flow_area
    EQ_E25 = Val(SHELL_LIQUID)      'SHELL_FLOW LIQUID
    EQ_E30 = EQ_E25 / EQ_E29
    EQ_E8 = D20 * 3.6               'SHELL_OUT(5) * 3.6
    EQ_E32 = EQ_E31 * EQ_E30 / EQ_E8
    Q_E22 = EQ_E31 * 1000
    Q_E17 = D19                      'Density, kg/m3
    Q_E27 = k47                      'Shell flow velocity, m/s
    Q_E18 = D20                      'SHELL viscosity, cP
    Q_E28 = Q_E22 * Q_E17 * Q_E27 / D20
    SHELL_OUT(6) = Format(Q_E28, "##,##0")

'CALCULATING SHELL SIDE PRESSURE DROP
    If Check_P_DROP_S = Unchecked Then
        'Pressure drop (tubes)
        P_E17 = D19              'SHELL_OUT(4)
        P_E22 = EQ_E31 * 1000    'Equivalent diameter, mm
        P_E27 = Q_E27            'Shell flow velocity, m/s
        P_E23 = XD55 * lungh     'T_len, mm
        P_E28 = Q_E28
        P_E29 = 0.44 * P_E28 ^ -0.19
        P_E30 = 4 * P_E29 * P_E23 * P_E27 ^ 2 / (P_E22 * 2 * 9.8) * P_E17 * 0.000096784 * 101.325
    'Pressure drop (sheet)
        P_E9 = XD59               'SHELL PASSES
        P_E31 = 3 * P_E9 * P_E27 ^ 2 / 2 / 9.8 * P_E17 * 0.000096784 * 101.325
        P_E32 = P_E30 + P_E31
        SHELL_OUT(9).Text = Format(P_E32, "0.00")      ' KPa
    Else
        P_E32 = (Val(SHELL_P_IN) - Val(SHELL_P_OUT))
        SHELL_OUT(9).Text = Format(P_E32 * 100, "0.00")      ' KPa
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
XD18L = SHELL_LIQUID            'Shell liquid fluid flowrate,Kg/h
XD19 = SHELL_TEMP_IN            'Shell fluid temperature in,ºC
XD20 = SHELL_TEMP_OUT           'Shell fluid temperature out,ºC
XD52M = Val(SHELL_TUBES_PITCH)  'Pitch, mm
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
XD85 = Val(T_ID) / 1000         'Tube Inlet diameter, m
XD84 = XD85 / 0.3048            'Tube Inlet diameter,ft
XD83 = XD85 / 25.4 * 1000       'Tube Inlet diameter,inches
XD112 = SHELL_FF                'Process side fouling factor [(hm^2ºC)/Kcal]*10^4

'WATER
'TUBES TEMP IN/OUT
    D11 = XD7                           'TUBES T_IN, °C
    D12 = D11 * 1.8 + 32                'Water T_IN, °F
    D13 = XD8                           'TUBES T_OUT, °C

'TUBES RANGE
    RANGE_T = D13 - D11
    
'SHELL TEMP IN/OUT
    XD19 = XD19                          'SHELL T_IN, °C
    XD20 = XD20                          'SHELL T_OUT, °C

'SHELL RANGE
    RANGE_S = XD19 - XD20

''TUBES flowrate INLET
'    D9 = TUBES_FLOW.Text                'Water flow rate, kg/h
'    D15 = D9 / D19

'SHELL flowrate INLET
    D29 = Val(SHELL_FLOW.Text)          'kg/h
    D30 = D29 * 2.20462                 'lb/h
    D15S = D29 / D19

If PROP = "TUBES" Then
    'TUBES Caloric temperature,ºC
        If Combo_T_FLUID.Text = "Water" Then
            D17 = D11 + (D13 - D11) / 2
        Else
            D17 = D13 + (D11 - D13) / 2
        End If
        XD9 = D17                     'TUBES Caloric temperature,ºC
ElseIf PROP = "SHELL" Then
    'SHELL Caloric temperature,ºC
        If Combo_S_FLUID.Text = "Water" Then
            D17 = XD20 + (XD19 - XD20) / 2
        Else
            D17 = XD19 + (XD20 - XD19) / 2
        End If
        XD9_S = D17                       'Caloric shell temperature
End If

'Thermal conductivity, Kcal/h m ºC
TH_C = 0.00000000592317 * D17 ^ 3 - 0.0000080425 * D17 ^ 2 + 0.0018262 * D17 + 0.478535

'Viscosity of water, cP
    D20 = (100 / (2.1482 * ((273.16 + D17 - 281.435) + Sqr(8078.4 + (273.16 + D17 - 281.435) ^ 2)) - 120))
  
'Caloric tubes-side temperature,°C
    D17_2 = Int(D17 / 2)
'Caloric tubes-side temperature,°F
    D18 = D17 * 1.8 + 32
    
'Density of water, kg/m3
Select Case D17_2
    Case 1: D19 = 999.94
    Case 2: D19 = 999.97
    Case 3: D19 = 999.94
    Case 4: D19 = 999.85
    Case 5: D19 = 999.7
    Case 6: D19 = 999.497
    Case 7: D19 = 999.244
    Case 8: D19 = 998.943
    Case 9: D19 = 998.595
    Case 10: D19 = 998.204
    Case 11: D19 = 997.77
    Case 12: D19 = 997.296
    Case 13: D19 = 996.783
    Case 14: D19 = 996.233
    Case 15: D19 = 995.647
    Case 16: D19 = 995.026
    Case 17: D19 = 994.371
    Case 18: D19 = 993.684
    Case 19: D19 = 992.965
    Case 20: D19 = 992.215
    Case 21: D19 = 991.436
    Case 22: D19 = 990.628
    Case 23: D19 = 989.792
    Case 24: D19 = 988.928
    Case 25: D19 = 988.037
    Case 26: D19 = 987.12
    Case 27: D19 = 986.177
    Case 28: D19 = 985.219
    Case 29: D19 = 984.217
End Select
    
'TUBES flowrate INLET
    D9 = TUBES_FLOW.Text                'Water flow rate, kg/h
    D15 = D9 / D19
    
'Specific heat of water
Select Case D17_2
    Case 1: D21 = 1.00636
    Case 2: D21 = 1.00495
    Case 3: D21 = 1.00378
    Case 4: D21 = 1.00277
    Case 5: D21 = 1.00194
    Case 6: D21 = 1.00124
    Case 7: D21 = 1.00067
    Case 8: D21 = 1.00019
    Case 9: D21 = 0.999978
    Case 10: D21 = 0.99947
    Case 11: D21 = 0.99921
    Case 12: D21 = 0.99902
    Case 13: D21 = 0.99885
    Case 14: D21 = 0.99873
    Case 15: D21 = 0.99866
    Case 16: D21 = 0.99861
    Case 17: D21 = 0.99859
    Case 18: D21 = 0.99861
    Case 19: D21 = 0.99864
    Case 20: D21 = 0.99869
    Case 21: D21 = 0.99876
    Case 22: D21 = 0.99883
    Case 23: D21 = 0.99895
    Case 24: D21 = 0.99907
    Case 25: D21 = 0.99919
    Case 26: D21 = 0.99935
    Case 27: D21 = 0.9995
    Case 28: D21 = 0.99969
    Case 29: D21 = 0.99988
End Select

End Sub
Private Sub FOULING()
On Error Resume Next
Dim S_side(40), T_side(10), P_FF(40)

'Process fouling
'D40 = 0
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
    S_side(13) = "High-boiling hydrocarbons"
    S_side(14) = "Low-boiling hydrocarbons"
    S_side(15) = "Steam"
    S_side(16) = "Steam condensing"
    S_side(17) = "Air, N2 etc (compressed)"
    S_side(18) = "Propane, Butane, etc."
    S_side(19) = "Water"
    S_side(20) = "Other"

    T_side(1) = "Water"
    T_side(2) = "Air, N2 etc (compressed)"
    T_side(3) = "Steam condensing"
    T_side(4) = "Feed Water"

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
    P_FF(17) = 0.001025
    P_FF(18) = 0.0003074
    P_FF(19) = 0.0006148
    P_FF(20) = 0.0001

    Shell_side = Combo_S_FLUID.Text
    Tube_side = Combo_T_FLUID.Text
     
    If Combo_S_FLUID.Text = "Water" And Tube_side = Combo_T_FLUID <> "Water" Then
        Shell_side = Combo_T_FLUID.Text
        Tube_side = Combo_S_FLUID.Text
    End If
    For i = 0 To 20
        If Shell_side = S_side(i) Then
            If Tube_side = "Water" Then
                D40 = P_FF(i)
            ElseIf Shell_side = "Water" Then
                If Tube_eside = "Air, N2 etc (compressed)" Then
                    D40 = 0.001025
                ElseIf Tube_side = "Steam condensing" Then
                    D40 = 0.0003074
                Else
                    D40 = P_FF(i)
                End If
            Else
                D40 = P_FF(i)
            End If
        End If
    Next i
    If D40 <= 0 And foul = 0 Then
        D40 = 0.0001
'        MsgBox "Process fouling not found. Please enter the value"
    End If
    SHELL_FF.Text = Format(D40 * 10000, "0.00")
End If
End Sub
Private Sub Mechanical()
On Error Resume Next
'MECHANICAL DATA
    
D68 = Val(T_OD.Text) / 1000      'Tube Outlet Diameter (m)
D67 = D68 * 1000 / 25.4          'Tube Outlet Diameter (inches)
D69 = Val(Combo_BWG.Text)        'BWG

'BWG / Wall thickness
Select Case D69
    Case 7:  D70 = 0.18       'Wall Thickness (inches)
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
D71 = D70 * 25.4 * 10 ^ -3          'Wall Thickness (m)
D72 = D67 - 2 * D70                 'Tube inlet diameter (inches)
D73 = D72 * 25.4 * 10 ^ -3          'Tube inlet diameter (meters)
T_ID.Text = Format(Val(D73 * 1000), "0.00")
T_ID_E = Format(D73 * 1000, "0.00")

D74 = Val(T_NO.Text)                'Number of tubes
D75 = Val(T_len.Text)               'Tube lenght (m)
D76 = D75 / 0.3048                  'Tube lenght (inches)
D77 = Val(T_PASS.Text)              'Number of tube side passes
    
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
If D54 = 0 And Combo_S_FLUID = "Steam" Then
    D54 = Mat_factor.Text
    Mat_factor.ForeColor = &HFFFF&
    Mat_factor.BackColor = &HFF&
    lblLabels(16).Visible = True
    lblLabels(16).Caption = "Enter the value (suggested: 0.5 - 1.0)"
ElseIf D54 > 0 And Combo_S_FLUID = "Steam" Then
    Mat_factor.Text = D54                 ' Material factor
    lblLabels(16).Visible = False
    Mat_factor.ForeColor = &HC0&
    Mat_factor.BackColor = &HE0E0E0
ElseIf Check_MAT_FACTOR = Checked And Combo_S_FLUID = "Steam" Then
    D54 = Spin_MAT_FACTOR / 100
    Mat_factor = Spin_MAT_FACTOR / 100
    lblLabels(16).Visible = True
End If
End Sub
Private Sub Combo_Plant_1_LostFocus()
On Error Resume Next
        XXX = 1
        Data1.Recordset.MoveFirst
        While pos < 0
            Data1.Recordset.MoveNext
        Wend
        PPP1 = Combo_Plant_1.Text
        While PPP1 <> PPP2
            PPP2 = Data1.Recordset.Plant
            If PPP1 = PPP2 Then
                n_rec_a = Data1.Recordset.AbsolutePosition + 1
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
            Data1.Recordset.MoveNext
        Wend
10 End Sub
Private Sub Combo_UNIT_1_Lostfocus()
On Error Resume Next
        XXX = 1
        Data1.Recordset.MoveLast
        n_record = Data1.Recordset.RecordCount
        Data1.Recordset.MoveFirst
       
        While pos < 0
            Data1.Recordset.MoveNext
        Wend
        PPP1 = Combo_Plant_1.Text
        UUU1 = Combo_UNIT_1.Text
        Do Until Data1.Recordset.EOF
            n_rec_a = Data1.Recordset.AbsolutePosition + 1
            PPP2 = Data1.Recordset.Plant
            UUU2 = Data1.Recordset.Unit_name
            If UUU1 = UUU2 And PPP1 = PPP2 Then
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
        Loop
12  End Sub
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
Private Sub Com_Go_Click()
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
                ELEVATION = Data2.Recordset.ELEVATION
                PARALLEL_N = Data2.Recordset.PARALLEL_N
                SERIES_N = Data2.Recordset.SERIES_N
                Combo_CURRENT = Data2.Recordset.CURRENT
                 
                T_NO.Text = Val(Data2.Recordset.TUBES_NO)
                T_len.Text = Val(Data2.Recordset.TUBES_LE)
                T_PASS.Text = Val(Data2.Recordset.TUBES_PASSES)
                T_OD.Text = Val(Data2.Recordset.TUBES_OD)
                Combo_BWG.Text = Val(Data2.Recordset.TUBES_BWG)
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
                SHELL_P_IN.Text = Data2.Recordset.SHELL_P_IN
                SHELL_P_OUT.Text = Data2.Recordset.SHELL_P_OUT
                S_press_KP.Text = Data2.Recordset.Press_COND
                SHELL_OUT(10).Text = Data2.Recordset.SHELL_LATENT
                PROCESS_TARGET_T_OUT = Data2.Recordset.PROCESS_TARGET_TEMP
                
                VAP_PERC = Data2.Recordset.VAP_FRACTION
                YXY = 1
                Spin_VAP_P = VAP_PERC * 10
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
                Check_U_CLEAN.Value = Data2.Recordset.Check_U_CLEAN
                
                If Combo_S_FLUID.Text <> "Water" Then
                    SHELL_OUT(1) = Data2.Recordset.SHELL_T_COND
                    SHELL_OUT(3) = Data2.Recordset.SHELL_SPH
                    SHELL_OUT(4) = Data2.Recordset.SHELL_DENS
                    SHELL_OUT(5) = Data2.Recordset.SHELL_VISC
                End If
                
                Spin_ELEVATION.Value = ELEVATION
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
                HScroll_TUBES_P_IN.Value = TUBES_P_IN * 1000
                HScroll_TUBES_P_OUT.Value = TUBES_P_OUT * 1000
                
                Spin_SHELL_FLOW.Value = SHELL_FLOW / FACT_FLOW
                HScroll_SHELL_VAPOR.Value = SHELL_VAPOR / FACT_FLOW
                HScroll_SHELL_LIQUID.Value = SHELL_LIQUID / FACT_FLOW
                HScroll_SHELL_WATER.Value = SHELL_WATER / FACT_FLOW
                HScroll_SHELL_NON_COND.Value = SHELL_NON_COND / FACT_FLOW
                HScroll_SHELL_T_IN.Value = SHELL_T_IN * 100
                HScroll_SHELL_T_OUT.Value = SHELL_T_OUT * 100
                HScroll_SHELL_P_IN.Value = SHELL_P_IN * 1000
                HScroll_SHELL_P_OUT.Value = SHELL_P_OUT * 1000
                Spin_S_PRESS.Value = Val(S_press_KP) * 10
                HScroll_SHELL_SPH = SHELL_OUT(3) * 1000
                HScroll_SHELL_DENS = SHELL_OUT(4) * 10
                HScroll_SHELL_VISC = SHELL_OUT(5) * 1000
                
                LMTD = Data2.Recordset.LMTD
                TTD = Data2.Recordset.TTD
                MTDc = Data2.Recordset.MTDc
                SKIN_TEMP = Data2.Recordset.SKIN_TEMP
                C_Factor = Data2.Recordset.C_Factor
                TUBES_FF = Data2.Recordset.TUBES_FF
                U_COEFF_CLEAN = Data2.Recordset.Clean
                HScroll_U_CLEAN = U_COEFF_CLEAN
                Wet_steam = Data2.Recordset.Wet_steam
                HScroll_WET_STEAM = Wet_steam * 10
                WATER_FF = Data2.Recordset.WATER_FF
                FACT_FLOW = Data2.Recordset.FACT_FLOW
                Spin_FACT_FLOW.Value = FACT_FLOW
                
                If Check_PF = 1 Then
                    SHELL_FF.Text = Data2.Recordset.SHELL_FF
                    Spin_PF.Value = SHELL_FF * 100
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
                If Check_S_SPH = 1 Then
                    SHELL_OUT(3) = Data2.Recordset.SHELL_SPH
                End If
                If Check_S_DENS = 1 Then
                    SHELL_OUT(4) = Data2.Recordset.SHELL_DENS
                End If
                If Check_S_VISC = 1 Then
                    SHELL_OUT(5) = Data2.Recordset.SHELL_VISC
                End If
                If Check_S_TC = 1 Then
                    SHELL_OUT(1) = Data2.Recordset.SHELL_T_COND
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
    
    For i = 1 To 5
        If Combo_T_FLUID = "Water" Then
            If i = 2 Then i = 3
            TUBES_OUT(i).ForeColor = &HFF0000
            TUBES_OUT(i).BackColor = &HE0E0E0
        Else
            TUBES_OUT(i).ForeColor = &HC0&
            TUBES_OUT(i).BackColor = &HE0E0E0
        End If
    Next i
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Or Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam" Or Combo_S_FLUID.Text = "Water" Or Combo_S_FLUID = "Steam condensing" Then
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
    If Combo_S_FLUID.Text = "Water" Or Combo_S_FLUID = "Steam condensing" Then
        SHELL_OUT(3).ForeColor = &HFF0000
        SHELL_OUT(3).BackColor = &HE0E0E0
    Else
        SHELL_OUT(3).ForeColor = &HC0&
        SHELL_OUT(3).BackColor = &HE0E0E0
    End If
        
Data2.UpdateRecord

    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
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
        Check_des.Value = Unchecked
        Combo_PLANT_UNIT = "PLANT UNIT"
        Combo_PROCESS_DESCR = "PROCESS DESCRIPTION"
        Combo_PROCESS_STREAM = "PROCESS STREAM"
        Combo_COOL_TOWER = "TOWER"
        
        Combo_TEMA = "AES"
        Combo_POSITION = "Horizontal"
        ELEVATION.Text = 1
        PARALLEL_N.Text = 1
        SERIES_N.Text = 1
        Combo_CURRENT = "Counter-flow"
        
        Spin_PARALLEL_N.Value = Val(PARALLEL_N)
        Spin_SERIES_N.Value = Val(SERIES_N)
        
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
        Combo_S_FLUID = "Other fluid"
        
        TUBES_FLOW.Text = Format(125000, "0")
        TUBES_VAPOR.Text = Format(0, "0")
        TUBES_LIQUID = Format(125000, "0")
        TUBES_WATER = Format(125000, "0")
        TUBES_NON_COND = Format(0, "0")
        TUBES_OUT(0) = Format(Val(TUBES_FLOW.Text), "##,##0")
        TUBES_OUT(0).BackColor = &HE0E0E0
        TUBES_OUT(0).ForeColor = &HC0&
        
        Spin_FACT_FLOW.Value = 10
        FACT_FLOW = 10
        
        HScroll_TUBES_FLOW.Value = Val(TUBES_FLOW) / FACT_FLOW
        HScroll_TUBES_VAPOR = Val(TUBES_VAPOR) / FACT_FLOW
        HScroll_TUBES_LIQUID = Val(TUBES_LIQUID) / FACT_FLOW
        HScroll_TUBES_WATER = Val(TUBES_WATER) / FACT_FLOW
        HScroll_TUBES_NON_COND = Val(TUBES_NON_COND) / FACT_FLOW
        
        TUBES_TEMP_IN.Text = Format(25, "0.00")
        TUBES_TEMP_OUT.Text = Format(35, "0.00")
        TUBES_P_IN.Text = Format(4, "0.00")
        TUBES_P_OUT.Text = Format(3.5, "0.00")
    
        TUBES_OUT(9).Text = Data1.Recordset.TUBES_PRESS_DROP
        
        Spin_TUBES_T_IN.Value = Val(TUBES_TEMP_IN) * 100
        Spin_TUBES_T_OUT.Value = Val(TUBES_TEMP_OUT) * 100
        HScroll_TUBES_P_IN.Value = Val(TUBES_P_IN) * 1000
        HScroll_TUBES_P_OUT.Value = Val(TUBES_P_OUT) * 1000
        
        SHELL_FLOW.Text = Format(45040, "0")
        SHELL_VAPOR.Text = Format(0, "0")
        SHELL_LIQUID = Format(45040, "0")
        SHELL_WATER = Format(0, "0")
        SHELL_NON_COND = Format(0, "0")
        SHELL_OUT(0) = Format(Val(SHELL_FLOW.Text), "##,##0")
        
        Spin_SHELL_FLOW.Value = Val(SHELL_FLOW) / FACT_FLOW
        HScroll_SHELL_VAPOR = Val(SHELL_VAPOR) / FACT_FLOW
        HScroll_SHELL_LIQUID = Val(SHELL_LIQUID) / FACT_FLOW
        HScroll_SHELL_WATER = Val(SHELL_WATER) / FACT_FLOW
        HScroll_SHELL_NON_COND = Val(SHELL_NON_COND) / FACT_FLOW
        VAP_PERC = 0
        LIQ_PERC = 100

        SHELL_TEMP_IN.Text = Format(85, "0.00")
        SHELL_TEMP_OUT.Text = Format(40, "0.00")
        SHELL_P_IN.Text = Format(10, "0.00")
        SHELL_P_OUT.Text = Format(9.5, "0.00")
        S_press_KP = Format(0, "0.0")
        SHELL_FF = Format(1, "0.00")
        PROCESS_TARGET_T_OUT = 40
        Spin_TARGET_T.Value = 400
        
        HScroll_SHELL_T_IN.Value = Val(SHELL_TEMP_IN) * 100
        HScroll_SHELL_T_OUT.Value = Val(SHELL_TEMP_OUT) * 100
        HScroll_SHELL_P_IN.Value = Val(SHELL_P_IN) * 1000
        HScroll_SHELL_P_OUT.Value = Val(SHELL_P_OUT) * 1000
        Spin_S_PRESS.Value = Val(S_press_KP) * 10
        Spin_PF = Val(SHELL_FF) * 100
        
        SHELL_OUT(1) = Format(0.2, "0.000")
        SHELL_OUT(3) = Format(0.6, "0.000")
        SHELL_OUT(4) = Format(800, "0.0")
        SHELL_OUT(5) = Format(0.2, "0.000")
        SHELL_OUT(9).Text = Data1.Recordset.SHELL_PRESS_DROP
        SHELL_OUT(10) = Format(1000, "0")
        WATER_FF = Format(4, "0.00")
        
        For i = 1 To 5
            If Combo_T_FLUID = "Water" Then
                If i = 2 Then i = 3
                TUBES_OUT(i).ForeColor = &HFF0000
                TUBES_OUT(i).BackColor = &HE0E0E0
            Else
                TUBES_OUT(i).ForeColor = &HC0&
                TUBES_OUT(i).BackColor = &HE0E0E0
            End If
        Next i
        
        Wet_steam = 0
        HScroll_P_DROP_S = SHELL_OUT(9) * 10
        U_COEFF_CLEAN = 1150
    
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
        Check_WET_STEAM = Unchecked
        
Data1.UpdateRecord
Data1.Recordset.Bookmark = Data1.Recordset.LastModified

    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
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
Private Sub ToggleButton1_Click()
On Error Resume Next
    If ToggleButton1 = True Then
        Call COOLERS
        Exit Sub
    ElseIf ToggleButton1 = False Then
        Call COOLERS
    End If
End Sub
Private Sub Check_water_steam_Click()
On Error Resume Next
    If Check_water_steam = Checked Then
        TUBES_FLOW = TUBES_OUT(0)
        TUBES_LIQUID = TUBES_OUT(0)
        TUBES_WATER = TUBES_OUT(0)
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Command2_Click()
On Error Resume Next
Call CheckLockedStatus(temp)
    If temp = "locked" Then
        MsgBox "This feature is not allowed in the trial version."
        Exit Sub
    End If
    Data12.Refresh
    Frame_Search.Visible = True
End Sub
Private Sub Command3_Click()
On Error Resume Next
    Frame_Search.Visible = False
End Sub
Private Sub Thermal_bal_tubes_Click()
    If Thermal_bal_tubes = True Then
        Thermal_bal_tubes.BackColor = &H8000&
        Thermal_bal_tubes.ForeColor = &H80&
        Thermal_bal_shell = False
        Check_T_OUT = Unchecked
    Else
        Check_T_OUT = Unchecked
        Thermal_bal_tubes.BackColor = &H8000000F
        Thermal_bal_tubes.ForeColor = &H80&
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Thermal_bal_shell_Click()
    If Thermal_bal_shell = True Then
        Thermal_bal_shell.BackColor = &H8000&
        Thermal_bal_shell.ForeColor = &H80&
        Thermal_bal_tubes = False
        Check_T_FLOW = Unchecked
    Else
        Thermal_bal_shell.BackColor = &H8000000F
        Thermal_bal_shell.ForeColor = &H80&
    End If
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
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
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub
Private Sub Check_T_OUT_Click()
    If Check_T_OUT = 1 And Thermal_bal_tubes = True Then
        Check_T_FLOW = Unchecked
    Else
        Check_T_FLOW = Unchecked
        Check_T_OUT = Unchecked
    End If
    
    If Combo_S_FLUID = "Benzene" Or Combo_S_FLUID = "Ammonia" Or Combo_S_FLUID = "VCM" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Toluene" Or Combo_S_FLUID = "Propylene" Or Combo_S_FLUID = "Steam condensing" Then
        Call CONDENSER
    ElseIf Combo_S_FLUID = "Steam" Then
        Call Steam
    ElseIf Combo_CURRENT = "Condensation" Then
        Call CONDENSER
    Else
        Call COOLERS
    End If
End Sub

