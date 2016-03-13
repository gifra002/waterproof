VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "fm20.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSTEAM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WaterProof SSC - Steam turbine exhaust condenser - Data Input"
   ClientHeight    =   10515
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   14895
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
   ScaleHeight     =   10284.13
   ScaleMode       =   0  'User
   ScaleWidth      =   14000
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   "C:\Condensers\Database\steam.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   11820
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_Date"
      Top             =   9480
      Width           =   1755
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   "C:\Condensers\Database\steam.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   7500
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_Country"
      Top             =   9420
      Width           =   1815
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   "C:\Condensers\Database\steam.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   5580
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_LOC"
      Top             =   9420
      Width           =   1875
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   "C:\Condensers\Database\steam.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_Unit"
      Top             =   9420
      Width           =   1875
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "C:\Condensers\Database\steam.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   1860
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_Plant"
      Top             =   9420
      Width           =   1875
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Condensers\Database\steam.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   180
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_test"
      Top             =   9420
      Width           =   1695
   End
   Begin VB.Data Data1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Condensers\Database\steam.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   3420
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_test"
      Top             =   8880
      Width           =   5535
   End
   Begin VB.Frame Frame8 
      Caption         =   "Trend period"
      ForeColor       =   &H00FF0000&
      Height          =   1035
      Left            =   120
      TabIndex        =   254
      Top             =   7680
      Visible         =   0   'False
      Width           =   1395
      Begin MSComCtl2.DTPicker DTP_Inizio 
         DataField       =   "Date_Inizio"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   0
         TabIndex        =   255
         ToolTipText     =   "Chose the date for the search"
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
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
         CalendarBackColor=   -2147483626
         CalendarForeColor=   4210816
         CalendarTitleBackColor=   -2147483624
         CalendarTitleForeColor=   4210816
         CustomFormat    =   "01/01/01"
         Format          =   57540609
         CurrentDate     =   37925
         MinDate         =   36526
      End
      Begin MSComCtl2.DTPicker DTP_Fine 
         DataField       =   "Date_Fine"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   0
         TabIndex        =   256
         ToolTipText     =   "Chose the date for the search"
         Top             =   600
         Width           =   1395
         _ExtentX        =   2461
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
         CalendarBackColor=   -2147483626
         CalendarForeColor=   4210816
         CalendarTitleBackColor=   -2147483624
         CalendarTitleForeColor=   4210816
         CustomFormat    =   "01/01/01"
         Format          =   57540609
         CurrentDate     =   37925
         MinDate         =   36526
      End
   End
   Begin VB.TextBox PLANT_X 
      DataField       =   "PLANT_Z"
      DataSource      =   "Data2"
      Height          =   255
      Left            =   10500
      TabIndex        =   217
      Text            =   "PLANT_X"
      Top             =   9480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Comm_graph_reset 
      Caption         =   "Reset chart selection"
      Height          =   495
      Left            =   180
      TabIndex        =   216
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Frame Frame7 
      Caption         =   "Search"
      ForeColor       =   &H00FF0000&
      Height          =   3435
      Left            =   120
      TabIndex        =   213
      Top             =   3060
      Width           =   1395
      Begin VB.ComboBox Combo_Date_X 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00004080&
         Height          =   330
         Left            =   0
         TabIndex        =   257
         Text            =   "Search Date"
         Top             =   2640
         Width           =   1395
      End
      Begin VB.CommandButton Com_Go 
         Caption         =   "Go to date"
         Height          =   255
         Left            =   120
         TabIndex        =   253
         Top             =   3060
         Width           =   1095
      End
      Begin VB.ComboBox Combo_UNIT_1 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00004080&
         Height          =   330
         Left            =   0
         TabIndex        =   249
         Text            =   "Search unit"
         Top             =   1740
         Width           =   1395
      End
      Begin VB.ComboBox Combo_Plant_1 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00004080&
         Height          =   330
         Left            =   0
         TabIndex        =   248
         Text            =   "Search plant"
         Top             =   1080
         Width           =   1395
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
         TabIndex        =   252
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Label Label45 
         Caption         =   "Unit"
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   120
         TabIndex        =   251
         Top             =   1500
         Width           =   795
      End
      Begin VB.Label Label44 
         Caption         =   "Plant:"
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   120
         TabIndex        =   250
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
         TabIndex        =   215
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Caption         =   "Search by Plant and Unit selection"
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   60
         TabIndex        =   214
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
      Left            =   9420
      TabIndex        =   183
      Text            =   "UNIT_X"
      Top             =   9480
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame6 
      Caption         =   "Actual Output data"
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
      Height          =   4455
      Left            =   8160
      TabIndex        =   143
      Top             =   4320
      Width           =   6675
      Begin VB.TextBox CHT_act 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         DataField       =   "Tube_PAS"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox S_MW_act 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3540
         TabIndex        =   73
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CheckBox Check_CFAC 
         Caption         =   "Check1"
         DataField       =   "Check_CFAC"
         DataSource      =   "Data1"
         Height          =   210
         Left            =   3420
         TabIndex        =   84
         ToolTipText     =   "Check to select the parameter to see the trend."
         Top             =   720
         Width           =   135
      End
      Begin VB.TextBox Proc_FF_act 
         Height          =   255
         Left            =   5520
         TabIndex        =   229
         Text            =   "Proc_FF_act"
         Top             =   1620
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Water_press_drop_act_bar 
         Height          =   255
         Left            =   5940
         TabIndex        =   227
         Text            =   "Text1"
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox W_MW_act 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3540
         TabIndex        =   72
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CheckBox Check_PF_act 
         Caption         =   "X"
         DataField       =   "Check_PF_act"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   3540
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Click to enter a different value if available, using the cursors."
         Top             =   1920
         UseMaskColor    =   -1  'True
         Width           =   195
      End
      Begin VB.CheckBox Check_PP 
         Caption         =   "Check1"
         DataField       =   "Check_P_POWER"
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
         Height          =   195
         Left            =   3420
         TabIndex        =   93
         ToolTipText     =   "Check to select the parameter to see the trend."
         Top             =   4140
         Width           =   135
      End
      Begin VB.TextBox P_Power_act 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         DataField       =   "Pump_power"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   78
         Top             =   4140
         Width           =   1215
      End
      Begin VB.CheckBox Check_CF 
         Caption         =   "Check8"
         DataField       =   "Check_CF"
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
         Height          =   195
         Left            =   3420
         TabIndex        =   91
         ToolTipText     =   "Check to select the parameter to see the trend."
         Top             =   3600
         Width           =   135
      End
      Begin VB.CheckBox Check_TTD 
         Caption         =   "Check7"
         DataField       =   "Check_TD"
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
         Height          =   195
         Left            =   3420
         TabIndex        =   90
         ToolTipText     =   "Check to select the parameter to see the trend."
         Top             =   2400
         Width           =   135
      End
      Begin VB.CheckBox Check_LMTD 
         Caption         =   "Check6"
         DataField       =   "Check_LMTD"
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
         Height          =   195
         Left            =   3420
         TabIndex        =   89
         ToolTipText     =   "Check to select the parameter to see the trend."
         Top             =   2160
         Width           =   135
      End
      Begin VB.CheckBox Check_CT 
         Caption         =   "Check5"
         DataField       =   "Check_CT"
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
         Height          =   195
         Left            =   3420
         TabIndex        =   88
         ToolTipText     =   "Check to select the parameter to see the trend."
         Top             =   1680
         Width           =   135
      End
      Begin VB.CheckBox Check_RE 
         Caption         =   "Check4"
         DataField       =   "Check_RE"
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
         Height          =   195
         Left            =   3420
         TabIndex        =   87
         ToolTipText     =   "Check to select the parameter to see the trend."
         Top             =   1440
         Width           =   135
      End
      Begin VB.CheckBox Check_PD 
         Caption         =   "Check3"
         DataField       =   "Check_PD"
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
         Height          =   195
         Left            =   3420
         TabIndex        =   86
         ToolTipText     =   "Check to select the parameter to see the trend."
         Top             =   1200
         Width           =   135
      End
      Begin VB.CheckBox Check_VEL 
         Caption         =   "Check2"
         DataField       =   "Check_VEL"
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
         Height          =   195
         Left            =   3420
         TabIndex        =   85
         ToolTipText     =   "Check to select the parameter to see the trend."
         Top             =   960
         Width           =   135
      End
      Begin VB.CheckBox Check_FF 
         Caption         =   "Check1"
         DataField       =   "Check_FF"
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
         Height          =   195
         Left            =   3420
         TabIndex        =   92
         ToolTipText     =   "Check to select the parameter to see the trend."
         Top             =   3840
         Width           =   135
      End
      Begin VB.TextBox W_FF_act 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000C0&
         DataField       =   "FF"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox CF_act 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         DataField       =   "Clean"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   76
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox DHT_act 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         DataField       =   "Tube_PAS"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox S_duty_act 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         DataField       =   "Tube_PAS"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   146
         Text            =   "S_duty_act"
         Top             =   2940
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox Wet_steam_act 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox RE_act 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         DataField       =   "Reynolds_W"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox Latent_Heat_act 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox C_Factor_act 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         DataField       =   "C_Factor_act"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Water_vel_act 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         DataField       =   "Vel_W"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Water_press_drop_act 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         DataField       =   "PressDrop_W"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Steam_cond_temp_act 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         DataField       =   "Temp_C"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Proc_FF_act_KW 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         DataField       =   "Proc_F_act"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox LMTD_act 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         DataField       =   "LMTD"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox TTD_act 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         DataField       =   "TTD"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox W_duty_act 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         DataField       =   "Tube_PAS"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   144
         Text            =   "W_duty_act"
         Top             =   2700
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label43 
         Caption         =   "MW"
         ForeColor       =   &H00404080&
         Height          =   195
         Left            =   4860
         TabIndex        =   222
         Top             =   2940
         Width           =   435
      End
      Begin VB.Label Label42 
         Caption         =   "MW"
         ForeColor       =   &H00404080&
         Height          =   195
         Left            =   4860
         TabIndex        =   221
         Top             =   2640
         Width           =   375
      End
      Begin MSForms.SpinButton Spin_PF_act 
         Height          =   255
         Left            =   4620
         TabIndex        =   68
         Top             =   1920
         Width           =   195
         Size            =   "344;450"
         Max             =   50000
         SmallChange     =   10
      End
      Begin VB.Label Label34 
         Caption         =   "Cooling water pumping power (PP):"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   208
         Top             =   4140
         Width           =   2775
      End
      Begin VB.Label Label33 
         Caption         =   "kW"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4860
         TabIndex        =   207
         Top             =   4140
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Water side duty:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   39
         Left            =   180
         TabIndex        =   199
         Top             =   2655
         Width           =   2625
      End
      Begin VB.Label lblLabels 
         Caption         =   "Terminal temperature difference (TTD):"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   40
         Left            =   180
         TabIndex        =   198
         Top             =   2415
         Width           =   3165
      End
      Begin VB.Label lblLabels 
         Caption         =   "Log Mean Temperature Difference (LMTD):"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   41
         Left            =   180
         TabIndex        =   197
         Top             =   2175
         Width           =   3105
      End
      Begin VB.Label lblLabels 
         Caption         =   "Process side fouling factor :"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   42
         Left            =   180
         TabIndex        =   196
         Top             =   1920
         Width           =   2625
      End
      Begin VB.Label lblLabels 
         Caption         =   "Condensed steam  temperature (CT):"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   43
         Left            =   180
         TabIndex        =   195
         Top             =   1680
         Width           =   3045
      End
      Begin VB.Label lblLabels 
         Caption         =   "Pressure drop for 100% clean tube side (PD):"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   44
         Left            =   180
         TabIndex        =   194
         Top             =   1200
         Width           =   3285
      End
      Begin VB.Label lblLabels 
         Caption         =   "Water velocity through tubes (WV):"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   45
         Left            =   180
         TabIndex        =   193
         Top             =   960
         Width           =   3150
      End
      Begin VB.Label lblLabels 
         Caption         =   "C Factor (CFAC):"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   46
         Left            =   180
         TabIndex        =   192
         Top             =   720
         Width           =   2625
      End
      Begin VB.Label lblLabels 
         Caption         =   "Wet steam:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   47
         Left            =   180
         TabIndex        =   191
         Top             =   480
         Width           =   2625
      End
      Begin VB.Label lblLabels 
         Caption         =   "Latent Heat:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   48
         Left            =   180
         TabIndex        =   190
         Top             =   240
         Width           =   2625
      End
      Begin VB.Label lblLabels 
         Caption         =   "Reynolds number (RE):"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   49
         Left            =   180
         TabIndex        =   189
         Top             =   1440
         Width           =   2580
      End
      Begin VB.Label lblLabels 
         Caption         =   "Steam side duty:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   51
         Left            =   180
         TabIndex        =   188
         Top             =   2940
         Width           =   2625
      End
      Begin VB.Label Label15 
         Caption         =   "Overall CLEAN heat transfer coefficient:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   187
         Top             =   3180
         Width           =   3015
      End
      Begin VB.Label Label17 
         Caption         =   "Overall DIRTY heat transfer coefficient:Overall DIRTY heat transfer coefficient"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   186
         Top             =   3420
         Width           =   2955
      End
      Begin VB.Label Label19 
         Caption         =   "Cleanliness factor (CF):"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   185
         Top             =   3660
         Width           =   2835
      End
      Begin VB.Label Label21 
         Caption         =   "Water side fouling factor (FF):"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   184
         Top             =   3900
         Width           =   2775
      End
      Begin VB.Label Label22 
         Caption         =   "[( m^2 ºC)/kW].10^-3"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   25
         Left            =   4860
         TabIndex        =   175
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label Label22 
         Caption         =   "%"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   24
         Left            =   4860
         TabIndex        =   174
         Top             =   3600
         Width           =   1395
      End
      Begin VB.Label Label22 
         Caption         =   "kW/( m^2 ºC)"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   23
         Left            =   4860
         TabIndex        =   173
         Top             =   3420
         Width           =   1395
      End
      Begin VB.Label Label22 
         Caption         =   "KW/( m^2 ºC)"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   22
         Left            =   4860
         TabIndex        =   172
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label Label22 
         Caption         =   "kcal/h"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   21
         Left            =   6120
         TabIndex        =   171
         Top             =   2940
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label22 
         Caption         =   "kcal/h"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   20
         Left            =   6060
         TabIndex        =   170
         Top             =   2700
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label22 
         Caption         =   "°C"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   19
         Left            =   4860
         TabIndex        =   169
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label Label22 
         Caption         =   "°C"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   18
         Left            =   4860
         TabIndex        =   168
         Top             =   2160
         Width           =   1395
      End
      Begin VB.Label Label22 
         Caption         =   "°C"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   17
         Left            =   4860
         TabIndex        =   167
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label Label22 
         Caption         =   "kPa"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   15
         Left            =   4860
         TabIndex        =   166
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label Label22 
         Caption         =   "m/s"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   14
         Left            =   4860
         TabIndex        =   165
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label Label24 
         Caption         =   "m3/h/kPa"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   4860
         TabIndex        =   164
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label23 
         Caption         =   "%"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   4860
         TabIndex        =   163
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label Label22 
         Caption         =   "kJ/Kg"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   13
         Left            =   4860
         TabIndex        =   162
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label12 
         Caption         =   "[(m^2 ºC)/kW]*10^-3"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   4860
         TabIndex        =   161
         Top             =   1920
         Width           =   1755
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Actual Operating conditions"
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
      Height          =   2415
      Left            =   8160
      TabIndex        =   111
      Top             =   2100
      Width           =   6675
      Begin VB.TextBox S_press_act 
         DataField       =   "Steam_press_act"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   6120
         TabIndex        =   225
         Text            =   "Text1"
         Top             =   1320
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CheckBox Check_SF 
         Caption         =   "Check1"
         DataField       =   "Check_S_FLOW"
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
         Height          =   195
         Left            =   2580
         TabIndex        =   82
         ToolTipText     =   "Check to select the parameter to see the trend."
         Top             =   1020
         Width           =   135
      End
      Begin VB.CheckBox Check_WF 
         Caption         =   "Check1"
         DataField       =   "Check_W_FLOW"
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
         Height          =   195
         Left            =   2580
         TabIndex        =   79
         ToolTipText     =   "Check to select the parameter to see the trend."
         Top             =   240
         Width           =   135
      End
      Begin VB.CheckBox Check_CP 
         Caption         =   "Check11"
         DataField       =   "Check_CP"
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
         Height          =   195
         Left            =   2580
         TabIndex        =   83
         ToolTipText     =   "Check to select the parameter to see the trend."
         Top             =   1320
         Width           =   135
      End
      Begin VB.CheckBox Check_TOUT 
         Caption         =   "Check10"
         DataField       =   "Check_TOUT"
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
         Height          =   195
         Left            =   2580
         TabIndex        =   81
         ToolTipText     =   "Check to select the parameter to see the trend."
         Top             =   780
         Width           =   135
      End
      Begin VB.CheckBox Check_TIN 
         Caption         =   "Check9"
         DataField       =   "Check_TIN"
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
         Height          =   195
         Left            =   2580
         TabIndex        =   80
         ToolTipText     =   "Check to select the parameter to see the trend."
         Top             =   480
         Width           =   135
      End
      Begin VB.HScrollBar Spin_S_PRESS_act 
         Height          =   195
         LargeChange     =   100
         Left            =   3840
         Max             =   3000
         Min             =   50
         TabIndex        =   38
         Top             =   1380
         Value           =   1000
         Width           =   1515
      End
      Begin VB.HScrollBar Spin_W_T_OUT_act 
         Height          =   195
         LargeChange     =   100
         Left            =   3840
         Max             =   10000
         Min             =   5
         TabIndex        =   34
         Top             =   820
         Value           =   3500
         Width           =   1515
      End
      Begin VB.HScrollBar Spin_W_T_IN_act 
         Height          =   195
         LargeChange     =   100
         Left            =   3840
         Max             =   10000
         Min             =   5
         TabIndex        =   32
         Top             =   540
         Value           =   2500
         Width           =   1515
      End
      Begin VB.HScrollBar Spin_S_FLOW_IN_act 
         Height          =   195
         LargeChange     =   10
         Left            =   3840
         Max             =   30000
         Min             =   1
         TabIndex        =   36
         Top             =   1100
         Value           =   10000
         Width           =   1515
      End
      Begin VB.TextBox Text7 
         DataField       =   "Tube_PAS"
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
         Height          =   285
         Left            =   2835
         TabIndex        =   112
         Top             =   4230
         Width           =   1935
      End
      Begin VB.TextBox S_press_act_KP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   37
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   1305
         Width           =   1155
      End
      Begin VB.TextBox S_flow_IN_act 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "Steam_flow_act"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   35
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   1035
         Width           =   1155
      End
      Begin VB.TextBox W_T_OUT_act 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "Temp_OUT_act"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   33
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   750
         Width           =   1155
      End
      Begin VB.TextBox W_T_IN_act 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "Temp_IN_act"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   31
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   480
         Width           =   1155
      End
      Begin VB.TextBox W_flow_IN_act 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         DataField       =   "Water_flow_act"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   200
         Top             =   210
         Width           =   1155
      End
      Begin VB.Label Label29 
         Caption         =   "(by thermal balance)"
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
         Left            =   3900
         TabIndex        =   202
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label28 
         Caption         =   "kg/h"
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
         Left            =   5520
         TabIndex        =   182
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label27 
         Caption         =   "°C"
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
         Left            =   5520
         TabIndex        =   181
         Top             =   510
         Width           =   555
      End
      Begin VB.Label Label26 
         Caption         =   "°C"
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
         Left            =   5520
         TabIndex        =   180
         Top             =   780
         Width           =   555
      End
      Begin VB.Label Label25 
         Caption         =   "kg/h"
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
         Left            =   5520
         TabIndex        =   179
         Top             =   1050
         Width           =   555
      End
      Begin VB.Label Label14 
         Caption         =   "KPar(a)"
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
         Left            =   5520
         TabIndex        =   178
         Top             =   1320
         Width           =   555
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tube passes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   180
         TabIndex        =   119
         Top             =   4275
         Width           =   2625
      End
      Begin VB.Label lblLabels 
         Caption         =   "Steam cond. pressure (CP):"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   30
         Left            =   180
         TabIndex        =   118
         Top             =   1350
         Width           =   2325
      End
      Begin VB.Label lblLabels 
         Caption         =   "Steam flowrate INLET (SF):"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   29
         Left            =   180
         TabIndex        =   117
         Top             =   1080
         Width           =   2370
      End
      Begin VB.Label lblLabels 
         Caption         =   "Temperature OUT (TOUT):"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   28
         Left            =   180
         TabIndex        =   116
         Top             =   810
         Width           =   2265
      End
      Begin VB.Label lblLabels 
         Caption         =   "Temperature IN (TIN):"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   27
         Left            =   180
         TabIndex        =   115
         Top             =   525
         Width           =   2205
      End
      Begin VB.Label lblLabels 
         Caption         =   "Water flow rate (WF):"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   16
         Left            =   180
         TabIndex        =   114
         Top             =   240
         Width           =   2625
      End
      Begin MSForms.SpinButton SpinButton23 
         Height          =   330
         Left            =   4770
         TabIndex        =   113
         Top             =   4185
         Width           =   195
         Size            =   "344;582"
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   2295
      Left            =   8160
      TabIndex        =   105
      Top             =   0
      Width           =   6675
      Begin VB.ComboBox ComboT_OD 
         BackColor       =   &H80000018&
         ForeColor       =   &H000000C0&
         Height          =   330
         ItemData        =   "frmSTEAM.frx":0000
         Left            =   5220
         List            =   "frmSTEAM.frx":0022
         TabIndex        =   14
         Top             =   900
         Width           =   795
      End
      Begin VB.TextBox Pump_EFF 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         DataField       =   "Pump_EFF"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   17
         ToolTipText     =   "Use the cursors to enter the value."
         Top             =   1830
         Width           =   1095
      End
      Begin VB.ComboBox Combo_Mat 
         BackColor       =   &H80000018&
         DataField       =   "Tube_mat"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   330
         ItemData        =   "frmSTEAM.frx":005C
         Left            =   1500
         List            =   "frmSTEAM.frx":009F
         Sorted          =   -1  'True
         TabIndex        =   16
         ToolTipText     =   "Select from the list.Use TAB Key to save the new enter"
         Top             =   1500
         Width           =   2835
      End
      Begin VB.ComboBox Combo_BWG 
         BackColor       =   &H80000018&
         DataField       =   "BWG"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   330
         ItemData        =   "frmSTEAM.frx":023A
         Left            =   2640
         List            =   "frmSTEAM.frx":0253
         Sorted          =   -1  'True
         TabIndex        =   15
         ToolTipText     =   "Select from the list.Use TAB Key to save the new enter"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox T_OD 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         DataField       =   "Tube_OD"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "Use the cursors to enter the value."
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Area 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
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
         Left            =   5220
         Locked          =   -1  'True
         TabIndex        =   132
         Top             =   1860
         Width           =   795
      End
      Begin VB.TextBox T_PASS 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         DataField       =   "Tube_PAS"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "Use the cursors to enter the value."
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox T_IN 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
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
         Left            =   5220
         Locked          =   -1  'True
         TabIndex        =   203
         Top             =   1320
         Width           =   795
      End
      Begin VB.TextBox Mat_factor 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
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
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   128
         Top             =   2775
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Mat_cond 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
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
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   126
         Top             =   2520
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox T_ID 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
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
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   122
         Top             =   2235
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox T_len 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         DataField       =   "Tube_LE"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   8
         ToolTipText     =   "Use the cursors to enter the value."
         Top             =   480
         Width           =   1095
      End
      Begin VB.HScrollBar HScroll_T_NO 
         Height          =   200
         LargeChange     =   100
         Left            =   3720
         Max             =   30000
         Min             =   10
         TabIndex        =   7
         Top             =   300
         Value           =   2000
         Width           =   1755
      End
      Begin VB.TextBox T_NO 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         DataField       =   "Tube_NO"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label47 
         Caption         =   "inches"
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   6060
         TabIndex        =   259
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label Label37 
         Caption         =   "mm    O.D."
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   4380
         TabIndex        =   258
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label Label39 
         Caption         =   "%"
         Height          =   255
         Left            =   4380
         TabIndex        =   218
         Top             =   1860
         Width           =   315
      End
      Begin MSForms.SpinButton Spin_P_EFF 
         Height          =   225
         Left            =   3720
         TabIndex        =   18
         Top             =   1860
         Width           =   615
         Size            =   "1085;397"
         Min             =   50
         Max             =   1000
         Position        =   500
         Orientation     =   1
      End
      Begin VB.Label lblLabels 
         Caption         =   "Water pump efficiency:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   52
         Left            =   240
         TabIndex        =   205
         Top             =   1860
         Width           =   2325
      End
      Begin VB.Label Label30 
         Caption         =   "mm"
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
         Left            =   6060
         TabIndex        =   204
         Top             =   1380
         Width           =   435
      End
      Begin MSForms.SpinButton Spin_T_OD 
         Height          =   195
         Left            =   3720
         TabIndex        =   13
         Top             =   1020
         Width           =   615
         Size            =   "1085;344"
         Max             =   10000
         Position        =   100
         Orientation     =   1
      End
      Begin MSForms.SpinButton Spin_T_LEN 
         Height          =   195
         Left            =   3720
         TabIndex        =   9
         Top             =   541
         Width           =   615
         Size            =   "1085;344"
         Max             =   10000
         Position        =   100
         Orientation     =   1
      End
      Begin MSForms.SpinButton Spin_T_PAS 
         Height          =   200
         Left            =   3720
         TabIndex        =   11
         Top             =   777
         Width           =   615
         Size            =   "1085;353"
         Min             =   1
         Max             =   8
         Position        =   1
         Orientation     =   1
      End
      Begin VB.Label lblLabels 
         Caption         =   "Heat transfer surface:"
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
         Left            =   4860
         TabIndex        =   134
         Top             =   1680
         Width           =   1605
      End
      Begin VB.Label Label7 
         Caption         =   "m2"
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
         Left            =   6060
         TabIndex        =   133
         Top             =   1920
         Width           =   315
      End
      Begin VB.Label Label2 
         Caption         =   "m"
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
         Left            =   4380
         TabIndex        =   131
         Top             =   540
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tube material:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   35
         Left            =   240
         TabIndex        =   130
         Top             =   1540
         Width           =   2325
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
         Height          =   255
         Index           =   34
         Left            =   240
         TabIndex        =   129
         Top             =   2805
         Visible         =   0   'False
         Width           =   2985
      End
      Begin VB.Label lblLabels 
         Caption         =   "Thermal conductivity of tube material:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   33
         Left            =   240
         TabIndex        =   127
         Top             =   2580
         Visible         =   0   'False
         Width           =   2955
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tube passes:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   125
         Top             =   760
         Width           =   2265
      End
      Begin VB.Label Label5 
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   124
         Top             =   2220
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tube Inlet Diameter:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   32
         Left            =   240
         TabIndex        =   123
         Top             =   2280
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.Label Label4 
         Caption         =   "Kcal/(h m^2 ºC/m)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   121
         Top             =   2580
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "mm     I.D."
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
         Left            =   4380
         TabIndex        =   120
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label lblLabels 
         Caption         =   "BWG:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   109
         Top             =   1280
         Width           =   2265
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tube Outlet diameter:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   108
         Top             =   1020
         Width           =   2265
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tubes lenght:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   107
         Top             =   500
         Width           =   2265
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tubes Number:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   106
         Top             =   240
         Width           =   2205
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2295
      Left            =   180
      TabIndex        =   142
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   4048
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      MultiSelect     =   -1  'True
      Placement       =   1
      Separators      =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   8
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
            Caption         =   "Refresh"
            Object.ToolTipText     =   "Refresh the database"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Object.ToolTipText     =   "Delete this record"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Object.ToolTipText     =   "Print this form"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Object.ToolTipText     =   "Close this form"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Design Output data"
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
      Height          =   4455
      Left            =   1560
      TabIndex        =   104
      Top             =   4320
      Width           =   6660
      Begin VB.TextBox CHT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox S_MW 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3420
         TabIndex        =   52
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Proc_FF 
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   5220
         TabIndex        =   228
         Text            =   "Proc_FF"
         Top             =   1620
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.TextBox Water_press_drop_bar 
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   5880
         TabIndex        =   226
         Text            =   "P_Drop"
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox W_MW 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3420
         TabIndex        =   51
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CheckBox Check_PF 
         Caption         =   "X"
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
         Height          =   195
         Left            =   3420
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Click to enter a different value if available, using the cursors.."
         Top             =   1920
         UseMaskColor    =   -1  'True
         Width           =   195
      End
      Begin VB.TextBox P_Power 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   4140
         Width           =   1215
      End
      Begin VB.TextBox Proc_FF_KW 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         DataField       =   "Proc_F"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   1920
         Width           =   1035
      End
      Begin VB.TextBox Steam_cond_temp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox RE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox W_duty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   5340
         Locked          =   -1  'True
         TabIndex        =   177
         Text            =   "W_duty"
         Top             =   2700
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox TTD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox LMTD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Water_press_drop 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Water_vel 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox C_Factor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         DataField       =   "C_Factor"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Wet_steam 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Latent_heat 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox S_duty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   5340
         Locked          =   -1  'True
         TabIndex        =   176
         Text            =   "S_duty"
         Top             =   2940
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox DHT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox CF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox W_FF 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label31 
         Caption         =   "Cooling water pumping power:"
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
         Left            =   180
         TabIndex        =   247
         Top             =   4140
         Width           =   2775
      End
      Begin VB.Label Label20 
         Caption         =   "Water side fouling factor (TEMA norm=170):"
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
         Left            =   180
         TabIndex        =   246
         Top             =   3870
         Width           =   3135
      End
      Begin VB.Label Label18 
         Caption         =   "Cleanliness factor:"
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
         Left            =   195
         TabIndex        =   245
         Top             =   3600
         Width           =   2835
      End
      Begin VB.Label Label16 
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
         Height          =   195
         Left            =   195
         TabIndex        =   244
         Top             =   3360
         Width           =   2955
      End
      Begin VB.Label Label13 
         Caption         =   "Overall CLEAN heat transfer coefficient:"
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
         Left            =   195
         TabIndex        =   243
         Top             =   3120
         Width           =   3015
      End
      Begin VB.Label lblLabels 
         Caption         =   "Steam side duty:"
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
         Left            =   195
         TabIndex        =   242
         Top             =   2880
         Width           =   2625
      End
      Begin VB.Label lblLabels 
         Caption         =   "Water side duty:"
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
         Index           =   26
         Left            =   195
         TabIndex        =   241
         Top             =   2640
         Width           =   2625
      End
      Begin VB.Label lblLabels 
         Caption         =   "Terminal temperature difference:"
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
         TabIndex        =   240
         Top             =   2400
         Width           =   2625
      End
      Begin VB.Label lblLabels 
         Caption         =   "Log Mean Temperature Difference:"
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
         TabIndex        =   239
         Top             =   2160
         Width           =   2625
      End
      Begin VB.Label lblLabels 
         Caption         =   "Process side fouling factor (TEMA norm=85):"
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
         Index           =   23
         Left            =   195
         TabIndex        =   238
         Top             =   1920
         Width           =   3150
      End
      Begin VB.Label lblLabels 
         Caption         =   "Condensed steam  temperature :"
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
         Left            =   195
         TabIndex        =   237
         Top             =   1680
         Width           =   2625
      End
      Begin VB.Label lblLabels 
         Caption         =   "Total pressure drop for 100% clean tube side:"
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
         Left            =   195
         TabIndex        =   236
         Top             =   1200
         Width           =   3285
      End
      Begin VB.Label lblLabels 
         Caption         =   "Water velocity through tubes"
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
         Left            =   195
         TabIndex        =   235
         Top             =   960
         Width           =   2670
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
         TabIndex        =   234
         Top             =   720
         Width           =   2625
      End
      Begin VB.Label lblLabels 
         Caption         =   "Wet steam  (Norm 6 - 9):"
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
         Left            =   195
         TabIndex        =   233
         Top             =   480
         Width           =   2625
      End
      Begin VB.Label lblLabels 
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
         Left            =   195
         TabIndex        =   232
         Top             =   240
         Width           =   2625
      End
      Begin VB.Label lblLabels 
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
         Index           =   0
         Left            =   195
         TabIndex        =   231
         Top             =   1440
         Width           =   2580
      End
      Begin VB.Label Label41 
         Caption         =   "MW"
         ForeColor       =   &H00404080&
         Height          =   195
         Left            =   4740
         TabIndex        =   220
         Top             =   2940
         Width           =   315
      End
      Begin VB.Label Label40 
         Caption         =   "MW"
         ForeColor       =   &H00404080&
         Height          =   195
         Left            =   4740
         TabIndex        =   219
         Top             =   2700
         Width           =   375
      End
      Begin MSForms.SpinButton Spin_PF 
         Height          =   255
         Left            =   4440
         TabIndex        =   47
         Top             =   1920
         Width           =   195
         Size            =   "344;450"
         Max             =   50000
         SmallChange     =   10
      End
      Begin VB.Label Label32 
         Caption         =   "kW"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4740
         TabIndex        =   206
         Top             =   4140
         Width           =   615
      End
      Begin VB.Label Label22 
         Caption         =   "[( m^2 ºC)/kW].10^-3"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   12
         Left            =   4740
         TabIndex        =   160
         Top             =   3900
         Width           =   1695
      End
      Begin VB.Label Label22 
         Caption         =   "%"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   11
         Left            =   4740
         TabIndex        =   159
         Top             =   3660
         Width           =   1395
      End
      Begin VB.Label Label22 
         Caption         =   "kW/( m^2 ºC)"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   10
         Left            =   4740
         TabIndex        =   158
         Top             =   3420
         Width           =   1395
      End
      Begin VB.Label Label22 
         Caption         =   "kW/( m^2 ºC)"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   9
         Left            =   4740
         TabIndex        =   157
         Top             =   3180
         Width           =   1395
      End
      Begin VB.Label Label22 
         Caption         =   "kcal/h"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   8
         Left            =   5940
         TabIndex        =   156
         Top             =   2940
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label22 
         Caption         =   "kcal/h"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   7
         Left            =   5940
         TabIndex        =   155
         Top             =   2760
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label22 
         Caption         =   "°C"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   6
         Left            =   4740
         TabIndex        =   154
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label Label22 
         Caption         =   "°C"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   5
         Left            =   4740
         TabIndex        =   153
         Top             =   2160
         Width           =   1395
      End
      Begin VB.Label Label22 
         Caption         =   "°C"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   4
         Left            =   4740
         TabIndex        =   152
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label Label22 
         Caption         =   "kPa"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   4740
         TabIndex        =   151
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "m/s"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   4740
         TabIndex        =   150
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label Label24 
         Caption         =   "m3/h/kPa"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   4740
         TabIndex        =   149
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label23 
         Caption         =   "%"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   4740
         TabIndex        =   148
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label Label22 
         Caption         =   "kJ/Kg"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   4740
         TabIndex        =   147
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label12 
         Caption         =   "[(m^2 ºC)/kW]*10^-3"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   4740
         TabIndex        =   145
         Top             =   1920
         Width           =   1755
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Design Operating conditions"
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
      Height          =   2355
      Left            =   1560
      TabIndex        =   98
      Top             =   2100
      Width           =   6675
      Begin VB.TextBox S_press 
         DataField       =   "Steam_press"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2400
         TabIndex        =   224
         Text            =   "Text1"
         Top             =   1320
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.HScrollBar Spin_S_PRESS 
         Height          =   195
         LargeChange     =   100
         Left            =   4320
         Max             =   3000
         Min             =   50
         TabIndex        =   28
         Top             =   1380
         Value           =   1000
         Width           =   1515
      End
      Begin VB.HScrollBar Spin_W_T_OUT 
         Height          =   195
         LargeChange     =   100
         Left            =   4320
         Max             =   10000
         Min             =   5
         TabIndex        =   24
         Top             =   810
         Value           =   3500
         Width           =   1515
      End
      Begin VB.HScrollBar Spin_W_T_IN 
         Height          =   195
         LargeChange     =   100
         Left            =   4320
         Max             =   10000
         Min             =   5
         TabIndex        =   22
         Top             =   525
         Value           =   2500
         Width           =   1515
      End
      Begin VB.ComboBox Combo_T_side 
         BackColor       =   &H80000018&
         DataField       =   "Tube_fluid"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   330
         ItemData        =   "frmSTEAM.frx":0273
         Left            =   3000
         List            =   "frmSTEAM.frx":0280
         TabIndex        =   30
         ToolTipText     =   "Select from the list.Use TAB Key to save the new enter"
         Top             =   1920
         Width           =   2175
      End
      Begin VB.ComboBox Combo_S_side 
         BackColor       =   &H80000018&
         DataField       =   "Shell_fluid"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   330
         ItemData        =   "frmSTEAM.frx":02B7
         Left            =   3000
         List            =   "frmSTEAM.frx":02F1
         TabIndex        =   29
         ToolTipText     =   "Select from the list.Use TAB Key to save the new enter"
         Top             =   1620
         Width           =   2175
      End
      Begin VB.TextBox S_press_KP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   27
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox S_flow_IN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "Steam_flow"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   25
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   1035
         Width           =   1335
      End
      Begin VB.TextBox W_T_OUT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "Temp_OUT"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   23
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   765
         Width           =   1335
      End
      Begin VB.TextBox W_T_IN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "Temp_IN"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   21
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   480
         Width           =   1335
      End
      Begin VB.HScrollBar Spin_S_FLOW_IN 
         Height          =   195
         LargeChange     =   10
         Left            =   4320
         Max             =   30000
         Min             =   1
         TabIndex        =   26
         Top             =   1095
         Value           =   10000
         Width           =   1515
      End
      Begin VB.HScrollBar HScroll_W_FLOW 
         Height          =   195
         LargeChange     =   10
         Left            =   4320
         Max             =   30000
         Min             =   10
         TabIndex        =   20
         Top             =   240
         Value           =   10000
         Width           =   1515
      End
      Begin VB.TextBox W_flow_IN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         DataField       =   "Water_flow"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   19
         ToolTipText     =   "Use the scrollbar to enter the value."
         Top             =   210
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tube side fluid:"
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
         Index           =   38
         Left            =   180
         TabIndex        =   141
         Top             =   1920
         Width           =   2505
      End
      Begin VB.Label lblLabels 
         Caption         =   "Shell side fluid:"
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
         Index           =   37
         Left            =   180
         TabIndex        =   140
         Top             =   1620
         Width           =   2505
      End
      Begin VB.Label Label11 
         Caption         =   "kPa(a)"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5880
         TabIndex        =   139
         Top             =   1380
         Width           =   555
      End
      Begin VB.Label Label10 
         Caption         =   "kg/h"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5880
         TabIndex        =   138
         Top             =   1095
         Width           =   555
      End
      Begin VB.Label Label9 
         Caption         =   "°C"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5880
         TabIndex        =   137
         Top             =   810
         Width           =   555
      End
      Begin VB.Label Label8 
         Caption         =   "°C"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5880
         TabIndex        =   136
         Top             =   525
         Width           =   555
      End
      Begin VB.Label Label6 
         Caption         =   "kg/h"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5880
         TabIndex        =   135
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblLabels 
         Caption         =   "Water flow rate:"
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
         TabIndex        =   103
         Top             =   240
         Width           =   2505
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
         Height          =   255
         Index           =   7
         Left            =   180
         TabIndex        =   102
         Top             =   525
         Width           =   2625
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
         Height          =   255
         Index           =   8
         Left            =   180
         TabIndex        =   101
         Top             =   810
         Width           =   2625
      End
      Begin VB.Label lblLabels 
         Caption         =   "Saturated steam flowrate INLET"
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
         TabIndex        =   100
         Top             =   1080
         Width           =   2490
      End
      Begin VB.Label lblLabels 
         Caption         =   "Steam condensing pressure:"
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
         TabIndex        =   99
         Top             =   1350
         Width           =   2625
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
      Height          =   2205
      Left            =   1560
      TabIndex        =   59
      Top             =   0
      Width           =   6675
      Begin VB.CheckBox Check_des 
         Caption         =   "Check if you enter design data for new unit"
         DataField       =   "Check_Design"
         DataSource      =   "Data1"
         ForeColor       =   &H00004080&
         Height          =   210
         Left            =   3060
         TabIndex        =   5
         ToolTipText     =   "Check if you enter the design data of new unit also in the actual section"
         Top             =   1860
         Width           =   3495
      End
      Begin VB.ComboBox Combo_UNIT 
         BackColor       =   &H80000018&
         DataField       =   "Unit_name"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   3000
         TabIndex        =   4
         Text            =   "Unit"
         ToolTipText     =   "Use TAB Key to save the new enter. Do not enter same UNIT name for different PLANT name!"
         Top             =   1440
         Width           =   2235
      End
      Begin VB.ComboBox Combo_Country 
         BackColor       =   &H80000018&
         DataField       =   "Country"
         DataSource      =   "Data1"
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
         ForeColor       =   &H000000C0&
         Height          =   330
         ItemData        =   "frmSTEAM.frx":042D
         Left            =   3000
         List            =   "frmSTEAM.frx":042F
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
         Format          =   57540609
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
         TabIndex        =   201
         Top             =   180
         Visible         =   0   'False
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
         TabIndex        =   209
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
         TabIndex        =   210
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
         TabIndex        =   211
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
         TabIndex        =   212
         Text            =   "Unit"
         Top             =   1440
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label35 
         Caption         =   "Design"
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   240
         TabIndex        =   230
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Label lblLabels 
         Caption         =   "Test_n°:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   21
         Left            =   5340
         TabIndex        =   110
         Top             =   180
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "Date  (must be unique for same unit!):"
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   225
         TabIndex        =   97
         Top             =   285
         Width           =   2805
      End
      Begin VB.Label lblLabels 
         Caption         =   "Plant name:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   20
         Left            =   225
         TabIndex        =   96
         Top             =   615
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Location:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   19
         Left            =   225
         TabIndex        =   95
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Country:"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   18
         Left            =   225
         TabIndex        =   94
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Unit identification (must be unque):"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   17
         Left            =   225
         TabIndex        =   62
         Top             =   1500
         Width           =   2715
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Height          =   510
      Left            =   120
      TabIndex        =   223
      Top             =   8820
      Width           =   14715
      _ExtentX        =   25956
      _ExtentY        =   900
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
            TextSave        =   "13/07/2004"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "20.45"
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
End
Attribute VB_Name = "frmSTEAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public XXX As Integer
Public metal As String
Public A40, D11, D13, D17, D18, D19, D20, D21, D29, D38, D40, D43, D44, D54, D67, D68, D69, D70, D71, D72, D73, D74, D75, D76, D77, D78, D80 As Double
Private Sub Spin_txt_no_Change()
'     txt_num.Text = Spin_txt_no.Value
End Sub
Private Sub Check_PF_Click()
    If Check_PF = Checked Then
        Proc_FF_KW = Format(Spin_PF.Value / 100, "0.0")
        Proc_FF = Format(Proc_FF_KW / 859.8 * 10, "0.000")
    Else
        Proc_FF_KW = Data1.Recordset.Proc_F
        Proc_FF = Format(Proc_FF_KW / 859.8 * 10, "0.000")
    End If
    Call Calculation
    Call Calculation_act
End Sub
Private Sub Check_PF_act_Click()
    If Check_PF_act = Checked Then
        Proc_FF_act_KW.Text = Format(Spin_PF_act.Value / 100, "0.0")
        Proc_FF_act = Format(Proc_FF_act_KW / 859.8 * 10, "0.000")
    Else
        Proc_FF_act_KW.Text = Data1.Recordset.Proc_F_act
        Proc_FF_act.Text = Format(Proc_FF_act_KW / 859.8 * 10, "0.000")
    End If
    Call Calculation
    Call Calculation_act
End Sub
Private Sub Comm_graph_reset_Click()
    On Error Resume Next
    Dim Rs1 As Recordset
    Data1.DatabaseName = "C:\Condensers\Database\steam.mdb"
    Data1.RecordSource = "Select * From [Query_Test]"
'    Data1.Refresh
    Set Rs1 = Data1.Recordset
    Rs1.MoveLast
    n_rec = Rs1.RecordCount
    Rs1.MoveFirst
    
    If Data1.Recordset.RecordCount > 0 Then
       Do Until Rs1.EOF
            rec_x = Data1.Recordset.AbsolutePosition + 1
                Check_TIN.Value = Unchecked
                Check_TOUT.Value = Unchecked
                Check_CP.Value = Unchecked
                Check_VEL.Value = Unchecked
                Check_PD.Value = Unchecked
                Check_RE.Value = Unchecked
                Check_CT.Value = Unchecked
                Check_LMTD.Value = Unchecked
                Check_TTD.Value = Unchecked
                Check_CF.Value = Unchecked
                Check_FF.Value = Unchecked
                Check_WF.Value = Unchecked
                Check_SF.Value = Unchecked
                Check_PP.Value = Unchecked
                Check_CFAC.Value = Unchecked

          
         If rec_x = n_rec Then
            Exit Do
        End If
          If Data1.Recordset.EOF Then
                Exit Do
          End If
          Data1.Recordset.MoveNext
          rec_x = Data1.Recordset.AbsolutePosition + 1
       Loop
    End If
'    Data1.UpdateRecord
'    Data1.Recordset.Bookmark = Data1.Recordset.LastModified
End Sub
Private Sub Spin_PF_Change()
    If Check_PF = Checked Then
        Proc_FF_KW.Text = Format(Spin_PF / 100, "0.0")
        Proc_FF = Format(Proc_FF_KW / 859.8 * 10, "0.0000")
        D40 = Val(Proc_FF.Text / 10000)
    ElseIf Spin_PF = Unchecked Then
        Proc_FF_KW.Text = Format(Data1.Recordset.Proc_F, "0.0")
        Proc_FF = Format(Proc_FF_KW / 859.8 * 10, "0.0000")
'        D40 = Val(Proc_FF.Text/10000)
    End If
    Call Calculation
    Call Calculation_act
End Sub
Private Sub Spin_PF_act_Change()
    If Check_PF_act = Checked Then
        Proc_FF_act_KW.Text = Format(Spin_PF_act / 100, "0.0")
        Proc_FF_act = Format(Proc_FF_act_KW.Text / 859.8 * 10, "0.0000")
        A40 = Val(Proc_FF_act.Text / 10000)
    ElseIf Spin_PF_act = Unchecked Then
        Proc_FF_act_KW = Format(Data1.Recordset.Proc_F_act, "0.0")
        Proc_FF_act = Format(Proc_FF_act_KW.Text / 859.8 * 10, "0.000")
    End If
    Call Calculation
    Call Calculation_act
End Sub
Private Sub TabStrip1_Click()
On Error Resume Next
    If TabStrip1.SelectedItem = "Update" Then
        Data1.UpdateRecord
        Data1.Recordset.Bookmark = Data1.Recordset.LastModified
    ElseIf TabStrip1.SelectedItem = "Add" Then
            Data1.Recordset.AddNew
            Data1.UpdateRecord
            Data1.Recordset.Bookmark = Data1.Recordset.LastModified
            Data1.Recordset.MoveLast
            Call RESET
            Call RESET
            Data1.UpdateRecord
            Data1.Recordset.Bookmark = Data1.Recordset.LastModified

    ElseIf TabStrip1.SelectedItem = "Close" Then
'            Me.WindowState = vbMinimized
            Unload Me
            
    ElseIf TabStrip1.SelectedItem = "Delete" Then
            If Data1.Recordset.AbsolutePosition + 1 = 1 Then
                MsgBox (" You cannot delete this record")
                Exit Sub
            End If
            If Data1.Recordset.Check_Design = "Falso" Then
                Reply = MsgBox("Confirm to delete this record?", vbYesNo, "Delete record")
            Else
                Reply = MsgBox("This record contain the design data! Confirm to delete this record?", vbYesNo, "Delete record")
            End If
            If Reply = vbYes Then
                Data1.Recordset.Delete
                Data1.Recordset.MovePrevious
            End If
    ElseIf TabStrip1.SelectedItem = "Refresh" Then
            Data1.Refresh
    ElseIf TabStrip1.SelectedItem = "Print" Then
            frmSTEAM.PrintForm
    ElseIf TabStrip1.SelectedItem = "Get design" Then
            For i = 1 To 1
            UNIT1 = Data1.Recordset.Unit_name
            PLANT1 = Data1.Recordset.Plant
            C_DES1 = Data1.Recordset.Check_Design
            Dim Rs2 As Recordset
            Data2.DatabaseName = "C:\Condensers\Database\steam.mdb"
            Data2.RecordSource = "Select * From [Query_test]"
            Data2.Refresh
            Set Rs2 = Data2.Recordset
            Data2.Recordset.MoveFirst
            If Rs2.RecordCount > 0 Then
               Do Until Rs2.EOF
                  UNIT2 = Data2.Recordset.Unit_name
                  PLANT2 = Data2.Recordset.Plant
                  C_DES2 = Data2.Recordset.Check_Design
                  DES = Check_des.Value
                  If UNIT1 = UNIT2 And PLANT1 = PLANT2 And C_DES2 = "Vero" Then
                        Plant.Text = Data2.Recordset.Plant
                        Location.Text = Data2.Recordset.Location
                        Country.Text = Data2.Recordset.Country
                        Unit.Text = Data2.Recordset.Unit_name
                        Combo_PLANT.Text = Plant.Text
                        Combo_LOC.Text = Location.Text
                        Combo_Country.Text = Country.Text
                        Combo_UNIT.Text = Unit.Text
                        
                        T_NO.Text = Data2.Recordset.Tube_NO
                        T_len.Text = Data2.Recordset.Tube_LE
                        T_PASS.Text = Data2.Recordset.Tube_PAS
                        T_OD.Text = Data2.Recordset.Tube_OD
                        Combo_BWG.Text = Data2.Recordset.BWG
                        Combo_Mat.Text = Data2.Recordset.Tube_mat
                        Pump_EFF.Text = Data2.Recordset.Pump_EFF
                        
                        W_flow_IN.Text = Format(Data2.Recordset.Water_flow, "0,00")
                        W_T_IN.Text = Format(Data2.Recordset.Temp_IN, "0.00")
                        W_T_OUT.Text = Format(Data2.Recordset.Temp_OUT, "0.00")
                        S_flow_IN.Text = Format(Data2.Recordset.Steam_flow, "0,00")
                        S_press.Text = Format(Data2.Recordset.Steam_press, "0.0000")
                        S_press_KP.Text = Format(S_press.Text * 100, "0.00")
                        Combo_S_side.Text = Data2.Recordset.Shell_fluid
                        Combo_T_side.Text = Data2.Recordset.Tube_fluid
                        
                        CPp = Data2.Recordset.Check_PF
                        
                        If CPp = "Vero" Then
                            Check_PF.Value = 1
                            Proc_FF_KW.Text = Data2.Recordset.Proc_F
                            Spin_PF.Value = Proc_FF_KW.Text * 100
                            
                        Else
                            Check_PF.Value = 0
                        End If
                    Exit Do
                 End If
                 
                 If Data2.Recordset.EOF Then
                    MsgBox ("Not found the design data for this unit")
                    GoTo 10
                     Exit Do
                 End If
                 Data1.UpdateRecord
                 Rs2.MoveNext
                 If Data2.Recordset.EOF Then
                    MsgBox ("Not found the design data for this unit")
                    
                     Exit Do
                 End If
                 rr = Data2.Recordset.AbsolutePosition + 1
              Loop
                Call Calculation
                Call Calculation_act
            End If
       Next i
    ElseIf TabStrip1.SelectedItem = "Reset" Then
        Call RESET
    End If
10 End Sub
Private Sub RESET()
On Error Resume Next
                
            Data1.DatabaseName = "C:\Condensers\Database\steam.mdb"
            Data1.RecordSource = "Select * From [Query_test]"
            Set Rs1 = Data1.Recordset
                
                DTPicker1.Value = Date
                Plant.Text = "Plant"
                Location.Text = "Location"
                Country.Text = "Country"
                Unit.Text = "UNIT"

                Combo_PLANT.Text = "Plant"
                Combo_LOC.Text = "Location"
                Combo_Country.Text = "Country"
                Combo_UNIT.Text = "UNIT"

                T_NO.Text = 10000
                T_len.Text = 6.1
                T_PASS.Text = 1
                T_OD.Text = 19.05
                Combo_BWG.Text = 22
                Combo_Mat.Text = "Titanium"
                Pump_EFF.Text = Format(85, "0.0")
                
                HScroll_T_NO.Value = T_NO.Text
                Spin_T_LEN.Value = T_len.Text * 10
                Spin_T_PAS.Value = T_PASS.Text
                Spin_T_OD.Value = T_OD.Text * 100
                Spin_P_EFF.Value = Pump_EFF.Text * 10
                
                W_flow_IN.Text = Format(18000000, "0,00")
                W_T_IN.Text = Format(23, "0.00")
                W_T_OUT.Text = Format(29, "0.00")
                S_flow_IN.Text = Format(200000, "0,00")
                S_press.Text = Format(0.065, "0.000")
                S_press_KP.Text = Format(S_press.Text * 100, "0.00")
                
                Combo_S_side.Text = "Steam"
                Combo_T_side.Text = "Water"
            
                W_F = W_flow_IN.Text / 10000
                HScroll_W_FLOW.Value = Val(W_F)
                Spin_W_T_IN.Value = Val(W_T_IN.Text) * 100
                Spin_W_T_OUT.Value = Val(W_T_OUT.Text) * 100
                S_F = S_flow_IN.Text / 100
                Spin_S_FLOW_IN.Value = Val(S_F)
                Spin_S_PRESS.Value = S_press.Text * 100
                
                W_T_IN_act.Text = Format(23, "0.00")
                W_T_OUT_act.Text = Format(29, "0.00")
                S_flow_IN_act.Text = Format(200000, "0,00")
                S_press_act.Text = Format(0.065, "0.000")
                S_press_act_KP.Text = Format(S_press_act.Text * 100, "0.00")
                
                Spin_W_T_IN_act.Value = Val(W_T_IN_act.Text) * 100
                Spin_W_T_OUT_act.Value = Val(W_T_OUT_act.Text) * 100
                S_F_act = S_flow_IN_act.Text / 100
                Spin_S_FLOW_IN_act.Value = Val(S_F_act)
                Spin_S_PRESS_act.Value = S_press_act.Text * 100
                
        Data1.UpdateRecord
        Data1.Recordset.Bookmark = Data1.Recordset.LastModified
        
         Call Calculation
         Call Calculation_act

End Sub
Private Sub Data1_Error(DataErr As Integer, Response As Integer)
  'Posizione per il codice di gestione degli errori
  'Se si desidera ignorare gli errori, impostare come commento la riga successiva
  'Se si desidera intercettarli, aggiungere qui il codice di gestione
  MsgBox "Intercettato errore dei dati:" & Error$(DataErr)
  Response = 0  'Ignora l'errore
End Sub
Private Sub Data1_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  'Visualizza la posizione del record corrente
  'per Recordset di tipo Dynaset e Snapshot
  Data1.Caption = "Control: " & (Data1.Recordset.AbsolutePosition + 1)
  'Per l'oggetto tabella è necessario impostare la proprietà Index
  'al momento della creazione del Recordset e utilizzare la riga seguente
  'Data1.Caption = "Record: " & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1

'    Dim cn As New ADODB.Connection
''    cn.Open "Provider=Microsoft Jet 6.0 OLE DB Provider;Data source=c:\Condenser\Database\Steam.mdb;Jet & OLEDB:Database Password=gifra"
'    cn.Provider = "Microsoft Jet 4.0 OLE DB Provider"
'    cn.ConnectionString = "Data source=c:\Condensers\Database\Steam.mdb"
'    cn.Properties("Jet OLEDB: :Database Password") = "gifra"
'    cn.Open

If XXX = 1 Then
GoTo 100
End If

'    UNIT_REC = Data1.Recordset.Unit_name
'    PLANT_REC = Data1.Recordset.Plant
    Dim Rs2 As Recordset
    Data2.DatabaseName = "C:\Condensers\Database\steam.mdb"
    Data2.RecordSource = "Select * From [Query_test]"
    Data2.Refresh
    Set Rs2 = Data2.Recordset
    Data2.Recordset.MoveFirst
    UNIT_REC = Data1.Recordset.Unit_name
    PLANT_REC = Data1.Recordset.Plant
    
    If Rs2.RecordCount > 0 Then
       Do Until Rs2.EOF
          UNIT_X.Text = UNIT_REC
          PLANT_X.Text = PLANT_REC
          If Data2.Recordset.EOF Then
            Exit Do
          End If
            Data1.UpdateRecord
            Rs2.MoveNext
            rr = Data2.Recordset.AbsolutePosition
            Loop
    End If

     txt_num.Text = Data1.Recordset.AbsolutePosition + 1
    
    T_NO.Text = Val(Data1.Recordset.Tube_NO)
    T_len.Text = Val(Data1.Recordset.Tube_LE)
    T_PASS.Text = Val(Data1.Recordset.Tube_PAS)
    T_OD.Text = Val(Data1.Recordset.Tube_OD)
    Combo_BWG.Text = Val(Data1.Recordset.BWG)
    Combo_Mat.Text = Data1.Recordset.Tube_mat
    Pump_EFF.Text = Data1.Recordset.Pump_EFF
    
    HScroll_T_NO.Value = T_NO.Text
    Spin_T_LEN.Value = T_len.Text * 10
    Spin_T_PAS.Value = T_PASS.Text
    Spin_T_OD.Value = T_OD.Text * 100
    Spin_P_EFF.Value = Pump_EFF.Text * 10
    Check_PF = Data1.Recordset.Check_PF
    
    Combo_PLANT.Text = Data1.Recordset.Plant
    Combo_LOC.Text = Data1.Recordset.Location
    Combo_Country.Text = Data1.Recordset.Country
    Combo_UNIT.Text = Data1.Recordset.Unit_name
    
    W_flow_IN.Text = Format(Data1.Recordset.Water_flow, "0,00")
    W_T_IN.Text = Format(Val(Data1.Recordset.Temp_IN), "0.00")
    W_T_OUT.Text = Format(Val(Data1.Recordset.Temp_OUT), "0.00")
    S_flow_IN.Text = Format(Data1.Recordset.Steam_flow, "0,00")
    S_press.Text = Format(Val(Data1.Recordset.Steam_press), "0.0000")
    Combo_S_side.Text = Data1.Recordset.Shell_fluid
    Combo_T_side.Text = Data1.Recordset.Tube_fluid
    
    W_F = W_flow_IN.Text / 10000
    HScroll_W_FLOW.Value = Val(W_F)
    Spin_W_T_IN.Value = Val(W_T_IN.Text) * 100
    Spin_W_T_OUT.Value = Val(W_T_OUT.Text) * 100
'    S_F = S_flow_IN.Text / 100
    Spin_S_FLOW_IN.Value = Val(S_flow_IN.Text / 100)
    Spin_S_PRESS.Value = S_press.Text * 10000
    
    W_flow_IN_act.Text = Format(Data1.Recordset.Water_flow_act, "0,00")
    W_T_IN_act.Text = Format(Val(Data1.Recordset.Temp_IN_act), "0.00")
    W_T_OUT_act.Text = Format(Val(Data1.Recordset.Temp_OUT_act), "0.00")
    S_flow_IN_act.Text = Format(Data1.Recordset.Steam_flow_act, "0,00")
    S_press_act.Text = Format(Val(Data1.Recordset.Steam_press_act), "0.0000")
    Combo_S_side_act.Text = Data1.Recordset.Shell_fluid_act
    Combo_T_side_act.Text = Data1.Recordset.Tube_fluid_act

    Spin_W_T_IN_act.Value = Val(W_T_IN_act.Text) * 100
    Spin_W_T_OUT_act.Value = Val(W_T_OUT_act.Text) * 100
'    S_F_act = S_flow_IN_act.Text / 100
    Spin_S_FLOW_IN_act.Value = Val(S_flow_IN_act.Text / 100)
    Spin_S_PRESS_act.Value = S_press_act.Text * 10000
    Check_PF_act = Data1.Recordset.Check_PF_act


    PF = Format(Data1.Recordset.Proc_F * 100, "0.00")
    Spin_PF.Value = PF
    PF_act = Format(Data1.Recordset.Proc_F_act * 100, "0.00")
    Spin_PF_act.Value = PF_act
    
Call Calculation
Call Calculation_act

100 End Sub
Private Sub Data1_Validate(Action As Integer, Save As Integer)
  'Posizione per il codice di convalida
  'Questo evento viene richiamato quando si verificano le seguenti azioni
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
End Sub
Private Sub Form_Load()
   Width = frmMain.Width * 0.98 ' Imposta la larghezza del form.
   Height = frmMain.Height * 0.89     ' Imposta l'altezza del form.
   Left = 50 '(frmMain.Width - Width) / 2 ' Centra il form orizzontalmente.
   Top = 0 '(frmMain.Height - Height) / 2 ' Centra il form verticalmente.

End Sub
Private Sub Form_Initialize()
On Error Resume Next
    Dim Date_X As Date
    Dim Rs3 As Recordset
    
    Data3.DatabaseName = "C:\Condensers\Database\steam.mdb"
    Data3.RecordSource = "Select * From [QUERY_Plant]"
    Data3.Refresh
    Set Rs3 = Data3.Recordset
    If Rs3.RecordCount > 0 Then
       Do Until Rs3.EOF
          PPP = Data3.Recordset.Plant
'          Date_X = Data3.Recordset.Date_test
          Combo_PLANT.AddItem PPP
          Combo_Plant_1.AddItem PPP
'          Combo_Date_X.AddItem Date_X
          
          Rs3.MoveNext
       Loop
    Else
       MsgBox "Plant not found"
    End If
'        Combo_PLANT.ListIndex = 0
        
    On Error Resume Next
    Dim Rs4 As Recordset
    Data4.DatabaseName = "C:\Condensers\Database\steam.mdb"
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
    
    
    Data7.DatabaseName = "C:\Condensers\Database\steam.mdb"
    Data7.RecordSource = "Select * From [QUERY_Date]"
    Data7.Refresh
    Set Rs7 = Data7.Recordset
    If Rs7.RecordCount > 0 Then
       Do Until Rs7.EOF
          Date_X = Data7.Recordset.Date_test
          Combo_Date_X.AddItem Date_X
          Rs7.MoveNext
       Loop
'    Else
'       MsgBox "Date not found"
    End If
'        Combo_Date_X.ListIndex = 0

        
    On Error Resume Next
    Dim Rs5 As Recordset
    Data5.DatabaseName = "C:\Condensers\Database\steam.mdb"
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
        
    On Error Resume Next
    Dim Rs6 As Recordset
    Data6.DatabaseName = "C:\Condensers\Database\steam.mdb"
    Data6.RecordSource = "Select * From [Query_Country]"
    Data6.Refresh
    Set Rs6 = Data6.Recordset
'    Combo_UNIT.Clear
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
    
End Sub
Private Sub Combo_Plant_LostFocus()
    Plant.Text = Combo_PLANT.Text
End Sub
Private Sub Combo_Country_LostFocus()
    Country.Text = Combo_Country.Text
End Sub
Private Sub Combo_LOC_LostFocus()
    Location.Text = Combo_LOC.Text
End Sub
Private Sub Combo_UNIT_LostFocus()
    Unit.Text = Combo_UNIT.Text
End Sub

Private Sub Combo_S_side_LostFocus()
    Call Calculation
    Call Calculation_act
End Sub
Private Sub Combo_T_side_LostFocus()
    Call Calculation
    Call Calculation_act
End Sub
Private Sub Spin_S_FLOW_IN_Change()
    SFIN_D = Spin_S_FLOW_IN
    S_flow_IN.Text = Format(SFIN_D * 100, "0,00")
    Call Calculation
    Call Calculation_act
End Sub
Private Sub Spin_S_PRESS_Change()
    S_press_KP = Format(Spin_S_PRESS / 100, "0.00")
    S_press = S_press_KP.Text / 100
    Call Calculation
    Call Calculation_act
End Sub
Private Sub HScroll_T_NO_Change()
    T_NO.Text = Val(HScroll_T_NO.Value)
    Call Calculation
    Call Calculation_act
End Sub
Private Sub Spin_T_LEN_Change()
    T_len.Text = Spin_T_LEN.Value / 10
    Call Calculation
    Call Calculation_act
End Sub
Private Sub Spin_T_PAS_Change()
    T_PASS.Text = Spin_T_PAS.Value
    Call Calculation
    Call Calculation_act
End Sub
Private Sub Spin_T_OD_Change()
    T_OD = Spin_T_OD / 100
    Call Calculation
    Call Calculation_act
End Sub
Private Sub ComboT_OD_LostFocus()
On Error Resume Next
Dim OD_inches As String

OD_inches = ComboT_OD.Text
    
        Select Case OD_inches
        Case "1/2"
            T_OD = 12.7               'OD  (mm)
        Case "3/4"
            T_OD = 19.05              'OD  (mm)
        Case "7/8"
            T_OD = 22.225               'OD  (mm)
        Case "1"
            T_OD = 25.4               'OD  (mm)
        Case "1-1/8"
            T_OD = 28.575               'OD  (mm)
        Case "1-1/4"
            T_OD = 31.75               'OD  (mm)
        Case "1-1/2"
            T_OD = 38.1               'OD  (mm)
        Case "1-3/4"
            T_OD = 44.45               'OD  (mm)
        Case "1-7/8"
            T_OD = 47.625               'OD  (mm)
        Case "2"
            T_OD = 50.8               'OD  (mm)
        End Select
    Spin_T_OD.Value = T_OD * 100
    Call Calculation
    Call Calculation_act

    
End Sub
Private Sub Combo_BWG_LostFocus()
    Call Calculation
    Call Calculation_act
End Sub
Private Sub Combo_Mat_LostFocus()
    metal = Combo_Mat.Text
    
    Call Calculation
    Call Calculation_act

End Sub
Private Sub Spin_P_EFF_Change()
    Pump_EFF.Text = Format(Spin_P_EFF.Value / 10, "0.0")
    Call Calculation
    Call Calculation_act
End Sub
Private Sub HScroll_W_FLOW_Change()
'On Error Resume Next
    WFIN = HScroll_W_FLOW
    W_flow_IN.Text = Format(Val(WFIN) * 10000, "0,00")
    Call Calculation
    Call Calculation_act
    
End Sub
Private Sub Spin_W_T_IN_Change()
    W_T_IN = Format(Spin_W_T_IN / 100, "0.00")
    Call Calculation
    Call Calculation_act
End Sub
Private Sub Spin_W_T_OUT_Change()
    W_T_OUT = Format(Spin_W_T_OUT / 100, "0.00")
    Call Calculation
    Call Calculation_act

End Sub
Private Sub Calculation()
On Error Resume Next

Call Mechanical
    Mat_cond.Text = Val(D78)                                                    'Thermal conductivity of tube material
    If D54 > 0 Then
        Mat_factor.Text = Val(D54)                                                   ' Material factor
        CHT.ForeColor = &H80&
        CHT.BackColor = &H8000000F
    Else
        CHT.Text = "NA"
        CHT.ForeColor = &HFFFF&
        CHT.BackColor = &HFF&
'        MsgBox "Metal factor not available - Overall U clean coefficient cannot be calculated"
    End If
    D79 = 3.14159 * D67 * 25.4 * 10 ^ -3 * D75 * D74          ' m^2
    D80 = D79 / (0.3048 ^ 2)                                                       'inch^2
    Area.Text = Format(Val(D79), "0.0")

    D13 = Val(W_T_OUT.Text)                                                    ' Water T_OUT, °C
    D11 = Val(W_T_IN.Text)                                                         ' Water T_IN, °C
    D9 = W_flow_IN.Text                                                                'Water flow rate, kg/h
    D12 = D11 * 1.8 + 32                                                               ' Water T_IN, °F

'Saturated steam flowrate INLET condenser
    D29 = S_flow_IN.Text                                                               'kg/h
    D30 = D29 * 2.20462                                                                'lb/h

'Steam loading
    D31 = D30 / D80                                                                        'lb/h ft2

'Steam condensation pressure
    D32 = Val(S_press.Text)

Call Properties

'Latent heat
    I9 = 0.168682569821809
    J9 = -1.80896828868017E-04
    J3 = -38.2917529410035
    D38 = (-I9 - Sqr(I9 ^ 2 - 4 * J9 * (J3 - Log(D32) / 2.3))) / (2 * J9)
    Latent_heat.Text = Format(D38 * 4.1868, "0.00")

'Water flow rate m3/h
    D15 = D9 / D19

'Steam side duty
    D43 = D9 * D21 * (D13 - D11)                                                        'Water side duty
    D33 = 100 - ((D43 * 100) / (D29 * D38))                                       '% wet steam
    D44 = D29 * D38 * (100 - D33) / 100                                             'Steam side duty
    D14 = D44 / (D21 * (D13 - D11))                                                     'Water flow rate, kg/h
    D10 = D14 * 4.4028                                                                           'Water flow rate, gpm

'DELTA DUTY test
    D16 = (D44 - D43) / 1000

'Water velocity through tubes
    D22 = (D15 / (3600 * (3.14159 / 4) * D73 ^ 2 * (D74 / D77)))             'm/s
    D23 = D22 / 0.3048                                                                                    'fps
 
 'Reynolds
    D24 = D22 * D73 * D19 / (D20 / 1000)
    RE.Text = Format(D24, "0,00")
    
'Pressure drop through tubes (Hazen-Williams with C=130) Related to CS
    D25 = ((6.05 * 10 ^ 5 * ((D15 * 1000 / 60) / (D74 / D77)) ^ 1.85) / (130 ^ 1.85 * (D73 * 1000) ^ 4.87)) * D75 * D77

'Pressure drop due to return (estimated four velocity heads)
    D26 = ((4 * (D22 ^ 2 / (2 * 9.81))) / 10) * 0.9807
    
'Total pressure drop for 100% clean tube side
    D27 = D25 + D26
    Water_press_drop.Text = Format(D27 * 100, "0.00")                            ' KPa
    Water_press_drop_bar.Text = Format(D27, "0.00")                              ' bar

'C Factor
C_Factor.Text = Format(D15 / (D27 * 100) ^ (1 / 2), "0.0")                      'm3/h/kPa


'% of wet steam
    D33 = 100 - ((D43 * 100) / (D29 * D38))
    Wet_steam.Text = Format(D33, "0.0")
    
'Condensed steam temperature if 0.023<P_c<0.067
    J3 = -4.15371836241458
    I9 = 1500.11867695013
    J9 = -25402.3185815265
    K9 = 259023.283197929
    L9 = -1093047.81628863
    
    If D32 >= 0.023 And D32 <= 0.067 Then
        D34 = J3 + I9 * D32 + J9 * D32 ^ 2 + K9 * D32 ^ 3 + L9 * D32 ^ 4
    Else
        D34 = 0
    End If
    
'Condensed steam temperature if 0.067<P_c<0.1
    J3 = 6.95024724711694
    I9 = 748.012008925124
    J9 = -5810.38128441877
    K9 = 27768.5506851489
    L9 = -55837.443148253
    
    If D32 > 0.067 And D32 <= 0.1 Then
        D35 = J3 + I9 * D32 + J9 * D32 ^ 2 + K9 * D32 ^ 3 + L9 * D32 ^ 4
    Else
        D35 = 0
    End If

'Condensed steam temperature if P_c>0.1
    J3 = 17.4992653754406
    I9 = 402.36973631205
    J9 = -1491.36622741945
    K9 = 3336.98098975695
    L9 = -3080.83798502275
    
    If D32 > 0.1 Then
        D36 = J3 + I9 * D32 + J9 * D32 ^ 2 + K9 * D32 ^ 3 + L9 * D32 ^ 4
    Else
        D36 = 0
    End If

'Condensed steam  temperature
    D37 = D34 + D35 + D36
    Steam_cond_temp.Text = Format(D37, "0.00")

'Process fouling
Dim S_side(30), T_side(10), P_FF(30)

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
    S_side(16) = "Air, N2 etc (compressed)"
    S_side(17) = "Propane, Butane, etc."
    S_side(18) = "Water"

    T_side(1) = "Water"
    T_side(2) = "Air, N2 etc (compressed)"
    T_side(3) = "Steam condensing"
    T_side(4) = "Feed Water"

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
    P_FF(16) = 0.001025
    P_FF(17) = 0.0003074
    P_FF(18) = 0.0006148

Shell_side = Combo_S_side.Text
Tube_side = Combo_T_side.Text

For i = 1 To 18
    If Shell_side = S_side(i) Then
        If Tube_side = "Water" Then
            D40 = P_FF(i)
        ElseIf Shell_side = "Water" Then
            If Tube_eside = "Air, N2 etc (compressed)" Then
                    D40 = 0.001025
            ElseIf Tube_side = "Steam condensing" Then
                    D40 = 0.0003074
            End If
        End If
    End If
Next i
    
Water_den.Text = Format(Val(D19), "0.0")
Water_vis.Text = Format(Val(D20), "0.000")
Water_heat.Text = Format(Val(D21), "0.000")
Water_vel.Text = Format(Val(D22), "0.000")

    If Check_PF = Checked Then
        Proc_FF_KW = Format(Spin_PF.Value / 100, "0.0")
        Proc_FF = Format(Proc_FF_KW / 859.8 * 10, "0.00000")
'        Proc_FF = Spin_PF.Value / 100
        D40 = Val(Proc_FF.Text / 10000)
    Else
        Proc_FF_KW.Text = Format(D40 * 859.8 * 1000, "0.0")
    End If
    
'    Proc_FF_KW.Text = Format(D40 / 0.001163 * 1000, "0.000")
    Proc_FF.Text = Format(D40 * 10000, "0.000")

    

'Log Mean Temperature Difference
    D41 = ((D37 - D11) - (D37 - D13)) / Log((D37 - D11) / (D37 - D13))
    LMTD.Text = Format(D41, "0.00")

'Terminal temperature difference
    D42 = D37 - D13
    TTD.Text = Format(D42, "0.00")

'Water side duty
    D43 = D9 * D21 * (D13 - D11)
    W_duty.Text = Format(D43, "0,000")                                                  'Kcal/h
    W_MW.Text = Format(D43 / 859.845 / 1000, "0.00")                        'MW
    
'Steam side duty
    D44 = D29 * D38 * (100 - D33) / 100
    S_duty.Text = Format(D44, "0,000")
    S_MW.Text = Format(D44 / 859.845 / 1000, "0.00")

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

'Overall CLEAN heat transfer coefficient:
    D55 = D48 * D52 * D53 * D54 * 4.882
    CHT.Text = Format(Val(D55 * 0.001163) * 1, "0.000")

'Overall DIRTY heat transfer coefficient
    D56 = D44 / (D41 * D79)
    DHT.Text = Format(Val(D56 * 0.001163) * 1, "0.000")

'CLEANLINESS FACTOR
    D57 = D56 * 100 / D55
    CF.Text = Format(Val(D57), "0.0")
        
    If CF.Text > 100 Or CF.Text <= 0 Then
        CF.BackColor = &HFF&
    Else
        CF.BackColor = &H8000000F
    End If
    
'Water side individual heat transfer coeficient referred to ext. surface
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
    If D64 - (D59 + D60 + D61 + D62) < 0 Then
        D63 = 0
        diff = D59 + D60 + D61 + D62
    Else
        D63 = D64 - (D59 + D60 + D61 + D62)
    End If

'Water side fouling factor
    D65 = D63 * (D73 / D68)
'    W_FF.Text = Format(Val(D65), "0.000")                                      '[(h m^2 ºC)/Kcal]*10^-4
    W_FF.Text = Format(Val(D65 / 0.01163), "0.000")                     '[(m^2 ºC)/KW]*10^-4

'Pumping power, KW
    P_Power.Text = Format((D9 * D27 * 100000 / 3600 / D19 / Val(Pump_EFF.Text / 100)) / 1000, "0")

End Sub
Private Sub Mechanical()
On Error Resume Next
'MECHANICAL DATA
'Dim D67, D68, D69, D70, D71 As Double
    
    D68 = Val(T_OD.Text) / 1000             'Tube Outlet Diameter (m)
    D67 = D68 * 1000 / 25.4                     'Tube Outlet Diameter (inches)
    D69 = Val(Combo_BWG.Text)           'BWG

'BWG / Wall thickness
        Select Case D69
        Case 7
            D70 = 0.18              'Wall Thickness (inches)
        Case 8
            D70 = 0.165
        Case 9
            D70 = 0.148
        Case 10
            D70 = 0.134
        Case 11
            D70 = 0.12
        Case 12
            D70 = 0.109
        Case 13
            D70 = 0.095
        Case 14
            D70 = 0.083
        Case 15
            D70 = 0.072
        Case 16
            D70 = 0.065
        Case 17
            D70 = 0.058
        Case 18
            D70 = 0.049
        Case 19
            D70 = 0.042
        Case 20
            D70 = 0.035
        Case 22
            D70 = 0.028
        Case 24
            D70 = 0.022
        Case 26
            D70 = 0.018
    End Select
    
'Tube inlet diameter
    D71 = D70 * 25.4 * 10 ^ -3          'Wall Thickness (m)
    D72 = D67 - 2 * D70                    'Tube inlet diameter (inches)
    D73 = D72 * 25.4 * 10 ^ -3          'Tube inlet diameter (meters)
    T_ID.Text = Format(Val(D73) * 1000, "0.00")
    D74 = Val(T_NO.Text)                  'Number of tubes
    D75 = Val(T_len.Text)                  'Tube lenght (m)
    D76 = D75 / 0.3048                      'Tube lenght (inches)
    D77 = Val(T_PASS.Text)              'Number of tube side passes
    T_IN.Text = Format(D73 * 1000, "0.00")
    
'Tube Material Factor
Dim mat(21), Fac(21, 7), Material_cond(21)

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
    
    Material_cond(1) = 24.8
    Material_cond(2) = 38.4
    Material_cond(3) = 95.4
    Material_cond(4) = 86.8
    Material_cond(5) = 60.7
    Material_cond(6) = 84
    Material_cond(7) = 44.6
    Material_cond(8) = 38.4
    Material_cond(9) = 34
    Material_cond(10) = 76
    Material_cond(11) = 108.7
    Material_cond(12) = 100.4
    Material_cond(13) = 14
    Material_cond(14) = 21.4
    Material_cond(15) = 21.9
    Material_cond(16) = 14.1
    Material_cond(17) = 6.7
    Material_cond(18) = 6.2
    Material_cond(19) = 172.3
    Material_cond(20) = 102.9
    Material_cond(21) = 11.8
    
    Fac(1, 1) = 0.64
    Fac(1, 2) = 0.71
    Fac(1, 3) = 0.77
    Fac(1, 4) = 0.82
    Fac(1, 5) = 0.87
    Fac(1, 6) = 0.9
    Fac(1, 7) = 0.93
    
    Fac(2, 1) = 0.74
    Fac(2, 2) = 0.8
    Fac(2, 3) = 0.85
    Fac(2, 4) = 0.9
    Fac(2, 5) = 0.94
    Fac(2, 6) = 0.97
    Fac(2, 7) = 0.99
    
    Fac(3, 1) = 0.87
    Fac(3, 2) = 0.92
    Fac(3, 3) = 0.96
    Fac(3, 4) = 1
    Fac(3, 5) = 1.02
    Fac(3, 6) = 1.04
    Fac(3, 7) = 1.06
    
    Fac(4, 1) = 0.84
    Fac(4, 2) = 0.9
    Fac(4, 3) = 0.94
    Fac(4, 4) = 0.97
    Fac(4, 5) = 1
    Fac(4, 6) = 1.02
    Fac(4, 7) = 1.03
    
    Fac(5, 1) = 0.89
    Fac(5, 2) = 0.9
    Fac(5, 3) = 0.94
    Fac(5, 4) = 0.97
    Fac(5, 5) = 1
    Fac(5, 6) = 1.02
    Fac(5, 7) = 1.03
    
    Fac(6, 1) = 0.87
    Fac(6, 2) = 0.92
    Fac(6, 3) = 0.96
    Fac(6, 4) = 1
    Fac(6, 5) = 1.02
    Fac(6, 6) = 1.04
    Fac(6, 7) = 1.06
    
    Fac(7, 1) = 0.74
    Fac(7, 2) = 0.8
    Fac(7, 3) = 0.86
    Fac(7, 4) = 0.91
    Fac(7, 5) = 0.95
    Fac(7, 6) = 0.98
    Fac(7, 7) = 1
    
    Fac(8, 1) = 0.74
    Fac(8, 2) = 0.8
    Fac(8, 3) = 0.86
    Fac(8, 4) = 0.91
    Fac(8, 5) = 0.95
    Fac(8, 6) = 0.98
    Fac(8, 7) = 1

    Fac(9, 1) = 0.74
    Fac(9, 2) = 0.8
    Fac(9, 3) = 0.86
    Fac(9, 4) = 0.91
    Fac(9, 5) = 0.95
    Fac(9, 6) = 0.98
    Fac(9, 7) = 1
    
    Fac(10, 1) = 0.87
    Fac(10, 2) = 0.92
    Fac(10, 3) = 0.96
    Fac(10, 4) = 1
    Fac(10, 5) = 1.02
    Fac(10, 6) = 1.04
    Fac(10, 7) = 1.06
    
    Fac(11, 1) = 0
    Fac(11, 2) = 0
    Fac(11, 3) = 0
    Fac(11, 4) = 0
    Fac(11, 5) = 0
    Fac(11, 6) = 0
    Fac(11, 7) = 0
    
    Fac(12, 1) = 0.87
    Fac(12, 2) = 0.92
    Fac(12, 3) = 0.96
    Fac(12, 4) = 1
    Fac(12, 5) = 1.02
    Fac(12, 6) = 1.04
    Fac(12, 7) = 1.06
    
    Fac(13, 1) = 0.49
    Fac(13, 2) = 0.56
    Fac(13, 3) = 0.63
    Fac(13, 4) = 0.69
    Fac(13, 5) = 0.75
    Fac(13, 6) = 0.79
    Fac(13, 7) = 0.83
    
    Fac(14, 1) = 0.59
    Fac(14, 2) = 0.65
    Fac(14, 3) = 0.7
    Fac(14, 4) = 0.76
    Fac(14, 5) = 0.82
    Fac(14, 6) = 0.85
    Fac(14, 7) = 0.88

    Fac(15, 1) = 0.87
    Fac(15, 2) = 0.92
    Fac(15, 3) = 0.96
    Fac(15, 4) = 1
    Fac(15, 5) = 1.02
    Fac(15, 6) = 1.04
    Fac(15, 7) = 1.06

    Fac(16, 1) = 0.71
    Fac(16, 2) = 0.71
    Fac(16, 3) = 0.71
    Fac(16, 4) = 0.71
    Fac(16, 5) = 0.77
    Fac(16, 6) = 0.81
    Fac(16, 7) = 0.85

    Fac(17, 1) = 0.71
    Fac(17, 2) = 0.71
    Fac(17, 3) = 0.71
    Fac(17, 4) = 0.71
    Fac(17, 5) = 0.77
    Fac(17, 6) = 0.81
    Fac(17, 7) = 0.85

    Fac(18, 1) = 0.71
    Fac(18, 2) = 0.71
    Fac(18, 3) = 0.71
    Fac(18, 4) = 0.71
    Fac(18, 5) = 0.77
    Fac(18, 6) = 0.81
    Fac(18, 7) = 0.85

    Fac(19, 1) = 0.54
    Fac(19, 2) = 0.6
    Fac(19, 3) = 0.65
    Fac(19, 4) = 0.69
    Fac(19, 5) = 0.74
    Fac(19, 6) = 0.76
    Fac(19, 7) = 0.78

    Fac(20, 1) = 0.87
    Fac(20, 2) = 0.92
    Fac(20, 3) = 0.96
    Fac(20, 4) = 1
    Fac(20, 5) = 1.02
    Fac(20, 6) = 1.04
    Fac(20, 7) = 1.06

    Fac(21, 1) = 0
    Fac(21, 2) = 0
    Fac(21, 3) = 0
    Fac(21, 4) = 0
    Fac(21, 5) = 0
    Fac(21, 6) = 0
    Fac(21, 7) = 0

metal = Combo_Mat.Text
'metal = Data1.Recordset.Tube_mat
For i = 1 To 21
    If metal = mat(i) Then
    D78 = Material_cond(i)
        If D69 = 12 Then
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

End Sub
Private Sub Properties()
On Error Resume Next

'Water Viscosity, cP
    D17 = (D11 + D13) / 2
    D20 = (100 / (2.1482 * ((273.16 + D17 - 281.435) + Sqr(8078.4 + (273.16 + D17 - 281.435) ^ 2)) - 120))

'Water density, kg/m3
    D17_2 = Int(D17 / 2)
    D18 = D17 * 1.8 + 32
    
    Select Case D17_2
        Case 1
            D19 = 999.94
        Case 2
            D19 = 999.97
        Case 3
            D19 = 999.94
        Case 4
            D19 = 999.85
        Case 5
            D19 = 999.7
        Case 6
            D19 = 999.497
        Case 7
            D19 = 999.244
        Case 8
            D19 = 998.943
        Case 9
            D19 = 998.595
        Case 10
            D19 = 998.204
        Case 11
            D19 = 997.77
        Case 12
            D19 = 997.296
        Case 13
            D19 = 996.783
        Case 14
            D19 = 996.233
        Case 15
            D19 = 995.647
        Case 16
            D19 = 995.026
        Case 17
            D19 = 994.371
        Case 18
            D19 = 993.684
        Case 19
            D19 = 992.965
        Case 20
            D19 = 992.215
        Case 21
            D19 = 991.436
        Case 22
            D19 = 990.628
        Case 23
            D19 = 989.792
        Case 24
            D19 = 988.928
        Case 25
            D19 = 988.037
        Case 26
            D19 = 987.12
        Case 27
            D19 = 986.177
        Case 28
            D19 = 985.219
        Case 29
            D19 = 984.217
    End Select
    
    'Average specific heat of water
    Select Case D17_2
        Case 1
            D21 = 1.00636
        Case 2
            D21 = 1.00495
        Case 3
            D21 = 1.00378
        Case 4
            D21 = 1.00277
        Case 5
            D21 = 1.00194
        Case 6
            D21 = 1.00124
        Case 7
            D21 = 1.00067
        Case 8
            D21 = 1.00019
        Case 9
            D21 = 0.999978
        Case 10
            D21 = 0.99947
        Case 11
            D21 = 0.99921
        Case 12
            D21 = 0.99902
        Case 13
            D21 = 0.99885
        Case 14
            D21 = 0.99873
        Case 15
            D21 = 0.99866
        Case 16
            D21 = 0.99861
        Case 17
            D21 = 0.99859
        Case 18
            D21 = 0.99861
        Case 19
            D21 = 0.99864
        Case 20
            D21 = 0.99869
        Case 21
            D21 = 0.99876
        Case 22
            D21 = 0.99883
        Case 23
            D21 = 0.99895
        Case 24
            D21 = 0.99907
        Case 25
            D21 = 0.99919
        Case 26
            D21 = 0.99935
        Case 27
            D21 = 0.9995
        Case 28
            D21 = 0.99969
        Case 29
            D21 = 0.99988
    End Select
End Sub
Private Sub Spin_S_FLOW_IN_act_Change()
    SFIN_A = Spin_S_FLOW_IN_act
    S_flow_IN_act.Text = Format(SFIN_A * 100, "0,00")
    Call Calculation_act
End Sub
Private Sub Spin_W_T_IN_act_Change()
    W_T_IN_act = Format(Spin_W_T_IN_act / 100, "0.00")
    Call Calculation_act
End Sub
Private Sub Spin_W_T_OUT_act_Change()
    W_T_OUT_act = Format(Spin_W_T_OUT_act / 100, "0.00")
    Call Calculation_act
End Sub
Private Sub Spin_S_PRESS_act_Change()
    S_press_act_KP = Format(Spin_S_PRESS_act / 100, "0.00")
    S_press_act = S_press_act_KP.Text / 100
    Call Calculation_act
End Sub
Private Sub Calculation_act()
On Error Resume Next

Call Mechanical
    Mat_cond.Text = Val(D78)                                                    'Thermal conductivity of tube material
    If D54 > 0 Then
        Mat_factor.Text = Val(D54)                                                   ' Material factor
        CHT_act.ForeColor = &H80&
        CHT_act.BackColor = &H8000000F
    Else
        CHT_act.Text = "NA"
        CHT_act.ForeColor = &HFFFF&
        CHT_act.BackColor = &HFF&
    End If
    D79 = 3.14159 * D67 * 25.4 * 10 ^ -3 * D75 * D74          ' m^2
    D80 = D79 / (0.3048 ^ 2)                                                       'inch^2
    Area.Text = Format(Val(D79), "0.0")

    A13 = Val(W_T_OUT_act.Text)                                                    ' Water T_OUT, °C
    A11 = Val(W_T_IN_act.Text)                                                         ' Water T_IN, °C
    A9 = W_flow_IN_act.Text                                                                'Water flow rate, kg/h
    A12 = A11 * 1.8 + 32                                                                       ' Water T_IN, °F

'Saturated steam flowrate INLET condenser
    A29 = S_flow_IN_act.Text                                                               'kg/h
    A30 = A29 * 2.20462                                                                        'lb/h

'Steam loading
    A31 = A30 / D80                                                                                 'lb/h ft2

'Steam condensation pressure
    A32 = Val(S_press_act.Text)

    D11 = A11
    D13 = A13
Call Properties

'Latent heat
    I9 = 0.168682569821809
    J9 = -1.80896828868017E-04
    J3 = -38.2917529410035
    A38 = (-I9 - Sqr(I9 ^ 2 - 4 * J9 * (J3 - Log(A32) / 2.3))) / (2 * J9)
    Latent_Heat_act.Text = Format(A38 * 4.1868, "0.00")

'Steam side duty
    A43 = A9 * D21 * (A13 - A11)
    A33 = 100 - ((D43 * 100) / (D29 * D38))
    A44 = A29 * A38 * (100 - A33) / 100
    A14 = D44 / (D21 * (A13 - A11))                                              'Water flow rate, kg/h
    A10 = A14 * 4.4028                                                                   'Water flow rate, gpm

'DELTA DUTY test
    
2770 W = 10
2780 j = 10
2790 HE = W: GoSub 2310
2800 Y = X: HE = j + W
2810 GoSub 2310
2820 G = W: W = G - j * Y / (X - Y)
2830 If Abs(G - W) >= 0.0000001 Then GoTo 2790
2840 W = HE: GoTo 2380

2310   A9 = HE
            A43 = A9 * D21 * (A13 - A11)
            X = (A44 - A43) / 1000
Return

2380           E15 = W
                    E9 = (E15 * 1)
                     W_flow_IN_act.Text = Format(E9, "0,00")

'Water flow rate m3/h
    A15 = A9 / D19

'Water velocity through tubes
    A22 = (A15 / (3600 * (3.14 / 4) * D73 ^ 2 * (D74 / D77)))                       'm/s
    A23 = A22 / 0.3048                                                                                       'fps
 
 'Reynolds
    A24 = A22 * D73 * D19 / (D20 * 10 ^ -3)
    RE_act.Text = Format(A24, "0,000")
    
'Pressure drop through tubes (Hazen-Williams with C=130) Related to CS
    A25 = ((6.05 * 10 ^ 5 * ((A15 * 1000 / 60) / (D74 / D77)) ^ 1.85) / (130 ^ 1.85 * (D73 * 1000) ^ 4.87)) * D75 * D77

'Pressure drop due to return (estimated four velocity heads)
    A26 = ((4 * (A22 ^ 2 / (2 * 9.81))) / 10) * 0.9807
    
'Total pressure drop for 100% clean tube side
    A27 = A25 + A26
   Water_press_drop_act.Text = Format(A27 * 100, "0.00")                                       'KPa
   Water_press_drop_act_bar.Text = Format(A27, "0.00")                                          'bar

'C Factor
C_Factor_act.Text = Format(A15 / (A27 * 100) ^ (1 / 2), "0.0")                          'm3/h/kPa


'% of wet steam
    A33 = 100 - ((A43 * 100) / (A29 * A38))
    Wet_steam_act.Text = Format(A33, "0.0")

    
'Condensed steam temperature if 0.023<P_c<0.067
    J3 = -4.15371836241458
    I9 = 1500.11867695013
    J9 = -25402.3185815265
    K9 = 259023.283197929
    L9 = -1093047.81628863
    
    If A32 >= 0.023 And A32 <= 0.067 Then
        A34 = J3 + I9 * A32 + J9 * A32 ^ 2 + K9 * A32 ^ 3 + L9 * A32 ^ 4
    Else
        A34 = 0
    End If
    
'Condensed steam temperature if 0.067<P_c<0.1
    J3 = 6.95024724711694
    I9 = 748.012008925124
    J9 = -5810.38128441877
    K9 = 27768.5506851489
    L9 = -55837.443148253
    
    If A32 > 0.067 And A32 <= 0.1 Then
        A35 = J3 + I9 * A32 + J9 * A32 ^ 2 + K9 * A32 ^ 3 + L9 * A32 ^ 4
    Else
        A35 = 0
    End If

'Condensed steam temperature if P_c>0.1
    J3 = 17.4992653754406
    I9 = 402.36973631205
    J9 = -1491.36622741945
    K9 = 3336.98098975695
    L9 = -3080.83798502275
    
    If A32 > 0.1 Then
        A36 = J3 + I9 * A32 + J9 * A32 ^ 2 + K9 * A32 ^ 3 + L9 * A32 ^ 4
    Else
        A36 = 0
    End If

'Condensed steam  temperature
    A37 = A34 + A35 + A36
    Steam_cond_temp_act.Text = Format(A37, "0.00")

'Process fouling
Dim S_side(30), T_side(10), P_FF(30)

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
    S_side(16) = "Air, N2 etc (compressed)"
    S_side(17) = "Propane, Butane, etc."
    S_side(18) = "Water"

    T_side(1) = "Water"
    T_side(2) = "Air, N2 etc (compressed)"
    T_side(3) = "Steam condensing"
    T_side(4) = "Feed Water"

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
    P_FF(16) = 0.001025
    P_FF(17) = 0.0003074
    P_FF(18) = 0.0006148

Shell_side = Combo_S_side.Text
Tube_side = Combo_T_side.Text

For i = 1 To 18
    If Shell_side = S_side(i) Then
        If Tube_side = "Water" Then
            A40 = P_FF(i)
        ElseIf Shell_side = "Water" Then
            If Tube_eside = "Air, N2 etc (compressed)" Then
                            A40 = 0.001025
            ElseIf Tube_side = "Steam condensing" Then
                            A40 = 0.0003074
            End If
        End If
    End If
Next i

Water_den_act.Text = Format(Val(D19), "0.0")
Water_vis_act.Text = Format(Val(D20), "0.000")
Water_heat_act.Text = Format(Val(D21), "0.000")
Water_vel_act.Text = Format(Val(A22), "0.000")

    If Check_PF_act = Checked Then
        Proc_FF_act_KW = Format(Spin_PF_act.Value / 100, "0.0")
        Proc_FF_act = Format(Proc_FF_act_KW / 859.8 * 10, "0.000")
        A40 = Val(Proc_FF_act.Text / 10000)
    Else
        Proc_FF_act_KW.Text = Format(A40 * 859.8 * 1000, "0.0")
    End If
    
'    Proc_FF_act_KW.Text = Format(A40 / 0.001163 * 1000, "0.0")
    Proc_FF_act.Text = Format(A40 * 10000, "0.000")

'Log Mean Temperature Difference
    A41 = ((A37 - A11) - (A37 - A13)) / Log((A37 - A11) / (A37 - A13))
    LMTD_act.Text = Format(A41, "0.00")

'Terminal temperature difference
    A42 = A37 - A13
    TTD_act.Text = Format(A42, "0.00")

'Water side duty
    A43 = A9 * D21 * (A13 - A11)
    W_duty_act.Text = Format(A43, "0,000")
    W_MW_act.Text = Format(A43 / 859.845 / 1000, "0.00")
    
'Steam side duty
    A44 = A29 * A38 * (100 - A33) / 100
  S_duty_act.Text = Format(A44, "0,000")
  S_MW_act.Text = Format(A44 / 859.845 / 1000, "0.00")

'Water temperature correction factor
    If A12 <= 50 Then
        A45 = 0.1228463 + 1.483184 * 10 ^ -2 * A12 - 2.17211 * 10 ^ -5 * A12 ^ 2
    Else
        A45 = 0
    End If
    If A12 > 50 And A12 < 70 Then
        A46 = 4.269782 * 10 ^ -2 + 1.931213 * 10 ^ -2 * A12 - 8.034383 * 10 ^ -5 * A12 ^ 2
    Else
        A46 = 0
    End If
    If A12 >= 70 Then
        A47 = 0.4470833 + 1.121063 * 10 ^ -2 * A12 - 4.696334 * 10 ^ -5 * A12 ^ 2
    Else
        A47 = 0
    End If
    A48 = A45 + A46 + A47

'Constant for tube outside diameter
    If D67 <= 0.75 Then
        A52 = 267 * Sqr(A23)
    ElseIf D67 > 0.75 And D67 <= 1 Then
        A52 = 263 * Sqr(A23)
    ElseIf D67 > 1 Then
        A52 = 259 * Sqr(A23)
    Else
        A52 = 0
    End If

'Loading correction factor
    A53 = (A31 / 8) ^ 0.25

'Overall CLEAN heat transfer coefficient:
    A55 = A48 * A52 * A53 * D54 * 4.882
    CHT_act.Text = Format(Val(A55 * 0.001163) * 1, "0.000")

'Overall DIRTY heat transfer coefficient
    A56 = A44 / (A41 * D79)
    DHT_act.Text = Format(Val(A56 * 0.001163) * 1, "0.000")

'CLEANLINESS FACTOR
    A57 = A56 * 100 / A55
    CF_act.Text = Format(Val(A57), "0.0")
    
    If CF_act.Text > 100 Or CF_act.Text <= 0 Then
        CF_act.BackColor = &HFF&
    Else
        CF_act.BackColor = &H8000000F
    End If

    
'Water side individual heat transfer coeficient referred to ext. surface
    A58 = (150 * (1 + 0.011 * A18) * (A23 ^ 0.8 / D72 ^ 0.2)) * 4.882 * (D72 / D67)

'Heat transfer resistance due to water flowing inside the tubes
    A59 = 10000 / A58

'Heat transfer resistance due to  the wall, referred to ext. surface
    A60 = (D68 * Log(D68 / D73) / (2 * D78)) * 10000

'Heat transfer resistance due to the condensing steam film
    A61 = (10000 / A55) - A59 - A60

'Heat transfer resistance due to outside fouling factor
    A62 = A40 * 10000

'Total heat transfer resistance, referred to external surface
    A64 = 10000 / A56

'Heat transfer resistance due to inside fouling factor referred to ext surface
    If A64 - (A59 + A60 + A61 + A62) < 0 Then
        A63 = 0
    Else
        A63 = A64 - (A59 + A60 + A61 + A62)
    End If

'Water side fouling factor
    A65 = A63 * (D73 / D68)
'    W_FF_act.Text = Format(Val(A65), "0.000")                                        '[(h m^2 ºC)/Kcal]*10^-4
    W_FF_act.Text = Format(Val(A65 / 0.01163), "0.000")                       '[(m^2 ºC)/KW]*10^-3

'Pumping power, KW
P_Power_act.Text = Format((A9 * A27 * 100000 / 3600 / D19 / Val(Pump_EFF.Text / 100)) / 1000, "0")
End Sub
Private Sub Combo_Plant_1_LostFocus()
On Error Resume Next
        Data1.Recordset.MoveFirst
        XXX = 1
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
Private Sub Combo_UNIT_1_LostFocus()
On Error Resume Next
        Data1.Recordset.MoveLast
        n_record = Data1.Recordset.RecordCount
        Data1.Recordset.MoveFirst
       XXX = 1
        While pos < 0
            Data1.Recordset.MoveNext
        Wend
        UUU1 = Combo_UNIT_1.Text
        While UUU1 <> UUU2
            n_rec_a = Data1.Recordset.AbsolutePosition + 1
            UUU2 = Data1.Recordset.Unit_name
            If UUU1 = UUU2 Then
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
Private Sub Combo_Unit_1_GotFocus()
    On Error Resume Next
    Dim Rs4 As Recordset
    Data4.DatabaseName = "C:\Condensers\Database\Steam.mdb"
    Data4.RecordSource = "Select * From [Query_Unit]"
    Data4.Refresh
    Set Rs4 = Data4.Recordset
    Combo_UNIT_1.Clear
    If Rs4.RecordCount > 0 Then
       Do Until Rs4.EOF
            PPP1 = Combo_Plant_1
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
            Date_2 = Data1.Recordset.Date_test
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

