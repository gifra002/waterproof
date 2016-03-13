VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "fm20.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Chart_ALL_UNITS 
   BackColor       =   &H00E0E0E0&
   Caption         =   "WaterProof HX - Heat Exchangers Performance Monitoring - Trend charts "
   ClientHeight    =   10245
   ClientLeft      =   60
   ClientTop       =   450
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
   MDIChild        =   -1  'True
   ScaleHeight     =   10245
   ScaleWidth      =   14895
   Begin VB.Frame Frame_PARAMETER 
      Caption         =   "Parameters"
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
      Height          =   4635
      Left            =   10500
      TabIndex        =   133
      Top             =   480
      Visible         =   0   'False
      Width           =   2415
      Begin VB.CheckBox Check_S_FLOW 
         Caption         =   "Check2"
         Height          =   210
         Left            =   2040
         TabIndex        =   18
         Top             =   3660
         Width           =   195
      End
      Begin VB.CheckBox Check_WATER_FOUL 
         Caption         =   "Check2"
         Height          =   210
         Left            =   2040
         TabIndex        =   17
         Top             =   3270
         Width           =   195
      End
      Begin VB.CheckBox Check_TTD 
         Caption         =   "Check2"
         Height          =   210
         Left            =   2040
         TabIndex        =   13
         Top             =   2190
         Width           =   195
      End
      Begin VB.CheckBox Check_CFACTOR 
         Caption         =   "Check2"
         Height          =   210
         Left            =   2040
         TabIndex        =   16
         Top             =   3000
         Width           =   195
      End
      Begin VB.CheckBox Check_CT 
         Caption         =   "Check2"
         Height          =   210
         Left            =   2040
         TabIndex        =   19
         Top             =   3930
         Width           =   195
      End
      Begin VB.CheckBox Check_LMTD 
         Caption         =   "Check2"
         Height          =   210
         Left            =   2040
         TabIndex        =   14
         Top             =   2460
         Width           =   195
      End
      Begin VB.CheckBox Check_SKIN 
         Caption         =   "Check2"
         Height          =   210
         Left            =   2040
         TabIndex        =   15
         Top             =   2730
         Width           =   195
      End
      Begin VB.CheckBox Check_CP 
         Caption         =   "Check2"
         Height          =   210
         Left            =   2040
         TabIndex        =   20
         Top             =   4245
         Width           =   195
      End
      Begin VB.CheckBox Check_DUTY 
         Caption         =   "Check2"
         Height          =   210
         Left            =   2040
         TabIndex        =   12
         Top             =   1920
         Width           =   195
      End
      Begin VB.CheckBox Check_T_FLOW 
         Caption         =   "Check2"
         Height          =   210
         Left            =   2040
         TabIndex        =   6
         Top             =   300
         Width           =   195
      End
      Begin VB.CheckBox Check_T_VEL 
         Caption         =   "Check2"
         Height          =   210
         Left            =   2040
         TabIndex        =   9
         Top             =   1110
         Width           =   195
      End
      Begin VB.CheckBox Check_T_PD 
         Caption         =   "Check2"
         Height          =   210
         Left            =   2040
         TabIndex        =   11
         Top             =   1650
         Width           =   195
      End
      Begin VB.CheckBox Check_T_RE 
         Caption         =   "Check2"
         Height          =   210
         Left            =   2040
         TabIndex        =   10
         Top             =   1380
         Width           =   195
      End
      Begin VB.CheckBox Check_T_TEMP_IN 
         Caption         =   "Check2"
         Height          =   210
         Left            =   2040
         TabIndex        =   7
         Top             =   570
         Width           =   195
      End
      Begin VB.CheckBox Check_T_TEMP_OUT 
         Caption         =   "Check2"
         Height          =   210
         Left            =   2040
         TabIndex        =   8
         Top             =   840
         Width           =   195
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         BorderStyle     =   6  'Inside Solid
         X1              =   120
         X2              =   2280
         Y1              =   3540
         Y2              =   3540
      End
      Begin VB.Label Label4 
         Caption         =   "Shell-side flow"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   13
         Left            =   180
         TabIndex        =   148
         Top             =   3660
         Width           =   1755
      End
      Begin VB.Label Label4 
         Caption         =   "Approach temperature"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   147
         Top             =   2195
         Width           =   1755
      End
      Begin VB.Label Label4 
         Caption         =   "Condensing  temperature"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   146
         Top             =   3945
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "MTDc"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   145
         Top             =   2470
         Width           =   1755
      End
      Begin VB.Label Label4 
         Caption         =   "Skin temperature"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   144
         Top             =   2745
         Width           =   1755
      End
      Begin VB.Label Label4 
         Caption         =   "Tube-side duty"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   143
         Top             =   1920
         Width           =   1755
      End
      Begin VB.Label Label4 
         Caption         =   "Tube-side flow rate"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   142
         Top             =   270
         Width           =   1755
      End
      Begin VB.Label Label4 
         Caption         =   "Tube-side flow velocity"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   6
         Left            =   180
         TabIndex        =   141
         Top             =   1095
         Width           =   1755
      End
      Begin VB.Label Label4 
         Caption         =   "Tube-side press. drop (clean"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   7
         Left            =   180
         TabIndex        =   140
         Top             =   1645
         Width           =   1755
      End
      Begin VB.Label Label4 
         Caption         =   "Tube-side Reynolds number"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   8
         Left            =   180
         TabIndex        =   139
         Top             =   1370
         Width           =   1755
      End
      Begin VB.Label Label4 
         Caption         =   "Tube-side temp. IN"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   9
         Left            =   180
         TabIndex        =   138
         Top             =   545
         Width           =   1755
      End
      Begin VB.Label Label4 
         Caption         =   "Tube-side temp. OUT"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   10
         Left            =   180
         TabIndex        =   137
         Top             =   820
         Width           =   1755
      End
      Begin VB.Label Label4 
         Caption         =   "Water side fouling factor"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   11
         Left            =   180
         TabIndex        =   136
         Top             =   3240
         Width           =   1755
      End
      Begin VB.Label Label4 
         Caption         =   "Condensing pressure"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   12
         Left            =   180
         TabIndex        =   135
         Top             =   4245
         Width           =   1755
      End
      Begin VB.Label Label3 
         Caption         =   "C_Factor"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   180
         TabIndex        =   134
         Top             =   3020
         Width           =   1455
      End
   End
   Begin VB.Frame Frame_Summary_ACT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Summary actual"
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
      Height          =   3975
      Left            =   120
      TabIndex        =   149
      Top             =   4500
      Visible         =   0   'False
      Width           =   12435
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   2940
         TabIndex        =   209
         Text            =   "Text1"
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   1
         Left            =   2940
         TabIndex        =   208
         Text            =   "Text1"
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   2
         Left            =   2940
         TabIndex        =   207
         Text            =   "Text1"
         Top             =   2100
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   3
         Left            =   2940
         TabIndex        =   206
         Text            =   "Text1"
         Top             =   2460
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   4
         Left            =   2940
         TabIndex        =   205
         Text            =   "Text1"
         Top             =   2820
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   5
         Left            =   2940
         TabIndex        =   204
         Text            =   "Text1"
         Top             =   3180
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   6
         Left            =   2940
         TabIndex        =   203
         Text            =   "Text1"
         Top             =   3540
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   202
         Text            =   "Text1"
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   201
         Text            =   "Text1"
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   200
         Text            =   "Text1"
         Top             =   2100
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   199
         Text            =   "Text1"
         Top             =   2460
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   198
         Text            =   "Text1"
         Top             =   2820
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   5
         Left            =   3960
         TabIndex        =   197
         Text            =   "Text1"
         Top             =   3180
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   6
         Left            =   3960
         TabIndex        =   196
         Text            =   "Text1"
         Top             =   3540
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   7
         Left            =   6180
         TabIndex        =   195
         Text            =   "Text1"
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   8
         Left            =   6180
         TabIndex        =   194
         Text            =   "Text1"
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   9
         Left            =   6180
         TabIndex        =   193
         Text            =   "Text1"
         Top             =   2100
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   10
         Left            =   6180
         TabIndex        =   192
         Text            =   "Text1"
         Top             =   2460
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   11
         Left            =   6180
         TabIndex        =   191
         Text            =   "Text1"
         Top             =   2820
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   12
         Left            =   6180
         TabIndex        =   190
         Text            =   "Text1"
         Top             =   3180
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   13
         Left            =   6180
         TabIndex        =   189
         Text            =   "Text1"
         Top             =   3540
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   7
         Left            =   5160
         TabIndex        =   188
         Text            =   "Text1"
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   8
         Left            =   5160
         TabIndex        =   187
         Text            =   "Text1"
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   9
         Left            =   5160
         TabIndex        =   186
         Text            =   "Text1"
         Top             =   2100
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   10
         Left            =   5220
         TabIndex        =   185
         Text            =   "Text1"
         Top             =   2460
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   11
         Left            =   5160
         TabIndex        =   184
         Text            =   "Text1"
         Top             =   2820
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   12
         Left            =   5160
         TabIndex        =   183
         Text            =   "Text1"
         Top             =   3180
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   13
         Left            =   5160
         TabIndex        =   182
         Text            =   "Text1"
         Top             =   3540
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   14
         Left            =   8460
         TabIndex        =   181
         Text            =   "Text1"
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   15
         Left            =   8460
         TabIndex        =   180
         Text            =   "Text1"
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   16
         Left            =   8460
         TabIndex        =   179
         Text            =   "Text1"
         Top             =   2100
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   17
         Left            =   8460
         TabIndex        =   178
         Text            =   "Text1"
         Top             =   2460
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   18
         Left            =   8460
         TabIndex        =   177
         Text            =   "Text1"
         Top             =   2820
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   19
         Left            =   8460
         TabIndex        =   176
         Text            =   "Text1"
         Top             =   3180
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   20
         Left            =   8460
         TabIndex        =   175
         Text            =   "Text1"
         Top             =   3540
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   14
         Left            =   7440
         TabIndex        =   174
         Text            =   "Text1"
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   15
         Left            =   7440
         TabIndex        =   173
         Text            =   "Text1"
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   16
         Left            =   7440
         TabIndex        =   172
         Text            =   "Text1"
         Top             =   2100
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   17
         Left            =   7440
         TabIndex        =   171
         Text            =   "Text1"
         Top             =   2460
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   18
         Left            =   7440
         TabIndex        =   170
         Text            =   "Text1"
         Top             =   2820
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   19
         Left            =   7440
         TabIndex        =   169
         Text            =   "Text1"
         Top             =   3180
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   20
         Left            =   7440
         TabIndex        =   168
         Text            =   "Text1"
         Top             =   3540
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   21
         Left            =   10740
         TabIndex        =   167
         Text            =   "Text1"
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   22
         Left            =   10740
         TabIndex        =   166
         Text            =   "Text1"
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   23
         Left            =   10740
         TabIndex        =   165
         Text            =   "Text1"
         Top             =   2100
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   24
         Left            =   10740
         TabIndex        =   164
         Text            =   "Text1"
         Top             =   2460
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   25
         Left            =   10740
         TabIndex        =   163
         Text            =   "Text1"
         Top             =   2820
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   26
         Left            =   10740
         TabIndex        =   162
         Text            =   "Text1"
         Top             =   3180
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   27
         Left            =   10740
         TabIndex        =   161
         Text            =   "Text1"
         Top             =   3540
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   21
         Left            =   9720
         TabIndex        =   160
         Text            =   "Text1"
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   22
         Left            =   9720
         TabIndex        =   159
         Text            =   "Text1"
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   23
         Left            =   9720
         TabIndex        =   158
         Text            =   "Text1"
         Top             =   2100
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   24
         Left            =   9720
         TabIndex        =   157
         Text            =   "Text1"
         Top             =   2460
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   25
         Left            =   9720
         TabIndex        =   156
         Text            =   "Text1"
         Top             =   2820
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   26
         Left            =   9720
         TabIndex        =   155
         Text            =   "Text1"
         Top             =   3180
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T_ACT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   27
         Left            =   9720
         TabIndex        =   154
         Text            =   "Text1"
         Top             =   3540
         Width           =   915
      End
      Begin VB.TextBox Text_ACT 
         Height          =   255
         Index           =   0
         Left            =   2940
         TabIndex        =   153
         Top             =   660
         Width           =   1875
      End
      Begin VB.TextBox Text_ACT 
         Height          =   255
         Index           =   1
         Left            =   5220
         TabIndex        =   152
         Top             =   660
         Width           =   1875
      End
      Begin VB.TextBox Text_ACT 
         Height          =   255
         Index           =   2
         Left            =   7500
         TabIndex        =   151
         Top             =   660
         Width           =   1875
      End
      Begin VB.TextBox Text_ACT 
         Height          =   255
         Index           =   3
         Left            =   9780
         TabIndex        =   150
         Top             =   660
         Width           =   1875
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Unit:"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   53
         Left            =   10380
         TabIndex        =   236
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tower:"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   52
         Left            =   5700
         TabIndex        =   235
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Unit-plant:"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   51
         Left            =   8160
         TabIndex        =   234
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Plant:"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   50
         Left            =   3540
         TabIndex        =   233
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total fluid flow"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   49
         Left            =   300
         TabIndex        =   232
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "units"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   48
         Left            =   2040
         TabIndex        =   231
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "TUBE-SIDE"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   47
         Left            =   3000
         TabIndex        =   230
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "SHELL-SIDE"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   46
         Left            =   3960
         TabIndex        =   229
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Vapor"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   45
         Left            =   300
         TabIndex        =   228
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Liquid"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   44
         Left            =   300
         TabIndex        =   227
         Top             =   2160
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Water"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   43
         Left            =   300
         TabIndex        =   226
         Top             =   2520
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Non-condensable"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   42
         Left            =   300
         TabIndex        =   225
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Duty"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   41
         Left            =   300
         TabIndex        =   224
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "kg/h"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   40
         Left            =   2040
         TabIndex        =   223
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "kg/h"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   39
         Left            =   2040
         TabIndex        =   222
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "kg/h"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   38
         Left            =   2040
         TabIndex        =   221
         Top             =   2160
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "kg/h"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   37
         Left            =   2040
         TabIndex        =   220
         Top             =   2520
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "kg/h"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   36
         Left            =   2040
         TabIndex        =   219
         Top             =   2880
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "KW"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   35
         Left            =   2040
         TabIndex        =   218
         Top             =   3240
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "kPa"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   34
         Left            =   2040
         TabIndex        =   217
         Top             =   3600
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pressure drop"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   33
         Left            =   300
         TabIndex        =   216
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "SHELL-SIDE"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   32
         Left            =   6180
         TabIndex        =   215
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "TUBE-SIDE"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   31
         Left            =   5220
         TabIndex        =   214
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "SHELL-SIDE"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   30
         Left            =   8460
         TabIndex        =   213
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "TUBE-SIDE"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   29
         Left            =   7500
         TabIndex        =   212
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "SHELL-SIDE"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   28
         Left            =   10740
         TabIndex        =   211
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "TUBE-SIDE"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   27
         Left            =   9780
         TabIndex        =   210
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   3900
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_Unit_1"
      Top             =   10020
      Width           =   1755
   End
   Begin VB.Frame Frame_Summary_DES 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Summary design"
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
      Height          =   3975
      Left            =   120
      TabIndex        =   44
      Top             =   600
      Visible         =   0   'False
      Width           =   12435
      Begin VB.TextBox Text1 
         Height          =   255
         Index           =   3
         Left            =   9780
         TabIndex        =   132
         Top             =   660
         Width           =   1875
      End
      Begin VB.TextBox Text1 
         Height          =   255
         Index           =   2
         Left            =   7500
         TabIndex        =   131
         Top             =   660
         Width           =   1875
      End
      Begin VB.TextBox Text1 
         Height          =   255
         Index           =   1
         Left            =   5220
         TabIndex        =   130
         Top             =   660
         Width           =   1875
      End
      Begin VB.TextBox Text1 
         Height          =   255
         Index           =   0
         Left            =   2940
         TabIndex        =   129
         Top             =   660
         Width           =   1875
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   27
         Left            =   9720
         TabIndex        =   126
         Text            =   "Text1"
         Top             =   3540
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   26
         Left            =   9720
         TabIndex        =   125
         Text            =   "Text1"
         Top             =   3180
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   25
         Left            =   9720
         TabIndex        =   124
         Text            =   "Text1"
         Top             =   2820
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   24
         Left            =   9720
         TabIndex        =   123
         Text            =   "Text1"
         Top             =   2460
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   23
         Left            =   9720
         TabIndex        =   122
         Text            =   "Text1"
         Top             =   2100
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   22
         Left            =   9720
         TabIndex        =   121
         Text            =   "Text1"
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   21
         Left            =   9720
         TabIndex        =   120
         Text            =   "Text1"
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   27
         Left            =   10740
         TabIndex        =   119
         Text            =   "Text1"
         Top             =   3540
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   26
         Left            =   10740
         TabIndex        =   118
         Text            =   "Text1"
         Top             =   3180
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   25
         Left            =   10740
         TabIndex        =   117
         Text            =   "Text1"
         Top             =   2820
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   24
         Left            =   10740
         TabIndex        =   116
         Text            =   "Text1"
         Top             =   2460
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   23
         Left            =   10740
         TabIndex        =   115
         Text            =   "Text1"
         Top             =   2100
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   22
         Left            =   10740
         TabIndex        =   114
         Text            =   "Text1"
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   21
         Left            =   10740
         TabIndex        =   113
         Text            =   "Text1"
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   20
         Left            =   7440
         TabIndex        =   110
         Text            =   "Text1"
         Top             =   3540
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   19
         Left            =   7440
         TabIndex        =   109
         Text            =   "Text1"
         Top             =   3180
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   18
         Left            =   7440
         TabIndex        =   108
         Text            =   "Text1"
         Top             =   2820
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   17
         Left            =   7440
         TabIndex        =   107
         Text            =   "Text1"
         Top             =   2460
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   16
         Left            =   7440
         TabIndex        =   106
         Text            =   "Text1"
         Top             =   2100
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   15
         Left            =   7440
         TabIndex        =   105
         Text            =   "Text1"
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   14
         Left            =   7440
         TabIndex        =   104
         Text            =   "Text1"
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   20
         Left            =   8460
         TabIndex        =   103
         Text            =   "Text1"
         Top             =   3540
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   19
         Left            =   8460
         TabIndex        =   102
         Text            =   "Text1"
         Top             =   3180
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   18
         Left            =   8460
         TabIndex        =   101
         Text            =   "Text1"
         Top             =   2820
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   17
         Left            =   8460
         TabIndex        =   100
         Text            =   "Text1"
         Top             =   2460
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   16
         Left            =   8460
         TabIndex        =   99
         Text            =   "Text1"
         Top             =   2100
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   15
         Left            =   8460
         TabIndex        =   98
         Text            =   "Text1"
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   14
         Left            =   8460
         TabIndex        =   97
         Text            =   "Text1"
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   13
         Left            =   5160
         TabIndex        =   94
         Text            =   "Text1"
         Top             =   3540
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   12
         Left            =   5160
         TabIndex        =   93
         Text            =   "Text1"
         Top             =   3180
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   11
         Left            =   5160
         TabIndex        =   92
         Text            =   "Text1"
         Top             =   2820
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   10
         Left            =   5160
         TabIndex        =   91
         Text            =   "Text1"
         Top             =   2460
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   9
         Left            =   5160
         TabIndex        =   90
         Text            =   "Text1"
         Top             =   2100
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   8
         Left            =   5160
         TabIndex        =   89
         Text            =   "Text1"
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   7
         Left            =   5160
         TabIndex        =   88
         Text            =   "Text1"
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   13
         Left            =   6180
         TabIndex        =   87
         Text            =   "Text1"
         Top             =   3540
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   12
         Left            =   6180
         TabIndex        =   86
         Text            =   "Text1"
         Top             =   3180
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   11
         Left            =   6180
         TabIndex        =   85
         Text            =   "Text1"
         Top             =   2820
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   10
         Left            =   6180
         TabIndex        =   84
         Text            =   "Text1"
         Top             =   2460
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   9
         Left            =   6180
         TabIndex        =   83
         Text            =   "Text1"
         Top             =   2100
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   8
         Left            =   6180
         TabIndex        =   82
         Text            =   "Text1"
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   7
         Left            =   6180
         TabIndex        =   81
         Text            =   "Text1"
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   6
         Left            =   3960
         TabIndex        =   80
         Text            =   "Text1"
         Top             =   3540
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   5
         Left            =   3960
         TabIndex        =   79
         Text            =   "Text1"
         Top             =   3180
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   78
         Text            =   "Text1"
         Top             =   2820
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   77
         Text            =   "Text1"
         Top             =   2460
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   76
         Text            =   "Text1"
         Top             =   2100
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   75
         Text            =   "Text1"
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox TOTAL_S 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   74
         Text            =   "Text1"
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   6
         Left            =   2940
         TabIndex        =   71
         Text            =   "Text1"
         Top             =   3540
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   5
         Left            =   2940
         TabIndex        =   70
         Text            =   "Text1"
         Top             =   3180
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   4
         Left            =   2940
         TabIndex        =   69
         Text            =   "Text1"
         Top             =   2820
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   3
         Left            =   2940
         TabIndex        =   68
         Text            =   "Text1"
         Top             =   2460
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   2
         Left            =   2940
         TabIndex        =   67
         Text            =   "Text1"
         Top             =   2100
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   1
         Left            =   2940
         TabIndex        =   66
         Text            =   "Text1"
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox TOTAL_T 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   2940
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   1380
         Width           =   915
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "TUBE-SIDE"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   26
         Left            =   9780
         TabIndex        =   128
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "SHELL-SIDE"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   25
         Left            =   10740
         TabIndex        =   127
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "TUBE-SIDE"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   24
         Left            =   7500
         TabIndex        =   112
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "SHELL-SIDE"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   23
         Left            =   8460
         TabIndex        =   111
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "TUBE-SIDE"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   22
         Left            =   5220
         TabIndex        =   96
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "SHELL-SIDE"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   21
         Left            =   6180
         TabIndex        =   95
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pressure drop"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   20
         Left            =   300
         TabIndex        =   72
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "kPa"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   19
         Left            =   2040
         TabIndex        =   64
         Top             =   3600
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "KW"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   18
         Left            =   2040
         TabIndex        =   63
         Top             =   3240
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "kg/h"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   17
         Left            =   2040
         TabIndex        =   62
         Top             =   2880
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "kg/h"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   16
         Left            =   2040
         TabIndex        =   61
         Top             =   2520
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "kg/h"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   15
         Left            =   2040
         TabIndex        =   60
         Top             =   2160
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "kg/h"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   14
         Left            =   2040
         TabIndex        =   59
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "kg/h"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   13
         Left            =   2040
         TabIndex        =   58
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Duty"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   12
         Left            =   300
         TabIndex        =   57
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Non-condensable"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   11
         Left            =   300
         TabIndex        =   56
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Water"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   10
         Left            =   300
         TabIndex        =   55
         Top             =   2520
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Liquid"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   9
         Left            =   300
         TabIndex        =   54
         Top             =   2160
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Vapor"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   8
         Left            =   300
         TabIndex        =   53
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "SHELL-SIDE"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   7
         Left            =   3960
         TabIndex        =   52
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "TUBE-SIDE"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   6
         Left            =   3000
         TabIndex        =   51
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "units"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   5
         Left            =   2040
         TabIndex        =   50
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total fluid flow"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   4
         Left            =   300
         TabIndex        =   49
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Plant:"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   3
         Left            =   3540
         TabIndex        =   48
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Unit-plant:"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   2
         Left            =   8160
         TabIndex        =   47
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tower:"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   5700
         TabIndex        =   46
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Unit:"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   0
         Left            =   10380
         TabIndex        =   45
         Top             =   420
         Width           =   675
      End
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   2100
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_test"
      Top             =   9960
      Width           =   1755
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_Plant"
      Top             =   9900
      Width           =   1815
   End
   Begin VB.ComboBox Combo_PLANT 
      ForeColor       =   &H00000080&
      Height          =   330
      ItemData        =   "Chart_ALL_UNITS.frx":0000
      Left            =   13260
      List            =   "Chart_ALL_UNITS.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "Combo_PLANT"
      ToolTipText     =   "Select different UNIT from the list."
      Top             =   120
      Width           =   1395
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
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
      RecordSource    =   "Query_Plant_UNIT"
      Top             =   9600
      Width           =   1755
   End
   Begin VB.ComboBox Combo_PLANT_UNIT 
      ForeColor       =   &H00000080&
      Height          =   330
      ItemData        =   "Chart_ALL_UNITS.frx":0004
      Left            =   13260
      List            =   "Chart_ALL_UNITS.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   3
      Text            =   "Combo_PLANT_UNIT"
      ToolTipText     =   "Select different UNIT from the list."
      Top             =   840
      Width           =   1395
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
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
      RecordSource    =   "Query_TOWER"
      Top             =   9600
      Width           =   1755
   End
   Begin VB.ComboBox Combo_TOWER 
      ForeColor       =   &H00000080&
      Height          =   330
      ItemData        =   "Chart_ALL_UNITS.frx":0008
      Left            =   13260
      List            =   "Chart_ALL_UNITS.frx":000A
      Sorted          =   -1  'True
      TabIndex        =   2
      Text            =   "Combo_TOWER"
      ToolTipText     =   "Select different UNIT from the list."
      Top             =   480
      Width           =   1395
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      RecordSource    =   "Query_Unit"
      Top             =   9600
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Top             =   9600
      Width           =   1815
   End
   Begin VB.ComboBox Combo_UNIT 
      ForeColor       =   &H00000080&
      Height          =   330
      ItemData        =   "Chart_ALL_UNITS.frx":000C
      Left            =   13260
      List            =   "Chart_ALL_UNITS.frx":000E
      Sorted          =   -1  'True
      TabIndex        =   4
      Text            =   "Combo_UNIT"
      ToolTipText     =   "Select different UNIT from the list."
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scale span"
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Left            =   13035
      TabIndex        =   33
      Top             =   6600
      Width           =   1575
      Begin VB.TextBox Text_MAX 
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   60
         TabIndex        =   35
         Text            =   "Max"
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox Check_MAX 
         Caption         =   "Click to change:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   60
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text_MIN 
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   60
         TabIndex        =   34
         Text            =   "Min"
         Top             =   1560
         Width           =   1095
      End
      Begin MSForms.SpinButton Spin_MAX 
         Height          =   615
         Left            =   1140
         TabIndex        =   25
         Top             =   540
         Width           =   375
         Size            =   "661;1085"
         Max             =   10000
      End
      Begin VB.Label Label1 
         Caption         =   "Max"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   600
         TabIndex        =   37
         Top             =   600
         Width           =   435
      End
      Begin MSForms.SpinButton Spin_MIN 
         Height          =   615
         Left            =   1140
         TabIndex        =   26
         Top             =   1260
         Width           =   375
         Size            =   "661;1085"
         Max             =   10000
      End
      Begin VB.Label Label2 
         Caption         =   "Min"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   660
         TabIndex        =   36
         Top             =   1320
         Width           =   315
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Trends charts"
      ForeColor       =   &H00FF0000&
      Height          =   3135
      Left            =   13035
      TabIndex        =   29
      Top             =   3540
      Width           =   1575
      Begin VB.CommandButton Comm_view 
         Caption         =   "View new period"
         Height          =   375
         Left            =   60
         TabIndex        =   23
         Top             =   2340
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTP_Fine 
         DataField       =   "Date_Fine"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   60
         TabIndex        =   22
         Top             =   1980
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
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
         Format          =   58589185
         CurrentDate     =   37770
      End
      Begin MSComCtl2.DTPicker DTP_Inizio 
         DataField       =   "Date_Inizio"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   60
         TabIndex        =   21
         Top             =   1380
         Width           =   1335
         _ExtentX        =   2355
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
         CustomFormat    =   "01/01/01"
         Format          =   58589185
         CurrentDate     =   37770
         MinDate         =   36526
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Caption         =   "Select the desired period for the trend chart "
         ForeColor       =   &H00000080&
         Height          =   675
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label38 
         Caption         =   "End date:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   60
         TabIndex        =   31
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label37 
         Caption         =   "Start date:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   60
         TabIndex        =   30
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "2D LINE / 3D BAR View"
      Height          =   315
      Left            =   1980
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   60
      Width           =   2115
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      Style           =   1
      HotTracking     =   -1  'True
      MultiSelect     =   -1  'True
      Placement       =   1
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Object.ToolTipText     =   "Close this form"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Object.ToolTipText     =   "Print this graph"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Update"
            Object.ToolTipText     =   "Update last changes"
            ImageVarType    =   2
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
   Begin MSChart20Lib.MSChart MSChart_ALL_UNITS 
      Height          =   6555
      Left            =   120
      OleObjectBlob   =   "Chart_ALL_UNITS.frx":0010
      TabIndex        =   0
      Top             =   1200
      Width           =   11415
   End
   Begin MSForms.ToggleButton Toggle_PARMETER 
      Height          =   435
      Left            =   10500
      TabIndex        =   5
      Top             =   60
      Width           =   2415
      BackColor       =   -2147483638
      ForeColor       =   192
      DisplayStyle    =   6
      Size            =   "4260;767"
      Value           =   "0"
      Caption         =   "Select parameter(s)"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ToggleButton Toggle_summary 
      Height          =   435
      Left            =   13020
      TabIndex        =   73
      Top             =   2100
      Width           =   1635
      BackColor       =   -2147483633
      ForeColor       =   128
      DisplayStyle    =   6
      Size            =   "2884;767"
      Value           =   "0"
      Caption         =   "See SUMMARY"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ToggleButton Toggle_UNIT_DATE 
      Height          =   255
      Left            =   13020
      TabIndex        =   43
      ToolTipText     =   "Choosing by date or by unit in the X-axis"
      Top             =   1620
      Width           =   255
      BackColor       =   14737632
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "450;450"
      Value           =   "0"
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ToggleButton Toggle_TOWER 
      Height          =   255
      Left            =   13020
      TabIndex        =   42
      ToolTipText     =   "Check to see the trend by tower"
      Top             =   540
      Width           =   255
      BackColor       =   14737632
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "450;450"
      Value           =   "0"
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ToggleButton Toggle_PLANT_UNIT 
      Height          =   255
      Left            =   13020
      TabIndex        =   41
      ToolTipText     =   "Check to see the trend by plant unit"
      Top             =   900
      Width           =   255
      BackColor       =   14737632
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "450;450"
      Value           =   "0"
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ToggleButton Toggle_UNIT 
      Height          =   255
      Left            =   13020
      TabIndex        =   40
      ToolTipText     =   "Check to see the trend by unit"
      Top             =   1260
      Width           =   255
      BackColor       =   14737632
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "450;450"
      Value           =   "0"
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ToggleButton Toggle_PLANT 
      Height          =   255
      Left            =   13020
      TabIndex        =   39
      ToolTipText     =   "Check to see the trend by plant"
      Top             =   180
      Width           =   255
      BackColor       =   14737632
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "450;450"
      Value           =   "0"
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Lab_Unit 
      Alignment       =   1  'Right Justify
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
      Left            =   13380
      TabIndex        =   38
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "Chart_ALL_UNITS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public www
Private Sub Combo_Plant_LostFocus()
    If Toggle_PLANT = True Then
        Text1(0) = Combo_PLANT
        If Toggle_summary = True Then
            Call Summary_TABLE_DES
        ElseIf Toggle_summary = False Then
            Call chart
        End If
    End If

End Sub
Private Sub Combo_TOWER_GotFocus()
    On Error Resume Next
    Dim Rs3 As Recordset
    Data3.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
    Data3.RecordSource = "Select * From [Query_TOWER]"
    Data3.Refresh
    Set Rs3 = Data3.Recordset
    Combo_TOWER.Clear
    If Rs3.RecordCount > 0 Then
       Do Until Rs3.EOF
            PPP1 = Combo_PLANT
            PPP2 = Data3.Recordset.Plant
            TTT1 = Data3.Recordset.COOL_TOWER
            If PPP1 = PPP2 Then
                Combo_TOWER.AddItem TTT1
            End If
            Rs3.MoveNext
        Loop
    End If
End Sub
Private Sub Combo_TOWER_LostFocus()
    If Toggle_TOWER = True Then
        Text1(1) = Combo_TOWER
        If Toggle_summary = True Then
            Call Summary_TABLE_DES
        ElseIf Toggle_summary = False Then
            Call chart
        End If
    End If
End Sub
Private Sub Combo_Plant_UNIT_GotFocus()
    On Error Resume Next
    Dim Rs4 As Recordset
    Data4.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
    Data4.RecordSource = "Select * From [Query_PLANT_UNIT]"
    Data4.Refresh
    Set Rs4 = Data4.Recordset
    Combo_PLANT_UNIT.Clear
    If Rs4.RecordCount > 0 Then
       Do Until Rs4.EOF
            PPP1 = Combo_PLANT
            TTT1 = Combo_TOWER
            PPP2 = Data4.Recordset.Plant
            TTT2 = Data4.Recordset.COOL_TOWER
            UUU1 = Data4.Recordset.PLANT_UNIT
            If PPP1 = PPP2 And TTT1 = TTT2 Then
                Combo_PLANT_UNIT.AddItem UUU1
            End If
            Rs4.MoveNext
        Loop
    End If
End Sub
Private Sub Combo_PLANT_UNIT_LostFocus()
    If Toggle_PLANT_UNIT = True Then
        Text1(2) = Combo_PLANT_UNIT
        If Toggle_summary = True Then
            Call Summary_TABLE_DES
        ElseIf Toggle_summary = False Then
            Call chart
        End If
    End If
End Sub
Private Sub Combo_UNIT_GotFocus()
    On Error Resume Next
    Dim Rs7 As Recordset
    Data7.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
    Data7.RecordSource = "Select * From [Query_UNIT_1]"
    Data7.Refresh
    Set Rs7 = Data7.Recordset
    Combo_UNIT.Clear
    If Rs7.RecordCount > 0 Then
       Do Until Rs7.EOF
            PPP1 = Combo_PLANT
            TTT1 = Combo_TOWER
            SSS1 = Combo_PLANT_UNIT
            PPP2 = Data7.Recordset.Plant
            TTT2 = Data7.Recordset.COOL_TOWER
            SSS2 = Data7.Recordset.PLANT_UNIT
            UUU1 = Data7.Recordset.Unit_name
            If PPP1 = PPP2 And TTT1 = TTT2 And SSS1 = SSS2 Then
                Combo_UNIT.AddItem UUU1
            End If
            Rs7.MoveNext
        Loop
    End If
End Sub
Private Sub Combo_UNIT_LostFocus()
    If Toggle_UNIT = True Then
        Text1(3) = Combo_UNIT
        If Toggle_summary = True Then
            Call Summary_TABLE_DES
        ElseIf Toggle_summary = False Then
            Call chart
        End If
    End If
End Sub


Private Sub Spin_MAX_Change()
    MM2 = Spin_MAX.Value
    Text_MAX.Text = Spin_MAX
    Call chart
End Sub
Private Sub Spin_MIN_Change()
    Min = Spin_MIN.Value
    Text_MIN.Text = Spin_MIN
    Call chart
End Sub
Private Sub Check_UNIT_Click()
If Check_UNIT = Unchecked Then
    Lab_Default.Visible = True
    Lab_NEW.Visible = False
Else
    Lab_Default.Visible = False
    Lab_NEW.Visible = True
End If
    Call chart
End Sub
Private Sub Check1_Click()
    Call chart
End Sub
Private Sub Comm_view_Click()
    Call chart
End Sub
Private Sub Form_Load()
On Error Resume Next
   Width = frmMain.Width * 0.98 ' Imposta la larghezza del form.
   Height = frmMain.Height * 0.89     ' Imposta l'altezza del form.
   Left = 50 '(frmMain.Width - Width) / 2 ' Centra il form orizzontalmente.
   Top = 0 '(frmMain.Height - Height) / 2 ' Centra il form verticalmente.

'    Dim Rs2 As Recordset
'    Data2.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
'    Data2.RecordSource = "Select * From [Query_Unit]"
'    Data2.Refresh
'    Set Rs2 = Data2.Recordset
'    If Rs2.RecordCount > 0 Then
'       Do Until Rs2.EOF
'            UUU2 = Data2.Recordset.Unit_name
'                Combo_UNIT.AddItem UUU2
'            Rs2.MoveNext
'    Loop
'    Else
'       MsgBox "No Units found"
'    End If

'    Dim Rs3 As Recordset
'    Data3.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
'    Data3.RecordSource = "Select * From [Query_TOWER]"
'    Data3.Refresh
'    Set Rs3 = Data3.Recordset
'    If Rs3.RecordCount > 0 Then
'       Do Until Rs3.EOF
'            UUU3 = Data3.Recordset.COOL_TOWER
'                Combo_TOWER.AddItem UUU3
'            Rs3.MoveNext
'    Loop
'    Else
'       MsgBox "No COOLING TOWER found"
'    End If
    
'    Dim Rs4 As Recordset
'    Data4.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
'    Data4.RecordSource = "Select * From [Query_Plant_UNIT]"
'    Data4.Refresh
'    Set Rs4 = Data4.Recordset
'    If Rs4.RecordCount > 0 Then
'       Do Until Rs4.EOF
'            UUU4 = Data4.Recordset.Plant_UNIT
'                Combo_PLANT_UNIT.AddItem UUU4
'            Rs4.MoveNext
'    Loop
'    Else
'       MsgBox "No COOLING TOWER found"
'    End If
    
    Dim Rs5 As Recordset
    Data5.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
    Data5.RecordSource = "Select * From [Query_Plant]"
    Data5.Refresh
    Set Rs5 = Data5.Recordset
    If Rs5.RecordCount > 0 Then
       Do Until Rs5.EOF
            UUU5 = Data5.Recordset.Plant
                Combo_PLANT.AddItem UUU5
            Rs5.MoveNext
    Loop
    Else
       MsgBox "No COOLING TOWER found"
    End If

    FFX = 0
    SKX = 0
    LMTDX = 0
    TTDX = 0
    REX = 0
    VELX = 0
    PDX = 0
    CPX = 0
    CTX = 0
    TINX = 0
    TOUTX = 0
    WFX = 0
    SFX = 0
    DUTYX = 0
    CFACX = 0

    Toggle_UNIT_DATE = False
    Lab_Unit.Caption = "by Date"
    Frame_Summary.Visible = False
Call chart
End Sub
Private Sub TabStrip1_Click()
On Error Resume Next
    If TabStrip1.SelectedItem = "Print" Then
        Me.PrintForm
    ElseIf TabStrip1.SelectedItem = "Close" Then
        Unload Me
    ElseIf TabStrip1.SelectedItem = "Update" Then
        Data1.UpdateRecord
        Data1.Recordset.Bookmark = Data1.Recordset.LastModified
        Call chart
    End If
End Sub
Private Sub Data1_Error(DataErr As Integer, Response As Integer)
  MsgBox "Intercettato errore dei dati:" & Error$(DataErr)
  Response = 0  'Ignora l'errore
End Sub
Private Sub Data1_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  Data1.Caption = "Record: " & (Data1.Recordset.AbsolutePosition + 1)
End Sub
Private Sub Data6_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  Data6.Caption = "Record: " & (Data1.Recordset.AbsolutePosition + 1)
End Sub
Private Sub chart()
On Error Resume Next
Dim CX(10000, 16), SY(10000, 16), LE(16), ET(10000), c, r, ETX(10000)
Dim Rs1 As Recordset
Dim Rs6 As Recordset
Dim uni(100)
Dim VIEW(15)

Data1.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
Data1.RecordSource = "Select * From [QUERY_TEST]"
Data1.Refresh
Set Rs1 = Data1.Recordset
    
qqq = 0
www = 0
XXX = 0
INIZIO = Data1.Recordset.Date_Inizio
FINE = Data1.Recordset.Date_Fine
DTP_Inizio.Value = INIZIO
DTP_Fine.Value = FINE
If Rs1.RecordCount > 0 Then
    Do Until Rs1.EOF
        PLANT1 = Data1.Recordset.Plant
        TOWER1 = Data1.Recordset.COOL_TOWER
        PLANT_U1 = Data1.Recordset.PLANT_UNIT
        UNIT1 = Data1.Recordset.Unit_name
        PLANT2 = Combo_PLANT
        TOWER2 = Combo_TOWER
        PLANT_U2 = Combo_PLANT_UNIT
        UNIT2 = Combo_UNIT
        If Toggle_PLANT = True And Toggle_TOWER = False And Toggle_PLANT_UNIT = False And Toggle_UNIT = False Then
            If PLANT1 = PLANT2 Then
                qqq = qqq + 1
                uni(qqq) = Data1.Recordset.Unit_name
                Unit_Y = Data1.Recordset.Unit_name
                If qqq > 1 Then
                    For i = 2 To qqq
                        If Unit_Y = uni(i - 1) Then
                            Exit For
                        Else
                            test = 0
                        End If
                    Next i
                End If
                If i = qqq + 1 Or qqq = 1 Then
                Data6.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
                Data6.RecordSource = "Select * From [query_test]"
                Data6.Refresh
                Set Rs6 = Data6.Recordset
                Do Until Rs6.EOF
                    UNIT_X = Data6.Recordset.Unit_name
                    If UNIT_X = Unit_Y Then
                        www = www + 1
                        ETX(www) = Data1.Recordset.date_test
                        If ETX(www) >= INIZIO And ETX(www) <= FINE Then
                            XXX = XXX + 1
                            SY(XXX, 1) = Data6.Recordset.TUBES_FF
                            SY(XXX, 2) = Data6.Recordset.SKIN_TEMP
                            SY(XXX, 3) = Data6.Recordset.MTDc
                            SY(XXX, 4) = Data6.Recordset.TTD
                            SY(XXX, 5) = Data6.Recordset.TUBES_REYNOLDS / 1000
                            SY(XXX, 6) = Data6.Recordset.TUBES_VEL
                            SY(XXX, 7) = Data6.Recordset.TUBES_PRESS_DROP
                            SY(XXX, 8) = Data6.Recordset.Press_COND * 100
                            SY(XXX, 9) = Data6.Recordset.Temp_COND
                            SY(XXX, 10) = Data6.Recordset.TUBES_TEMP_IN
                            SY(XXX, 11) = Data6.Recordset.TUBES_TEMP_OUT
                            SY(XXX, 12) = Data6.Recordset.TUBES_FLOW / 1000
                            SY(XXX, 13) = Data6.Recordset.SHELL_FLOW / 1000
                            SY(XXX, 14) = Data6.Recordset.TUBES_DUTY
                            SY(XXX, 15) = Data6.Recordset.C_Factor
                            C_DES = Data6.Recordset.CHECK_DESIGN
                            If C_DES = True Then
                                ET(XXX) = "Des " & Data6.Recordset.Unit_name
                            ElseIf Toggle_UNIT_DATE = False Then
                                ET(XXX) = Data6.Recordset.date_test
                            ElseIf Toggle_UNIT_DATE = True Then
                                ET(XXX) = Data6.Recordset.Unit_name
                            End If
                        End If
                    End If
                rec = Data6.Recordset.AbsolutePosition + 1
                Rs6.MoveNext
                Loop
                End If
            End If
        End If
        If Toggle_PLANT = True And Toggle_TOWER = True And Toggle_PLANT_UNIT = False And Toggle_UNIT = False Then
            If PLANT1 = PLANT2 And TOWER1 = TOWER2 Then
                qqq = qqq + 1
                uni(qqq) = Data1.Recordset.Unit_name
                Unit_Y = Data1.Recordset.Unit_name
                If qqq > 1 Then
                    For i = 2 To qqq
                        If Unit_Y = uni(i - 1) Then
                            Exit For
                        Else
                            test = 0
                        End If
                    Next i
                End If
                If i = qqq + 1 Or qqq = 1 Then
                Data6.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
                Data6.RecordSource = "Select * From [query_test]"
                Data6.Refresh
                Set Rs6 = Data6.Recordset
                Do Until Rs6.EOF
                    UNIT_X = Data6.Recordset.Unit_name
                    If UNIT_X = Unit_Y Then
                        
                        www = www + 1
                        ETX(www) = Data1.Recordset.date_test
                        If ETX(www) >= INIZIO And ETX(www) <= FINE Then
                            XXX = XXX + 1
                            SY(XXX, 1) = Data6.Recordset.TUBES_FF
                            SY(XXX, 2) = Data6.Recordset.SKIN_TEMP
                            SY(XXX, 3) = Data6.Recordset.MTDc
                            SY(XXX, 4) = Data6.Recordset.TTD
                            SY(XXX, 5) = Data6.Recordset.TUBES_REYNOLDS / 1000
                            SY(XXX, 6) = Data6.Recordset.TUBES_VEL
                            SY(XXX, 7) = Data6.Recordset.TUBES_PRESS_DROP
                            SY(XXX, 8) = Data6.Recordset.Press_COND * 100
                            SY(XXX, 9) = Data6.Recordset.Temp_COND
                            SY(XXX, 10) = Data6.Recordset.TUBES_TEMP_IN
                            SY(XXX, 11) = Data6.Recordset.TUBES_TEMP_OUT
                            SY(XXX, 12) = Data6.Recordset.TUBES_FLOW / 1000
                            SY(XXX, 13) = Data6.Recordset.SHELL_FLOW / 1000
                            SY(XXX, 14) = Data6.Recordset.TUBES_DUTY
                            SY(XXX, 15) = Data6.Recordset.C_Factor
                            C_DES = Data6.Recordset.CHECK_DESIGN
                            If C_DES = True Then
                                ET(XXX) = "Des " & Data6.Recordset.Unit_name
                            ElseIf Toggle_UNIT_DATE = False Then
                                ET(XXX) = Data6.Recordset.date_test
                            ElseIf Toggle_UNIT_DATE = True Then
                                ET(XXX) = Data6.Recordset.Unit_name
                            End If
                        End If
                    End If
                rec = Data6.Recordset.AbsolutePosition + 1
                Rs6.MoveNext
                Loop
                End If
            End If
        End If
        If Toggle_PLANT = True And Toggle_TOWER = True And Toggle_PLANT_UNIT = True And Toggle_UNIT = False Then
            If PLANT1 = PLANT2 And TOWER1 = TOWER2 And PLANT_U1 = PLANT_U2 Then
                qqq = qqq + 1
                uni(qqq) = Data1.Recordset.Unit_name
                Unit_Y = Data1.Recordset.Unit_name
                If qqq > 1 Then
                    For i = 2 To qqq
                        If Unit_Y = uni(i - 1) Then
                            Exit For
                        Else
                            test = 0
                        End If
                    Next i
                End If
                If i = qqq + 1 Or qqq = 1 Then
                Data6.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
                Data6.RecordSource = "Select * From [query_test]"
                Data6.Refresh
                Set Rs6 = Data6.Recordset
                Do Until Rs6.EOF
                    UNIT_X = Data6.Recordset.Unit_name
                    If UNIT_X = Unit_Y Then
                        
                        www = www + 1
                        ETX(www) = Data1.Recordset.date_test
                        If ETX(www) >= INIZIO And ETX(www) <= FINE Then
                            XXX = XXX + 1
                            SY(XXX, 1) = Data6.Recordset.TUBES_FF
                            SY(XXX, 2) = Data6.Recordset.SKIN_TEMP
                            SY(XXX, 3) = Data6.Recordset.MTDc
                            SY(XXX, 4) = Data6.Recordset.TTD
                            SY(XXX, 5) = Data6.Recordset.TUBES_REYNOLDS / 1000
                            SY(XXX, 6) = Data6.Recordset.TUBES_VEL
                            SY(XXX, 7) = Data6.Recordset.TUBES_PRESS_DROP
                            SY(XXX, 8) = Data6.Recordset.Press_COND * 100
                            SY(XXX, 9) = Data6.Recordset.Temp_COND
                            SY(XXX, 10) = Data6.Recordset.TUBES_TEMP_IN
                            SY(XXX, 11) = Data6.Recordset.TUBES_TEMP_OUT
                            SY(XXX, 12) = Data6.Recordset.TUBES_FLOW / 1000
                            SY(XXX, 13) = Data6.Recordset.SHELL_FLOW / 1000
                            SY(XXX, 14) = Data6.Recordset.TUBES_DUTY
                            SY(XXX, 15) = Data6.Recordset.C_Factor
                            C_DES = Data6.Recordset.CHECK_DESIGN
                            If C_DES = True Then
                                ET(XXX) = "Des " & Data6.Recordset.Unit_name
                            ElseIf Toggle_UNIT_DATE = False Then
                                ET(XXX) = Data6.Recordset.date_test
                            ElseIf Toggle_UNIT_DATE = True Then
                                ET(XXX) = Data6.Recordset.Unit_name
                            End If
                        End If
                    End If
                rec = Data6.Recordset.AbsolutePosition + 1
                Rs6.MoveNext
                Loop
                End If
            End If
        End If
        If Toggle_PLANT = True And Toggle_TOWER = True And Toggle_PLANT_UNIT = True And Toggle_UNIT = True Then
            If PLANT1 = PLANT2 And TOWER1 = TOWER2 And PLANT_U1 = PLANT_U2 And UNIT1 = UNIT2 Then
                qqq = qqq + 1
                uni(qqq) = Data1.Recordset.Unit_name
                Unit_Y = Data1.Recordset.Unit_name
                If qqq > 1 Then
                    For i = 2 To qqq
                        If Unit_Y = uni(i - 1) Then
                            Exit For
                        Else
                            test = 0
                        End If
                    Next i
                End If
                If i = qqq + 1 Or qqq = 1 Then
                Data6.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
                Data6.RecordSource = "Select * From [query_test]"
                Data6.Refresh
                Set Rs6 = Data6.Recordset
                Do Until Rs6.EOF
                    UNIT_X = Data6.Recordset.Unit_name
                    If UNIT_X = Unit_Y Then
                        
                        www = www + 1
                        ETX(www) = Data1.Recordset.date_test
                        If ETX(www) >= INIZIO And ETX(www) <= FINE Then
                            XXX = XXX + 1
                            SY(XXX, 1) = Data6.Recordset.TUBES_FF
                            SY(XXX, 2) = Data6.Recordset.SKIN_TEMP
                            SY(XXX, 3) = Data6.Recordset.MTDc
                            SY(XXX, 4) = Data6.Recordset.TTD
                            SY(XXX, 5) = Data6.Recordset.TUBES_REYNOLDS / 1000
                            SY(XXX, 6) = Data6.Recordset.TUBES_VEL
                            SY(XXX, 7) = Data6.Recordset.TUBES_PRESS_DROP
                            SY(XXX, 8) = Data6.Recordset.Press_COND * 100
                            SY(XXX, 9) = Data6.Recordset.Temp_COND
                            SY(XXX, 10) = Data6.Recordset.TUBES_TEMP_IN
                            SY(XXX, 11) = Data6.Recordset.TUBES_TEMP_OUT
                            SY(XXX, 12) = Data6.Recordset.TUBES_FLOW / 1000
                            SY(XXX, 13) = Data6.Recordset.SHELL_FLOW / 1000
                            SY(XXX, 14) = Data6.Recordset.TUBES_DUTY
                            SY(XXX, 15) = Data6.Recordset.C_Factor
                            C_DES = Data6.Recordset.CHECK_DESIGN
                            If C_DES = True Then
                                ET(XXX) = "Des " & Data6.Recordset.Unit_name
                            ElseIf Toggle_UNIT_DATE = False Then
                                ET(XXX) = Data6.Recordset.date_test
                            ElseIf Toggle_UNIT_DATE = True Then
                                ET(XXX) = Data6.Recordset.Unit_name
                            End If
                        End If
                    End If
                rec = Data6.Recordset.AbsolutePosition + 1
                Rs6.MoveNext
                Loop
                End If
            End If
        End If

        Rs1.MoveNext
    Loop
End If
Rs1.MoveFirst

xWF = Check_T_FLOW
xTIN = Check_T_TEMP_IN
xTOUT = Check_T_TEMP_OUT
xFF = Check_WATER_FOUL
xSK = Check_SKIN
xLMTD = Check_LMTD
xTTD = Check_TTD
xRE = Check_T_RE
xVEL = Check_T_VEL
xPD = Check_T_PD
xCP = Check_CP
xCT = Check_CT
xSF = Check_S_FLOW
xDUTY = Check_DUTY
xCFAC = Check_CFACTOR

If xFF = 1 Then
    FFX = 1
End If
If xSK = 1 Then
    SKX = 1
End If
If xCFAC = 1 Then
    CFACX = 1
End If
If xLMTD = 1 Then
    LMTDX = 1
End If
If xTTD = 1 Then
    TTDX = 1
End If
If xRE = 1 Then
    REX = 1
End If
If xVEL = 1 Then
    VELX = 1
End If
If xPD = 1 Then
    PDX = 1
End If
If xCP = 1 Then
    CPX = 1
End If
If xCT = 1 Then
    CTX = 1
End If
If xTIN = 1 Then
    TINX = 1
End If
If xTOUT = 1 Then
    TOUTX = 1
End If
If xWF = 1 Then
    WFX = 1
End If
If xSF = 1 Then
    SFX = 1
End If
If xDUTY = 1 Then
    DUTYX = 1
 End If

MM2 = 0
MM1 = 0
If FFX = 1 Then
    ZZ = ZZ + 1
    For i = 1 To XXX
            MM2 = MM1
            MM1 = SY(i, 1)
            If MM1 > MM2 Then
                MM2 = MM1
                MM1 = MM2
            Else
                MM1 = MM2
            End If
    Next i
End If
If SKX = 1 Then
    ZZ = ZZ + 1
    For i = 1 To XXX
            MM2 = MM1
            MM1 = SY(i, 2)
            If MM1 > MM2 Then
                MM2 = MM1
                MM1 = MM2
            Else
                MM1 = MM2
            End If
    Next i
End If
If LMTDX = 1 Then
    ZZ = ZZ + 1
    For i = 1 To XXX
            MM2 = MM1
            MM1 = SY(i, 3)
            If MM1 > MM2 Then
                MM2 = MM1
                MM1 = MM2
            Else
                MM1 = MM2
            End If
    Next i
End If
If TTDX = 1 Then
    ZZ = ZZ + 1
    For i = 1 To XXX
            MM2 = MM1
            MM1 = SY(i, 4)
            If MM1 > MM2 Then
                MM2 = MM1
                MM1 = MM2
            Else
                MM1 = MM2
            End If
    Next i
End If
If REX = 1 Then
    ZZ = ZZ + 1
    For i = 1 To XXX
            MM2 = MM1
            MM1 = SY(i, 5)
            If MM1 > MM2 Then
                MM2 = MM1
                MM1 = MM2
            Else
                MM1 = MM2
            End If
    Next i
End If
If VELX = 1 Then
    ZZ = ZZ + 1
    For i = 1 To XXX
            MM2 = MM1
            MM1 = SY(i, 6)
            If MM1 > MM2 Then
                MM2 = MM1
                MM1 = MM2
            Else
                MM1 = MM2
            End If
    Next i
End If
If PDX = 1 Then
    ZZ = ZZ + 1
    For i = 1 To XXX
            MM2 = MM1
            MM1 = SY(i, 7)
            If MM1 > MM2 Then
                MM2 = MM1
                MM1 = MM2
            Else
                MM1 = MM2
            End If
    Next i
End If
If CPX = 1 Then
    ZZ = ZZ + 1
    For i = 1 To XXX
            MM2 = MM1
            MM1 = SY(i, 8)
            If MM1 > MM2 Then
                MM2 = MM1
                MM1 = MM2
            Else
                MM1 = MM2
            End If
    Next i
End If
If CTX = 1 Then
    ZZ = ZZ + 1
    For i = 1 To XXX
            MM2 = MM1
            MM1 = SY(i, 9)
            If MM1 > MM2 Then
                MM2 = MM1
                MM1 = MM2
            Else
                MM1 = MM2
            End If
    Next i
End If
If TINX = 1 Then
    ZZ = ZZ + 1
    For i = 1 To XXX
            MM2 = MM1
            MM1 = SY(i, 10)
            If MM1 > MM2 Then
                MM2 = MM1
                MM1 = MM2
            Else
                MM1 = MM2
            End If
    Next i
End If
If TOUTX = 1 Then
    ZZ = ZZ + 1
    For i = 1 To XXX
            MM2 = MM1
            MM1 = SY(i, 11)
            If MM1 > MM2 Then
                MM2 = MM1
                MM1 = MM2
            Else
                MM1 = MM2
            End If
    Next i
End If
If WFX = 1 Then
    ZZ = ZZ + 1
    For i = 1 To XXX
            MM2 = MM1
            MM1 = SY(i, 12)
            If MM1 > MM2 Then
                MM2 = MM1
                MM1 = MM2
            Else
                MM1 = MM2
            End If
    Next i
End If
If SFX = 1 Then
    ZZ = ZZ + 1
    For i = 1 To XXX
            MM2 = MM1
            MM1 = SY(i, 13)
            If MM1 > MM2 Then
                MM2 = MM1
                MM1 = MM2
            Else
                MM1 = MM2
            End If
    Next i
End If
If DUTYX = 1 Then
    ZZ = ZZ + 1
    For i = 1 To XXX
            MM2 = MM1
            MM1 = SY(i, 14)
            If MM1 > MM2 Then
                MM2 = MM1
                MM1 = MM2
            Else
                MM1 = MM2
            End If
    Next i
End If

If CFACX = 1 Then
    ZZ = ZZ + 1
    For i = 1 To XXX
            MM2 = MM1
            MM1 = SY(i, 15)
            If MM1 > MM2 Then
                MM2 = MM1
                MM1 = MM2
            Else
                MM1 = MM2
            End If
    Next i
End If
  
With MSChart_ALL_UNITS
.Width = 12800
.Height = 8500
.Top = 500
.Left = 100
If Check1 = Checked Then
   .chartType = VtChChartType2dLine
ElseIf Check1 = Unchecked Then
   .chartType = VtChChartType3dBar
End If
.ColumnCount = XXX
.RowCount = 15
     With MSChart_ALL_UNITS.Plot
          .AngleUnit = VtAngleUnitsDegrees
          .Projection = VtProjectionTypeOrthogonal
          .Axis(VtChAxisIdY).CategoryScale.Auto = False
          .Axis(VtChAxisIdY).ValueScale.Minimum = 0
          If Check_MAX = Unchecked Then
                MM2 = Int(MM2 * 1000 + MM2 * 200) / 1000
                Min = 0
                Spin_MAX.Value = MM2
                Spin_MIN.Value = Min
          Else
                MM2 = Spin_MAX
                Min = Spin_MIN
          End If
              .Axis(VtChAxisIdY).ValueScale.Maximum = Format(MM2, "0.0")
              .Axis(VtChAxisIdY).ValueScale.Minimum = Min
              .Axis(VtChAxisIdY).ValueScale.MajorDivision = 5
              .Axis(VtChAxisIdY2).ValueScale.MajorDivision = 5
              .Axis(VtChAxisIdY).ValueScale.MinorDivision = 2
              .Axis(VtChAxisIdY2).ValueScale.MinorDivision = 2
              .DepthToHeightRatio = 1.5
              .WidthToHeightRatio = 1.5
              .xGap = 0.8
              .zGap = 0.8
     End With
       
        If FFX = 1 Then
          For c = 1 To XXX
             r = 1
                  .Column = c
                  .Row = r
                  If SY(c, r) > 0 Then
                      .Data = SY(c, r)
                  Else
                      .Data = 0
                  End If
                  .ColumnLabel = ET(c)
                  .RowLabel = "FF,[°C.m^2/KW (10^-3)"
          Next c
        .Plot.SeriesCollection.Item(1).Position.Hidden = False
        Else
        .Plot.SeriesCollection.Item(1).Position.Hidden = True
        End If
        
        If SKX = 1 Then
          For c = 1 To XXX
             r = 2
                  .Column = c
                  .Row = r
                  If SY(c, r) > 0 Then
                      .Data = SY(c, r)
                  Else
                      .Data = 0
                  End If
                  .ColumnLabel = ET(c)
                  .RowLabel = "Skin-T,°C"
          Next c
        .Plot.SeriesCollection.Item(2).Position.Hidden = False
        Else
        .Plot.SeriesCollection.Item(2).Position.Hidden = True
        End If
        
        If LMTDX = 1 Then
          For c = 1 To XXX
             r = 3
                  .Column = c
                  .Row = r
                  If SY(c, r) > 0 Then
                      .Data = SY(c, r)
                  Else
                      .Data = 0
                  End If
                  .ColumnLabel = ET(c)
                  .RowLabel = "MTDc,°C"
          Next c
        .Plot.SeriesCollection.Item(3).Position.Hidden = False
        Else
        .Plot.SeriesCollection.Item(3).Position.Hidden = True
        End If
        
        If TTDX = 1 Then
          For c = 1 To XXX
             r = 4
                  .Column = c
                  .Row = r
                  If SY(c, r) > 0 Then
                      .Data = SY(c, r)
                  Else
                      .Data = 0
                  End If
                  .ColumnLabel = ET(c)
                  .RowLabel = "Appr-T,°C"
          Next c
        .Plot.SeriesCollection.Item(4).Position.Hidden = False
        Else
        .Plot.SeriesCollection.Item(4).Position.Hidden = True
        End If
        
        If REX = 1 Then
          For c = 1 To XXX
             r = 5
                  .Column = c
                  .Row = r
                  If SY(c, r) > 0 Then
                      .Data = SY(c, r)
                  Else
                      .Data = 0
                  End If
                  .ColumnLabel = ET(c)
                  .RowLabel = "T-RE (10^3)"
          Next c
        .Plot.SeriesCollection.Item(5).Position.Hidden = False
        Else
        .Plot.SeriesCollection.Item(5).Position.Hidden = True
        End If
        
        If VELX = 1 Then
          For c = 1 To XXX
             r = 6
                  .Column = c
                  .Row = r
                  If SY(c, r) > 0 Then
                      .Data = SY(c, r)
                  Else
                      .Data = 0
                  End If
                  .ColumnLabel = ET(c)
                  .RowLabel = "T-Vel,m/s"
          Next c
        .Plot.SeriesCollection.Item(6).Position.Hidden = False
        Else
        .Plot.SeriesCollection.Item(6).Position.Hidden = True
        End If
        
        If PDX = 1 Then
          For c = 1 To XXX
             r = 7
                  .Column = c
                  .Row = r
                  If SY(c, r) > 0 Then
                      .Data = SY(c, r)
                  Else
                      .Data = 0
                  End If
                  .ColumnLabel = ET(c)
                  .RowLabel = "P-Drop,KPa"
          Next c
        .Plot.SeriesCollection.Item(7).Position.Hidden = False
        Else
        .Plot.SeriesCollection.Item(7).Position.Hidden = True
        End If
        
        If CPX = 1 Then
          For c = 1 To XXX
             r = 8
                  .Column = c
                  .Row = r
                  If SY(c, r) > 0 Then
                      .Data = SY(c, r)
                  Else
                      .Data = 0
                  End If
                  .ColumnLabel = ET(c)
                  .RowLabel = "Cond-P,KPa(a)"
          Next c
        .Plot.SeriesCollection.Item(8).Position.Hidden = False
        Else
        .Plot.SeriesCollection.Item(8).Position.Hidden = True
        End If
        
        If CTX = 1 Then
          For c = 1 To XXX
             r = 9
                  .Column = c
                  .Row = r
                  If SY(c, r) > 0 Then
                      .Data = SY(c, r)
                  Else
                      .Data = 0
                  End If
                  .ColumnLabel = ET(c)
                  .RowLabel = "Cond-T,°C"
          Next c
        .Plot.SeriesCollection.Item(9).Position.Hidden = False
        Else
        .Plot.SeriesCollection.Item(9).Position.Hidden = True
        End If
        
        If TINX = 1 Then
          For c = 1 To XXX
             r = 10
                  .Column = c
                  .Row = r
                  If SY(c, r) > 0 Then
                      .Data = Format(Val(SY(c, r)), "0.0")
                  Else
                      .Data = 0
                  End If
                  .ColumnLabel = ET(c)
                  .RowLabel = "T-IN,°C"
          Next c
        .Plot.SeriesCollection.Item(10).Position.Hidden = False
        Else
        .Plot.SeriesCollection.Item(10).Position.Hidden = True
        End If
        
        If TOUTX = 1 Then
          For c = 1 To XXX
             r = 11
                  .Column = c
                  .Row = r
                  If SY(c, r) > 0 Then
                      .Data = Format(Val(SY(c, r)), "0.0")
                  Else
                      .Data = 0
                  End If
                  .ColumnLabel = ET(c)
                  .RowLabel = "T-OUT,°C"
          Next c
        .Plot.SeriesCollection.Item(11).Position.Hidden = False
        Else
        .Plot.SeriesCollection.Item(11).Position.Hidden = True
        End If
        
        If WFX = 1 Then
          For c = 1 To XXX
             r = 12
                  .Column = c
                  .Row = r
                  If SY(c, r) > 0 Then
                      .Data = Format(Val(SY(c, r)), "0.0")
                  Else
                      .Data = 0
                  End If
                  .ColumnLabel = ET(c)
                  .RowLabel = "T-Flow,kg/h (10^3)"
          Next c
        .Plot.SeriesCollection.Item(12).Position.Hidden = False
        Else
        .Plot.SeriesCollection.Item(12).Position.Hidden = True
        End If
        
        If SFX = 1 Then
          For c = 1 To XXX
             r = 13
                  .Column = c
                  .Row = r
                  If SY(c, r) > 0 Then
                      .Data = Format(Val(SY(c, r)), "0.0")
                  Else
                      .Data = 0
                  End If
                  .ColumnLabel = ET(c)
                  .RowLabel = "S-Flow,kg/h (10^3)"
          Next c
        .Plot.SeriesCollection.Item(13).Position.Hidden = False
        Else
        .Plot.SeriesCollection.Item(13).Position.Hidden = True
        End If
        
        If DUTYX = 1 Then
          For c = 1 To XXX
             r = 14
                  .Column = c
                  .Row = r
                  If SY(c, r) > 0 Then
                      .Data = Format(Val(SY(c, r)), "0.0")
                  Else
                      .Data = 0
                  End If
                  .ColumnLabel = ET(c)
                  .RowLabel = "DUTY,KW"
          Next c
        .Plot.SeriesCollection.Item(14).Position.Hidden = False
        Else
        .Plot.SeriesCollection.Item(14).Position.Hidden = True
        End If
        
        If CFACX = 1 Then
          For c = 1 To XXX
             r = 15
                  .Column = c
                  .Row = r
                  If SY(c, r) > 0 Then
                      .Data = SY(c, r)
                  Else
                      .Data = 0
                  End If
                  .ColumnLabel = ET(c)
                  .RowLabel = "CFAC,m3/h.kPa"
          Next c
        .Plot.SeriesCollection.Item(15).Position.Hidden = False
        Else
        .Plot.SeriesCollection.Item(15).Position.Hidden = True
        End If
End With
End Sub

Private Sub Toggle_PARMETER_Click()
    If Toggle_PARMETER = True Then
        Frame_PARAMETER.Visible = True
    ElseIf Toggle_PARMETER = False Then
        Frame_PARAMETER.Visible = False
    End If
End Sub
Private Sub Check_1_Click()
    Call chart
End Sub
Private Sub Check_CFACTOR_Click()
    Call chart
End Sub
Private Sub Check_CP_Click()
    Call chart
End Sub
Private Sub Check_CT_Click()
    Call chart
End Sub
Private Sub Check_DUTY_Click()
    Call chart
End Sub
Private Sub Check_LMTD_Click()
    Call chart
End Sub
Private Sub Check_S_FLOW_Click()
    Call chart
End Sub
Private Sub Check_SKIN_Click()
    Call chart
End Sub
Private Sub Check_T_FLOW_Click()
    Call chart
End Sub
Private Sub Check_T_PD_Click()
    Call chart
End Sub
Private Sub Check_T_RE_Click()
    Call chart
End Sub
Private Sub Check_T_TEMP_IN_Click()
    Call chart
End Sub
Private Sub Check_T_TEMP_OUT_Click()
    Call chart
End Sub
Private Sub Check_T_VEL_Click()
    Call chart
End Sub
Private Sub Check_TTD_Click()
    Call chart
End Sub
Private Sub Check_WATER_FOUL_Click()
    Call chart
End Sub
Private Sub Toggle_PLANT_Click()
    If Toggle_PLANT = True Then
        Toggle_PLANT.BackColor = &HC000&
    Else
        Toggle_PLANT.BackColor = &HE0E0E0
    End If
    If Toggle_summary = True Then
        Call Summary_TABLE_DES
    ElseIf Toggle_summary = False Then
        Call chart
    End If
End Sub
Private Sub Toggle_PLANT_UNIT_Click()
    If Toggle_PLANT_UNIT = True Then
        Toggle_PLANT_UNIT.BackColor = &HC000&
    Else
        Toggle_PLANT_UNIT.BackColor = &HE0E0E0
    End If
    If Toggle_summary = True Then
        Call Summary_TABLE_DES
    ElseIf Toggle_summary = False Then
        Call chart
    End If
End Sub
Private Sub Toggle_TOWER_Click()
    If Toggle_TOWER = True Then
        Toggle_TOWER.BackColor = &HC000&
    Else
        Toggle_TOWER.BackColor = &HE0E0E0
    End If
    If Toggle_summary = True Then
        Call Summary_TABLE_DES
    ElseIf Toggle_summary = False Then
        Call chart
    End If
End Sub
Private Sub Toggle_UNIT_Click()
    If Toggle_UNIT = True Then
        Toggle_UNIT.BackColor = &HC000&
    Else
        Toggle_UNIT.BackColor = &HE0E0E0
    End If
    If Toggle_summary = True Then
        Call Summary_TABLE_DES
    ElseIf Toggle_summary = False Then
        Call chart
    End If
End Sub

Private Sub Toggle_UNIT_DATE_Click()
    If Toggle_UNIT_DATE = True Then
        Lab_Unit.Caption = "By Unit name"
        Toggle_UNIT_DATE.BackColor = &HC000&
    Else
        Lab_Unit.Caption = "By Date"
        Toggle_UNIT_DATE.BackColor = &HE0E0E0
    End If
    If Toggle_summary = True Then
        Call Summary_TABLE_DES
    ElseIf Toggle_summary = False Then
        Call chart
    End If
End Sub
Private Sub Summary_TABLE_DES()
Dim TY(10, 1000), SY(10, 1000)
Dim T_FLOW(5), T_VAPOR(5), T_LIQUID(5), T_WATER(5), T_NC(5), T_DUTY(5), T_P_DROP(5)
Dim S_FLOW(5), S_VAPOR(5), S_LIQUID(5), S_WATER(5), S_NC(5), S_duty(5), S_P_DROP(5)

Text1(0).Text = Combo_PLANT.Text
Text1(1).Text = Combo_TOWER.Text
Text1(2).Text = Combo_PLANT_UNIT.Text
Text1(3).Text = Combo_UNIT.Text

Data6.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
Data6.RecordSource = "Select * From [query_test]"
Data6.Refresh
Set Rs6 = Data6.Recordset
If Rs6.RecordCount > 0 Then
    Do Until Rs6.EOF
        PLANT1 = Data6.Recordset.Plant
        TOWER1 = Data6.Recordset.COOL_TOWER
        PLANT_U1 = Data6.Recordset.PLANT_UNIT
        UNIT1 = Data6.Recordset.Unit_name
        CHK = Data6.Recordset.CHECK_DESIGN
        PLANT2 = Combo_PLANT
        TOWER2 = Combo_TOWER
        PLANT_U2 = Combo_PLANT_UNIT
        UNIT2 = Combo_UNIT
        rec = Data6.Recordset.AbsolutePosition + 1
        i = i + 1
        If Toggle_PLANT = True And CHK = True Then
            If PLANT1 = PLANT2 Then
                TY(1, i) = Data6.Recordset.TUBES_FLOW
                    T_FLOW(1) = T_FLOW(1) + TY(1, i)
                TY(2, i) = Data6.Recordset.TUBES_VAPOR
                    T_VAPOR(1) = T_VAPOR(1) + TY(2, i)
                TY(3, i) = Data6.Recordset.TUBES_LIQUID
                    T_LIQUID(1) = T_LIQUID(1) + TY(3, i)
                TY(4, i) = Data6.Recordset.TUBES_WATER
                    T_WATER(1) = T_WATER(1) + TY(4, i)
                TY(5, i) = Data6.Recordset.TUBES_NON_COND
                    T_NC(1) = T_NC(1) + TY(5, i)
                TY(6, i) = Data6.Recordset.TUBES_DUTY
                    T_DUTY(1) = T_DUTY(1) + TY(6, i)
                TY(7, i) = Data6.Recordset.TUBES_PRESS_DROP
                    T_P_DROP(1) = T_P_DROP(1) + TY(7, i)
                
                SY(1, i) = Data6.Recordset.SHELL_FLOW
                    S_FLOW(1) = S_FLOW(1) + SY(1, i)
                SY(2, i) = Data6.Recordset.SHELL_VAPOR
                    S_VAPOR(1) = S_VAPOR(1) + SY(2, i)
                SY(3, i) = Data6.Recordset.SHELL_LIQUID
                    S_LIQUID(1) = S_LIQUID(1) + SY(3, i)
                SY(4, i) = Data6.Recordset.SHELL_WATER
                    S_WATER(1) = S_WATER(1) + SY(4, i)
                SY(5, i) = Data6.Recordset.SHELL_NON_COND
                    S_NC(1) = S_NC(1) + SY(5, i)
                SY(6, i) = Data6.Recordset.SHELL_DUTY
                    S_duty(1) = S_duty(1) + SY(6, i)
                SY(7, i) = Data6.Recordset.SHELL_PRESS_DROP
                    S_P_DROP(1) = S_P_DROP(1) + SY(7, i)
            End If
        End If
        If Toggle_PLANT = True And Toggle_TOWER = True And CHK = True Then
            If PLANT1 = PLANT2 And TOWER1 = TOWER2 Then
                TY(1, i) = Data6.Recordset.TUBES_FLOW
                    T_FLOW(2) = T_FLOW(2) + TY(1, i)
                TY(2, i) = Data6.Recordset.TUBES_VAPOR
                    T_VAPOR(2) = T_VAPOR(2) + TY(2, i)
                TY(3, i) = Data6.Recordset.TUBES_LIQUID
                    T_LIQUID(2) = T_LIQUID(2) + TY(3, i)
                TY(4, i) = Data6.Recordset.TUBES_WATER
                    T_WATER(2) = T_WATER(2) + TY(4, i)
                TY(5, i) = Data6.Recordset.TUBES_NON_COND
                    T_NC(2) = T_NC(2) + TY(5, i)
                TY(6, i) = Data6.Recordset.TUBES_DUTY
                    T_DUTY(2) = T_DUTY(2) + TY(6, i)
                TY(7, i) = Data6.Recordset.TUBES_PRESS_DROP
                    T_P_DROP(2) = T_P_DROP(2) + TY(7, i)
                
                SY(1, i) = Data6.Recordset.SHELL_FLOW
                    S_FLOW(2) = S_FLOW(2) + SY(1, i)
                SY(2, i) = Data6.Recordset.SHELL_VAPOR
                    S_VAPOR(2) = S_VAPOR(2) + SY(2, i)
                SY(3, i) = Data6.Recordset.SHELL_LIQUID
                    S_LIQUID(2) = S_LIQUID(2) + SY(3, i)
                SY(4, i) = Data6.Recordset.SHELL_WATER
                    S_WATER(2) = S_WATER(2) + SY(4, i)
                SY(5, i) = Data6.Recordset.SHELL_NON_COND
                    S_NC(2) = S_NC(2) + SY(5, i)
                SY(6, i) = Data6.Recordset.SHELL_DUTY
                    S_duty(2) = S_duty(2) + SY(6, i)
                SY(7, i) = Data6.Recordset.SHELL_PRESS_DROP
                    S_P_DROP(2) = S_P_DROP(2) + SY(7, i)
            End If
        End If
        If Toggle_PLANT = True And Toggle_PLANT_UNIT = True And CHK = True Then
            If PLANT1 = PLANT2 And PLANT_U1 = PLANT_U2 Then
                TY(1, i) = Data6.Recordset.TUBES_FLOW
                    T_FLOW(3) = T_FLOW(3) + TY(1, i)
                TY(2, i) = Data6.Recordset.TUBES_VAPOR
                    T_VAPOR(3) = T_VAPOR(3) + TY(2, i)
                TY(3, i) = Data6.Recordset.TUBES_LIQUID
                    T_LIQUID(3) = T_LIQUID(3) + TY(3, i)
                TY(4, i) = Data6.Recordset.TUBES_WATER
                    T_WATER(3) = T_WATER(3) + TY(4, i)
                TY(5, i) = Data6.Recordset.TUBES_NON_COND
                    T_NC(3) = T_NC(3) + TY(5, i)
                TY(6, i) = Data6.Recordset.TUBES_DUTY
                    T_DUTY(3) = T_DUTY(3) + TY(6, i)
                TY(7, i) = Data6.Recordset.TUBES_PRESS_DROP
                    T_P_DROP(3) = T_P_DROP(3) + TY(7, i)
                
                SY(1, i) = Data6.Recordset.SHELL_FLOW
                    S_FLOW(3) = S_FLOW(3) + SY(1, i)
                SY(2, i) = Data6.Recordset.SHELL_VAPOR
                    S_VAPOR(3) = S_VAPOR(3) + SY(2, i)
                SY(3, i) = Data6.Recordset.SHELL_LIQUID
                    S_LIQUID(3) = S_LIQUID(3) + SY(3, i)
                SY(4, i) = Data6.Recordset.SHELL_WATER
                    S_WATER(3) = S_WATER(3) + SY(4, i)
                SY(5, i) = Data6.Recordset.SHELL_NON_COND
                    S_NC(3) = S_NC(3) + SY(5, i)
                SY(6, i) = Data6.Recordset.SHELL_DUTY
                    S_duty(3) = S_duty(3) + SY(6, i)
                SY(7, i) = Data6.Recordset.SHELL_PRESS_DROP
                    S_P_DROP(3) = S_P_DROP(3) + SY(7, i)
            End If
        End If
        If Toggle_PLANT = True And Toggle_UNIT = True And CHK = True Then
            If PLANT1 = PLANT2 And UNIT1 = UNIT2 Then
                TY(1, i) = Data6.Recordset.TUBES_FLOW
                    T_FLOW(4) = T_FLOW(4) + TY(1, i)
                TY(2, i) = Data6.Recordset.TUBES_VAPOR
                    T_VAPOR(4) = T_VAPOR(4) + TY(2, i)
                TY(3, i) = Data6.Recordset.TUBES_LIQUID
                    T_LIQUID(4) = T_LIQUID(4) + TY(3, i)
                TY(4, i) = Data6.Recordset.TUBES_WATER
                    T_WATER(4) = T_WATER(4) + TY(4, i)
                TY(5, i) = Data6.Recordset.TUBES_NON_COND
                    T_NC(4) = T_NC(4) + TY(5, i)
                TY(6, i) = Data6.Recordset.TUBES_DUTY
                    T_DUTY(4) = T_DUTY(4) + TY(6, i)
                TY(7, i) = Data6.Recordset.TUBES_PRESS_DROP
                    T_P_DROP(4) = T_P_DROP(4) + TY(7, i)
                
                SY(1, i) = Data6.Recordset.SHELL_FLOW
                    S_FLOW(4) = S_FLOW(4) + SY(1, i)
                SY(2, i) = Data6.Recordset.SHELL_VAPOR
                    S_VAPOR(4) = S_VAPOR(4) + SY(2, i)
                SY(3, i) = Data6.Recordset.SHELL_LIQUID
                    S_LIQUID(4) = S_LIQUID(4) + SY(3, i)
                SY(4, i) = Data6.Recordset.SHELL_WATER
                    S_WATER(4) = S_WATER(4) + SY(4, i)
                SY(5, i) = Data6.Recordset.SHELL_NON_COND
                    S_NC(4) = S_NC(4) + SY(5, i)
                SY(6, i) = Data6.Recordset.SHELL_DUTY
                    S_duty(4) = S_duty(4) + SY(6, i)
                SY(7, i) = Data6.Recordset.SHELL_PRESS_DROP
                    S_P_DROP(4) = S_P_DROP(4) + SY(7, i)
            End If
        End If
        rec = Data6.Recordset.AbsolutePosition + 1
        Rs6.MoveNext
    Loop
End If
For j = 1 To 1
    TOTAL_T(0) = Format(T_FLOW(j), "##,##0")
    TOTAL_T(1) = Format(T_VAPOR(j), "##,##0")
    TOTAL_T(2) = Format(T_LIQUID(j), "##,##0")
    TOTAL_T(3) = Format(T_WATER(j), "##,##0")
    TOTAL_T(4) = Format(T_NC(j), "##,##0")
    TOTAL_T(5) = Format(T_DUTY(j), "##,##0")
    TOTAL_T(6) = Format(T_P_DROP(j), "##,##0")
Next j
For j = 2 To 2
    TOTAL_T(7) = Format(T_FLOW(j), "##,##0")
    TOTAL_T(8) = Format(T_VAPOR(j), "##,##0")
    TOTAL_T(9) = Format(T_LIQUID(j), "##,##0")
    TOTAL_T(10) = Format(T_WATER(j), "##,##0")
    TOTAL_T(11) = Format(T_NC(j), "##,##0")
    TOTAL_T(12) = Format(T_DUTY(j), "##,##0")
    TOTAL_T(13) = Format(T_P_DROP(j), "##,##0")
Next j
For j = 3 To 3
    TOTAL_T(14) = Format(T_FLOW(j), "##,##0")
    TOTAL_T(15) = Format(T_VAPOR(j), "##,##0")
    TOTAL_T(16) = Format(T_LIQUID(j), "##,##0")
    TOTAL_T(17) = Format(T_WATER(j), "##,##0")
    TOTAL_T(18) = Format(T_NC(j), "##,##0")
    TOTAL_T(19) = Format(T_DUTY(j), "##,##0")
    TOTAL_T(20) = Format(T_P_DROP(j), "##,##0")
Next j
For j = 4 To 4
    TOTAL_T(21) = Format(T_FLOW(j), "##,##0")
    TOTAL_T(22) = Format(T_VAPOR(j), "##,##0")
    TOTAL_T(23) = Format(T_LIQUID(j), "##,##0")
    TOTAL_T(24) = Format(T_WATER(j), "##,##0")
    TOTAL_T(25) = Format(T_NC(j), "##,##0")
    TOTAL_T(26) = Format(T_DUTY(j), "##,##0")
    TOTAL_T(27) = Format(T_P_DROP(j), "##,##0")
Next j

For j = 1 To 1
    TOTAL_S(0) = Format(S_FLOW(j), "##,##0")
    TOTAL_S(1) = Format(S_VAPOR(j), "##,##0")
    TOTAL_S(2) = Format(S_LIQUID(j), "##,##0")
    TOTAL_S(3) = Format(S_WATER(j), "##,##0")
    TOTAL_S(4) = Format(S_NC(j), "##,##0")
    TOTAL_S(5) = Format(S_duty(j), "##,##0")
    TOTAL_S(6) = Format(S_P_DROP(j), "##,##0")
Next j
For j = 2 To 2
    TOTAL_S(7) = Format(S_FLOW(j), "##,##0")
    TOTAL_S(8) = Format(S_VAPOR(j), "##,##0")
    TOTAL_S(9) = Format(S_LIQUID(j), "##,##0")
    TOTAL_S(10) = Format(S_WATER(j), "##,##0")
    TOTAL_S(11) = Format(S_NC(j), "##,##0")
    TOTAL_S(12) = Format(S_duty(j), "##,##0")
    TOTAL_S(13) = Format(S_P_DROP(j), "##,##0")
Next j
For j = 3 To 3
    TOTAL_S(14) = Format(S_FLOW(j), "##,##0")
    TOTAL_S(15) = Format(S_VAPOR(j), "##,##0")
    TOTAL_S(16) = Format(S_LIQUID(j), "##,##0")
    TOTAL_S(17) = Format(S_WATER(j), "##,##0")
    TOTAL_S(18) = Format(S_NC(j), "##,##0")
    TOTAL_S(19) = Format(S_duty(j), "##,##0")
    TOTAL_S(20) = Format(S_P_DROP(j), "##,##0")
Next j
For j = 4 To 4
    TOTAL_S(21) = Format(S_FLOW(j), "##,##0")
    TOTAL_S(22) = Format(S_VAPOR(j), "##,##0")
    TOTAL_S(23) = Format(S_LIQUID(j), "##,##0")
    TOTAL_S(24) = Format(S_WATER(j), "##,##0")
    TOTAL_S(25) = Format(S_NC(j), "##,##0")
    TOTAL_S(26) = Format(S_duty(j), "##,##0")
    TOTAL_S(27) = Format(S_P_DROP(j), "##,##0")
Next j
Call Summary_TABLE_ACT
End Sub
Private Sub Summary_TABLE_ACT()
Dim TY(10, 1000), SY(10, 1000)
Dim T_FLOW(5), T_VAPOR(5), T_LIQUID(5), T_WATER(5), T_NC(5), T_DUTY(5), T_P_DROP(5)
Dim S_FLOW(5), S_VAPOR(5), S_LIQUID(5), S_WATER(5), S_NC(5), S_duty(5), S_P_DROP(5)

Text_ACT(0).Text = Combo_PLANT.Text
Text_ACT(1).Text = Combo_TOWER.Text
Text_ACT(2).Text = Combo_PLANT_UNIT.Text
Text_ACT(3).Text = Combo_UNIT.Text

Data6.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
Data6.RecordSource = "Select * From [query_test]"
Data6.Refresh
Set Rs6 = Data6.Recordset
If Rs6.RecordCount > 0 Then
    Do Until Rs6.EOF
        PLANT1 = Data6.Recordset.Plant
        TOWER1 = Data6.Recordset.COOL_TOWER
        PLANT_U1 = Data6.Recordset.PLANT_UNIT
        UNIT1 = Data6.Recordset.Unit_name
        CHK = Data6.Recordset.CHECK_ACTUAL
        PLANT2 = Combo_PLANT
        TOWER2 = Combo_TOWER
        PLANT_U2 = Combo_PLANT_UNIT
        UNIT2 = Combo_UNIT
        rec = Data6.Recordset.AbsolutePosition + 1
        i = i + 1
        If Toggle_PLANT = True And CHK = True Then
            If PLANT1 = PLANT2 Then
                TY(1, i) = Data6.Recordset.TUBES_FLOW
                    T_FLOW(1) = T_FLOW(1) + TY(1, i)
                TY(2, i) = Data6.Recordset.TUBES_VAPOR
                    T_VAPOR(1) = T_VAPOR(1) + TY(2, i)
                TY(3, i) = Data6.Recordset.TUBES_LIQUID
                    T_LIQUID(1) = T_LIQUID(1) + TY(3, i)
                TY(4, i) = Data6.Recordset.TUBES_WATER
                    T_WATER(1) = T_WATER(1) + TY(4, i)
                TY(5, i) = Data6.Recordset.TUBES_NON_COND
                    T_NC(1) = T_NC(1) + TY(5, i)
                TY(6, i) = Data6.Recordset.TUBES_DUTY
                    T_DUTY(1) = T_DUTY(1) + TY(6, i)
                TY(7, i) = Data6.Recordset.TUBES_PRESS_DROP
                    T_P_DROP(1) = T_P_DROP(1) + TY(7, i)
                
                SY(1, i) = Data6.Recordset.SHELL_FLOW
                    S_FLOW(1) = S_FLOW(1) + SY(1, i)
                SY(2, i) = Data6.Recordset.SHELL_VAPOR
                    S_VAPOR(1) = S_VAPOR(1) + SY(2, i)
                SY(3, i) = Data6.Recordset.SHELL_LIQUID
                    S_LIQUID(1) = S_LIQUID(1) + SY(3, i)
                SY(4, i) = Data6.Recordset.SHELL_WATER
                    S_WATER(1) = S_WATER(1) + SY(4, i)
                SY(5, i) = Data6.Recordset.SHELL_NON_COND
                    S_NC(1) = S_NC(1) + SY(5, i)
                SY(6, i) = Data6.Recordset.SHELL_DUTY
                    S_duty(1) = S_duty(1) + SY(6, i)
                SY(7, i) = Data6.Recordset.SHELL_PRESS_DROP
                    S_P_DROP(1) = S_P_DROP(1) + SY(7, i)
            End If
        End If
        If Toggle_PLANT = True And Toggle_TOWER = True And CHK = True Then
            If PLANT1 = PLANT2 And TOWER1 = TOWER2 Then
                TY(1, i) = Data6.Recordset.TUBES_FLOW
                    T_FLOW(2) = T_FLOW(2) + TY(1, i)
                TY(2, i) = Data6.Recordset.TUBES_VAPOR
                    T_VAPOR(2) = T_VAPOR(2) + TY(2, i)
                TY(3, i) = Data6.Recordset.TUBES_LIQUID
                    T_LIQUID(2) = T_LIQUID(2) + TY(3, i)
                TY(4, i) = Data6.Recordset.TUBES_WATER
                    T_WATER(2) = T_WATER(2) + TY(4, i)
                TY(5, i) = Data6.Recordset.TUBES_NON_COND
                    T_NC(2) = T_NC(2) + TY(5, i)
                TY(6, i) = Data6.Recordset.TUBES_DUTY
                    T_DUTY(2) = T_DUTY(2) + TY(6, i)
                TY(7, i) = Data6.Recordset.TUBES_PRESS_DROP
                    T_P_DROP(2) = T_P_DROP(2) + TY(7, i)
                
                SY(1, i) = Data6.Recordset.SHELL_FLOW
                    S_FLOW(2) = S_FLOW(2) + SY(1, i)
                SY(2, i) = Data6.Recordset.SHELL_VAPOR
                    S_VAPOR(2) = S_VAPOR(2) + SY(2, i)
                SY(3, i) = Data6.Recordset.SHELL_LIQUID
                    S_LIQUID(2) = S_LIQUID(2) + SY(3, i)
                SY(4, i) = Data6.Recordset.SHELL_WATER
                    S_WATER(2) = S_WATER(2) + SY(4, i)
                SY(5, i) = Data6.Recordset.SHELL_NON_COND
                    S_NC(2) = S_NC(2) + SY(5, i)
                SY(6, i) = Data6.Recordset.SHELL_DUTY
                    S_duty(2) = S_duty(2) + SY(6, i)
                SY(7, i) = Data6.Recordset.SHELL_PRESS_DROP
                    S_P_DROP(2) = S_P_DROP(2) + SY(7, i)
            End If
        End If
        If Toggle_PLANT = True And Toggle_PLANT_UNIT = True And CHK = True Then
            If PLANT1 = PLANT2 And PLANT_U1 = PLANT_U2 Then
                TY(1, i) = Data6.Recordset.TUBES_FLOW
                    T_FLOW(3) = T_FLOW(3) + TY(1, i)
                TY(2, i) = Data6.Recordset.TUBES_VAPOR
                    T_VAPOR(3) = T_VAPOR(3) + TY(2, i)
                TY(3, i) = Data6.Recordset.TUBES_LIQUID
                    T_LIQUID(3) = T_LIQUID(3) + TY(3, i)
                TY(4, i) = Data6.Recordset.TUBES_WATER
                    T_WATER(3) = T_WATER(3) + TY(4, i)
                TY(5, i) = Data6.Recordset.TUBES_NON_COND
                    T_NC(3) = T_NC(3) + TY(5, i)
                TY(6, i) = Data6.Recordset.TUBES_DUTY
                    T_DUTY(3) = T_DUTY(3) + TY(6, i)
                TY(7, i) = Data6.Recordset.TUBES_PRESS_DROP
                    T_P_DROP(3) = T_P_DROP(3) + TY(7, i)
                
                SY(1, i) = Data6.Recordset.SHELL_FLOW
                    S_FLOW(3) = S_FLOW(3) + SY(1, i)
                SY(2, i) = Data6.Recordset.SHELL_VAPOR
                    S_VAPOR(3) = S_VAPOR(3) + SY(2, i)
                SY(3, i) = Data6.Recordset.SHELL_LIQUID
                    S_LIQUID(3) = S_LIQUID(3) + SY(3, i)
                SY(4, i) = Data6.Recordset.SHELL_WATER
                    S_WATER(3) = S_WATER(3) + SY(4, i)
                SY(5, i) = Data6.Recordset.SHELL_NON_COND
                    S_NC(3) = S_NC(3) + SY(5, i)
                SY(6, i) = Data6.Recordset.SHELL_DUTY
                    S_duty(3) = S_duty(3) + SY(6, i)
                SY(7, i) = Data6.Recordset.SHELL_PRESS_DROP
                    S_P_DROP(3) = S_P_DROP(3) + SY(7, i)
            End If
        End If
        If Toggle_PLANT = True And Toggle_UNIT = True And CHK = True Then
            If PLANT1 = PLANT2 And UNIT1 = UNIT2 Then
                TY(1, i) = Data6.Recordset.TUBES_FLOW
                    T_FLOW(4) = T_FLOW(4) + TY(1, i)
                TY(2, i) = Data6.Recordset.TUBES_VAPOR
                    T_VAPOR(4) = T_VAPOR(4) + TY(2, i)
                TY(3, i) = Data6.Recordset.TUBES_LIQUID
                    T_LIQUID(4) = T_LIQUID(4) + TY(3, i)
                TY(4, i) = Data6.Recordset.TUBES_WATER
                    T_WATER(4) = T_WATER(4) + TY(4, i)
                TY(5, i) = Data6.Recordset.TUBES_NON_COND
                    T_NC(4) = T_NC(4) + TY(5, i)
                TY(6, i) = Data6.Recordset.TUBES_DUTY
                    T_DUTY(4) = T_DUTY(4) + TY(6, i)
                TY(7, i) = Data6.Recordset.TUBES_PRESS_DROP
                    T_P_DROP(4) = T_P_DROP(4) + TY(7, i)
                
                SY(1, i) = Data6.Recordset.SHELL_FLOW
                    S_FLOW(4) = S_FLOW(4) + SY(1, i)
                SY(2, i) = Data6.Recordset.SHELL_VAPOR
                    S_VAPOR(4) = S_VAPOR(4) + SY(2, i)
                SY(3, i) = Data6.Recordset.SHELL_LIQUID
                    S_LIQUID(4) = S_LIQUID(4) + SY(3, i)
                SY(4, i) = Data6.Recordset.SHELL_WATER
                    S_WATER(4) = S_WATER(4) + SY(4, i)
                SY(5, i) = Data6.Recordset.SHELL_NON_COND
                    S_NC(4) = S_NC(4) + SY(5, i)
                SY(6, i) = Data6.Recordset.SHELL_DUTY
                    S_duty(4) = S_duty(4) + SY(6, i)
                SY(7, i) = Data6.Recordset.SHELL_PRESS_DROP
                    S_P_DROP(4) = S_P_DROP(4) + SY(7, i)
            End If
        End If
        rec = Data6.Recordset.AbsolutePosition + 1
        Rs6.MoveNext
    Loop
End If
For j = 1 To 1
    TOTAL_T_ACT(0) = Format(T_FLOW(j), "##,##0")
    TOTAL_T_ACT(1) = Format(T_VAPOR(j), "##,##0")
    TOTAL_T_ACT(2) = Format(T_LIQUID(j), "##,##0")
    TOTAL_T_ACT(3) = Format(T_WATER(j), "##,##0")
    TOTAL_T_ACT(4) = Format(T_NC(j), "##,##0")
    TOTAL_T_ACT(5) = Format(T_DUTY(j), "##,##0")
    TOTAL_T_ACT(6) = Format(T_P_DROP(j), "##,##0")
Next j
For j = 2 To 2
    TOTAL_T_ACT(7) = Format(T_FLOW(j), "##,##0")
    TOTAL_T_ACT(8) = Format(T_VAPOR(j), "##,##0")
    TOTAL_T_ACT(9) = Format(T_LIQUID(j), "##,##0")
    TOTAL_T_ACT(10) = Format(T_WATER(j), "##,##0")
    TOTAL_T_ACT(11) = Format(T_NC(j), "##,##0")
    TOTAL_T_ACT(12) = Format(T_DUTY(j), "##,##0")
    TOTAL_T_ACT(13) = Format(T_P_DROP(j), "##,##0")
Next j
For j = 3 To 3
    TOTAL_T_ACT(14) = Format(T_FLOW(j), "##,##0")
    TOTAL_T_ACT(15) = Format(T_VAPOR(j), "##,##0")
    TOTAL_T_ACT(16) = Format(T_LIQUID(j), "##,##0")
    TOTAL_T_ACT(17) = Format(T_WATER(j), "##,##0")
    TOTAL_T_ACT(18) = Format(T_NC(j), "##,##0")
    TOTAL_T_ACT(19) = Format(T_DUTY(j), "##,##0")
    TOTAL_T_ACT(20) = Format(T_P_DROP(j), "##,##0")
Next j
For j = 4 To 4
    TOTAL_T_ACT(21) = Format(T_FLOW(j), "##,##0")
    TOTAL_T_ACT(22) = Format(T_VAPOR(j), "##,##0")
    TOTAL_T_ACT(23) = Format(T_LIQUID(j), "##,##0")
    TOTAL_T_ACT(24) = Format(T_WATER(j), "##,##0")
    TOTAL_T_ACT(25) = Format(T_NC(j), "##,##0")
    TOTAL_T_ACT(26) = Format(T_DUTY(j), "##,##0")
    TOTAL_T_ACT(27) = Format(T_P_DROP(j), "##,##0")
Next j

For j = 1 To 1
    TOTAL_S_ACT(0) = Format(S_FLOW(j), "##,##0")
    TOTAL_S_ACT(1) = Format(S_VAPOR(j), "##,##0")
    TOTAL_S_ACT(2) = Format(S_LIQUID(j), "##,##0")
    TOTAL_S_ACT(3) = Format(S_WATER(j), "##,##0")
    TOTAL_S_ACT(4) = Format(S_NC(j), "##,##0")
    TOTAL_S_ACT(5) = Format(S_duty(j), "##,##0")
    TOTAL_S_ACT(6) = Format(S_P_DROP(j), "##,##0")
Next j
For j = 2 To 2
    TOTAL_S_ACT(7) = Format(S_FLOW(j), "##,##0")
    TOTAL_S_ACT(8) = Format(S_VAPOR(j), "##,##0")
    TOTAL_S_ACT(9) = Format(S_LIQUID(j), "##,##0")
    TOTAL_S_ACT(10) = Format(S_WATER(j), "##,##0")
    TOTAL_S_ACT(11) = Format(S_NC(j), "##,##0")
    TOTAL_S_ACT(12) = Format(S_duty(j), "##,##0")
    TOTAL_S_ACT(13) = Format(S_P_DROP(j), "##,##0")
Next j
For j = 3 To 3
    TOTAL_S_ACT(14) = Format(S_FLOW(j), "##,##0")
    TOTAL_S_ACT(15) = Format(S_VAPOR(j), "##,##0")
    TOTAL_S_ACT(16) = Format(S_LIQUID(j), "##,##0")
    TOTAL_S_ACT(17) = Format(S_WATER(j), "##,##0")
    TOTAL_S_ACT(18) = Format(S_NC(j), "##,##0")
    TOTAL_S_ACT(19) = Format(S_duty(j), "##,##0")
    TOTAL_S_ACT(20) = Format(S_P_DROP(j), "##,##0")
Next j
For j = 4 To 4
    TOTAL_S_ACT(21) = Format(S_FLOW(j), "##,##0")
    TOTAL_S_ACT(22) = Format(S_VAPOR(j), "##,##0")
    TOTAL_S_ACT(23) = Format(S_LIQUID(j), "##,##0")
    TOTAL_S_ACT(24) = Format(S_WATER(j), "##,##0")
    TOTAL_S_ACT(25) = Format(S_NC(j), "##,##0")
    TOTAL_S_ACT(26) = Format(S_duty(j), "##,##0")
    TOTAL_S_ACT(27) = Format(S_P_DROP(j), "##,##0")
Next j

End Sub

Private Sub Toggle_summary_Click()
    If Toggle_summary = True Then
        Call Summary_TABLE_DES
        Frame_Summary_DES.Visible = True
        Frame_Summary_ACT.Visible = True
    ElseIf Toggle_summary = False Then
        Frame_Summary_DES.Visible = False
        Frame_Summary_ACT.Visible = False
    End If
End Sub
