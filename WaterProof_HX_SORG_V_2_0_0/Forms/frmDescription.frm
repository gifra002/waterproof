VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDescription 
   Caption         =   "WaterProof HX - Heat Exchangers Performance Monitoring - Description"
   ClientHeight    =   10185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   15240
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   435
      Left            =   7980
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "HELP"
      Top             =   7560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      DataField       =   "DESCRIPTION"
      DataSource      =   "Data1"
      Height          =   8775
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   15478
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmDescription.frx":0000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   9420
      TabIndex        =   7
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   9360
      TabIndex        =   6
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   9360
      TabIndex        =   5
      Top             =   5940
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   13500
      TabIndex        =   4
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   13500
      TabIndex        =   3
      Top             =   3540
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   13680
      TabIndex        =   2
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "TEMA: SHELL TYPE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   8160
      TabIndex        =   1
      Top             =   540
      Width           =   6555
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   5925
      Left            =   7980
      Picture         =   "frmDescription.frx":0082
      Stretch         =   -1  'True
      Top             =   1140
      Width           =   7200
   End
End
Attribute VB_Name = "frmDescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Width = frmMain.Width * 0.987  ' Imposta la larghezza del form.
   Height = frmMain.Height * 0.88    ' Imposta l'altezza del form.
   Left = 0 '(frmMain.Width - Width)   ' Centra il form orizzontalmente.
   Top = 0 '(frmMain.Height - Height)   ' Centra il form verticalmente.

End Sub

