VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSteam 
   Caption         =   "WaterProof HX - Heat Exchangers Performance Monitoring - Description"
   ClientHeight    =   9660
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
   ScaleHeight     =   9660
   ScaleWidth      =   15240
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   435
      Left            =   60
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "HELP"
      Top             =   9240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      DataField       =   "INSTRUCTIONS_steam"
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
      TextRTF         =   $"frmSteamInstructions.frx":0000
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4725
      Left            =   7800
      Picture         =   "frmSteamInstructions.frx":0082
      Stretch         =   -1  'True
      Top             =   480
      Width           =   7485
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   3855
      Left            =   7800
      Picture         =   "frmSteamInstructions.frx":39FAC
      Stretch         =   -1  'True
      Top             =   5220
      Width           =   7365
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Steam surface condenser"
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
      Top             =   120
      Width           =   6555
   End
End
Attribute VB_Name = "frmSteam"
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

