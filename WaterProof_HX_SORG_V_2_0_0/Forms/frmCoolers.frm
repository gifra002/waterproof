VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCoolers 
   Caption         =   "WaterProof HX - Heat Exchangers Performance Monitoring - Description"
   ClientHeight    =   10185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14010
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
   ScaleWidth      =   14010
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   435
      Left            =   6960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "HELP"
      Top             =   9540
      Visible         =   0   'False
      Width           =   2535
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      DataField       =   "INSTRUCTIONS"
      DataSource      =   "Data1"
      Height          =   8775
      Left            =   840
      TabIndex        =   0
      Top             =   420
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   15478
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmCoolers.frx":0000
   End
End
Attribute VB_Name = "frmCoolers"
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

