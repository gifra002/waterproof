VERSION 5.00
Begin VB.Form frmInstruction 
   Caption         =   "WaterProof HX - Heat Exchangers Performance Monitoring -IiNSTRUCTIONS"
   ClientHeight    =   8910
   ClientLeft      =   1440
   ClientTop       =   735
   ClientWidth     =   15180
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
   ScaleHeight     =   9862.336
   ScaleMode       =   0  'User
   ScaleWidth      =   23403.46
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Condensers\Database\steam.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Table1"
      Top             =   9480
      Visible         =   0   'False
      Width           =   2715
   End
End
Attribute VB_Name = "frmInstruction"
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

