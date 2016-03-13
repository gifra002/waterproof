VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form List 
   Caption         =   "WaterProof CTS - Cooling tower systems treatment - Data list"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   15240
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "c:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_Unit_sort"
      Top             =   7440
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "List.frx":0000
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   12726
      _Version        =   393216
      Cols            =   6
      ForeColor       =   128
      ForeColorSel    =   16711680
      FocusRect       =   2
      FillStyle       =   1
      AllowUserResizing=   1
      MousePointer    =   5
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
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   List1.Clear               ' Clears the list box.
End Sub
Private Sub Command2_Click()
   Unload Me
End Sub
Private Sub Data1_Reposition()
  Data1.Caption = "     COOLING WATER OPEN SYSTEM TREATMENT - Case: " & (Data1.Recordset.AbsolutePosition + 1)
End Sub
Private Sub Form_Load()
On Error Resume Next
   Width = frmMain.Width * 0.98 ' Imposta la larghezza del form.
   Height = frmMain.Height * 0.89     ' Imposta l'altezza del form.
   Left = 50 '(frmMain.Width - Width) / 2 ' Centra il form orizzontalmente.
   Top = 0 '(frmMain.Height - Height) / 2 ' Centra il form verticalmente.
End Sub


