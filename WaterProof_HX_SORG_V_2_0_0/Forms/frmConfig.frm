VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfig 
   Caption         =   "WaterProof HX - Heat Exchangers Performance Monitoring - Heat exchangers configuration"
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   9885
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   8415
      Index           =   8
      Left            =   4320
      Picture         =   "frmConfig.frx":0000
      ScaleHeight     =   8355
      ScaleWidth      =   6675
      TabIndex        =   9
      Top             =   1200
      Width           =   6735
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   735
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   1296
      MultiRow        =   -1  'True
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   11
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear all configurations"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "p1-s1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "p1-s2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "p2-s1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "p2-s2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "p1-s1-2pass"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "p1-s2-2pass"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "p2-s1-2pass"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "p2-s2-2pass"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "TEMA"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   6015
      Index           =   7
      Left            =   2040
      Picture         =   "frmConfig.frx":B3F2A
      ScaleHeight     =   5955
      ScaleWidth      =   11955
      TabIndex        =   5
      Top             =   1440
      Width           =   12015
   End
   Begin VB.PictureBox Picture1 
      Height          =   5895
      Index           =   6
      Left            =   2040
      Picture         =   "frmConfig.frx":193568
      ScaleHeight     =   5835
      ScaleWidth      =   10875
      TabIndex        =   4
      ToolTipText     =   "p2s1_2pass"
      Top             =   1440
      Width           =   10935
   End
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Index           =   5
      Left            =   2040
      Picture         =   "frmConfig.frx":25BD4A
      ScaleHeight     =   4275
      ScaleWidth      =   11355
      TabIndex        =   6
      Top             =   1440
      Width           =   11415
   End
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Index           =   4
      Left            =   2040
      Picture         =   "frmConfig.frx":2F0E50
      ScaleHeight     =   4275
      ScaleWidth      =   8955
      TabIndex        =   7
      Top             =   1440
      Width           =   9015
   End
   Begin VB.PictureBox Picture1 
      Height          =   5895
      Index           =   3
      Left            =   2040
      Picture         =   "frmConfig.frx":367132
      ScaleHeight     =   5835
      ScaleWidth      =   11235
      TabIndex        =   3
      ToolTipText     =   "p2s2"
      Top             =   1440
      Width           =   11295
   End
   Begin VB.PictureBox Picture1 
      Height          =   5655
      Index           =   2
      Left            =   2040
      Picture         =   "frmConfig.frx":446770
      ScaleHeight     =   5595
      ScaleWidth      =   10755
      TabIndex        =   2
      ToolTipText     =   "p2s1"
      Top             =   1440
      Width           =   10815
   End
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Index           =   1
      Left            =   2040
      Picture         =   "frmConfig.frx":503EA2
      ScaleHeight     =   4275
      ScaleWidth      =   10755
      TabIndex        =   1
      ToolTipText     =   "p1s2"
      Top             =   1440
      Width           =   10815
   End
   Begin VB.PictureBox Picture1 
      Height          =   5055
      Index           =   0
      Left            =   2040
      Picture         =   "frmConfig.frx":59243C
      ScaleHeight     =   4995
      ScaleWidth      =   9315
      TabIndex        =   0
      ToolTipText     =   "p1s1"
      Top             =   1440
      Visible         =   0   'False
      Width           =   9375
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
    Width = frmMain.Width * 0.985 ' Imposta la larghezza del form.
    Height = frmMain.Height * 0.95    ' Imposta l'altezza del form.
    Left = 0 '(frmMain.Width - Width) / 2 ' Centra il form orizzontalmente.
    Top = 0 '(frmMain.Height - Height) / 2 ' Centra il form verticalmente.
    frmHX.WindowState = 2

    For i = 0 To 8
        Picture1(i).Visible = False
    Next i

End Sub
Private Sub Cancella()
For i = 0 To 8
    Picture1(i).Visible = False
Next i
End Sub
Private Sub TabStrip1_Click()
On Error Resume Next
    If TabStrip1.SelectedItem = "Close" Then
        Unload Me
    ElseIf TabStrip1.SelectedItem = "Clear all configurations" Then
        Call Cancella
    ElseIf TabStrip1.SelectedItem = "p1-s1" Then
        Picture1(0).Visible = True
    ElseIf TabStrip1.SelectedItem = "p1-s2" Then
        Picture1(1).Visible = True
    ElseIf TabStrip1.SelectedItem = "p2-s1" Then
        Picture1(2).Visible = True
    ElseIf TabStrip1.SelectedItem = "p2-s2" Then
        Picture1(3).Visible = True
    ElseIf TabStrip1.SelectedItem = "p1-s1-2pass" Then
        Picture1(4).Visible = True
    ElseIf TabStrip1.SelectedItem = "p1-s2-2pass" Then
        Picture1(5).Visible = True
    ElseIf TabStrip1.SelectedItem = "p2-s1-2pass" Then
        Picture1(6).Visible = True
    ElseIf TabStrip1.SelectedItem = "p2-s2-2pass" Then
        Picture1(7).Visible = True
    ElseIf TabStrip1.SelectedItem = "TEMA" Then
        Picture1(8).Visible = True
    End If
End Sub
