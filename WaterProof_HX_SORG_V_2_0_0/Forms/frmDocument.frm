VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmDocument 
   Caption         =   "WaterProof SSC - Steam surface condenser"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   6030
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   5775
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   10186
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmDocument.frx":0000
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Form_Resize
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    rtfText.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
    rtfText.RightMargin = rtfText.Width - 400
End Sub

