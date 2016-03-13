VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opzioni"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Opzioni"
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2490
      TabIndex        =   1
      Tag             =   "OK"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annulla"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Tag             =   "Annulla"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Applica"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Tag             =   "&Applica"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Esempio 4"
         Height          =   2022
         Left            =   505
         TabIndex        =   11
         Tag             =   "Esempio 4"
         Top             =   502
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Esempio 3"
         Height          =   2022
         Left            =   406
         TabIndex        =   10
         Tag             =   "Esempio 3"
         Top             =   403
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Esempio 2"
         Height          =   2022
         Left            =   307
         TabIndex        =   8
         Tag             =   "Esempio 2"
         Top             =   305
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   0
      Left            =   210
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample1 
         Caption         =   "Esempio 1"
         Height          =   2022
         Left            =   208
         TabIndex        =   4
         Tag             =   "Esempio 1"
         Top             =   207
         Width           =   2033
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   4245
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7488
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gruppo 1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gruppo 2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gruppo 3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gruppo 4"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApply_Click()
    'Da fare: aggiungere il codice per "cmdApply_Click".
    MsgBox "Inserire in questo punto il codice per impostare le opzioni senza chiudere la finestra."
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOK_Click()
    'Da fare: aggiungere il codice per "cmdOK_Click".
    MsgBox "Inserire in questo punto il codice per impostare le opzioni e chiudere la finestra."
    Unload Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    i = tbsOptions.SelectedItem.Index
    'Gestisce CTRL+TAB per il passaggio alla scheda successiva.
    If (Shift And 3) = 2 And KeyCode = vbKeyTab Then
        If i = tbsOptions.Tabs.Count Then
            'Ultima scheda. Torna alla scheda 1.
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'Incrementa il conteggio delle schede.
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    ElseIf (Shift And 3) = 3 And KeyCode = vbKeyTab Then
        If i = 1 Then
            'Scheda 1. Passa all'ultima scheda.
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(tbsOptions.Tabs.Count)
        Else
            'Decrementa il conteggio delle schede.
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i - 1)
        End If
    End If
End Sub


Private Sub tbsOptions_Click()
    

    Dim i As Integer
    'Visualizza e attiva i controlli della scheda selezionata
    'e nasconde e disattiva tutti gli altri.
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 210
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next
    

End Sub

