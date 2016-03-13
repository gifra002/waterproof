VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "WaterProof HX - Heat Exchangers Performance Monitoring"
   ClientHeight    =   8190
   ClientLeft      =   1065
   ClientTop       =   1755
   ClientWidth     =   12885
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   7920
      Width           =   12885
      _ExtentX        =   22728
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17066
            Text            =   "Stato"
            TextSave        =   "Stato"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "13/03/2016"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "3.54"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1740
      Top             =   1350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1740
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0112
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0224
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0336
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0448
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":055A
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":066C
            Key             =   "Paste"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Print setup"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEnter 
      Caption         =   "Data Input"
   End
   Begin VB.Menu mnuOutput 
      Caption         =   "Trends"
      Begin VB.Menu mnuOverall 
         Caption         =   "Overall trends"
      End
      Begin VB.Menu mnuFF 
         Caption         =   "Fouling factor"
      End
      Begin VB.Menu mnuCF 
         Caption         =   "Skin temperature"
      End
      Begin VB.Menu mnuCFAC 
         Caption         =   "C Factor"
      End
      Begin VB.Menu mnuLNTD 
         Caption         =   "LMTD"
      End
      Begin VB.Menu mnuTTD 
         Caption         =   "Approach temp."
      End
      Begin VB.Menu mnuRE 
         Caption         =   "Reynolds number"
      End
      Begin VB.Menu mnuVEL 
         Caption         =   "Water flow velocity"
      End
      Begin VB.Menu mnuPRESS 
         Caption         =   "Tube side pressure drop"
      End
      Begin VB.Menu mnuSCP 
         Caption         =   "Condensing pressure"
      End
      Begin VB.Menu mnuSCT 
         Caption         =   "Condensing temperature"
      End
      Begin VB.Menu mnuWT 
         Caption         =   "Tube-side temperatures"
      End
   End
   Begin VB.Menu mnuList 
      Caption         =   "List"
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "HX Configuration"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuOverview 
         Caption         =   "Overview"
      End
      Begin VB.Menu mnuInstructions 
         Caption         =   "Instructions"
      End
      Begin VB.Menu mnuSteam 
         Caption         =   "Steam Turbine Exhaust Condenser"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private Sub MDIForm_Load()
    Width = Screen.Width * 1  ' Imposta la larghezza del form.
    Height = Screen.Height * 0.96  ' Imposta l'altezza del form.
    Left = 0 '(Screen.Width - Width) / 2   ' Centra il form orizzontalmente.
    Top = 0 '(Screen.Height - Height) / 2   ' Centra il form verticalmente.
    frmMain.WindowState = 2

End Sub
Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer
    'Se il progetto non include un file della Guida, visualizza un messaggio per
    'l'utente. È possibile impostare il file della Guida per l'applicazione nella
    'finestra di dialogo Proprietà progetto.
    If Len(App.HelpFile) = 0 Then
        MsgBox "Impossibile visualizzare il Sommario della Guida. Nessun file della Guida associato al progetto.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub

Private Sub mnuConfig_Click()
    frmConfig.Show
End Sub
Private Sub mnuOverview_Click()
    frmDescription.Show
End Sub
Private Sub mnuSteam_Click()
    frmSteam.Show
End Sub
Private Sub mnuInstructions_Click()
    frmCoolers.Show
End Sub
Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub
Private Sub mnuFileExit_Click()
    'Scarica il form.
    Unload Me
End Sub
Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Imposta stampante"
        .CancelError = True
        .ShowPrinter
    End With
End Sub
Private Sub mnuEnter_Click()
    Dim f As New frmHX
    f.Show
End Sub
Private Sub mnuList_Click()
    List.Show
End Sub

Private Sub mnuOVERALL_Click()
    Chart_ALL_TRENDS.Show
End Sub
Private Sub mnuFF_Click()
    Chart_FF.Show
End Sub
Private Sub mnuCF_Click()
        Chart_SKIN.Show
End Sub
Private Sub mnuCFAC_Click()
    Call CheckLockedStatus(temp)
''        If temp = "locked" Then
'            MsgBox "This feature is not allowed in the trial version."
'            Exit Sub
''        End If
        Chart_CFAC.Show
End Sub
Private Sub mnuLNTD_Click()
    Call CheckLockedStatus(temp)
''        If temp = "locked" Then
'            MsgBox "This feature is not allowed in the trial version."
'            Exit Sub
''        End If
        Chart_LMTD.Show
End Sub
Private Sub mnuPRESS_Click()
    Call CheckLockedStatus(temp)
''        If temp = "locked" Then
'            MsgBox "This feature is not allowed in the trial version."
'            Exit Sub
''        End If
        Chart_PD.Show
End Sub
Private Sub mnuRE_Click()
    Call CheckLockedStatus(temp)
''        If temp = "locked" Then
'            MsgBox "This feature is not allowed in the trial version."
'            Exit Sub
''        End If
            Chart_RE.Show
End Sub
Private Sub mnuSCP_Click()
    Call CheckLockedStatus(temp)
''        If temp = "locked" Then
'            MsgBox "This feature is not allowed in the trial version."
'            Exit Sub
''        End If
        Chart_CP.Show
End Sub
Private Sub mnuSCT_Click()
    Call CheckLockedStatus(temp)
''        If temp = "locked" Then
'            MsgBox "This feature is not allowed in the trial version."
'            Exit Sub
''        End If
        Chart_CT.Show
End Sub
Private Sub mnuTTD_Click()
    Call CheckLockedStatus(temp)
''        If temp = "locked" Then
'            MsgBox "This feature is not allowed in the trial version."
'            Exit Sub
''        End If
        Chart_TTD.Show
End Sub
Private Sub mnuVEL_Click()
    Call CheckLockedStatus(temp)
''        If temp = "locked" Then
'            MsgBox "This feature is not allowed in the trial version."
'            Exit Sub
''        End If
        Chart_VEL.Show
End Sub
Private Sub mnuWT_Click()
    Call CheckLockedStatus(temp)
''        If temp = "locked" Then
'            MsgBox "This feature is not allowed in the trial version."
'            Exit Sub
'        End If
        Chart_TW.Show
End Sub
