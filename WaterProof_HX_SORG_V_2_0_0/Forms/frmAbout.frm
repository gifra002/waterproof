VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About WaterProof HX"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   ClipControls    =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Informazioni su STEAM"
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   180
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1380
      Width           =   540
      Begin VB.PictureBox picLogo 
         AutoRedraw      =   -1  'True
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   0
         Picture         =   "frmAbout.frx":030A
         ScaleHeight     =   20.373
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   28.046
         TabIndex        =   7
         Top             =   -60
         Width           =   1650
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4470
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   2625
      Width           =   1245
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4500
      TabIndex        =   1
      Tag             =   "&System Info..."
      Top             =   3075
      Width           =   1215
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   555
      Left            =   60
      Top             =   60
      Width           =   675
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Developed and written by G.F. Mazzani"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   1500
      TabIndex        =   10
      Top             =   2220
      Width           =   3435
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Left            =   120
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WP"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   285
      Left            =   180
      TabIndex        =   9
      Tag             =   "Societ‡Prodotto"
      Top             =   180
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "WaterProof HX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   435
      Left            =   1080
      TabIndex        =   8
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      Caption         =   $"frmAbout.frx":E606
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   645
      Left            =   1470
      TabIndex        =   6
      Tag             =   "Descrizione applicazione"
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label lblTitle 
      Caption         =   "HEAT EXCHANGERS THERMAL PERFORMANCE "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   480
      Left            =   1050
      TabIndex        =   5
      Tag             =   "Titolo applicazione"
      Top             =   780
      Width           =   4695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   225
      X2              =   5657
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   5657
      Y1              =   2445
      Y2              =   2445
   End
   Begin VB.Label lblVersion 
      Caption         =   "Versione"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   1020
      TabIndex        =   4
      Tag             =   "Versione"
      Top             =   1260
      Width           =   4095
   End
   Begin VB.Label lblDisclaimer 
      Alignment       =   2  'Center
      Caption         =   $"frmAbout.frx":E690
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1020
      Left            =   75
      TabIndex        =   3
      Tag             =   "Avviso: ..."
      Top             =   2505
      Width           =   4080
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Opzioni di protezione delle chiavi del registro di configurazione.
Const KEY_ALL_ACCESS = &H2003F
                                          

' Chiavi di primo livello del registro di configurazione.
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Stringa Unicode a terminazione Null.
Const REG_DWORD = 4                      ' Numero a 32 bit.


Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"


Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub Form_Load()
    lblVersion.Caption = "Versione " & App.Major & "." & App.Minor & "." & App.Revision
'    lblTitle.Caption = App.Title
End Sub



Private Sub cmdSysInfo_Click()
        Call StartSysInfo
End Sub


Private Sub cmdOK_Click()
        Unload Me
End Sub


Public Sub StartSysInfo()
    On Error GoTo SysInfoErr


        Dim rc As Long
        Dim SysInfoPath As String
        

        ' Tenta di recuperare il percorso\nome del programma System Info dal
        ' registro di configurazione.
        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Tenta di recuperare solo il percorso del programma System Info dal
        ' registro di configurazione.
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
                ' Convalida l'esistenza di una versione nota a 32 bit del file.
                If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                        

                ' Errore - Impossibile trovare il file.
                Else
                        GoTo SysInfoErr
                End If
        ' Errore - Impossibile trovare la voce del registro di configurazione.
        Else
                GoTo SysInfoErr
        End If
        

        Call Shell(SysInfoPath, vbNormalFocus)
        

        Exit Sub
SysInfoErr:
        MsgBox "Informazioni sul sistema non disponibili in questa fase.", vbOKOnly
End Sub


Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long                                           ' Contatore del ciclo.
        Dim rc As Long                                          ' Codice restituito.
        Dim hKey As Long                                        ' Handle a una chiave del registro di configurazione aperta.
        Dim hDepth As Long                                      '
        Dim KeyValType As Long                                  ' Tipo di dati di una chiave del registro di configurazione.
        Dim tmpVal As String                                    ' Variabile per la memorizzazione temporanea del valore di una chiave del registro di configurazione.
        Dim KeyValSize As Long                                  ' Dimensioni della variabile della chiave del registro di configurazione.
        '------------------------------------------------------------
        ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Apre la chiave del registro di configurazione.
        

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Gestisce l'errore.
        

        tmpVal = String$(1024, 0)                             ' Assegna spazio per la variabile.
        KeyValSize = 1024                                       ' Imposta le dimensioni della variabile.
        

        '----------------------------------------------------------------
        ' Recupera il valore della chiave del registro di configurazione.
        '----------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Recupera/crea il valore della chiave.
                                                

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Gestisce gli errori.
        

        tmpVal = VBA.Left(tmpVal, InStr(tmpVal, VBA.Chr(0)) - 1)
        '------------------------------------------------------------
        ' Determina il tipo di valore della chiave per la conversione.
        '------------------------------------------------------------
        Select Case KeyValType                                  ' Cerca i tipi di dati.
        Case REG_SZ                                             ' Tipo di dati String per la chiave del registro di configurazione.
                KeyVal = tmpVal                                     ' Copia il valore String.
        Case REG_DWORD                                          ' Tipo di dati Double Word per la chiave del registro di configurazione.
                For i = Len(tmpVal) To 1 Step -1                    ' Converte ogni bit.
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Crea il valore carattere per carattere.
                Next
                KeyVal = Format$("&h" + KeyVal)                     ' Converte Double Word in String.
        End Select
        

        GetKeyValue = True                                      ' Esito positivo.
        rc = RegCloseKey(hKey)                                  ' Chiude la chiave del registro di configurazione.
        Exit Function                                           ' Esce.
        

GetKeyError:    ' Svuota in seguito a un errore.
        KeyVal = ""                                             ' Imposta il valore restituito su una stringa vuota.
        GetKeyValue = False                                     ' Esito negativo.
        rc = RegCloseKey(hKey)                                  ' Chiude la chiave del registro di configurazione.
End Function

