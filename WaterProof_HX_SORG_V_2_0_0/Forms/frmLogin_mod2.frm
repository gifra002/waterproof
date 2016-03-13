VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WaterProof HX - Heat Exchangers Performance Monitoring"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Accesso"
   Visible         =   0   'False
   Begin VB.CheckBox Check_pass 
      Caption         =   "ChecK to save the password"
      Height          =   195
      Left            =   1740
      TabIndex        =   29
      Top             =   4380
      Width           =   2355
   End
   Begin VB.TextBox Key_Unlock 
      Height          =   285
      Left            =   3540
      TabIndex        =   28
      Text            =   "Key_Unlock"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox Libero_3 
      DataField       =   "Libero_3"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      TabIndex        =   26
      Text            =   "Libero_3"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   3420
      TabIndex        =   22
      Top             =   3720
      Width           =   2055
      Begin VB.TextBox Security_2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Security_2"
         Top             =   180
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Expiring time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   25
         Tag             =   "&Password:"
         Top             =   165
         Width           =   990
      End
      Begin VB.Label lblLabels 
         Caption         =   "days"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   3
         Left            =   1560
         TabIndex        =   24
         Tag             =   "&Password:"
         Top             =   165
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   60
      TabIndex        =   18
      Top             =   3720
      Width           =   1035
      Begin VB.TextBox Libero_1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "Libero_1"
         DataSource      =   "Data1"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   420
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "1"
         Top             =   180
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "/10"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   660
         TabIndex        =   21
         Top             =   180
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Trial"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   60
         TabIndex        =   20
         Top             =   180
         Width           =   495
      End
   End
   Begin VB.TextBox Libero_2 
      DataField       =   "Libero_2"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   0
      TabIndex        =   17
      Text            =   "Libero_2"
      Top             =   5160
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "HELP"
      Top             =   4800
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2280
      TabIndex        =   16
      Text            =   "Text2"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox Txt_IDname 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   1635
      TabIndex        =   13
      Text            =   "ID Name"
      ToolTipText     =   "Send the ID_Number to the administrator to receive the password."
      Top             =   1370
      Width           =   3435
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   3690
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   90
      Width           =   1335
   End
   Begin VB.TextBox Key_Trial 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3540
      TabIndex        =   8
      Text            =   "Key_Trial"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox DECODIFICA 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Text            =   "DECODIFICA"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox txtComputerName 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "ComputerName"
      Top             =   1000
      Width           =   3435
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2940
      TabIndex        =   3
      Tag             =   "Annulla"
      Top             =   3180
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Enter"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1650
      TabIndex        =   2
      Tag             =   "OK"
      Top             =   3180
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1620
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Enter the password"
      Top             =   1740
      Width           =   3450
   End
   Begin VB.TextBox txtUserName 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "User Name"
      Top             =   630
      Width           =   3450
   End
   Begin VB.Label Label8 
      Caption         =   "Attention: this is the last trial"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1200
      TabIndex        =   27
      Top             =   3840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   555
      Left            =   60
      Top             =   60
      Width           =   615
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WP"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   15
      Tag             =   "SocietàProdotto"
      Top             =   180
      Width           =   375
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "ID Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Left            =   420
      TabIndex        =   14
      Top             =   1425
      Width           =   1140
   End
   Begin VB.Label Label3 
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   3195
      TabIndex        =   12
      Top             =   135
      Width           =   465
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmLogin_mod2.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   885
      Left            =   240
      TabIndex        =   10
      Top             =   2100
      Width           =   5160
   End
   Begin VB.Label Label1 
      Caption         =   "PC Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Left            =   420
      TabIndex        =   9
      Top             =   1065
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Index           =   1
      Left            =   420
      TabIndex        =   0
      Tag             =   "&Password:"
      Top             =   1785
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "User name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Index           =   0
      Left            =   420
      TabIndex        =   4
      Tag             =   "&Nome utente:"
      Top             =   645
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
'Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'Public OK As Boolean, MyDate, NewDate, Oggi, FirstDate, OLD_date, PASS_2, PASS_X, PASS_Y
'Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'Public Unlocked As Boolean
'Public FirstUse As String
'Public Conto2, RRR1, VOLTA, RRR, Libero1, Number As Integer
'Public Function GetSerial(str As String) As Long
'    Dim Buf$, Name$, Flags&, Length&
'    Dim Serial As Long
'    GetVolumeInformation str, Buf$, 255, Serial, Length, Flags, Name$, 255
'    GetSerial = Serial
'End Function
'Private Sub cmdAdd_Click()
'  Data1.Recordset.AddNew
'End Sub
'Private Sub cmdClose_Click()
'  Unload Me
'End Sub
'Private Sub Data1_Error(DataErr As Integer, Response As Integer)
'  MsgBox "Intercettato errore dei dati:" & Error$(DataErr)
'  Response = 0
'End Sub
'Private Sub Data1_Validate(Action As Integer, Save As Integer)
'  Select Case Action
'    Case vbDataActionMoveFirst
'    Case vbDataActionMovePrevious
'    Case vbDataActionMoveNext
'    Case vbDataActionMoveLast
'    Case vbDataActionAddNew
'    Case vbDataActionUpdate
'    Case vbDataActionDelete
'    Case vbDataActionFind
'    Case vbDataActionBookmark
'    Case vbDataActionClose
'  End Select
'  Screen.MousePointer = vbDefault
'End Sub
'Private Sub Data1_Reposition()
'  Screen.MousePointer = vbDefault
'  On Error Resume Next
'  Data1.Caption = "Record: " & (Data1.Recordset.AbsolutePosition + 1)
'    Call CheckLockedStatus(temp)
'    'temp = "unlocked"
'    If temp = "unlocked" Then
'        Frame1.Visible = False
'        Frame2.Visible = False
'        Label8.Visible = False
'        Check1.Visible = False
'        Libero_1.Text = 0
'        RRR1 = ""
'        VOLTE = ""
'        Libero_2.Text = ""
'        Libero_3.Text = ""
'        Data1.UpdateRecord
'    End If
'End Sub
'Private Sub Form_Load()
'On Error Resume Next
'    Dim sBuffer As String
'    Dim lSize As Long
'    Dim Mess
'    Dim FirstDate, Oggi As Date
'    Dim IntervalType As String
'    Dim Msg
'
'    MyDate = Date
'    Text1.Text = MyDate
'    sBuffer = Space$(255)
'    lSize = Len(sBuffer)
'    Call GetUserName(sBuffer, lSize)
'    If lSize > 0 Then
'        txtUserName.Text = Left$(sBuffer, lSize)
'    Else
'        txtUserName.Text = vbNullString
'    End If
'    Codice_1 = Left$(txtUserName, lSize)
'    upCodice_1 = UCase(Codice_1)
'    cod_test_1 = Left$(upCodice_1, 5)
'
'    sBuffer = Space$(255)
'    lSize = Len(sBuffer)
'        Call GetComputerName(sBuffer, lSize)
'    If lSize > 0 Then
'        txtComputerName.Text = Left$(sBuffer, lSize)
'    Else
'        txtUserName.Text = vbNullString
'    End If
'    Codice_2 = Left$(txtComputerName, lSize)
'    upCodice_2 = UCase(Codice_2)
'    cod_test_2 = Left$(upCodice_2, 5)
'
'    Dim objs
'    Dim obj
'    Dim WMI
'    Set WMI = GetObject("WinMgmts:")
'    Set objs = WMI.InstancesOf("Win32_BaseBoard")
'    For Each obj In objs
'       Text2.Text = obj.SerialNumber
'       board = obj.SerialNumber
'    Next
'    xSize = Len(board)
'    If xSize > 0 Then
'        Codice_3 = Left$(board, xSize)
'        upCodice_3 = UCase(Codice_3)
'        Txt_IDname.Text = upCodice_3
'    Else
'        Txt_IDname.Text = vbNullString
'    End If
'    cod_test_3 = Left$(upCodice_3, 5)
'
'    hard = GetSerial("C:\")
'    xSize = Len(hard)
'    If xSize > 0 Then
'        Codice_4 = Left$(hard, xSize)
'        upCodice_4 = UCase(Codice_4)
'        Txt_IDname.Text = upCodice_4
'    Else
'        Txt_IDname.Text = vbNullString
'    End If
'    cod_test_4 = Left$(upCodice_4, 5)
'    If cod_test_4 <> "" And cod_test_1 <> "" Then
'        IDname = Left$(upCodice_4, 2) + Left$(upCodice_2, 2)
'    ElseIf cod_test_4 <> "" And cod_test_2 <> "" Then
'        IDname = Left$(upCodice_4, 2) + Left$(upCodice_1, 2)
'    ElseIf cod_test_4 <> "" And cod_test_3 <> "" Then
'        IDname = Left$(upCodice_4, 2) + Left$(upCodice_3, 2)
'    ElseIf cod_test_4 <> "" Then
'        IDname = Left$(upCodice_4, 4)
'    ElseIf cod_test_4 = "" And cod_test_1 <> "" Then
'        IDname = Left$(upCodice_1, 4)
'    ElseIf cod_test_4 = "" And cod_test_2 <> "" Then
'        IDname = Left$(upCodice_2, 4)
'    ElseIf cod_test_4 = "" And cod_test_3 <> "" Then
'        IDname = Left$(upCodice_3, 4)
''    ElseIf cod_test_1 <> "" And cod_test_2 <> "" And cod_test_4 = "" Then
''        IDname = Left$(upCodice_1, 2) + Left$(upCodice_2, 2)
''    ElseIf cod_test_1 <> "" And cod_test_2 = "" And cod_test_4 <> "" Then
''        IDname = Left$(upCodice_1, 2) + Left$(upCodice_4, 2)
''    ElseIf cod_test_1 = "" And cod_test_2 <> "" And cod_test_4 = "" Then
''        IDname = Left$(upCodice_2, 4)
''    ElseIf cod_test_1 = "" And cod_test_2 <> "" And cod_test_4 <> "" Then
''        IDname = Left$(upCodice_2, 2) + Left$(upCodice_4, 2)
'    ElseIf cod_test_3 = "" And cod_test_1 = "" And cod_test_2 = "" And cod_test_4 = "" Then
'        IDname = "H7ZW"
'    End If
'    Codice_1 = IDname
'    Text2 = IDname
'    xSize = Len(Codice_1)
'
''CODIFICA 1
'    ALFABETO = "ABCDEFGHIJKLMNOPQRSTUVWXYZ._"
'    ReDim az1(xSize), Numero(xSize)
'    For i = 1 To xSize
'        az1(i) = Mid(Codice_1, i, 1)
'        Numero(i) = Asc(az1(i))
'        let1 = Val(Left$(Numero(i), 1)) + Val(Right$(Numero(i), 1))
'        If let1 < 1 Then
'            let1 = 1
'        ElseIf let1 > 27 Then
'            let2 = 27
'        End If
'        PASS_1 = Left$(Numero(i), 2) & Mid$(ALFABETO, let1, 1)
'        PASS_2 = PASS_2 & PASS_1
'        PASS_1 = 0
'    Next i
'    PASS_X = Right(PASS_2, 10)
'    Codice_2 = PASS_2
'    xSize = Len(Codice_2)
'    Txt_IDname.Text = PASS_2
'    PASS_2 = ""
'
''CODIFICA 2
'    ALFABETO = "ABCDEFGHIJKLMNOPQRSTUVWXYZ._"
'    ReDim az1(xSize), Numero(xSize)
'        For i = 1 To xSize
'            az1(i) = Mid(Codice_2, i, 1)
'            Numero(i) = Asc(az1(i))
'            let1 = Val(Left$(Numero(i), 1)) + Val(Right$(Numero(i), 1))
'            If let1 < 1 Then
'                let1 = 1
'            ElseIf let1 > 27 Then
'                let2 = 27
'            End If
'            PASS_1 = Left$(Numero(i), 2) & Mid$(ALFABETO, let1, 1)
'            PASS_2 = PASS_2 & PASS_1
'            PASS_1 = 0
'        Next i
''Assigning trial and unlock keys
'        PASS_X = Right(PASS_2, 10)
'        Key_Trial.Text = PASS_X
'        PASS_Y = Left(PASS_2, 10)
'        Key_Unlock.Text = PASS_Y
'
''DECODIFICA 1
'    pSize = Len(PASS_2)
'    ReDim NOME(pSize)
'    For i = 1 To pSize Step 3
'        j = j + 1
'        NOME(i) = Chr(Mid$(PASS_2, i, 2))
'        NOM1 = NOME(i)
'        NOM2 = NOM2 & NOM1
'        NOM1 = 0
'        DECODIFICA.Text = NOM2
'    Next i
'    PASS_2 = NOM2
'    NOM2 = ""
'
''DECODIFICA 2
'    pSize = Len(PASS_2)
'    ReDim NOME(pSize)
'    For i = 1 To pSize Step 3
'        j = j + 1
'        NOME(i) = Chr(Mid$(PASS_2, i, 2))
'        NOM1 = NOME(i)
'        NOM2 = NOM2 & NOM1
'        NOM1 = 0
'        DECODIFICA.Text = NOM2
'    Next i
'
'    Dim getcont As String
'    Dim f As Integer
'    f = FreeFile
'    Open App.Path + "\trialf.ini" For Input As f
'    Do
'    Line Input #f, getcont
'    Loop Until EOF(f)
'    Close f
'    txtPassword.Text = getcont
'End Sub
'Private Sub cmdCancel_Click()
'    OK = False
'    Me.Hide
'End Sub
'Private Sub cmdOK_Click()
'On Error Resume Next
'GoTo 300
''Test for trial or unlock key
'300 If txtPassword.Text = PASS_Y Then
'        SaveSetting APP_NAME, SECTION_NAME, KEY_NAME, "unlocked"
'        Unlocked = True
'        MsgBox "Valid unlock password.", , "Enter"
'        OK = True
'        Call CONTROLLO
'        Me.Hide
'    Else
'        MsgBox "Password not valid. Check again.", , "Enter"
'        txtPassword.SetFocus
'        txtPassword.SelStart = 0
'        txtPassword.SelLength = Len(txtPassword.Text)
'    End If
'End Sub
'Private Sub CONTROLLO()
'On Error Resume Next
'Dim Oggi As Date
'Dim FirstUse As String
'Const APP_NAME = "WaterProof_HX"
'Const SECTION_NAME = "Gifra_HX"
'Const KEY_NAME = "Deltagifra_HX"
'
'FirstUse = GetSetting(APP_NAME, SECTION_NAME, "First used")
'If FirstUse = "" Then
'    SaveSetting APP_NAME, SECTION_NAME, "First used", CStr(Date)
'    FirstUse = GetSetting(APP_NAME, SECTION_NAME, "First used")
'End If
'
'Call CheckLockedStatus(temp)
'If Check_pass = 1 Then
'    Dim getcont As String
'    Dim f As Integer
'    f = FreeFile
'    getcont = txtPassword.Text
'    Open App.Path + "\trialf.ini" For Output As f
'    Print #f, getcont
'    Close f
'End If
'
'If temp = "locked" Then
'    Frame2.Visible = True
'    MyDate = Date                                 'Today
'    Oggi = FirstUse                               'Starting date
'    IntervalType = "d"
'    Number = 15                                   'Expiring days
'    NewDate = DateAdd(IntervalType, Number, Oggi)
'    If MyDate < Oggi Then
'        Oggi = NewDate
'    Else
'        Oggi = MyDate
'    End If
'    Differenza = Abs(Oggi - NewDate)
'    LastDate = "01/02/2015"
'    Security_2.Text = Val(Differenza)
'
'    If Oggi >= NewDate Then
'        Msg = MsgBox("The program passed the expiring date. You should order the unlocked version of WaterProof HX. The program will be closed.")
'        End
'    ElseIf Differenza <= 15 And Differenza > 10 Then
'        Msg = MsgBox("REMINDER!  You have <15 days before the program will expire. You should order the unlimited version of WaterProof HX")
'    ElseIf Differenza <= 10 And Differenza > 5 Then
'        Msg = MsgBox("REMINDER!  You have <10 days before the program will expire. You should order the unlimited version of WaterProof HX")
'        Security_2.ForeColor = 192
'        Libero_1.ForeColor = 192
'    ElseIf Differenza <= 5 And Differenza > 2 Then
'        Msg = MsgBox("REMINDER!  You have <5 days before the program will expire. You should order the unlimited version of WaterProof HX")
'        Security_2.ForeColor = 255
'        Libero_1.ForeColor = 255
'    ElseIf Differenza <= 2 And Differenza > 0 Then
'        Msg = MsgBox("You have <2 days before the program will expire. You should order the unlimited version of WaterProof HX")
'        Security_2.ForeColor = 255
'        Libero_1.ForeColor = 255
'    End If
'
'    Lib = Val(Libero1)
'    If Conto2 - RRR1 <> Lib Or VOLTA - RRR < 1 Then
'        MsgBox ("Some parameters have been tampered. The program will be closed")
'        End
'    End If
'    If VOLTA - RRR = 10 Then
'        Label8.Caption = "Attention: this is the last trial!"
'        Label8.Visible = True
'    End If
'    If VOLTA - RRR = 11 Then
'        Label8.Caption = "Attention: You must enter the unlock password!"
'        Label8.Visible = True
'    End If
'    If VOLTA - RRR > 11 Then
'        Label8.Caption = "Attention: You must enter the unlock password!"
'        Label8.Visible = True
'        MsgBox ("You passed trial runs. The program will be closed!")
'        End
'    End If
'End If
'End Sub
