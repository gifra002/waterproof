VERSION 5.00
Begin VB.Form frmLogin_B 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Access"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Accesso"
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annulla"
      Height          =   360
      Left            =   2100
      TabIndex        =   5
      Tag             =   "Annulla"
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   495
      TabIndex        =   4
      Tag             =   "OK"
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1305
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   525
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1305
      TabIndex        =   3
      Top             =   135
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   248
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Tag             =   "&Password:"
      Top             =   540
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Nome utente:"
      Height          =   248
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Tag             =   "&Nome utente:"
      Top             =   150
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin_B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public OK As Boolean, MyDate, NewDate, PASS_2, PASS_X
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Sub cmdAdd_Click()
  Data1.Recordset.AddNew
End Sub
Private Sub cmdClose_Click()
  Unload Me
End Sub
Private Sub Data1_Error(DataErr As Integer, Response As Integer)
  MsgBox "Intercettato errore dei dati:" & Error$(DataErr)
  Response = 0
End Sub
Private Sub Data1_Validate(Action As Integer, Save As Integer)
  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
  End Select
  Screen.MousePointer = vbDefault
End Sub
Private Sub Data1_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  Data1.Caption = "Record: " & (Data1.Recordset.AbsolutePosition + 1)
  'Per l'oggetto tabella è necessario impostare la proprietà Index
  'al momento della creazione del Recordset e utilizzare la riga seguente
  'Data1.Caption = "Record: " & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
End Sub
Private Sub Command1_Click()
    txtPassword.Text = PASS_Y
End Sub
Private Sub Form_Load()
On Error Resume Next
    
    Dim sBuffer As String
    Dim lSize As Long
    Dim Mess
    
    MyDate = Date
    Text1.Text = MyDate
    Dim FirstDate As Date
    Dim IntervalType As String
    Dim Number As Integer
    Dim Msg
    IntervalType = "m"
    FirstDate = "01/06/2003"
    Text1.Text = MyDate
    Number = 12
    NewDate = DateAdd(IntervalType, Number, FirstDate)
    Differenza = MyDate - NewDate
    Security_2.Text = Val(Differenza)
    
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    If lSize > 0 Then
        txtUserName.Text = Left$(sBuffer, lSize)
    Else
        txtUserName.Text = vbNullString
    End If
    
    Dim tBuffer As String
    Dim xSize As Long
    tBuffer = Space$(255)
    xSize = Len(tBuffer)
    Call GetComputerName(tBuffer, xSize)
    If xSize > 0 Then
        CODICE = Left$(tBuffer, xSize)
        upCODICE = UCase(CODICE)
        txtComputerName.Text = upCODICE
    Else
        txtComputerName.Text = vbNullString
    End If

'CODIFICA
    ALFABETO = "ABCDEFGHIJKLMNOPQRSTUVWXYZ._"
    ReDim az1(xSize), NUMERO(xSize)
        For i = 1 To xSize
            az1(i) = Mid(upCODICE, i, 1)
            NUMERO(i) = Asc(az1(i))
            let1 = Val(Left(NUMERO(i), 1)) + Val(Right(NUMERO(i), 1))
            If let1 < 1 Then
                let1 = 1
            ElseIf let1 > 27 Then
                let2 = 27
            End If
            PASS_1 = NUMERO(i) & Mid(ALFABETO, let1, 1)
            PASS_2 = PASS_2 & PASS_1
            PASS_1 = 0
        Next i
        PASS_X = Right(PASS_2, 8)
        Security_1.Text = PASS_X

'DECODIFICA
    pSize = Len(PASS_2)
    ReDim NOME(pSize)
    For i = 1 To pSize Step 3
        j = j + 1
        NOME(i) = Chr(Mid(PASS_2, i, 2))
        NOM1 = NOME(i)
        NOM2 = NOM2 & NOM1
        NOM1 = 0
        DECODIFICA.Text = NOM2
    Next i
1000 End Sub
Private Sub cmdCancel_Click()
    OK = False
    Me.Hide
End Sub
Private Sub cmdOK_Click()
On Error Resume Next
    
    Differenza = MyDate - NewDate
    If Differenza > -30 And Differenza < -20 Then
        Msg = MsgBox("REMINDER!  You have <30 days before the program will expire. Call the Administrator to get the updated version of WaterProof")
    ElseIf Differenza > -20 And Differenza < -10 Then
        Msg = MsgBox("REMINDER!  You have <20 days before the program will expire. Call the Administrator to get the updated version of WaterProof")
    ElseIf Differenza > -10 And Differenza < 0 Then
        Msg = MsgBox("REMINDER!  You have <10 days before the program will expire. Call the Administrator to get the updated version of WaterProof")
    ElseIf MyDate > NewDate Then
        Msg = MsgBox("The program passed the expiring date. Call the Administrator to get the updated version of the program")
        GoTo 2000
    End If

    If txtPassword.Text = PASS_X Then
        OK = True
        Me.Hide
    Else
        MsgBox "Password not valid. Check again.", , "Enter"
        txtPassword.SetFocus
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)
    End If

2000 End Sub

