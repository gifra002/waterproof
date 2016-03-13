VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "fm20.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Chart_ALL 
   Caption         =   "WaterProof SSC - Steam surface condenser - Trend charts "
   ClientHeight    =   10245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14895
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
   MDIChild        =   -1  'True
   ScaleHeight     =   10245
   ScaleWidth      =   14895
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\WaterProof_HX\Database\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_Unit"
      Top             =   9900
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\WaterProof_HX\Database\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_test"
      Top             =   9900
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Index           =   2
      Left            =   12360
      TabIndex        =   25
      Text            =   "Plant:"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Plant_N 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   12900
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   60
      Width           =   1635
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000F&
      DataField       =   "PLANT_Z"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   60
      Width           =   1755
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Index           =   1
      Left            =   4620
      TabIndex        =   22
      Text            =   "Plant:"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Index           =   0
      Left            =   6900
      TabIndex        =   21
      Text            =   "Unit:"
      Top             =   120
      Width           =   435
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      DataField       =   "UNIT_Z"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   7380
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   60
      Width           =   1755
   End
   Begin VB.ComboBox Combo_UNIT 
      ForeColor       =   &H00000080&
      Height          =   330
      ItemData        =   "Chart_ALL.frx":0000
      Left            =   10500
      List            =   "Chart_ALL.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   19
      ToolTipText     =   "Select different UNIT from the list."
      Top             =   60
      Width           =   1875
   End
   Begin VB.CheckBox Check_UNIT 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   9660
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Select first different UNIT from the list and check this button.."
      Top             =   60
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scale span"
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Left            =   13035
      TabIndex        =   10
      Top             =   4260
      Width           =   1575
      Begin VB.TextBox Text_MAX 
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   60
         TabIndex        =   13
         Text            =   "Max"
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox Check_MAX 
         Caption         =   "Click to change:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   60
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text_MIN 
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   60
         TabIndex        =   11
         Text            =   "Min"
         Top             =   1560
         Width           =   1095
      End
      Begin MSForms.SpinButton Spin_MAX 
         Height          =   615
         Left            =   1140
         TabIndex        =   17
         Top             =   540
         Width           =   375
         Size            =   "661;1085"
         Max             =   10000
      End
      Begin VB.Label Label1 
         Caption         =   "Max"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   600
         Width           =   435
      End
      Begin MSForms.SpinButton Spin_MIN 
         Height          =   615
         Left            =   1140
         TabIndex        =   15
         Top             =   1260
         Width           =   375
         Size            =   "661;1085"
         Max             =   10000
      End
      Begin VB.Label Label2 
         Caption         =   "Min"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   660
         TabIndex        =   14
         Top             =   1320
         Width           =   315
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Trends charts"
      ForeColor       =   &H00FF0000&
      Height          =   3135
      Left            =   13035
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
      Begin VB.CommandButton Comm_view 
         Caption         =   "View new period"
         Height          =   375
         Left            =   60
         TabIndex        =   4
         Top             =   2340
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTP_Fine 
         DataField       =   "Date_Fine"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   1980
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   57868289
         CurrentDate     =   37770
      End
      Begin MSComCtl2.DTPicker DTP_Inizio 
         DataField       =   "Date_Inizio"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   60
         TabIndex        =   6
         Top             =   1380
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "01/01/01"
         Format          =   57868289
         CurrentDate     =   37770
         MinDate         =   36526
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Caption         =   "Select the desired period for the trend chart "
         ForeColor       =   &H00000080&
         Height          =   675
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label38 
         Caption         =   "End date:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   60
         TabIndex        =   8
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label37 
         Caption         =   "Start date:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   60
         TabIndex        =   7
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "2D LINE / 3D BAR View"
      Height          =   315
      Left            =   1980
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   2115
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      Style           =   1
      HotTracking     =   -1  'True
      MultiSelect     =   -1  'True
      Placement       =   1
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Object.ToolTipText     =   "Close this form"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Object.ToolTipText     =   "Print this graph"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Update"
            Object.ToolTipText     =   "Update last changes"
            ImageVarType    =   2
         EndProperty
      EndProperty
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
   Begin MSChart20Lib.MSChart MSChart_ALL 
      Height          =   6555
      Left            =   120
      OleObjectBlob   =   "Chart_ALL.frx":0004
      TabIndex        =   0
      Top             =   1200
      Width           =   11415
   End
   Begin VB.Label Lab_Unit 
      Alignment       =   1  'Right Justify
      Caption         =   "Unit:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Left            =   10080
      TabIndex        =   28
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Lab_Default 
      Caption         =   "<----"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9180
      TabIndex        =   27
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Lab_NEW 
      Alignment       =   1  'Right Justify
      Caption         =   "---> "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10080
      TabIndex        =   26
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Chart_ALL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public www

Private Sub Check_1_Click()
    Call chart
End Sub
Private Sub Spin_MAX_Change()
    MM2 = Spin_MAX.Value
    Text_MAX.Text = Spin_MAX
    Call chart
End Sub
Private Sub Spin_MIN_Change()
    Min = Spin_MIN.Value
    Text_MIN.Text = Spin_MIN
    Call chart
End Sub
Private Sub Check_UNIT_Click()
If Check_UNIT = Unchecked Then
    Lab_Default.Visible = True
    Lab_NEW.Visible = False
Else
    Lab_Default.Visible = False
    Lab_NEW.Visible = True
End If
    Call chart
End Sub
Private Sub Check1_Click()
    Call chart
End Sub
Private Sub Combo_UNIT_LostFocus()
    Call chart
End Sub
Private Sub Comm_view_Click()
    Call chart
End Sub
Private Sub Form_Load()
    On Error Resume Next
   Width = frmMain.Width * 0.98 ' Imposta la larghezza del form.
   Height = frmMain.Height * 0.89     ' Imposta l'altezza del form.
   Left = 50 '(frmMain.Width - Width) / 2 ' Centra il form orizzontalmente.
   Top = 0 '(frmMain.Height - Height) / 2 ' Centra il form verticalmente.

    Dim Rs2 As Recordset
    Data2.DatabaseName = "C:\WaterProof_HX\Database\HX.mdb"
    Data2.RecordSource = "Select * From [Query_Unit]"
    Data2.Refresh
    Set Rs2 = Data2.Recordset
    If Rs2.RecordCount > 0 Then
       Do Until Rs2.EOF
            UUU1 = Data2.Recordset.Unit_NAME
                Combo_UNIT.AddItem UUU1
            Rs2.MoveNext
    Loop
    Else
       MsgBox "No Units found"
    End If

    Lab_Default.Visible = True
    Lab_NEW.Visible = False


    FFX = 0
    SKX = 0
    LMTDX = 0
    TTDX = 0
    REX = 0
    VELX = 0
    PDX = 0
    CPX = 0
    CTX = 0
    TINX = 0
    TOUTX = 0
    WFX = 0
    SFX = 0
    DUTYX = 0
    CFACX = 0
Call chart
End Sub

Private Sub TabStrip1_Click()
On Error Resume Next
    If TabStrip1.SelectedItem = "Print" Then
        Me.PrintForm
    ElseIf TabStrip1.SelectedItem = "Close" Then
        Unload Me
    ElseIf TabStrip1.SelectedItem = "Update" Then
        Data1.UpdateRecord
        Data1.Recordset.Bookmark = Data1.Recordset.LastModified
        Call chart
    End If
End Sub
Private Sub Data1_Error(DataErr As Integer, Response As Integer)
  'Posizione per il codice di gestione degli errori
  'Se si desidera ignorare gli errori, impostare come commento la riga successiva
  'Se si desidera intercettarli, aggiungere qui il codice di gestione
  MsgBox "Intercettato errore dei dati:" & Error$(DataErr)
  Response = 0  'Ignora l'errore
End Sub
Private Sub Data1_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  'Visualizza la posizione del record corrente
  'per Recordset di tipo Dynaset e Snapshot
  Data1.Caption = "Record: " & (Data1.Recordset.AbsolutePosition + 1)
  'Per l'oggetto tabella è necessario impostare la proprietà Index
  'al momento della creazione del Recordset e utilizzare la riga seguente
  'Data1.Caption = "Record: " & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1

End Sub
Private Sub chart()
On Error Resume Next
Dim CX(10000, 16), SY(10000, 16), LE(16), ET(10000), C, R, ETX(10000)

    Dim Rs1 As Recordset
    Data1.DatabaseName = "C:\WaterProof_HX\Database\HX.mdb"
    Data1.RecordSource = "Select * From [QUERY_TEST]"
    Data1.Refresh
    Set Rs1 = Data1.Recordset
    
    www = 0
    XXX = 0
    INIZIO = Data1.Recordset.Date_Inizio
    FINE = Data1.Recordset.Date_Fine
    DTP_Inizio.Value = INIZIO
    DTP_Fine.Value = FINE
    If Rs1.RecordCount > 0 Then
       Do Until Rs1.EOF

            If Check_UNIT = Checked And Combo_UNIT.Text <> "" Then
                Lab_Unit.Visible = False
                UUU1 = Combo_UNIT.Text
                PPP1 = Data1.Recordset.Plant
            Else
                Lab_Unit.Visible = True
                Lab_Default.Visible = True
                Lab_NEW.Visible = False
                UUU1 = Data1.Recordset.UNIT_Z
                Plant_N.Text = ""
            End If

'    If Rs1.RecordCount > 0 Then
'       Do Until Rs1.EOF
            UUU2 = Data1.Recordset.Unit_NAME
'                PPP1 = Data1.Recordset.Plant
            If UUU1 = UUU2 Then
                    www = www + 1
                    Plant_N.Text = Data1.Recordset.Plant
                    ETX(www) = Data1.Recordset.Date_test

                    If ETX(www) >= INIZIO And ETX(www) <= FINE Then
                           XXX = XXX + 1
                            SY(XXX, 1) = Data1.Recordset.TUBES_FF
                            SY(XXX, 2) = Data1.Recordset.SKIN_TEMP
                            SY(XXX, 3) = Data1.Recordset.MTDc
                            SY(XXX, 4) = Data1.Recordset.TTD
                            SY(XXX, 5) = Data1.Recordset.TUBES_REYNOLDS / 1000
                            SY(XXX, 6) = Data1.Recordset.TUBES_VEL
                            SY(XXX, 7) = Data1.Recordset.TUBES_PRESS_DROP
                            SY(XXX, 8) = Data1.Recordset.Press_COND * 100
                            SY(XXX, 9) = Data1.Recordset.Temp_COND
                            SY(XXX, 10) = Data1.Recordset.TUBES_TEMP_IN
                            SY(XXX, 11) = Data1.Recordset.TUBES_TEMP_OUT
                            SY(XXX, 12) = Data1.Recordset.TUBES_FLOW / 1000
                            SY(XXX, 13) = Data1.Recordset.SHELL_FLOW / 1000
                            SY(XXX, 14) = Data1.Recordset.TUBES_DUTY
                            SY(XXX, 15) = Data1.Recordset.C_Factor
                            
                            ET(XXX) = Data1.Recordset.Date_test
                        
                        End If
                    End If

                            FF = Data1.Recordset.Check_FF
                            SK = Data1.Recordset.Check_SKIN
                            LMTD = Data1.Recordset.Check_LMTD
                            TTD = Data1.Recordset.Check_TD
                            RE = Data1.Recordset.Check_RE
                            VEL = Data1.Recordset.Check_VEL
                            PD = Data1.Recordset.Check_PD
                            CP = Data1.Recordset.Check_CP
                            CT = Data1.Recordset.Check_CT
                            TIN = Data1.Recordset.Check_TIN
                            TOUT = Data1.Recordset.Check_TOUT
                            WF = Data1.Recordset.Check_W_FLOW
                            SF = Data1.Recordset.Check_S_FLOW
                            DUTY = Data1.Recordset.Check_DUTY
                            CFAC = Data1.Recordset.Check_CFAC
                            If FF = True Then
                                FFX = 1
                            End If
                            If SK = True Then
                                SKX = 1
                            End If
                            If CFAC = True Then
                                CFACX = 1
                            End If
                            If LMTD = True Then
                                LMTDX = 1
                            End If
                            If TTD = True Then
                                TTDX = 1
                            End If
                            If RE = True Then
                                REX = 1
                            End If
                            If VEL = True Then
                                VELX = 1
                            End If
                            If PD = True Then
                                PDX = 1
                            End If
                            If CP = True Then
                                CPX = 1
                            End If
                            If CT = True Then
                                CTX = 1
                            End If
                            If TIN = True Then
                                TINX = 1
                            End If
                            If TOUT = True Then
                                TOUTX = 1
                            End If
                            If WF = True Then
                                WFX = 1
                            End If
                            If SF = True Then
                                SFX = 1
                            End If
                            If DUTY = True Then
                                DUTYX = 1
                             End If
                    
                    Rs1.MoveNext
            Loop
     End If
            Rs1.MoveFirst

        MM2 = 0
        MM1 = 0
    If FFX = 1 Then
        ZZ = ZZ + 1
        For i = 1 To XXX
                MM2 = MM1
                MM1 = SY(i, 1)
                If MM1 > MM2 Then
                    MM2 = MM1
                    MM1 = MM2
                Else
                    MM1 = MM2
                End If
        Next i
    End If
    If SKX = 1 Then
        ZZ = ZZ + 1
        For i = 1 To XXX
                MM2 = MM1
                MM1 = SY(i, 2)
                If MM1 > MM2 Then
                    MM2 = MM1
                    MM1 = MM2
                Else
                    MM1 = MM2
                End If
        Next i
    End If
    If LMTDX = 1 Then
        ZZ = ZZ + 1
        For i = 1 To XXX
                MM2 = MM1
                MM1 = SY(i, 3)
                If MM1 > MM2 Then
                    MM2 = MM1
                    MM1 = MM2
                Else
                    MM1 = MM2
                End If
        Next i
    End If
    If TTDX = 1 Then
        ZZ = ZZ + 1
        For i = 1 To XXX
                MM2 = MM1
                MM1 = SY(i, 4)
                If MM1 > MM2 Then
                    MM2 = MM1
                    MM1 = MM2
                Else
                    MM1 = MM2
                End If
        Next i
    End If
    If REX = 1 Then
        ZZ = ZZ + 1
        For i = 1 To XXX
                MM2 = MM1
                MM1 = SY(i, 5)
                If MM1 > MM2 Then
                    MM2 = MM1
                    MM1 = MM2
                Else
                    MM1 = MM2
                End If
        Next i
    End If
    If VELX = 1 Then
        ZZ = ZZ + 1
        For i = 1 To XXX
                MM2 = MM1
                MM1 = SY(i, 6)
                If MM1 > MM2 Then
                    MM2 = MM1
                    MM1 = MM2
                Else
                    MM1 = MM2
                End If
        Next i
    End If
    If PDX = 1 Then
        ZZ = ZZ + 1
        For i = 1 To XXX
                MM2 = MM1
                MM1 = SY(i, 7)
                If MM1 > MM2 Then
                    MM2 = MM1
                    MM1 = MM2
                Else
                    MM1 = MM2
                End If
        Next i
    End If
    If CPX = 1 Then
        ZZ = ZZ + 1
        For i = 1 To XXX
                MM2 = MM1
                MM1 = SY(i, 8)
                If MM1 > MM2 Then
                    MM2 = MM1
                    MM1 = MM2
                Else
                    MM1 = MM2
                End If
        Next i
    End If
    If CTX = 1 Then
        ZZ = ZZ + 1
        For i = 1 To XXX
                MM2 = MM1
                MM1 = SY(i, 9)
                If MM1 > MM2 Then
                    MM2 = MM1
                    MM1 = MM2
                Else
                    MM1 = MM2
                End If
        Next i
    End If
    If TINX = 1 Then
        ZZ = ZZ + 1
        For i = 1 To XXX
                MM2 = MM1
                MM1 = SY(i, 10)
                If MM1 > MM2 Then
                    MM2 = MM1
                    MM1 = MM2
                Else
                    MM1 = MM2
                End If
        Next i
    End If
    If TOUTX = 1 Then
        ZZ = ZZ + 1
        For i = 1 To XXX
                MM2 = MM1
                MM1 = SY(i, 11)
                If MM1 > MM2 Then
                    MM2 = MM1
                    MM1 = MM2
                Else
                    MM1 = MM2
                End If
        Next i
    End If
    If WFX = 1 Then
        ZZ = ZZ + 1
        For i = 1 To XXX
                MM2 = MM1
                MM1 = SY(i, 12)
                If MM1 > MM2 Then
                    MM2 = MM1
                    MM1 = MM2
                Else
                    MM1 = MM2
                End If
        Next i
    End If
    If SFX = 1 Then
        ZZ = ZZ + 1
        For i = 1 To XXX
                MM2 = MM1
                MM1 = SY(i, 13)
                If MM1 > MM2 Then
                    MM2 = MM1
                    MM1 = MM2
                Else
                    MM1 = MM2
                End If
        Next i
    End If
    If DUTYX = 1 Then
        ZZ = ZZ + 1
        For i = 1 To XXX
                MM2 = MM1
                MM1 = SY(i, 14)
                If MM1 > MM2 Then
                    MM2 = MM1
                    MM1 = MM2
                Else
                    MM1 = MM2
                End If
        Next i
    End If
    
    If CFACX = 1 Then
        ZZ = ZZ + 1
        For i = 1 To XXX
                MM2 = MM1
                MM1 = SY(i, 15)
                If MM1 > MM2 Then
                    MM2 = MM1
                    MM1 = MM2
                Else
                    MM1 = MM2
                End If
        Next i
    End If
    
With MSChart_ALL
        .Width = 12800
        .Height = 8500
        .Top = 500
        .Left = 100
        If Check1 = Checked Then
            .chartType = VtChChartType2dLine
        ElseIf Check1 = Unchecked Then
            .chartType = VtChChartType3dBar
        End If
        .ColumnCount = XXX
        .RowCount = 15
       With MSChart_ALL.Plot
            .AngleUnit = VtAngleUnitsDegrees
            .Projection = VtProjectionTypeOrthogonal
            .Axis(VtChAxisIdY).CategoryScale.Auto = False
            .Axis(VtChAxisIdY).ValueScale.Minimum = 0
            If Check_MAX = Unchecked Then
                  MM2 = Int(MM2 * 1000 + MM2 * 200) / 1000
                  Min = 0
                  Spin_MAX.Value = MM2
                  Spin_MIN.Value = Min
            Else
                  MM2 = Spin_MAX
                  Min = Spin_MIN
            End If
                .Axis(VtChAxisIdY).ValueScale.Maximum = Format(MM2, "0.0")
                .Axis(VtChAxisIdY).ValueScale.Minimum = Min
                .Axis(VtChAxisIdY).ValueScale.MajorDivision = 5
                .Axis(VtChAxisIdY2).ValueScale.MajorDivision = 5
                .Axis(VtChAxisIdY).ValueScale.MinorDivision = 2
                .Axis(VtChAxisIdY2).ValueScale.MinorDivision = 2
                .DepthToHeightRatio = 1.5
                .WidthToHeightRatio = 1.5
                .xGap = 0.8
                .zGap = 0.8
      End With
                If FFX = 1 Then
                    For C = 1 To XXX
                       R = 1
                            .Column = C
                            .Row = R
                            If SY(C, R) > 0 Then
                                .Data = SY(C, R)
                            Else
                                .Data = 0
                            End If
                            .ColumnLabel = ET(C)
                            .RowLabel = "FF,[°C.m^2/KW (10^-3)"
                    Next C
                Else
                  .Plot.SeriesCollection.Item(1).Position.Hidden = True
                End If
               
                If SKX = 1 Then
                    For C = 1 To XXX
                       R = 2
                            .Column = C
                            .Row = R
                            If SY(C, R) > 0 Then
                                .Data = SY(C, R)
                            Else
                                .Data = 0
                            End If
                            .ColumnLabel = ET(C)
                            .RowLabel = "Skin-T,°C"
                    Next C
                Else
                  .Plot.SeriesCollection.Item(2).Position.Hidden = True
                End If
               
                If LMTDX = 1 Then
                    For C = 1 To XXX
                       R = 3
                            .Column = C
                            .Row = R
                            If SY(C, R) > 0 Then
                                .Data = SY(C, R)
                            Else
                                .Data = 0
                            End If
                            .ColumnLabel = ET(C)
                            .RowLabel = "MTDc,°C"
                    Next C
                Else
                  .Plot.SeriesCollection.Item(3).Position.Hidden = True
                End If
            
                If TTDX = 1 Then
                    For C = 1 To XXX
                       R = 4
                            .Column = C
                            .Row = R
                            If SY(C, R) > 0 Then
                                .Data = SY(C, R)
                            Else
                                .Data = 0
                            End If
                            .ColumnLabel = ET(C)
                            .RowLabel = "Appr-T,°C"
                    Next C
                Else
                  .Plot.SeriesCollection.Item(4).Position.Hidden = True
                End If
               
                If REX = 1 Then
                    For C = 1 To XXX
                       R = 5
                            .Column = C
                            .Row = R
                            If SY(C, R) > 0 Then
                                .Data = SY(C, R)
                            Else
                                .Data = 0
                            End If
                            .ColumnLabel = ET(C)
                            .RowLabel = "T-RE (10^3)"
                    Next C
                Else
                  .Plot.SeriesCollection.Item(5).Position.Hidden = True
                End If
               
                If VELX = 1 Then
                    For C = 1 To XXX
                       R = 6
                            .Column = C
                            .Row = R
                            If SY(C, R) > 0 Then
                                .Data = SY(C, R)
                            Else
                                .Data = 0
                            End If
                            .ColumnLabel = ET(C)
                            .RowLabel = "T-Vel,m/s"
                    Next C
                Else
                  .Plot.SeriesCollection.Item(6).Position.Hidden = True
                End If
               
                If PDX = 1 Then
                    For C = 1 To XXX
                       R = 7
                            .Column = C
                            .Row = R
                            If SY(C, R) > 0 Then
                                .Data = SY(C, R)
                            Else
                                .Data = 0
                            End If
                            .ColumnLabel = ET(C)
                            .RowLabel = "P-Drop,KPa"
                    Next C
                Else
                  .Plot.SeriesCollection.Item(7).Position.Hidden = True
                End If
              
                If CPX = 1 Then
                    For C = 1 To XXX
                       R = 8
                            .Column = C
                            .Row = R
                            If SY(C, R) > 0 Then
                                .Data = SY(C, R)
                            Else
                                .Data = 0
                            End If
                            .ColumnLabel = ET(C)
                            .RowLabel = "Cond-P,KPa(a)"
                    Next C
                Else
                  .Plot.SeriesCollection.Item(8).Position.Hidden = True
                End If
               
                If CTX = 1 Then
                    For C = 1 To XXX
                       R = 9
                            .Column = C
                            .Row = R
                            If SY(C, R) > 0 Then
                                .Data = SY(C, R)
                            Else
                                .Data = 0
                            End If
                            .ColumnLabel = ET(C)
                            .RowLabel = "Cond-T,°C"
                    Next C
                Else
                  .Plot.SeriesCollection.Item(9).Position.Hidden = True
                End If
               
                If TINX = 1 Then
                    For C = 1 To XXX
                       R = 10
                            .Column = C
                            .Row = R
                            If SY(C, R) > 0 Then
                                .Data = Format(Val(SY(C, R)), "0.0")
                            Text3.Text = .Data
                            Else
                                .Data = 0
                            End If
                            .ColumnLabel = ET(C)
                            .RowLabel = "T-IN,°C"
                    Next C
                Else
                  .Plot.SeriesCollection.Item(10).Position.Hidden = True
                End If
               
                If TOUTX = 1 Then
                    For C = 1 To XXX
                       R = 11
                            .Column = C
                            .Row = R
                            If SY(C, R) > 0 Then
                                .Data = Format(Val(SY(C, R)), "0.0")
                            Else
                                .Data = 0
                            End If
                            .ColumnLabel = ET(C)
                            .RowLabel = "T-OUT,°C"
                    Next C
                Else
                  .Plot.SeriesCollection.Item(11).Position.Hidden = True
                End If
            
                If WFX = 1 Then
                    For C = 1 To XXX
                       R = 12
                            .Column = C
                            .Row = R
                            If SY(C, R) > 0 Then
                                .Data = Format(Val(SY(C, R)), "0.0")
                            Else
                                .Data = 0
                            End If
                            .ColumnLabel = ET(C)
                            .RowLabel = "T-Flow,kg/h (10^3)"
                    Next C
                Else
                  .Plot.SeriesCollection.Item(12).Position.Hidden = True
                End If
            
                If SFX = 1 Then
                    For C = 1 To XXX
                       R = 13
                            .Column = C
                            .Row = R
                            If SY(C, R) > 0 Then
                                .Data = Format(Val(SY(C, R)), "0.0")
                            Else
                                .Data = 0
                            End If
                            .ColumnLabel = ET(C)
                            .RowLabel = "S-Flow,kg/h (10^3)"
                    Next C
                Else
                  .Plot.SeriesCollection.Item(13).Position.Hidden = True
                End If
            
                If DUTYX = 1 Then
                    For C = 1 To XXX
                       R = 14
                            .Column = C
                            .Row = R
                            If SY(C, R) > 0 Then
                                .Data = Format(Val(SY(C, R)), "0.0")
                            Else
                                .Data = 0
                            End If
                            .ColumnLabel = ET(C)
                            .RowLabel = "DUTY,KW"
                    Next C
                Else
                  .Plot.SeriesCollection.Item(14).Position.Hidden = True
                End If
                
                If CFACX = 1 Then
                    For C = 1 To XXX
                       R = 15
                            .Column = C
                            .Row = R
                            If SY(C, R) > 0 Then
                                .Data = SY(C, R)
                            Else
                                .Data = 0
                            End If
                            .ColumnLabel = ET(C)
                            .RowLabel = "CFAC,m3/h.kPa"
                    Next C
                Else
                  .Plot.SeriesCollection.Item(15).Position.Hidden = True
                End If
                
   End With
End Sub


