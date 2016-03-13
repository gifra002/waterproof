VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Chart_VEL 
   Caption         =   "WaterProof HX - Heat Exchangers Performance Monitoring - Tube-side flow velocity trend chart"
   ClientHeight    =   9720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14715
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
   ScaleHeight     =   9720
   ScaleWidth      =   14715
   Begin VB.ComboBox Combo_UNIT_1 
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
      Height          =   330
      Left            =   7260
      TabIndex        =   26
      ToolTipText     =   "Select different UNIT from the list."
      Top             =   60
      Width           =   1815
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   300
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_Unit"
      Top             =   9120
      Width           =   2535
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\WaterProof HX\HX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   300
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Query_test"
      Top             =   9420
      Width           =   2535
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   24
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
      Left            =   4500
      TabIndex        =   23
      Text            =   "Plant:"
      Top             =   120
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scale span"
      ForeColor       =   &H00FF0000&
      Height          =   1995
      Left            =   12840
      TabIndex        =   15
      Top             =   3780
      Width           =   1635
      Begin VB.TextBox Text_MAX 
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   60
         TabIndex        =   18
         Text            =   "Max"
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox Check_MAX 
         Caption         =   "Click to change:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   60
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text_MIN 
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   60
         TabIndex        =   16
         Text            =   "Min"
         Top             =   1560
         Width           =   1095
      End
      Begin MSForms.SpinButton Spin_MAX 
         Height          =   615
         Left            =   1140
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   600
         Width           =   435
      End
      Begin MSForms.SpinButton Spin_MIN 
         Height          =   615
         Left            =   1140
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   1320
         Width           =   315
      End
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
      Left            =   6780
      TabIndex        =   12
      Text            =   "Unit:"
      Top             =   120
      Width           =   435
   End
   Begin VB.ComboBox Combo_UNIT_2 
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
      Height          =   330
      Left            =   10380
      TabIndex        =   11
      ToolTipText     =   "Select different UNIT from the list."
      Top             =   60
      Width           =   1935
   End
   Begin VB.CheckBox Check_UNIT 
      Caption         =   "Y"
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
      Left            =   9540
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Select first different UNIT from the list and check this button.."
      Top             =   60
      Width           =   375
   End
   Begin VB.Frame Frame7 
      Caption         =   "Trends charts"
      ForeColor       =   &H00FF0000&
      Height          =   3075
      Left            =   12800
      TabIndex        =   3
      Top             =   600
      Width           =   1635
      Begin VB.CommandButton Comm_view 
         Caption         =   "View new period"
         Height          =   375
         Left            =   60
         TabIndex        =   4
         Top             =   2520
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTP_Fine 
         DataField       =   "Date_Fine"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1980
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   37770
      End
      Begin MSComCtl2.DTPicker DTP_Inizio 
         DataField       =   "Date_Inizio"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   1380
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "01/01/01"
         Format          =   16580609
         CurrentDate     =   37770
         MinDate         =   36526
      End
      Begin VB.Label Label37 
         Caption         =   "Start date:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   60
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
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
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Caption         =   "Select the desired period for the trend chart"
         ForeColor       =   &H00000080&
         Height          =   795
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "2D LINE / 3D BAR View"
      Height          =   315
      Left            =   1980
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   2055
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
   End
   Begin MSChart20Lib.MSChart MSChart_VEL 
      Height          =   6555
      Left            =   120
      OleObjectBlob   =   "Chart_VEL.frx":0000
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
      Left            =   9960
      TabIndex        =   25
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Lab_Default 
      Caption         =   "<----"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9060
      TabIndex        =   14
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Lab_NEW 
      Alignment       =   1  'Right Justify
      Caption         =   "---> "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9960
      TabIndex        =   13
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "Chart_VEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public www
Private Sub Combo_UNIT_1_Lostfocus()
If Check_UNIT = Unchecked Then
    Lab_Unit.Visible = True
    Lab_Default.Visible = True
    Lab_NEW.Visible = False
Else
    Lab_Unit.Visible = False
    Lab_Default.Visible = False
    Lab_NEW.Visible = True
End If
    Call chart
End Sub
Private Sub Combo_UNIT_2_Lostfocus()
If Check_UNIT = Unchecked Then
    Lab_Unit.Visible = True
    Lab_Default.Visible = True
    Lab_NEW.Visible = False
Else
    Lab_Unit.Visible = False
    Lab_Default.Visible = False
    Lab_NEW.Visible = True
End If
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
    Lab_Unit.Visible = True
    Lab_Default.Visible = True
    Lab_NEW.Visible = False
Else
    Lab_Unit.Visible = False
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
   Width = frmMain.Width * 0.98 ' Imposta la larghezza del form.
   Height = frmMain.Height * 0.89     ' Imposta l'altezza del form.
   Left = 50 '(frmMain.Width - Width) / 2 ' Centra il form orizzontalmente.
   Top = 0 '(frmMain.Height - Height) / 2 ' Centra il form verticalmente.

    Dim Rs2 As Recordset
    Data2.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
    Data2.RecordSource = "Select * From [Query_Unit]"
    Data2.Refresh
    Set Rs2 = Data2.Recordset
    If Rs2.RecordCount > 0 Then
        Do Until Rs2.EOF
            UUU1 = Data2.Recordset.Unit_name
            Combo_UNIT_1.AddItem UUU1
            Combo_UNIT_2.AddItem UUU1
            Rs2.MoveNext
        Loop
    Else
       MsgBox "No Units found"
    End If

    Lab_Default.Visible = True
    Lab_NEW.Visible = False
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
Dim CX(10000, 12), SY(10000, 12), LE(12), ET(10000), c, r, ETX(10000)
    Dim Rs1 As Recordset
    Data1.DatabaseName = "C:\Program Files\WaterProof HX\HX.mdb"
    Data1.RecordSource = "Select * From [QUERY_TEST]"
    Data1.Refresh
    Set Rs1 = Data1.Recordset
    CONTO = Rs1.RecordCount
    www = 0
    XXX = 0
    INIZIO = DTP_Inizio.Value
    FINE = DTP_Fine
    If Check_UNIT = Unchecked And Combo_UNIT_1.Text <> "" Then
        Lab_Unit.Visible = True
        UUU1 = Combo_UNIT_1.Text
    ElseIf Check_UNIT = Checked And Combo_UNIT_2.Text <> "" Then
        Lab_Unit.Visible = False
        UUU1 = Combo_UNIT_2.Text
    End If
    If Rs1.RecordCount > 0 Then
       Do Until Rs1.EOF
            UUU2 = Data1.Recordset.Unit_name
            If UUU1 = UUU2 Then
            CONTO = Rs1.RecordCount
                www = www + 1
                Plant_N.Text = Data1.Recordset.Plant
                ETX(www) = Data1.Recordset.date_test
                    If ETX(www) >= INIZIO And ETX(www) <= FINE Then
                        XXX = XXX + 1
                        SY(XXX, 1) = Data1.Recordset.TUBES_VEL
                        C_DES = Data1.Recordset.CHECK_DESIGN
                        If C_DES = -1 Then
                            ET(XXX) = "Design"
                        Else
                            ET(XXX) = Data1.Recordset.date_test
                        End If
                    End If
            End If
            Rs1.MoveNext
    Loop
    Else
       MsgBox "No Units found"
    End If
    Data1.Recordset.MoveFirst

With MSChart_VEL
        .Width = 12700
        .Height = 8500
        .Top = 500
        .Left = 100
        If Check1 = Checked Then
            .chartType = VtChChartType2dLine
        ElseIf Check1 = Unchecked Then
            .chartType = VtChChartType3dBar
        End If
        .ColumnCount = XXX
        .RowCount = 1
             With MSChart_VEL.Plot
                  .AngleUnit = VtAngleUnitsDegrees
                  .Projection = VtProjectionTypeOrthogonal
                  .Axis(VtChAxisIdY).CategoryScale.Auto = False
                  .Axis(VtChAxisIdY).ValueScale.Minimum = 0
              
              MM2 = 0
              MM1 = 0
              For i = 1 To XXX
                      For j = 1 To 1
                      MM2 = MM1
                      MM1 = SY(i, j)
                      If MM1 > MM2 Then
                          MM2 = MM1
                          MM1 = MM2
                      Else
                          MM1 = MM2
                      End If
                      Next j
              Next i
              
              
              If Check_MAX = Unchecked Then
                    MM2 = Int(MM2 / 5 + 1) * 5
                    Min = 0
                    Spin_MAX.Value = MM2
                    Spin_MIN.Value = Min
                Else
                    MM2 = Spin_MAX
                    Min = Spin_MIN
              End If
                  .Axis(VtChAxisIdY).ValueScale.Maximum = Format(MM2, "0.00")
                  .Axis(VtChAxisIdY).ValueScale.Minimum = Min
                  .Axis(VtChAxisIdY).ValueScale.MajorDivision = 5
                  .Axis(VtChAxisIdY).ValueScale.MinorDivision = 2
                  .DepthToHeightRatio = 1.5
                  .WidthToHeightRatio = 1.5
                  .xGap = 0.8
                  .zGap = 0.8
            End With
           
       LE(1) = "VEL"
       
        For c = 1 To XXX
           r = 1
                .Column = c
                .Row = r
                If SY(c, r) > 0 Then
                    .Data = SY(c, r)
                Else
                    .Data = 0
                End If
                .ColumnLabel = ET(c)
                .RowLabel = LE(r)
        Next c
  
   End With
End Sub
Private Sub Check_MAX_click()
    Call chart
End Sub

