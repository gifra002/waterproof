VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WaterProof HX - Heat Exchangers Performance Monitoring"
   ClientHeight    =   4710
   ClientLeft      =   1050
   ClientTop       =   1440
   ClientWidth     =   7755
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      Height          =   4590
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7560
      Begin VB.PictureBox picLogo 
         AutoRedraw      =   -1  'True
         DragMode        =   1  'Automatic
         Height          =   1215
         Left            =   360
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   20.373
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   28.046
         TabIndex        =   1
         Top             =   2460
         Width           =   1650
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "WP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   540
         Left            =   300
         TabIndex        =   7
         Tag             =   "Societ‡Prodotto"
         Top             =   420
         Width           =   720
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Height          =   795
         Left            =   180
         TabIndex        =   8
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "WaterProof HX"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   615
         Left            =   1440
         TabIndex        =   6
         Top             =   1140
         Width           =   5055
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "HEAT EXCHANGERS THERMAL PERFORMANCE "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   660
         TabIndex        =   5
         Tag             =   "Prodotto"
         Top             =   1800
         Width           =   6720
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   210
         Left            =   6570
         TabIndex        =   4
         Tag             =   "Versione"
         Top             =   4200
         Width           =   570
      End
      Begin VB.Label lblWarning 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Program developed based on literature data to monitor the thermal performance  of industrial shell&&tubes heat transfer units."
         ForeColor       =   &H00000080&
         Height          =   735
         Left            =   3240
         TabIndex        =   2
         Tag             =   "Avviso"
         Top             =   2700
         Width           =   3135
      End
      Begin VB.Label lblCompany 
         Caption         =   "Produced by WP"
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Tag             =   "Societ‡"
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Height          =   1035
         Left            =   60
         TabIndex        =   9
         Top             =   180
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
'    lblProductName.Caption = App.Title

   Width = frmMain.Width * 0.523  ' Imposta la larghezza del form.
   Height = frmMain.Height * 0.47  ' Imposta l'altezza del form.
   Left = (frmMain.Width - Width) / 2   ' Centra il form orizzontalmente.
   Top = (frmMain.Height - Height) / 3   ' Centra il form verticalmente.
End Sub

