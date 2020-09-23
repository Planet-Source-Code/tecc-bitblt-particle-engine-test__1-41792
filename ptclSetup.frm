VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ptclSetup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Engine Settings"
   ClientHeight    =   5355
   ClientLeft      =   2385
   ClientTop       =   2190
   ClientWidth     =   5145
   BeginProperty Font 
      Name            =   "Tahoma"
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
   ScaleHeight     =   5355
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   30
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply and restart engine"
      Height          =   375
      Left            =   2880
      TabIndex        =   29
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      Caption         =   "Other:"
      Height          =   855
      Left            =   60
      TabIndex        =   27
      Top             =   3960
      Width           =   4995
      Begin VB.CheckBox CHK_GRAPH 
         Caption         =   "Use Graphical Particles"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Top             =   540
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox CHK_DFS 
         Caption         =   "Dynamic Frame Size"
         Height          =   255
         Left            =   180
         TabIndex        =   28
         Top             =   240
         Value           =   1  'Checked
         Width           =   3555
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "General Settings"
      Height          =   1635
      Left            =   60
      TabIndex        =   17
      Top             =   60
      Width           =   4995
      Begin MSComDlg.CommonDialog col1 
         Left            =   1980
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox ptclColor 
         Height          =   255
         Left            =   2340
         ScaleHeight     =   195
         ScaleWidth      =   2415
         TabIndex        =   32
         Top             =   900
         Width           =   2475
      End
      Begin VB.ComboBox SST 
         Height          =   315
         Left            =   2340
         TabIndex        =   26
         Text            =   "Explosion"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtPartLife 
         Height          =   285
         Left            =   2340
         TabIndex        =   21
         Text            =   "0"
         Top             =   540
         Width           =   975
      End
      Begin VB.TextBox txtMaxPart 
         Height          =   285
         Left            =   2340
         TabIndex        =   20
         Text            =   "0"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "System Spawn Type:"
         Height          =   255
         Left            =   180
         TabIndex        =   25
         Top             =   1260
         Width           =   2475
      End
      Begin VB.Label Label11 
         Caption         =   "Particle Color:"
         Height          =   255
         Left            =   180
         TabIndex        =   22
         Top             =   900
         Width           =   1995
      End
      Begin VB.Label Label10 
         Caption         =   "Maximum particle life:"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   600
         Width           =   1995
      End
      Begin VB.Label Label9 
         Caption         =   "Maximum particles at once:"
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   300
         Width           =   1995
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Particle Forces"
      Height          =   2175
      Left            =   60
      TabIndex        =   0
      Top             =   1740
      Width           =   4995
      Begin VB.TextBox Grav 
         Height          =   285
         Left            =   1260
         TabIndex        =   24
         Text            =   "0"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox PFS_W 
         Height          =   285
         Left            =   3660
         TabIndex        =   16
         Text            =   "0"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox PFS_E 
         Height          =   285
         Left            =   3660
         TabIndex        =   15
         Text            =   "0"
         Top             =   900
         Width           =   975
      End
      Begin VB.TextBox PFS_S 
         Height          =   285
         Left            =   3660
         TabIndex        =   14
         Text            =   "0"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox PFS_N 
         Height          =   285
         Left            =   3660
         TabIndex        =   13
         Text            =   "0"
         Top             =   300
         Width           =   975
      End
      Begin VB.TextBox PF_W 
         Height          =   285
         Left            =   1260
         TabIndex        =   8
         Text            =   "0"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox PF_E 
         Height          =   285
         Left            =   1260
         TabIndex        =   7
         Text            =   "0"
         Top             =   900
         Width           =   975
      End
      Begin VB.TextBox PF_S 
         Height          =   285
         Left            =   1260
         TabIndex        =   6
         Text            =   "0"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox PF_N 
         Height          =   285
         Left            =   1260
         TabIndex        =   5
         Text            =   "0"
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Gravity:"
         Height          =   255
         Left            =   180
         TabIndex        =   23
         Top             =   1740
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "West Scatter:"
         Height          =   255
         Left            =   2580
         TabIndex        =   12
         Top             =   1260
         Width           =   2235
      End
      Begin VB.Label Label7 
         Caption         =   "East Scatter:"
         Height          =   255
         Left            =   2580
         TabIndex        =   11
         Top             =   960
         Width           =   2235
      End
      Begin VB.Label Label6 
         Caption         =   "South Scatter:"
         Height          =   255
         Left            =   2580
         TabIndex        =   10
         Top             =   660
         Width           =   2235
      End
      Begin VB.Label Label5 
         Caption         =   "North Scatter:"
         Height          =   255
         Left            =   2580
         TabIndex        =   9
         Top             =   360
         Width           =   2235
      End
      Begin VB.Label Label4 
         Caption         =   "Pull to West:"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   1260
         Width           =   2235
      End
      Begin VB.Label Label3 
         Caption         =   "Pull to East:"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   960
         Width           =   2235
      End
      Begin VB.Label Label2 
         Caption         =   "Pull to South:"
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   660
         Width           =   2235
      End
      Begin VB.Label Label1 
         Caption         =   "Pull to North:"
         Height          =   255
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   2235
      End
   End
   Begin VB.CheckBox CHK_APL 
      Caption         =   "Keep open"
      Height          =   255
      Left            =   60
      TabIndex        =   31
      Top             =   4980
      Width           =   1875
   End
End
Attribute VB_Name = "ptclSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApply_Click()
Dim IPP As POINTFX
Dim DMS As POINTFX

IPP.x = FrmMain.ScaleWidth / 2
IPP.y = FrmMain.ScaleHeight / 2
DMS.x = FrmMain.ScaleWidth
DMS.y = FrmMain.ScaleHeight

With pMain
    .NumOfParticles = Val(txtMaxPart.Text)
    .MaxLife = Val(txtPartLife.Text)
    .SColor = ptclColor.BackColor
    Select Case SST.Text
        Case "Explosion"
            .SystemTYP = Explosion
        Case "Spray"
            .SystemTYP = Spray
        Case Else
            .SystemTYP = Explosion
    End Select
    .FORCES.N = Val(PF_N.Text)
    .FORCES.E = Val(PF_E.Text)
    .FORCES.S = Val(PF_S.Text)
    .FORCES.W = Val(PF_W.Text)
    .RandomOffset.N = Val(PFS_N.Text)
    .RandomOffset.E = Val(PFS_E.Text)
    .RandomOffset.S = Val(PFS_S.Text)
    .RandomOffset.W = Val(PFS_W.Text)
    .Gravity = Abs(Val(Grav.Text))
    If CHK_DFS.Value = 1 Then
        .DynamicSize = True
    Else
        .DynamicSize = False
    End If
    If CHK_GRAPH.Value = 1 Then
        .GRaphicalParticles = True
    Else
        .GRaphicalParticles = False
    End If
End With
InitSystem FrmMain.hdc, IPP, DMS
If Me.CHK_APL.Value = 0 Then
    Unload Me
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
SST.AddItem "Explosion"
SST.AddItem "Spray"
On Error Resume Next
txtMaxPart.Text = UBound(pMain.Parts)
txtPartLife.Text = pMain.MaxLife
ptclColor.BackColor = pMain.SColor

Select Case pMain.SystemTYP
    Case SystemType.Explosion
        SST.Text = "Explosion"
    Case SystemType.Spray
        SST.Text = "Spray"
End Select

With pMain.FORCES
    PF_N.Text = .N
    PF_E.Text = .E
    PF_S.Text = .S
    PF_W.Text = .W
End With

With pMain.RandomOffset
    PFS_N.Text = .N
    PFS_E.Text = .E
    PFS_S.Text = .S
    PFS_W.Text = .W
End With

If pMain.DynamicSize = True Then
    CHK_DFS.Value = 1
Else
    CHK_DFS.Value = 0
End If

If pMain.ParticleBitmap.bitmapPath = "" Then
    CHK_GRAPH.Value = 0
Else
    CHK_GRAPH.Value = 1
End If


Grav.Text = Abs(pMain.Gravity)

End Sub

Private Sub ptclColor_DblClick()
col1.ShowColor
ptclColor.BackColor = col1.Color
End Sub
