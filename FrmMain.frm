VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00000000&
   Caption         =   "BitBlt Particle Engine beta"
   ClientHeight    =   5715
   ClientLeft      =   1650
   ClientTop       =   1830
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   381
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   439
   Begin VB.Menu mnuEngine 
      Caption         =   "&Engine"
      Begin VB.Menu mnuSEs 
         Caption         =   "&Start Engine"
      End
      Begin VB.Menu mnuSE 
         Caption         =   "S&top Engine"
      End
      Begin VB.Menu mnuB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCFGE 
         Caption         =   "&Configure engine"
      End
      Begin VB.Menu mnub2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLEC 
         Caption         =   "&Load Engine Config..."
      End
      Begin VB.Menu mnuSEC 
         Caption         =   "&Save Engine Config..."
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private old As POINTFX
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Changetype pMain, Explosion
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'set the emitter location on the mouse point
'for extra fun
pMain.SourcePoint.x = x
pMain.SourcePoint.y = y

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
pMain.Alive = False
End Sub

Private Sub Form_Resize()
pMain.TargetDimensions.x = Me.ScaleWidth
pMain.TargetDimensions.y = Me.ScaleHeight
LoadBB pMain
End Sub

Private Sub mnuCFGE_Click()
ptclSetup.Show
End Sub

Private Sub mnuSE_Click()
pMain.Alive = False
End Sub

Private Sub mnuSEs_Click()
'load and run a default particle engine configuration
Dim IPP As POINTFX
Dim DMS As POINTFX

IPP.x = Me.ScaleWidth / 2
IPP.y = Me.ScaleHeight / 2
DMS.x = Me.ScaleWidth
DMS.y = Me.ScaleHeight

pMain.MaxLife = 100
pMain.NumOfParticles = 900
pMain.SColor = vbWhite
pMain.FORCES.E = 0
pMain.FORCES.W = 0
pMain.FORCES.N = 0
pMain.FORCES.S = 0
pMain.RandomOffset.E = 0
pMain.RandomOffset.W = 0
pMain.RandomOffset.N = 0
pMain.RandomOffset.S = 0
pMain.Gravity = 0
pMain.SystemTYP = Explosion
pMain.DynamicSize = True
pMain.GRaphicalParticles = True
InitSystem Me.hdc, IPP, DMS
End Sub

Private Sub Timer1_Timer()
pMain.SourcePoint.x = Rnd * Me.ScaleWidth
pMain.SourcePoint.y = Rnd * Me.ScaleHeight
End Sub
