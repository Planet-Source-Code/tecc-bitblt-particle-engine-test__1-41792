Attribute VB_Name = "modParticleSys"
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

'bitblt declares
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

'point type
Public Type POINTFX
    x As Long
    y As Long
End Type

'used to determine forces on particles
 'north, south, east, and west
Public Type Pforce
    N As Double
    S As Double
    E As Double
    W As Double
End Type

'the particle type
Public Type Particle
    Life As Long              'how many frames until respawn occurs
    Position As POINTFX       'the particle's current position
    Dead As Boolean           'if the lifetime runs out, this is needed to be false
    GRAVSWNG As Byte          'GRAVITY setting, not used yet
    FORCES As Pforce          'used to govern forces on a single particle
End Type

'Bitblt mask type
Public Type MSK
    HASMASK As Boolean        'true if the particle bitmap has a mask companion
    xSrc As Single            'the source of the image (x)
    ySrc As Single            'the source of the image (x)
    Height As Long            'height of the mask image
    Width As Long             'width of the mask image
End Type

'bitblt object type
Public Type BBOBJ
    Loaded As Boolean         'signals weather the particle bitmap is loaded or not
    xSrc As Single            'Draw from this point
    ySrc As Single            'Draw from this point
    Height As Long            'Height of the particle bitmap
    Width As Long             'Width of the particle bitmap
    DCpointer As Long         'points to the bitmap device context
    OBJpointer As Long        'points to the image we will load
    bitmapPath As String      'path of the image to load
    RenderMode As RasterOpConstants  'how will we render the particle?
    Mask As MSK               'if the particle bitmap needs a mask...
    NonMaskedFrames As Byte   'If we have a dynamic particle size, we use frames to change size
End Type

'Particle system generation type
Public Enum SystemType
    Spray = &H901
    Explosion = &H902
End Enum

'type for the bitblt backbuffer
Public Type BBackBuffer
    DC As Long                'backbuffer device context
    OBJC As Long              'backbuffer compatible bitmap pointer
End Type

'the particle system type holds variables that govern how the particles will
'react in the application. all forces are defined in this system variable
Public Type ParticleSystem
    Parts() As Particle       'each particle has its own element in an array
    NumOfParticles As Integer 'maximum particles to render
    MaxLife As Integer        'maximum particle life (in frames)
    SourcePoint As POINTFX    'this is where the particles will be emmitted from
    RespawnDeadParticles As Boolean 'if the particles die, do we want to re-emmitt them?
    FORCES As Pforce          'the global forces
    RandomOffset As Pforce    'Scatters particles a bit for more realism in certain settings
    SColor As OLE_COLOR       'if we arent using graphical particles, we need to choose a pixel color for the particles
    ParticleBitmap As BBOBJ   'the bitblt object we will use for particle graphics
    TargetDC As Long          'the backfuffer will be rendered to this device context
    Alive As Boolean          'signals wheather the particle system is on or not
    Gravity As Byte           'gravity force on all particles
    SystemTYP As SystemType   'Type of particle emmitting system this is
    BackBuffer As BBackBuffer 'holds the pointer for the backbuffer
    TargetDimensions As POINTFX    'the dimensions of the object the backbuffer will render to
    DynamicSize As Boolean    'determines if particles use dynamic sized particle bitmaps
    GRaphicalParticles As Boolean  'determines if we will be rendering graphical particles
End Type

'this is the main particle system
Public pMain As ParticleSystem

'this sub will animate all the particles in a given system
Public Sub AniParticleSystem _
            (PSYS As ParticleSystem)

ReDim PSYS.Parts(PSYS.NumOfParticles) 'redimension the particle array for the set
                                      'amount of maximum particles at any given time
If PSYS.RespawnDeadParticles = True Then
KillAllParts PSYS                     'set death of particles so they will be dynamically
                                      'respawned
End If

Changetype PSYS, PSYS.SystemTYP       'change to the desired particle system emmitter type
                                      '(Explosion or spray)
LoadBB PSYS                           'Initiate the backbuffer for this system

Dim CURFRM As Long                    'used to calculate what frame we will use
                                      'when rendering dynamically sized particle graphics
Dim FrameP As POINTFX                 'the source point of the frame of a dynamic-sized bitmap
'start the animation loop
Do
    'clear the backbuffer by rendering 'whitespace' to it
    BitBlt PSYS.BackBuffer.DC, 0, 0, PSYS.TargetDimensions.x, PSYS.TargetDimensions.y, 0, 0, 0, vbWhiteness
    
    For i = 0 To UBound(PSYS.Parts) 'render particles 0 to however many there are
        With PSYS.Parts(i)
            If .Life <= PSYS.MaxLife Then 'if the particle's life hasnt run out, continue to
                                          'update its position and add to its lifetime.
            .Life = .Life + 1
            
            'adjust particle position based on forces defined within the particle
            'system variable
            .Position.x = (.Position.x + PSYS.FORCES.E) + .FORCES.E + (Rnd * PSYS.RandomOffset.E)
            .Position.x = (.Position.x - PSYS.FORCES.W) - .FORCES.W - (Rnd * PSYS.RandomOffset.W)
            .Position.y = (.Position.y - PSYS.FORCES.N) - .FORCES.N - (Rnd * PSYS.RandomOffset.N)
            .Position.y = (.Position.y + PSYS.FORCES.S) + .FORCES.S + (Rnd * PSYS.RandomOffset.S)

            'apply gravity to the particle
                .Position.y = .Position.y + (Rnd * PSYS.Gravity)

            Else
                .Dead = True 'the particle has reached its maximum lifetime, signal it to
                             'respawn
            End If
            
            If .Dead Then 'if the particle is dead, and we are respawning dead particles
                          'then respawn the particle without rendering it
                If PSYS.RespawnDeadParticles Then
                    .Dead = False
                    .Position = PSYS.SourcePoint 'reset its position
                    .Life = Rnd * PSYS.MaxLife   'reset its lifetime
                End If
            Else
                If PSYS.ParticleBitmap.Loaded Then
                    'if theres a particle bitmap (for graphical particles)
                    'the render it
                    If PSYS.ParticleBitmap.Mask.HASMASK Then
                        'render the mask, if there is one with the original
                        'particle graphic to create transparency
                    BitBlt PSYS.BackBuffer.DC, .Position.x - -(PSYS.ParticleBitmap.Width / 2), _
                    .Position.y - (PSYS.ParticleBitmap.Height / 2), PSYS.ParticleBitmap.Mask.Width _
                    , PSYS.ParticleBitmap.Mask.Height, _
                    PSYS.ParticleBitmap.DCpointer, _
                    PSYS.ParticleBitmap.Mask.xSrc, _
                    PSYS.ParticleBitmap.Mask.ySrc, vbSrcAnd
                    BitBlt PSYS.BackBuffer.DC, .Position.x, _
                    .Position.y, PSYS.ParticleBitmap.Width _
                    , PSYS.ParticleBitmap.Height, _
                    PSYS.ParticleBitmap.DCpointer, _
                    PSYS.ParticleBitmap.xSrc, _
                    PSYS.ParticleBitmap.ySrc, vbSrcPaint
                    Else
                        'theres no mask, so only render the particle graphic
                    If PSYS.DynamicSize Then
                        'for dynamic sized particles
                        
                        'the idea of dynamic sized particles stems from the fact that
                        'when fire (simulated by a particle system) goes throught it's
                        'lifespan, it looks as if its getting smaller, so we will simulate
                        'that in our particle system.
                        CURFRM = Int(.Life * 100) / PSYS.MaxLife 'calculate the percentage
                                                                 'of life the particle has
                                                                 'traveled.
                 
                        If CURFRM >= 81 Then 'life is at 81% or higher, render a Huge particle
                            FrameP.x = 4 * (PSYS.ParticleBitmap.Width / PSYS.ParticleBitmap.NonMaskedFrames)
                        GoTo re:
                        End If
                        If CURFRM >= 61 And CURFRM <= 80 Then 'life is at 61-80%, render a Large particle
                            FrameP.x = 3 * (PSYS.ParticleBitmap.Width / PSYS.ParticleBitmap.NonMaskedFrames)
                        GoTo re:
                        End If
                        If CURFRM >= 41 And CURFRM <= 60 Then
                            FrameP.x = 2 * (PSYS.ParticleBitmap.Width / PSYS.ParticleBitmap.NonMaskedFrames)
                        GoTo re:
                        End If
                        If CURFRM >= 21 And CURFRM <= 40 Then
                            FrameP.x = 1 * (PSYS.ParticleBitmap.Width / PSYS.ParticleBitmap.NonMaskedFrames)
                        GoTo re:
                        End If
                        If CURFRM >= 1 And CURFRM <= 20 Then 'life is at 1-20%, render a very small particle
                            FrameP.x = 0 * (PSYS.ParticleBitmap.Width / PSYS.ParticleBitmap.NonMaskedFrames)
                        GoTo re:
                        End If
re:

                'the framep point variable declares where the current frame is located
                'within the particle bitmap.
                    FrameP.y = 0
                    BitBlt PSYS.BackBuffer.DC, .Position.x - (PSYS.ParticleBitmap.Width / PSYS.ParticleBitmap.NonMaskedFrames), _
                    .Position.y - (PSYS.ParticleBitmap.Height / PSYS.ParticleBitmap.NonMaskedFrames), PSYS.ParticleBitmap.Width / PSYS.ParticleBitmap.NonMaskedFrames _
                    , PSYS.ParticleBitmap.Height, _
                    PSYS.ParticleBitmap.DCpointer, _
                    FrameP.x, _
                    FrameP.y, PSYS.ParticleBitmap.RenderMode
                    Else
                'theres no dynamicc sizing turned on, just render the particle!
                    BitBlt PSYS.BackBuffer.DC, .Position.x - (PSYS.ParticleBitmap.Width / 2), _
                    .Position.y - (PSYS.ParticleBitmap.Height / 2), PSYS.ParticleBitmap.Width _
                    , PSYS.ParticleBitmap.Height, _
                    PSYS.ParticleBitmap.DCpointer, _
                    PSYS.ParticleBitmap.xSrc, _
                    PSYS.ParticleBitmap.ySrc, PSYS.ParticleBitmap.RenderMode
                    End If
                    End If
                Else
                'theres no graphical bitmap loaded for rendering as a particle, simply render
                'a colored pixel in place of it.
                    SetPixelV PSYS.BackBuffer.DC, .Position.x _
                    , .Position.y, PSYS.SColor
                End If
            End If
        End With
    Next
'render the contents of the backbuffer
    BitBlt PSYS.TargetDC, 0, 0, PSYS.TargetDimensions.x, PSYS.TargetDimensions.y, PSYS.BackBuffer.DC, 0, 0, vbSrcCopy

DoEvents 'so we can input the position of the mouse and actually see whats going on!
Sleep 10 'prevents freezing and instability
Loop Until PSYS.Alive = False 'loop the engine until we kill it

DeleteDC PSYS.ParticleBitmap.DCpointer 'we need to get rid of the objects taking up
                                       'memory
DeleteObject PSYS.ParticleBitmap.OBJpointer
End Sub

Public Sub LoadBB(PSYS As ParticleSystem) 'load a backbuffer
PSYS.BackBuffer.DC = CreateCompatibleDC(GetDC(0))
PSYS.BackBuffer.OBJC = CreateCompatibleBitmap(GetDC(0), PSYS.TargetDimensions.x, PSYS.TargetDimensions.y)
SelectObject PSYS.BackBuffer.DC, PSYS.BackBuffer.OBJC
End Sub

Private Sub KillAllParts(PSYS As ParticleSystem) ' simulate all particles as dead
For i = 0 To UBound(PSYS.Parts)
    PSYS.Parts(i).Dead = True
Next
End Sub

Public Sub LoadDCBitmap(PSYS As ParticleSystem) 'load a graphical bitmap to be rendered
                                                'where a particle presides
On Error GoTo er:
If PSYS.ParticleBitmap.bitmapPath <> "" Then
    With PSYS.ParticleBitmap
    .DCpointer = CreateCompatibleDC(GetDC(0)) 'create an empty device context
    .OBJpointer = LoadImage(0, .bitmapPath, 0, 0, 0, &H10) 'load an image into an object
    SelectObject .DCpointer, .OBJpointer 'link the loaded image object to the device context
                                         'so that it points to the image
    .Loaded = True                          'signal that a graphical particle bitmap
                                            'is successfully loaded
    End With
End If
er:
End Sub

Public Sub Changetype(PSYS As ParticleSystem, TYP As SystemType, Optional explosionforce As Double = 1)
'changes emitter type
Select Case TYP
    Case SystemType.Spray
        For i = 0 To UBound(PSYS.Parts)
            With PSYS.Parts(i)
                .FORCES.E = PSYS.FORCES.E
                .FORCES.W = PSYS.FORCES.W
                .FORCES.N = PSYS.FORCES.N
                .FORCES.S = PSYS.FORCES.S
            End With
        Next
    Case SystemType.Explosion
        For i = 0 To UBound(PSYS.Parts)
            With PSYS.Parts(i)
                'we need to pull from a variety of angles to simulate
                'an explosion.
                .FORCES.E = IIf(Rnd * 10 <= 5, Rnd * explosionforce, -(Rnd * explosionforce))
                .FORCES.W = IIf(Rnd * 10 <= 5, Rnd * explosionforce, -(Rnd * explosionforce))
                .FORCES.N = IIf(Rnd * 10 <= 5, Rnd * explosionforce, -(Rnd * explosionforce))
                .FORCES.S = IIf(Rnd * 10 <= 5, Rnd * explosionforce, -(Rnd * explosionforce))
            End With
        Next
End Select
End Sub

Public Sub InitSystem(TRGT As Long, SRCp As POINTFX, DIMEN As POINTFX)
'initialize a particle system
pMain.SourcePoint = SRCp
pMain.RespawnDeadParticles = True
pMain.TargetDC = TRGT
pMain.Alive = True
pMain.TargetDimensions = DIMEN
If pMain.GRaphicalParticles Then
With pMain.ParticleBitmap
    .NonMaskedFrames = 5
    .bitmapPath = App.Path & "\" & "particle.bmp"
    .RenderMode = vbSrcPaint
    .xSrc = 0
    .ySrc = 0
    .Height = 45
    .Width = 225
    .Mask.HASMASK = False
    .Mask.xSrc = 32
    .Mask.ySrc = 32
    .Mask.Height = 32
    .Mask.Width = 32
End With
End If
LoadDCBitmap pMain
AniParticleSystem pMain
End Sub
