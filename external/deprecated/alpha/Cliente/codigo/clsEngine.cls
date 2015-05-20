VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEngineMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'************************************************************************************
' clsEngine.cls - Argentum Direct3D 8 Graphic and Sound Engine - v1.0 (First Release)
'
' Graphic & Sound Engine.
' Parra Engine v1.1 adapted for Argentum Online 1.12.2
' Designed and implemented by Parra (parra_real_93@hotmail.com)
'************************************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

'*****************************************************************
'Contributors History
'   When releasing modifications to this source file please add your
'   date of release, name, email, and any info to the top of this list.
'   Follow this template:
'    XX/XX/200X - Your Name Here (Your Email Here)
'       - Your Description Here
'       Sub Release Contributors:
'           XX/XX/2003 - Sub Contributor Name Here (SC Email Here)
'               - SC Description Here
'*****************************************************************
'
'08/09/2009 - Parra - (parra_real_93@hotmail.com)
'   - First Release
'
'**************************************************************************

'***************************
'Required Externals
'***************************
'Reference to dx8vb.dll
'   - URL: http://www.microsoft.com/directx
'***************************

''
' Declarations of all DirectX Graphics controls using Direct3D 8 Graphics
' Designed and implemented by Parra
Option Explicit
''
' Directx8 Objects
Dim DirectX     As DirectX8
Dim D3D         As Direct3D8
Dim D3DDevice   As Direct3DDevice8
Dim D3DX        As D3DX8

''
' Directx8 Manage Window
Dim D3DWindow As D3DPRESENT_PARAMETERS

''
' Directx8 Display Mode
Dim DispMode As D3DDISPLAYMODE

''
' Directx8 Device caps
Dim DevCaps As D3DCAPS8

''
' Directx8 Fonts
Private Fonts_Path  As String

Private Type FontInfo
    D3DFont     As D3DXFont
    desc        As IFont
    Format      As New StdFont
End Type: Private Font() As FontInfo

''
' Render Region Controller
Dim regRender    As RECT
Dim regRenderDX  As D3DRect

''
' FPS Counter
Private fpsLastCheck   As Long
Private fpsCounter     As Integer
Private fpsValue       As Integer

''
' Texture Struct
Private Type structGraphic
        FileNum    As Long
        Width  As Integer
        Height  As Integer
        lastAccess As Long
        D3DTexture As Direct3DTexture8
End Type: Private Graphic() As structGraphic

''
' Used to cache the Texture
Dim LastArrayIndex As Integer

''
' Memory bytes used in textures
Dim usedBytes As Long

''
' Max memory bytes used in textures
Dim maxUsedBytes As Long

'''
'' Engine Constants
'''

''
' Tile size view
Private WindowTileWidth  As Integer
Private WindowTileHeight As Integer

''
' Half Tile size view
Private HalfWindowTileWidth  As Integer
Private HalfWindowTileHeight As Integer

''
' Main View Offset
Private MainViewTop  As Integer
Private MainViewLeft As Integer

''
' Render buffer size
Private TileBufferSize As Integer

''
' Buffer offset pixels
Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

''
' Tile size in pixels
Private TilePixelHeight As Integer
Private TilePixelWidth  As Integer

''
' Switch On/Off the Game Engine
Public GameRun As Boolean

''
' Number of pixels the engine scrolls per frame (must divide evenly into pixels per tile)
Private ScrollPixelsPerFrameX As Integer
Private ScrollPixelsPerFrameY As Integer

''
' Mouse Position in Tiles
Private MouseTileX As Byte
Private MouseTileY As Byte

''
' Engine Base Speed
Dim engineBaseSpeed As Single

''
' Defines 2D Position
Private Type Point
    x As Long: y As Long
End Type

'''
' Grh Declarations
'''

''
' Points to a GrhData Structure and keeps the animation
Private Type Grh
    GrhIndex        As Integer
    FrameCounter    As Single
    Speed           As Single
    Started         As Byte
    Loops           As Integer
End Type

''
' Contains information about where you can find a size and animation grh
Private Type GrhData
    sX          As Integer
    sY          As Integer
    
    FileNum     As Long
    
    pixelWidth  As Integer
    pixelHeight As Integer
    
    TileWidth   As Single
    TileHeight  As Single
    
    NumFrames   As Integer
    Frames()    As Long
    
    Speed       As Single
End Type: Private GrhData() As GrhData

''
' Body List
Private Type BodyData
    Walk(Heading.North To Heading.West) As Grh
    HeadOffset As Point
End Type: Private BodyData() As BodyData

Private Type BodyList
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

''
' Head List
Private Type HeadData
    Head(Heading.North To Heading.West) As Grh
End Type: Private HeadData() As HeadData: Private HelmetData() As HeadData

Private Type HeadList
    Head(1 To 4) As Integer
End Type

''
' Fx List
Private Type FxData
    Animation As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type: Private FxData() As FxData

Private Type FxList
    Animacion As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type

''
' Weapon Anim List
Private Type WeaponAnimData
    WeaponWalk(Heading.North To Heading.West) As Grh
End Type: Private WeaponAnimData() As WeaponAnimData

''
' Shield Anim List
Private Type ShieldAnimData
    ShieldWalk(Heading.North To Heading.West) As Grh
End Type: Private ShieldAnimData() As ShieldAnimData

'''
' Map Declarations
'''

''
' Map sizes in tiles
Const XMaxMapSize As Byte = 100
Const XMinMapSize As Byte = 1
Const YMaxMapSize As Byte = 100
Const YMinMapSize As Byte = 1

''
' Map borders
Private MinXBorder As Byte: Private MaxXBorder As Byte
Private MinYBorder As Byte: Private MaxYBorder As Byte

''
' World Position Point
'Private Type WorldPos
'    Position As Point
'    Map      As Integer
'End Type

''
' Object Info
'Private Type ObjInfo
'    OBJIndex As Integer
'    Amount As Integer
'End Type

''
' Map Tile Info
Private Type mapBlock
    Graphic(1 To 4) As Grh
    CharIndex       As Integer
    
    ObjGrh          As Grh
    
    'not use..
    'ObjInfo         As ObjInfo
    'TileExit        As WorldPos
    
    Blocked         As Byte
    
    Trigger         As Integer
End Type: Private mapData() As mapBlock

''
' Map Area Controller
Private MinLimit As Point
Private MaxLimit As Point

''
' Map Ambient
Private mapRain()         As Byte

Public bRain              As Boolean 'está raineando?
Public RainBufferIndex    As Long

Public bFogata            As Boolean
Public FogataBufferIndex  As Long

Public bTecho             As Boolean 'hay techo?



'''
' Character Declarations
'''

Private Type CharacterAparence
     Heading            As Heading
     Pos                As Point
     
     scrollDirectionX   As Integer
     scrollDirectionY   As Integer
    
     Moving             As Byte
     MoveOffsetX        As Single
     MoveOffsetY        As Single
     
     Body        As BodyData: iBody       As Integer
     Head        As HeadData: iHead       As Integer
     Helmet      As HeadData
     Arma        As WeaponAnimData
     Escudo      As ShieldAnimData
     Fx          As Grh: FxIndex          As Integer
    
     Criminal    As Byte
    
     Name        As String
    
     Pie         As Boolean
     Muerto      As Boolean
     Invisible   As Boolean
     Priv        As Byte
     
     Active       As Byte

End Type

Private Const MAX_CHARS As Integer = 10000
Private lastChar As Integer
Private charList(1 To MAX_CHARS) As CharacterAparence


'''
' User Declarations
'''

'Status del user
Private UserMoving    As Byte
Private UserPos       As Point 'Posicion
Private AddtoUserPos  As Point 'Si se mueve

Public UserMap As Integer
Public UserMapName As String

Public UserCharIndex  As Integer


'[CODE 001]:MatuX
Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum

Public IsPlaying As Byte

''
' 64bit system counter
Dim timerElapsedTime    As Single
Dim timerTicksPerFrame  As Single

''
' Very percise API Counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

''
' Gets number of ticks since windows started
Private Declare Function GetTickCount Lib "kernel32" () As Long

''
' BitBlt Declarations
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

''
' Declarations of all audio controls using DirectSound8 and DirectMusic8
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' Modified and adapted for DirectX8 by Parra

'''
' General Audio Declarations
'''

''
' Set to true once the object is created to prevent multiple initializations of the same object.
Dim Inited As Boolean

'''
' General Audio_Music Declarations
'''

''
' DirectMusic's Performance object
Dim Performance As DirectMusicPerformance8

''
' Currently loaded segment
Dim Segment As DirectMusicSegment8

''
' The one and only DirectMusic Loader
Dim Loader As DirectMusicLoader8

''
' State of the currently loaded segment
Dim SegState As DirectMusicSegmentState8

''
' States wether music is supported by hardware or not
Dim Music As Boolean

''
' Directory from which to load music files by default
Dim Music_Path As String

''
' Last loaded music file. Used to be able to load it when music gets activated
Dim Music_CurrentFile As String

''
' States wether music is currently activated or not
Dim mActivated As Boolean

'''
' General Audio_Sound Declarations
'''

''
' Maximum number of sound buffers to hold in memory at any time.
Private Const Sound_MaxBuffers As Byte = 30

''
' The distance at which the sound is inaudible.
Private Const SOUND_MAX_DISTANCE_TO_SOURCE As Integer = 150

''
' The frequency variation coefficient for Doppler effect simulation.
Private Const SOUND_DELTA_FQ As Integer = 75

''
' Custom sound buffer structure.
'
' @param    FileName    The name of the file loaded in this buffer
' @param    looping     Set to true if the sound will loop until stopped
' @param    X           The X position of the sound in map coordinates
' @param    Y           The Y position of the sound in map coordinates
' @param    normalFq    The frequency of the sound when not affected by Doppler effect.
' @param    Buffer      The DirectSound_Buffer object to wich the file is loaded

Private Type Sound_Buffer
    FileName As String
    looping As Boolean
    x As Byte
    y As Byte
    normalFq As Long
    Buffer As DirectSoundSecondaryBuffer8
End Type

''
' States how to set a sound's looping state.
'
' @param    Default     If wave is in memory use it's looping setting, otherwise false
' @param    Disabled    Never loop
' @param    Enabled     Loop until stopped

Public Enum Sound_LoopStyle
    Default = 0
    Disabled = 1
    Enabled = 2
End Enum

''
' The one and only DirectSound object
Dim DirectSound As DirectSound8

''
' The DirectSound enumeration
Dim DirectSoundEnum As DirectSoundEnum8

''
' Array of all existing sound buffers
Dim Sound_Buffers(1 To Sound_MaxBuffers) As Sound_Buffer

''
' States wether sound is supported by hardware or not
Dim Sound As Boolean

''
' Directory from which to load sound files by default
Dim Sound_Path As String

''
' The volume for sounds (FXs)
Dim Sound_Vol As Long

''
' States wether sound is currently activated or not
Dim sActivated As Boolean

'''
' General Audio_SoundEffects Declarations
'''

''
' States wether sound effects are currently activated or not
Dim sEffectsActivated As Boolean

Private Type Sound_Pointer
    x As Integer
    y As Integer
End Type

''
' States the last position where the listener was in the X and Y axis
Dim SoundEffects_lastPos As Sound_Pointer


Private Type TLVERTEX    'NO TOCAR POR NADA EN EL MUNDO
    v As D3DVECTOR
    rhw As Single       'NO TOCAR POR NADA EN EL MUNDO
    'normal As D3DVECTOR
    color As Long       'NO TOCAR POR NADA EN EL MUNDO
    tu As Single        'NO TOCAR POR NADA EN EL MUNDO
    tv As Single        'NO TOCAR POR NADA EN EL MUNDO
End Type                'NO TOCAR POR NADA EN EL MUNDO


'***************************************************
' Default Functions of Class
'***************************************************

''
' Constructor. Performs no action at all.

Private Sub Class_Initialize()


End Sub

''
' Destructor. Releases all created objects assuring no memory-leaks.

Private Sub Class_Terminate()
    
    

    ' Destroy Audio Buffer and DirectX Objects
      Audio_Destroy
    
    ' Destroy DirectX Fonts
      Fonts_Destroy
    
    ' Destroy Direct3D 8 Objects
      DirectX_Destroy
    
    ' Destroy Textures
      Graphics_Destroy
      
End Sub


'***************************************************
' DirectX Graphics & Sound Functions
'***************************************************

Public Sub DirectX_Initialize(DeviceForm As Form)

    ' Create the root D3D objects
    Set DirectX = New DirectX8
    Set D3D = DirectX.Direct3DCreate()
    Set D3DX = New D3DX8

    Dim D3DWindow As D3DPRESENT_PARAMETERS
    Dim DispMode As D3DDISPLAYMODE
    
    Dim Windowed As Byte: Windowed = 0
    
'    If Windowed Then
'        DispMode.Width = 800
'        DispMode.Height = 600
'        DispMode.Format = D3DFMT_R5G6B5
'
'        D3DWindow.FullScreen_RefreshRateInHz = D3DPRESENT_RATE_DEFAULT
'        D3DWindow.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
'
'        D3DWindow.Windowed = False
'        D3DWindow.BackBufferCount = 1
'        D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
'    Else
'
'        D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
'        D3DWindow.Windowed = True
'        D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
'    End If
'

    With D3DWindow
        D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
        .Windowed = 1
        .SwapEffect = D3DSWAPEFFECT_COPY
        .hDeviceWindow = frmMain.hWnd
        .BackBufferWidth = DispMode.Width
        .BackBufferHeight = DispMode.Height
        .BackBufferFormat = DispMode.Format
        .MultiSampleType = D3DMULTISAMPLE_NONE
    End With
      
  '###################################
  '## CHECK THE DEVICE CAPABILITIES ##
  '###################################
    
    Dim DevCaps As D3DCAPS8
    
    D3D.GetDeviceCaps D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DevCaps
   
    If Err.Number = D3DERR_INVALIDDEVICE Then
        'We couldn't get data from the hardware device - probably doesn't exist...
        D3D.GetDeviceCaps D3DADAPTER_DEFAULT, D3DDEVTYPE_REF, DevCaps
        Err.Clear
    End If

    'Set the D3DDevices
    If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DeviceForm.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
    
    DeviceForm.Visible = False
    
    DirectX_RenderStates_Reset

    'Initialize Audio
    Audio_Initialize frmMain.hWnd, App.path & "\WAV\", App.path & "\MIDI\"
    
    'Enable / Disable audio
    Music_Activated = True: Sound_Activated = True: SoundEffects_Activated = True

    ' Initialize Textures
    Graphics_Initialize &O200
    
    'Initialize Directx Fonts
    Fonts_Initialize App.path & "\Fonts\fonts.ini"
    
End Sub

Private Sub DirectX_Destroy()
'***************************************************
'Author: •Parra
'Last Modification: 08/09/09
'Destroys all Direct3D 8 Objects
'***************************************************

    'Set no texture in the device to avoid memory leaks
    If Not D3DDevice Is Nothing Then
        D3DDevice.SetTexture 0, Nothing
    End If
    
    'Destroy objects
    Set DirectX = Nothing
    Set D3D = Nothing
    Set D3DX = Nothing
    Set D3DDevice = Nothing
    
End Sub

Public Sub DirectX_Scene_Begin()

    With D3DDevice
                    
        .Clear 0, ByVal 0&, D3DCLEAR_TARGET, 0, 1#, 0
        .BeginScene
    
    End With
    
End Sub
Public Sub DirectX_Scene_BeginCustom(ClearRectCount As Long, D3DRect As D3DRect, Flags As CONST_D3DCLEARFLAGS)

    With D3DDevice
                    
        .Clear ClearRectCount, D3DRect, D3DCLEAR_TARGET, 0, 1#, 0
        .BeginScene
    
    End With

End Sub
Public Sub DirectX_Scene_End()

    With D3DDevice
                    
        'Flip the backbuffer to the screen
        .Present regRenderDX, regRenderDX, frmMain.hWnd, ByVal 0&
        .EndScene
    
    End With
    
End Sub
Public Sub DirectX_Scene_EndCustom(RECT As RECT, hWnd As Long)

    With D3DDevice
                    
        'Flip the backbuffer to the screen
        .Present RECT, RECT, hWnd, ByVal 0&
        .EndScene
    
    End With
    
End Sub
Public Function Engine_FToDW(f As Single) As Long
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    Call DXCopyMemory(Engine_FToDW, f, 4)
End Function
Private Sub DirectX_RenderStates_Reset()
'
'    With D3DDevice
'
'        'Set the shader to be used
'        D3DDevice.SetVertexShader (D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE) 'D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
'
'        'Set the render states
'        .SetRenderState D3DRS_LIGHTING, False
'        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
'        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
'        .SetRenderState D3DRS_ALPHABLENDENABLE, True
'
'        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
'        .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
'        .SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_CURRENT
'
'        'Set the texture stage stats (filters)
'        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
'        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
'
'    End With
    With D3DDevice
        Call .SetVertexShader((D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE))
        Call .SetRenderState(D3DRS_LIGHTING, 0)
        Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
        Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        Call .SetRenderState(D3DRS_ALPHABLENDENABLE, 1)
        Call .SetRenderState(D3DRS_POINTSIZE, Engine_FToDW(32))
        Call .SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
        Call .SetRenderState(D3DRS_POINTSPRITE_ENABLE, 1)
        Call .SetRenderState(D3DRS_POINTSCALE_ENABLE, 0)
    End With
End Sub


'***************************************************
' Font Functions
'***************************************************

Private Sub Fonts_Initialize(ByRef path As String)
'**************************************************************
'Author: •Parra
'Last Modify Date: 04/09/09
'Initialize fonts
'**************************************************************
    
    Fonts_Path = path
    
    ReDim Font(1 To CInt(GetVar(Fonts_Path, "MAIN", "Size")))
    
    Dim i As Long
    
    ' Set configuration
    For i = 1 To UBound(Font)
        With Font(i)
            .Format.Name = GetVar(Fonts_Path, CStr(i), "Name")
            .Format.Size = GetVar(Fonts_Path, CStr(i), "Size")
            
            Set .desc = .Format
            Set .D3DFont = D3DX.CreateFont(D3DDevice, .desc.hFont)
        End With
    Next i

End Sub

Public Sub Fonts_Render(ByRef index As Long, ByVal lngXPos As Integer, ByVal lngYPos As Integer, ByVal Width As Integer, ByVal Height As Integer, ByRef strText As String, ByVal lngColor As Long)
    
    If strText = vbNullString Then Exit Sub
    
    'This defines where it will be
    Static fontRect As RECT
    
    With fontRect
        .Top = lngYPos + regRenderDX.X1
        .Left = lngXPos + regRenderDX.Y1
        .bottom = lngYPos + Height + regRenderDX.X1
        .Right = lngXPos + Width + regRenderDX.Y1
    End With
    
    D3DX.DrawText Font(index).D3DFont, lngColor, strText, fontRect, DT_LEFT

End Sub

Private Sub Fonts_Destroy()
'**************************************************************
'Author: •Parra
'Last Modify Date: 04/09/09
'Destroy DirectX Fonts
'**************************************************************

    Dim i As Long
    
    For i = 1 To UBound(Font)
        Set Font(i).D3DFont = Nothing
        Set Font(i).desc = Nothing
        Set Font(i).Format = Nothing
    Next i
    
    Erase Font
End Sub


'***************************************************
' Functions to load graphics
'***************************************************

Private Sub Graphics_Initialize(ByRef maxMemoryInMB As Long)
'***************************************************
'Author: •Parra
'Last Modification: 08/09/09
'Initialize textures
'***************************************************
    
    ReDim Graphic(0)

    maxUsedBytes = maxMemoryInMB * 1024

    usedBytes = 0
    
End Sub



Private Sub Graphics_Destroy()
'***************************************************
'Author: •Parra
'Last Modification: 08/09/09
'Destroys all textures
'***************************************************

    ' Terminate Textures
    Dim i As Long
    
    For i = 1 To UBound(Graphic)
        Set Graphic(i).D3DTexture = Nothing
    Next i
    
    Erase Graphic
    
End Sub

''
' Draws a Grh's portion to the given area of any Device Context using BitBlt DLL
'
' @param    hdc         Handle to the window to be used by BitBlt when drawing graphics.
' @param    grhIndex    Handle to the Grh will be used to draw the graph in the picture.

Public Sub Graphics_Draw_GrhToHdc(ByVal hdc As Long, ByVal GrhIndex As Integer)
'******************************************************
'Author: Parra
'Last Modification: 08/09/09
'Draw a graphic on any Object that has the property hdc
'******************************************************
    Dim hDCsrc As Long
 
    If GrhIndex <= 0 Then Exit Sub
        
        'If it's animated switch GrhIndex to first frame
        If GrhData(GrhIndex).NumFrames <> 1 Then
            GrhIndex = GrhData(GrhIndex).Frames(1)
        End If
           
        hDCsrc = CreateCompatibleDC(hdc)
        
        Call SelectObject(hDCsrc, LoadPicture(App.path & "\Graficos\" & GrhData(GrhIndex).FileNum & ".bmp"))

        'Draw
        BitBlt hdc, 0, 0, _
        GrhData(GrhIndex).pixelWidth, GrhData(GrhIndex).pixelWidth, _
        hDCsrc, _
        GrhData(GrhIndex).sX, GrhData(GrhIndex).sY, _
        vbSrcCopy

        DeleteDC hDCsrc
End Sub

''
' Draws a Grh's portion to the given area of any Device Context using DirectX8 Graphics
'
' @param    Grh         Pointer to Grh.
' @param    X           Horizontal drawing position.
' @param    Y           Vertical drawing position.
' @param    Center      If this variable is true, the graph will be centered.
' @param    Animate     If this variable is true, indicate that the drawing is a graphic animation, and display the relevant frame.

Private Sub Graphics_Draw_GrhToSurface(ByRef Grh As Grh, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte, ByVal Animate As Byte)
'*****************************************************************
'Author: Unknown
'Last Modification: 08/09/09
'Last Modification Author: Parra
'Draws a GRH to a X & Y Device position
'*****************************************************************
On Error GoTo Error

    Dim CurrentGrhIndex As Integer

    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> -1 Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If

        Dim color(3) As Long: Graphics_ResetColor color()
        
        'Draw
        Graphics_Draw_Texture .FileNum, x, y, .pixelWidth, .pixelHeight, .sX, .sY, color()

    End With
    
Exit Sub

Error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.Number & " ] Error"
        End
    End If
End Sub

Public Sub Graphics_Draw_Texture(ByVal FileNum As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, ByRef color() As Long)
    
    Dim index As Integer
    
    If FileNum = Graphic(LastArrayIndex).FileNum Then
        index = LastArrayIndex
    Else
        If Graphics_LoadTexture(FileNum, index) = False Then Exit Sub
    End If
    
    If Not Graphic(index).D3DTexture Is Nothing Then
    
        Dim vector(3) As TLVERTEX
        
        ' * v1      * v3
        ' |\        |
        ' |  \      |
        ' |    \    |
        ' |      \  |
        ' |        \|
        ' * v0      * v2
        
        'Set up the vector(3) vertices
        vector(0).v.x = dest_x
        vector(0).v.y = dest_y + src_height
        vector(0).color = color(0)
        vector(0).rhw = 1
        
        vector(1).v.x = dest_x
        vector(1).v.y = dest_y
        vector(1).color = color(1)
        vector(1).rhw = 1
        
        vector(2).v.x = dest_x + src_width
        vector(2).v.y = dest_y + src_height
        vector(2).color = color(2)
        vector(2).rhw = 1
        
        vector(3).v.x = dest_x + src_width
        vector(3).v.y = dest_y
        vector(3).color = color(3)
        vector(3).rhw = 1
        
        vector(0).tu = src_x / Graphic(index).Width
        vector(0).tv = (src_y + src_height + 1) / Graphic(index).Height
        vector(1).tu = src_x / Graphic(index).Width
        vector(1).tv = src_y / Graphic(index).Height
        vector(2).tu = (src_x + src_width + 1) / Graphic(index).Width
        vector(2).tv = (src_y + src_height + 1) / Graphic(index).Height
        vector(3).tu = (src_x + src_width + 1) / Graphic(index).Width
        vector(3).tv = src_y / Graphic(index).Height
    
    End If
    
    LastArrayIndex = index
    
    'Set Textures
    D3DDevice.SetTexture 0, Graphic(index).D3DTexture
    
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, vector(0), Len(vector(0))
    
End Sub
Private Function Graphics_LoadTexture(ByRef FileNumber As Long, ByRef ArrayIndex As Integer) As Boolean
'***************************************
'Author: Parra
'Last Modification: 08/09/09
'Load texture from file to memory
'***************************************
    Dim i As Long
    
    ' Binary Search
    For i = 0 To UBound(Graphic)
        If Graphic(i).FileNum = FileNumber Then
            ArrayIndex = i
            Graphic(i).lastAccess = GetTickCount()
            Graphics_LoadTexture = True
            Exit Function
        End If
    Next i
    
    ' If not graphic is loaded
    ArrayIndex = UBound(Graphic) + 1
    
    ReDim Preserve Graphic(0 To ArrayIndex) As structGraphic
    
    ' Get bitmap to byte array
    Dim FileData() As Byte
    
    Graphics_GetBitmap FileNumber, FileData()
    
    ' Contains texture basic info
    Dim SurfaceDesc As D3DSURFACE_DESC
    
    With Graphic(ArrayIndex)
    
        ' Load texture from file to memory
        'Set Graphic(ArrayIndex).D3DTexture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, FileData(0), UBound(FileData()) + 1, D3DX_DEFAULT, _
                                              D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                              D3DX_FILTER_POINT, D3DX_FILTER_POINT, D3DColorXRGB(0, 0, 0), ByVal 0, ByVal 0)
                                              
        Set .D3DTexture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, FileData(0), UBound(FileData()) + 1, _
                    D3DX_DEFAULT, D3DX_DEFAULT, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                    D3DX_FILTER_NONE, &HFF000000, ByVal 0, ByVal 0)
        Erase FileData

        'Get the width and height of the texture
        .D3DTexture.GetLevelDesc 0, SurfaceDesc
        .Width = SurfaceDesc.Width
        .Height = SurfaceDesc.Height
    
        ' Set last access
        .lastAccess = GetTickCount()
        
        ' Set file number
        .FileNum = FileNumber
    
    End With
    
    ' Update used bytes
    usedBytes = usedBytes + (SurfaceDesc.Size / 1024) ' Is this correct?
    
    'If we use more memory than allowed...
'    Do While usedBytes > maxUsedBytes
'
'        'Destroy the graphics less used
'        If Not Graphics_ReleaseMemory Then Exit Do
'
'    Loop
    
    Graphics_LoadTexture = True
End Function
Private Function Graphics_ReleaseMemory() As Boolean

    Dim SurfaceDesc As D3DSURFACE_DESC
    Dim lastAccess As Long
    Dim LessUsed As Integer
    
    Dim i As Long
    
    For i = 0 To UBound(Graphic)
        If lastAccess > Graphic(i).lastAccess Then
            LessUsed = i
            lastAccess = Graphic(i).lastAccess
        End If
    Next i
    
    If Not Graphic(LessUsed).D3DTexture Is Nothing Then
    
        'Retrieve the surface desc
        Graphic(LessUsed).D3DTexture.GetLevelDesc 0, SurfaceDesc
        
        'Remove it
        Set Graphic(LessUsed).D3DTexture = Nothing
        
    End If
    
    Graphic(LessUsed).FileNum = 0
    
    ' ¿Acomodamos el array?
    '...
    
    'Update the used bytes
    usedBytes = usedBytes - (SurfaceDesc.Size / 1024)
    
    Graphics_ReleaseMemory = True
            
End Function

Private Sub Graphics_GetBitmap(ByRef FileNumber As Long, FileData() As Byte)

    ' Get a free file number
    Dim Handle As Integer
    Handle = FreeFile()

    Open App.path & "\Graficos\" & CStr(FileNumber) & ".bmp" For Binary Access Read As #Handle
        
        ' Create an array just big enough to hold the whole file
        ReDim FileData(LOF(Handle) - 1) As Byte

        ' Read the file into that array
        Get #Handle, , FileData()

    Close #Handle

End Sub

''
' Sets Color Array to original graphic color
'
' @param    Color()     Pointer to original array

Public Sub Graphics_ResetColor(ByRef color() As Long)
'***************************************
'Author: Parra
'Last Modification: 08/09/09
'Sets Color Array
'***************************************
    color(0) = -1: color(1) = -1: color(2) = -1: color(3) = -1
End Sub

''
' Main Loop.

Public Sub General_MainLoop()
'***************************************************
'Author: Arron Perkins
'Last Modification: 03/09/09
'Last modified by: •Parra
'Updates the game's model and renders everything.
'***************************************************
    
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
    
    Do While prgRun
    
        'Sólo dibujamos si la ventana no está minimizada
        If frmMain.WindowState <> 1 And frmMain.Visible Then
            
            If GameRun Then
                If UserMoving Then
                    '****** Move screen Left and Right if needed ******
                    If AddtoUserPos.x <> 0 Then
                        OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.x * timerTicksPerFrame
                        If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.x) Then
                            OffsetCounterX = 0
                            AddtoUserPos.x = 0
                            UserMoving = 0
                        End If
                    End If
                    
                    '****** Move screen Up and Down if needed ******
                    If AddtoUserPos.y <> 0 Then
                        OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.y * timerTicksPerFrame
                        If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.y) Then
                            OffsetCounterY = 0
                            AddtoUserPos.y = 0
                            UserMoving = 0
                        End If
                    End If
                End If
                
                'Update mouse position within view area
                General_ConvertMouseToTile frmMain.MouseX, frmMain.MouseY, MouseTileX, MouseTileY
                
                '****** Update screen ******
                'If (testCooperative = False) Then Exit Sub

                DirectX_Scene_BeginCustom 1, regRenderDX, D3DCLEAR_TARGET

                    If UserCiego Then
                        D3DDevice.Clear 1, regRenderDX, D3DCLEAR_TARGET, 0, 1#, 0
                    Else
                        Map_Render OffsetCounterX, OffsetCounterY
                    End If

                    If IScombate Then Fonts_Render 1, 260, 260, 300, 270, "MODO COMBATE", D3DColorXRGB(255, 0, 0)
                            
                    Dialogos.Render
                    DibujarCartel
                            
                    DialogosClanes.Draw
                
                'DirectX_Scene_EndCustom regRender, frmMain.hWnd
                DirectX_Scene_End
    
                'General_LimitFPS
                
                General_UpdateFPS
                
                General_UpdateElapsedTime

            End If 'GameRun

            'Play ambient sounds
            Sound_RenderAmbient
            
            General_CheckKeys
            
        End If ' Window State
        
        ' If there is anything to be sent, we send it
        FlushBuffer
        
        DoEvents
    Loop

End Sub

''
' Return Elapsed Time

Private Function General_ElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    QueryPerformanceCounter start_time
    
    'Calculate elapsed time
    General_ElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    QueryPerformanceCounter end_time
End Function


''
' Update Elapsed time

Public Sub General_UpdateElapsedTime()
'***************************************************
'Author: Parra
'Last Modification: 09/09/09
'Update Elapsed time
'***************************************************
    'Get timing info
    timerElapsedTime = General_ElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

End Sub

''
' Return engine frames per second

Public Property Get General_FPS() As Integer
'***************************************************
'Author: Parra
'Last Modification: 09/09/09
'Return engine frames per second
'***************************************************

    General_FPS = fpsValue
End Property

''
' Calculate frames per second

Private Sub General_UpdateFPS()
'***************************************************
'Author: Parra
'Last Modification: 09/09/09
'Calculate frames per second
'***************************************************
    
    'FPS update
    If GetTickCount() - fpsLastCheck >= 1000 Then
        fpsValue = fpsCounter
            frmMain.Caption = CStr(fpsValue) 'Provisional
        fpsCounter = 0
        fpsLastCheck = GetTickCount()
    End If
    
    fpsCounter = fpsCounter + 1
    
End Sub

Sub General_ConvertMouseToTile(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Byte, ByRef tY As Byte)
'***************************************************
'Author: Unknown
'Last Modification: Unknown
'Converts where the mouse is in the main window to a tile position.
'***************************************************
    
    tX = UserPos.x + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
    tY = UserPos.y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2
End Sub

''
' Limits the frames per second to the vertical refresh rate monitor

Private Sub General_LimitFPS()
'***************************************************
'Author: Unknown
'Last Modification Author: Parra
'Last Modification Date: 14/09/09
'Limit Frames per Second
'***************************************************
    
    'Limit FPS to 100 (an easy number higher than monitor's vertical refresh rates)
    While (GetTickCount() - fpsLastCheck) / 9.8 < fpsCounter
        Sleep 5
    Wend
End Sub

''
' Checks keys and respond

Private Sub General_CheckKeys()
    Static lastMovement As Long
    
    'No input allowed while Argentum is not the active window
    If (GetActiveWindow <> 0) = False Then Exit Sub
    
    'No walking when in commerce or banking.
    If Comerciando Then Exit Sub
    
    'No walking while writting in the forum.
    If frmForo.Visible Then Exit Sub
    
    'If game is paused, abort movement.
    If pausa Then Exit Sub
    
    'Control movement interval (this enforces the 1 step loss when meditating / resting client-side)
    If GetTickCount - lastMovement > 56 Then
        lastMovement = GetTickCount
    Else
        Exit Sub
    End If
    
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then
            'Move Up
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
                User_Move Heading.North
                frmMain.Coord.Caption = "(" & UserMap & "[" & UserMapName & "]" & "," & UserPos.x & "," & UserPos.y & ")"
                Exit Sub
            End If
            
            'Move Right
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
                User_Move East
                frmMain.Coord.Caption = "(" & UserMap & "[" & UserMapName & "]" & "," & UserPos.x & "," & UserPos.y & ")"
                Exit Sub
            End If
        
            'Move down
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
                User_Move Heading.South
                frmMain.Coord.Caption = "(" & UserMap & "[" & UserMapName & "]" & "," & UserPos.x & "," & UserPos.y & ")"
                Exit Sub
            End If
        
            'Move left
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
                User_Move Heading.West
                frmMain.Coord.Caption = "(" & UserMap & "[" & UserMapName & "]" & "," & UserPos.x & "," & UserPos.y & ")"
                Exit Sub
            End If
            
            ' We haven't moved - Update 3D sounds!
            Call Engine.SoundEffects_MoveListener(UserPos.x, UserPos.y)
        Else
            
            If (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
                    GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
                User_Move RandomNumber(Heading.North, Heading.West)
            Else
                ' We haven't moved - Update 3D sounds!
                SoundEffects_MoveListener UserPos.x, UserPos.y
            End If

            frmMain.Coord.Caption = "(" & UserPos.x & "," & UserPos.y & ")"
        End If
    End If
End Sub

''
'Initializes the Audio object.
'
' @param    hWnd        Handle to the window to be used by DirectMusic when playing midi files.
' @param    sound_path  Directory in wich to search for sound files.
' @param    music_path  Directory in wich to search for music files.
'
' @return   True if no errors occured, false otherwise.

Public Function Audio_Initialize(ByVal hWnd As Long, ByVal SoundPath As String, ByVal MusicPath As String) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/09/09
'Last modified by: •Parra
'Creates all DS8 and DM8 objects
'***************************************************
On Error GoTo ErrHandler
    
    ' Prevent multiple creations
    If Inited Then Exit Function
    
    'Set default volume
    Sound_Vol = 0
    
    'Validate DX object
    If DirectX Is Nothing Then GoTo ErrHandler
    
    ''
    ' Initialize DirectSound
    Sound = Sound_Initialize(hWnd)
    
    'Set sound path
    Sound_Path = SoundPath
    
    sActivated = True
    
    
    ''
    ' Initialize DirectMusic
    Music = Music_Initialize(hWnd)
    
    'Set music path
    Music_Path = MusicPath

    mActivated = True
        
    Music_CurrentFile = vbNullString

    ' Audio is Initied correctly
    Inited = True
    
    Audio_Initialize = True
Exit Function

ErrHandler:
End Function

''
' Terminates the Audio object.

Private Sub Audio_Destroy()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/09/09
'Last Modification by: •Parra
'Destroy's all DirectMusic8 and DirectSound8 Objects
'***************************************************
On Local Error Resume Next
    Dim loopc As Long
    
    'Stop every channel being used and destroy the buffer
    For loopc = 1 To Sound_MaxBuffers
        Sound_Buffers(loopc).Buffer.Stop
        Set Sound_Buffers(loopc).Buffer = Nothing
    Next loopc
    
    'Stop music
    Performance.StopEx Segment, 0, DMUS_SEGF_DEFAULT
    
    Performance.CloseDown
    
    Set DirectSound = Nothing
    
    Set Segment = Nothing
    Set Performance = Nothing
    Set Loader = Nothing
    
    Inited = False
End Sub

''
' Creates and configures all DirectMusic objects.
'
' @param    DX          DirectX8 object to be used when creating the DirectSound and Directmusic objects.
' @param    hWnd        Handle to the window to be used by DirectSound when playing wav files.
'
' @return   True if no errors occured, false otherwise.

Private Function Sound_Initialize(ByVal hWnd As Long) As Boolean
'***************************************************
'Author: •Parra
'Last Modification: 03/09/09
'Initialize DS8. Returns false if music isn't supported by HW
'***************************************************
On Error GoTo ErrHandler
    
    Set DirectSoundEnum = DirectX.GetDSEnum
    Set DirectSound = DirectX.DirectSoundCreate(DirectSoundEnum.GetGuid(1))
 
    DirectSound.SetCooperativeLevel hWnd, DSSCL_PRIORITY
    
    Sound_Initialize = True
Exit Function

ErrHandler:
End Function

''
' Plays a new WAVE file.
'
' @param    FileName    The path to the file to be loaded. Must be a relative path from the default path set at initialization.
' @param    srcX        The X position of the sound source, in map coordinates. -1 if sound should not be altered by 3D alterations.
' @param    srcY        The Y position of the sound source, in map coordinates. -1 if sound should not be altered by 3D alterations.
' @param    LoopSound   The loop state for the sound.
'
' @return   The buffer index in which the sound is playing, or 0 in case of error.

Public Function Sound_Play(ByVal FileName As String, Optional ByVal srcX As Byte = 0, Optional ByVal srcY As Byte = 0, Optional ByVal LoopSound As Sound_LoopStyle = Sound_LoopStyle.Default) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 28/07/08
'Last Modified by: Franco Zeoli (Noich)
'Added support for 3D sound effects.
'***************************************************
On Error GoTo ErrHandler
    Dim bufferIndex As Long
    
    If Not Sound_Activated Then Exit Function
    
    'Get the buffer index were wave was loaded
    bufferIndex = Sound_Load(FileName, LoopSound)
    If bufferIndex = 0 Then Exit Function   'If an error ocurred abort
    
    With Sound_Buffers(bufferIndex)
        If .looping Then
            .Buffer.Play DSBPLAY_LOOPING
        Else
            .Buffer.Play DSBPLAY_DEFAULT
        End If
        
        'Apply volume
        .Buffer.SetVolume Sound_Vol
        
        'Store position
        .x = srcX
        .y = srcY
        
        'Get standard frequency
        .normalFq = .Buffer.GetFrequency()
    End With
    
    Sound_Play = bufferIndex
    
    If sEffectsActivated Then
        If srcX <> 0 And srcY <> 0 Then
            SoundEffects_Update3D bufferIndex, 0, 0
        End If
    End If
Exit Function

ErrHandler:
End Function

''
' Loads a new WAVE file.
'
' @param    FileName    The path to the file to be loaded. Must be a relative path from the default path set at initialization.
' @param    looping     The loop state for the sound.
'
' @return   The buffer index in which the sound is playing, or 0 in case of error.

Private Function Sound_Load(ByVal FileName As String, ByVal looping As Sound_LoopStyle) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/09/09
'Last modified by: •Parra
'Loads a new wave into memory (if it's not already loaded) and returns
'the buffer index
'***************************************************
On Error GoTo ErrHandler
    Dim bufferDesc As DSBUFFERDESC
    Dim waveFormat As WAVEFORMATEX
    Dim i As Long
    
    FileName = UCase$(FileName)
    
    If Not FileExist(Sound_Path & FileName, vbArchive) Then _
        Exit Function
    
    'Check if the buffer is in memory and not playing
    For i = 1 To Sound_MaxBuffers
        If Sound_Buffers(i).FileName = FileName Then
            If Sound_Buffers(i).Buffer.GetStatus = 0 Or Sound_Buffers(i).Buffer.GetStatus() = DSBSTATUS_BUFFERLOST Then
                'Found it!!! We just play this one :)
                Sound_Load = i
                
                'Set looping if needed
                If looping <> Sound_LoopStyle.Default Then Sound_Buffers(i).looping = (looping = Sound_LoopStyle.Enabled)
                
                'Reset default sound frecuency to avoid issues.
                Sound_Buffers(i).Buffer.SetFrequency Sound_Buffers(i).normalFq
                
                Exit Function
            End If
        End If
    Next i
    
    'Not in memory, search for an empty buffer
    For i = 1 To Sound_MaxBuffers
        If Sound_Buffers(i).Buffer Is Nothing Then
            Exit For
        End If
    Next i
    
    'No empty buffers?? Search for one not being played and use that slot
    If i > Sound_MaxBuffers Then
        For i = 1 To Sound_MaxBuffers
            If Sound_Buffers(i).Buffer.GetStatus() = DSBSTATUS_BUFFERLOST Or Sound_Buffers(i).Buffer.GetStatus() = 0 Then
                Exit For
            End If
        Next i
    End If
    
    'Still no free buffers?? Get the first non-looping sound
    If i > Sound_MaxBuffers Then
        For i = 1 To Sound_MaxBuffers
            If Not (Sound_Buffers(i).Buffer.GetStatus() And DSBSTATUS_LOOPING) Then
                Sound_Buffers(i).Buffer.Stop
                Exit For
            End If
        Next i
    End If
    
    'If every sound is looping, just ignore new sound unless it's looping too (in which case we set it at the first slot by default)
    If i > Sound_MaxBuffers Then
        If Not looping Then Exit Function
        i = 1
    End If
    
    'Create new sound
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    
    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2
    waveFormat.lSamplesPerSec = 22050
    waveFormat.nBitsPerSample = 16
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
    
    With Sound_Buffers(i)
        Set .Buffer = Nothing   'Get rid of any previous data
        
        .FileName = FileName
        .looping = (looping = Sound_LoopStyle.Enabled)
        
        Set .Buffer = DirectSound.CreateSoundBufferFromFile(Sound_Path & FileName, bufferDesc)
    End With
    
    Sound_Load = i
Exit Function

ErrHandler:
End Function

''
' Stops a given sound or all of them.
'
' @param    bufferIndex     The index of the sound file to stop playing. If 0, all sounds are stopped.

Public Sub Sound_Stop(Optional ByVal bufferIndex As Long = 0)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/09/09
'Last modified by: •Parra
'Stops the sound in the given buffer index, or every sound if
'index is 0
'***************************************************
    If bufferIndex > 0 And bufferIndex <= Sound_MaxBuffers Then
        If Sound_Buffers(bufferIndex).Buffer.GetStatus() And DSBSTATUS_PLAYING Then
            Sound_Buffers(bufferIndex).Buffer.Stop
        End If
    ElseIf bufferIndex = 0 Then
        Dim i As Long
        For i = 1 To Sound_MaxBuffers
            If Not Sound_Buffers(i).Buffer Is Nothing Then
                If Sound_Buffers(i).Buffer.GetStatus() <> 0 And Sound_Buffers(i).Buffer.GetStatus() <> DSBSTATUS_BUFFERLOST Then
                    Sound_Buffers(i).Buffer.Stop
                End If
            End If
        Next i
    End If
End Sub

''
' Retrieves wether there are sounds currentyl playing or not.
'
' @return   true if there is at least one sound playing, false otherwise.

Public Property Get Sound_Playing() As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Returns true if there is at least one sound playing
'***************************************************
    
    If Not Sound Then Exit Property

    Dim i As Long
    
    For i = 1 To Sound_MaxBuffers
        If Sound_Buffers(i).Buffer.GetStatus() And DSBSTATUS_PLAYING Then
            Sound_Playing = True
            Exit Property
        End If
    Next i
End Property

''
' Sets the volume of sound.
'
' @param    Volume  New value to be set, in percentage (0 to 100).

Public Property Let Sound_Volume(ByVal Volume As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Sets the sound's volume as a percentage
'***************************************************
    Dim i As Long
    
    If Volume > 100 Then Exit Property
    
    'Take percentage to actual value
    Sound_Vol = Volume * (DSBVOLUME_MAX - DSBVOLUME_MIN) / 100 + DSBVOLUME_MIN
    
    For i = 1 To Sound_MaxBuffers
        If Not Sound_Buffers(i).Buffer Is Nothing Then
            If Sound_Buffers(i).Buffer.GetStatus And DSBSTATUS_PLAYING Then
                Sound_Buffers(i).Buffer.SetVolume Sound_Vol
            End If
        End If
    Next i
End Property

''
' Gets the volume of sound.
'
' @return   The current sound volume in percentage.

Public Property Get Sound_Volume() As Byte
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Sets the music's volume as a percentage
'***************************************************
    If Not Sound Then Exit Property
    
    Sound_Volume = CByte((Sound_Vol - DSBVOLUME_MIN) * 100 / (DSBVOLUME_MAX - DSBVOLUME_MIN))
End Property

Private Function Sound_RenderAmbient()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 14/09/2009
'Last Modify By: Parra
'Update map Sounds
'**************************************************************
    
    Dim Location As Point

    If mapRain(UserMap) = 1 Then
        If bRain Then
            If bTecho Then
                If IsPlaying <> PlayLoop.plLluviain Then
                    If RainBufferIndex Then _
                        Call Engine.Sound_Stop(RainBufferIndex)
                    RainBufferIndex = Engine.Sound_Play("lluviain.wav", 0, 0, Sound_LoopStyle.Enabled)
                    IsPlaying = PlayLoop.plLluviain
                End If
            Else
                If IsPlaying <> PlayLoop.plLluviaout Then
                    If RainBufferIndex Then _
                        Call Engine.Sound_Stop(RainBufferIndex)
                    RainBufferIndex = Engine.Sound_Play("lluviaout.wav", 0, 0, Sound_LoopStyle.Enabled)
                    IsPlaying = PlayLoop.plLluviaout
                End If
            End If
        End If
    End If
        
    If bFogata Then
        bFogata = Map_CheckBonfire(Location)
        If Not bFogata Then
            Sound_Stop FogataBufferIndex
            FogataBufferIndex = 0
        End If
    Else
        bFogata = Map_CheckBonfire(Location)
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Sound_Play("fuego.wav", Location.x, Location.y, Sound_LoopStyle.Enabled)
    End If
    
End Function



Private Sub SoundEffects_Update3D(ByVal bufferIndex As Long, ByVal deltaX As Integer, ByVal deltaY As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/09/09
'Last modified by: •Parra
'
'Fixed Dopplers formula
'Refreshes the sound's arguments.
'Fixed the Sound_Vol formula
'Remodelation function for DirectX8
'***************************************************
    Dim linearDistanceOld As Single
    Dim linearDistanceNew As Single
    Dim deltaDistance As Single
    Dim distanceXOld As Long
    Dim distanceYOld As Long
    Dim distanceXNew As Long
    Dim distanceYNew As Long
    Dim perceivedFq As Long
    
    With Sound_Buffers(bufferIndex)
        distanceXOld = .x - SoundEffects_lastPos.x
        distanceYOld = .y - SoundEffects_lastPos.y
        
        distanceXNew = distanceXOld + deltaX
        distanceYNew = distanceYOld + deltaY
        
        linearDistanceOld = Sqr(distanceXOld * distanceXOld + distanceYOld * distanceYOld)
        linearDistanceNew = Sqr(distanceXNew * distanceXNew + distanceYNew * distanceYNew)
        
        deltaDistance = linearDistanceNew - linearDistanceOld
        
        'Update frequency emulating Doppler effect
        perceivedFq = .normalFq * (1 - deltaDistance / SOUND_DELTA_FQ)
        If perceivedFq < DSBFREQUENCY_MIN Then perceivedFq = DSBFREQUENCY_MIN
        
        .Buffer.SetFrequency perceivedFq
        
        'Set volumen amortiguation according to distance
        If linearDistanceNew > SOUND_MAX_DISTANCE_TO_SOURCE Then
            .Buffer.SetVolume Sound_Vol + (DSBVOLUME_MIN - Sound_Vol)
        Else
            .Buffer.SetVolume Sound_Vol + (linearDistanceNew / SOUND_MAX_DISTANCE_TO_SOURCE) * (DSBVOLUME_MIN - Sound_Vol)
        End If
        
        'Prevent division by zero
        If linearDistanceNew = 0 Then linearDistanceNew = 1
        
        'Set panning according to relative position of the source to the listener
        .Buffer.SetPan (distanceXNew / linearDistanceNew) * DSBPAN_RIGHT
    End With
End Sub

''
' Updates 3D sounds based on the movement of the listener.
'
' @param    x   The new coordinate of the listener in the x axis (consider 0 as the furthest point to the left).
' @param    y   The new coordinate of the listener in the y axis (consider 0 as the furthest point to the front).
'
' @return   The buffer index in which the sound is playing, or 0 in case of error.

Public Sub SoundEffects_MoveListener(ByVal x As Integer, ByVal y As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/09/09
'Last modified by: •Parra
'Moves the position of the listener updating all 3D sounds.
'***************************************************
    Dim i As Long
    Dim deltaX As Integer
    Dim deltaY As Integer
    
    If Not sEffectsActivated Then Exit Sub
    
    deltaX = x - SoundEffects_lastPos.x
    deltaY = y - SoundEffects_lastPos.y
    
    For i = 1 To Sound_MaxBuffers
        If Not Sound_Buffers(i).Buffer Is Nothing Then
            If Sound_Buffers(i).Buffer.GetStatus And DSBSTATUS_PLAYING Then
                If Sound_Buffers(i).x <> 0 And Sound_Buffers(i).y <> 0 Then
                    SoundEffects_Update3D i, deltaX, deltaY
                End If
            End If
        End If
    Next i
    
    SoundEffects_lastPos.x = x
    SoundEffects_lastPos.y = y
End Sub

''
' Sets wether sound effects are activated or not.
'
' @param nValue True if sound effects should be activated, false otherwise.

Public Property Let SoundEffects_Activated(ByVal nValue As Boolean)
'***************************************************
'Author: Franco Zeoli (Noich)
'Last Modification: 28/07/08
'Sets wether the sound effects are activated or not
'***************************************************
    sEffectsActivated = nValue
End Property

''
' Retrieves wether sound effects are activated or not.
'
' @return   True if sound effects are activated, false otherwise.

Public Property Get SoundEffects_Activated() As Boolean
'***************************************************
'Author: Franco Zeoli (Noich)
'Last Modification: 28/07/08
'Retrieves wether the sound effects are activated or not
'***************************************************
    SoundEffects_Activated = sEffectsActivated
End Property

''
' Creates and configures all DirectMusic objects.
'
' @param    hWnd        Handle to the window to be used by DirectMusic when playing midi files.
'
' @return   True if no errors occured, false otherwise.

Private Function Music_Initialize(ByVal hWnd As Long) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/09/09
'Last modified by: •Parra
'Initialize DM. Returns false if music isn't supported by HW
'***************************************************
On Error GoTo ErrHandler
    
    Dim musParams As DMUS_AUDIOPARAMS

    Set Loader = DirectX.DirectMusicLoaderCreate()
    
    Set Performance = DirectX.DirectMusicPerformanceCreate()
    Performance.InitAudio hWnd, DMUS_AUDIOF_ALL, musParams, Nothing, DMUS_APATH_DYNAMIC_STEREO, 128
    Performance.SetMasterAutoDownload True        'Enable auto download of instruments
        
    'Set tempo to 0 and volume of music
    Music_Tempo = 0
    Performance.SetMasterVolume 200
    
    Music_Initialize = True
Exit Function

ErrHandler:
End Function

''
' Plays a new MIDI file.
'
' @param    file    The path to the file to be loaded. Must be a relative path from the default path set at initialization.
' @param    Loops   Number of times to loop sound. For eternal loops set it to -1.
'
' @return   True if file was properly loaded and started playing, False otherwise.

Public Function Music_Play(Optional ByVal file As String = vbNullString, Optional ByVal Loops As Long = -1) As Boolean ' GSZ
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/09/09
'Last modified by: •Parra
'Loads a midi and start to play it
'***************************************************
On Error GoTo ErrHandler
    If LenB(file) > 0 Then _
        Music_CurrentFile = file
    
    If Not Music_Activated Then Exit Function
    
    If Not Music Then Exit Function
    
    If Music_Playing() Then Music_Stop
    
    If LenB(file) > 0 Then
        If Not Music_Load(file) Then Exit Function
    Else
        'Make sure we have a loaded segment
        If Segment Is Nothing Then Exit Function
    End If
    
    'Play it
    Segment.SetRepeats Loops
    
    Set SegState = Performance.PlaySegmentEx(Segment, DMUS_SEGF_DEFAULT, 0)
    
    Music_Play = True
Exit Function

ErrHandler:
End Function

''
' Loads a new MIDI file.
'
' @param    file    The path to the file to be loaded. Must be a relative path from the default path set at initialization.
'
' @return   True if file was properly loaded and started playing, False otherwise.

Private Function Music_Load(ByVal file As String) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Loads a midi to memory
'***************************************************
On Error GoTo ErrHandler
    If Not FileExist(Music_Path & file, vbArchive) Then Exit Function
    
    Music_Stop
    
    'Destroy old object
    Set Segment = Nothing
    
    Set Segment = Loader.LoadSegment(Music_Path & file)
    
    If Segment Is Nothing Then
        Exit Function
    End If
    
    Segment.SetStandardMidiFile
    
    Music_Load = True
Exit Function

ErrHandler:
End Function

''
' Stops playing the currently loaded MIDI file.

Public Sub Music_Stop()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/09/09
'Last modified by: •Parra
'Stops playing a midi
'***************************************************
On Error GoTo ErrHandler
    If Music_Playing Then
        Performance.StopEx Segment, 0, DMUS_SEGF_DEFAULT
    End If
Exit Sub

ErrHandler:
End Sub

''
' Sets the music's tempo.
'
' @param    Tempo   New Tempo value. Must be in the range of 0.25 to 2.0

Public Property Let Music_Tempo(ByVal Tempo As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Sets the music's tempo
'***************************************************
    If Not Music Then Exit Property
    
    If Tempo < 0.25 Or Tempo > 2# Then Exit Property
    
    Performance.SetMasterTempo Tempo
End Property

''
' Sets the music's volume.
'
' @param    Volume  Percentage to wich to set the music volume.

Public Property Let Music_Volume(ByVal Volume As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Sets the music's volume as a percentage
'***************************************************
    If Not Music Then Exit Property
    
    If Volume < 0 Or Volume > 100 Then Exit Property
    
    ' Volume ranges from -10000 to 10000
    Performance.SetMasterVolume Volume * 200 - 10000
End Property

''
' Gets the music's tempo.
'
' @return   Current music'tempo.

Public Property Get Music_Tempo() As Single
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Returns the music's tempo
'***************************************************
    If Not Music Then Exit Property
    
    Music_Tempo = Performance.GetMasterTempo()
End Property

''
' Gets the music's volume.
'
' @return   Current music's volume as a percentage.

Public Property Get Music_Volume() As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Returns the music's volume as a percentage
'***************************************************
    If Not Music Then Exit Property
    
    Music_Volume = (Performance.GetMasterVolume() + 10000) / 200
End Property

''
' Checks wether there is music currently playing.
'
' @return   True if there is a MIDI playing, false otherwise.

Public Property Get Music_Playing() As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Returns True if a midi is being played
'***************************************************
    If Not Music Or Segment Is Nothing Then Exit Property
    
    Music_Playing = Performance.IsPlaying(Segment, SegState)
End Property

''
' Retrieves the music's length.
'
' @return   Current music's length.

Public Property Get Music_Lenght() As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Returns the lenght of the current midi
'***************************************************
    If Not Music Then Exit Property
    
    Music_Lenght = Segment.GetLength()
End Property

''
' Retrieves wether music is activated or not.
'
' @return   True if music is activated, false otherwise.

Public Property Get Music_Activated() As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/03/07
'Retrieves wether the music is activated or not
'***************************************************
    Music_Activated = mActivated
End Property

''
' Sets wether music is activated or not.
'
' @param    nValue  True if music should be activated, false otherwise.

Public Property Let Music_Activated(ByVal nValue As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/03/07
'Sets wether the music is activated or not
'***************************************************
    If mActivated = nValue Then Exit Property
    
    mActivated = nValue
    
    If nValue Then
        Music_Play Music_CurrentFile
    Else
        Music_Stop
    End If
End Property

''
' Retrieves wether sound is activated or not.
'
' @return   True if sound is activated, false otherwise.

Public Property Get Sound_Activated() As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/03/07
'Retrieves wether the sound is activated or not
'***************************************************
    If Not Sound Then Exit Property

    Sound_Activated = sActivated
End Property

''
' Sets wether sound is activated or not.
'
' @param    nValue  True if sound should be activated, false otherwise.

Public Property Let Sound_Activated(ByVal nValue As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/03/07
'Sets wether the sound is activated or not
'***************************************************
    If sActivated = nValue Then Exit Property
    
    sActivated = nValue
    
    If Not nValue Then
        Sound_Stop
    End If
End Property


'***************************************************
' Grh Functions
'***************************************************

Private Sub Grh_Initialize(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    Grh.GrhIndex = GrhIndex
    
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = -1
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed
End Sub

''
' Loads grh data using the new file format.
'
' @return   True if the load was successfull, False otherwise.

Private Function Grh_Load_Data(ByVal IndexType As Byte) As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim Handle As Integer
    Dim fileVersion As Long
    
    'Open files
    Handle = FreeFile()
    
    Open App.path & "\INIT\Graficos" & CStr(IndexType) & ".ind" For Binary Access Read As Handle
    Seek #1, 1
    
    'Get file version
    Get Handle, , fileVersion
    
    'Get number of grhs
    Get Handle, , grhCount
    
    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData
    
    While Not EOF(Handle)
        Get Handle, , Grh
        
        With GrhData(Grh)
            'Get number of frames
            Get Handle, , .NumFrames
            If .NumFrames <= 0 Then GoTo ErrorHandler
            
            ReDim .Frames(1 To GrhData(Grh).NumFrames)
            
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                    Get Handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        GoTo ErrorHandler
                    End If
                Next Frame
                
                Get Handle, , .Speed
                
                If .Speed <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then GoTo ErrorHandler
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                Get Handle, , .FileNum
                If .FileNum <= 0 Then GoTo ErrorHandler
                
                Get Handle, , GrhData(Grh).sX
                If .sX < 0 Then GoTo ErrorHandler
                
                Get Handle, , .sY
                If .sY < 0 Then GoTo ErrorHandler
                
                Get Handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                Get Handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth
                
                .Frames(1) = Grh
            End If
        End With
    Wend
    
    Close Handle
    
    Grh_Load_Data = True
Exit Function

ErrorHandler:
    Grh_Load_Data = False
End Function

Private Function Grh_Load_Helmets() As Boolean
'***************************************************
'Author: Parra
'Last Modification: 10/11/09
'Loads the helmets list
'***************************************************

On Error GoTo ErrHandle

    Dim N As Integer
    Dim i As Long
    Dim NumHelmets As Integer
    
    Dim Miscabezas() As HeadList
    
    N = FreeFile()
    Open App.path & "\init\Cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumHelmets
    
    'Resize array
    ReDim HelmetData(0 To NumHelmets) As HeadData
    ReDim Miscabezas(0 To NumHelmets) As HeadList
    
    For i = 1 To NumHelmets
        Get #N, , Miscabezas(i)
    
        If Miscabezas(i).Head(1) Then
            Call Grh_Initialize(HelmetData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call Grh_Initialize(HelmetData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call Grh_Initialize(HelmetData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call Grh_Initialize(HelmetData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
    
ErrHandle:

End Function

Private Function Grh_Load_ShieldsAnim() As Boolean
'***************************************************
'Author: Parra
'Last Modification: 10/11/09
'Loads the Shields Anim list
'***************************************************

On Error GoTo ErrHandle

    Dim loopc As Long
    Dim arch As String

    arch = App.path & "\init\" & "escudos.dat"

    Dim NumShieldAnims As Integer

    NumShieldAnims = Val(GetVar(arch, "INIT", "NumEscudos"))

    ReDim ShieldAnimData(1 To NumShieldAnims) As ShieldAnimData

    For loopc = 1 To NumShieldAnims
        Grh_Initialize ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopc, "Dir1")), 0
        Grh_Initialize ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopc, "Dir2")), 0
        Grh_Initialize ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopc, "Dir3")), 0
        Grh_Initialize ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopc, "Dir4")), 0
    Next loopc

ErrHandle:

End Function
    
Private Function Grh_Load_Heads() As Boolean
'***************************************************
'Author: Parra
'Last Modification: 10/11/09
'Loads the heads list
'***************************************************

On Error GoTo ErrHandle
    
    Dim N As Integer
    Dim i As Long
    Dim Numheads As Integer

    Dim Miscabezas() As HeadList

    N = FreeFile()
    Open App.path & "\init\Cabezas.ind" For Binary Access Read As #N

    'cabecera
    Get #N, , MiCabecera

    'num de cabezas
    Get #N, , Numheads

    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As HeadList

    For i = 1 To Numheads
        Get #N, , Miscabezas(i)

        If Miscabezas(i).Head(1) Then
            Call Grh_Initialize(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call Grh_Initialize(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call Grh_Initialize(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call Grh_Initialize(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i

    Close #N

ErrHandle:

End Function

Private Function Grh_Load_Bodys() As Boolean
'***************************************************
'Author: Parra
'Last Modification: 10/11/09
'Loads the bodys list
'***************************************************

On Error GoTo ErrHandle

    Dim N As Integer
    Dim i As Long
    Dim NumCuerpos As Integer

    Dim MisCuerpos() As BodyList

    N = FreeFile()
    Open App.path & "\init\Personajes.ind" For Binary Access Read As #N

    'cabecera
    Get #N, , MiCabecera

    'num de cabezas
    Get #N, , NumCuerpos

    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As BodyList

    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)

        If MisCuerpos(i).Body(1) Then
            Grh_Initialize BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
            Grh_Initialize BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
            Grh_Initialize BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
            Grh_Initialize BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0

            BodyData(i).HeadOffset.x = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.y = MisCuerpos(i).HeadOffsetY
        End If
    Next i

    Close #N

ErrHandle:

End Function

Private Function Grh_Load_WeaponsAnim() As Boolean
'***************************************************
'Author: Parra
'Last Modification: 10/11/09
'Loads the weapons anim list
'***************************************************

On Error GoTo ErrHandle

On Error Resume Next

    Dim loopc As Long
    Dim arch As String

    arch = App.path & "\init\" & "armas.dat"

    Dim NumWeaponAnims As Integer

    NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))

    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData

    For loopc = 1 To NumWeaponAnims
        Grh_Initialize WeaponAnimData(loopc).WeaponWalk(1), Val(GetVar(arch, "ARMA" & loopc, "Dir1")), 0
        Grh_Initialize WeaponAnimData(loopc).WeaponWalk(2), Val(GetVar(arch, "ARMA" & loopc, "Dir2")), 0
        Grh_Initialize WeaponAnimData(loopc).WeaponWalk(3), Val(GetVar(arch, "ARMA" & loopc, "Dir3")), 0
        Grh_Initialize WeaponAnimData(loopc).WeaponWalk(4), Val(GetVar(arch, "ARMA" & loopc, "Dir4")), 0
    Next loopc

ErrHandle:

End Function

Private Function Grh_Load_Fxs() As Boolean
'***************************************************
'Author: Parra
'Last Modification: 10/11/09
'Loads the fxs list
'***************************************************

On Error GoTo ErrHandle

    Dim N As Integer
    Dim i As Long
    Dim NumFxs As Integer

    N = FreeFile()
    Open App.path & "\init\Fxs.ind" For Binary Access Read As #N

    'cabecera
    Get #N, , MiCabecera

    'num de cabezas
    Get #N, , NumFxs

    'Resize array
    ReDim FxData(1 To NumFxs) As FxData

    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i

    Close #N

ErrHandle:

End Function


'***************************************************
' Map Functions
'***************************************************

Public Function Map_Load(ByVal Map As Integer) As Boolean
'**************************************************************
'Formato de mapas optimizado para reducir el espacio que ocupan.
'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
'**************************************************************
On Error GoTo ErrHandle

    Dim y As Long
    Dim x As Long
    Dim TempInt As Integer
    Dim ByFlags As Byte
    Dim Handle As Integer
    
    Handle = FreeFile()
    
    Open App.path & "\MAPAS\Mapa" & Map & ".map" For Binary As Handle
    Seek Handle, 1
            
    'map Header
    Get Handle, , TempInt
    Get Handle, , MiCabecera ' Cabezera sucks
    Get Handle, , TempInt
    Get Handle, , TempInt
    Get Handle, , TempInt
    Get Handle, , TempInt
    
    'Load arrays
    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize
            Get Handle, , ByFlags
            
            mapData(x, y).Blocked = (ByFlags And 1)
            
            Get Handle, , mapData(x, y).Graphic(1).GrhIndex
            Grh_Initialize mapData(x, y).Graphic(1), mapData(x, y).Graphic(1).GrhIndex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get Handle, , mapData(x, y).Graphic(2).GrhIndex
                Grh_Initialize mapData(x, y).Graphic(2), mapData(x, y).Graphic(2).GrhIndex
            Else
                mapData(x, y).Graphic(2).GrhIndex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get Handle, , mapData(x, y).Graphic(3).GrhIndex
                Grh_Initialize mapData(x, y).Graphic(3), mapData(x, y).Graphic(3).GrhIndex
            Else
                mapData(x, y).Graphic(3).GrhIndex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get Handle, , mapData(x, y).Graphic(4).GrhIndex
                Grh_Initialize mapData(x, y).Graphic(4), mapData(x, y).Graphic(4).GrhIndex
            Else
                mapData(x, y).Graphic(4).GrhIndex = 0
            End If
            
            'Trigger used?
            If ByFlags And 16 Then
                Get Handle, , mapData(x, y).Trigger
            Else
                mapData(x, y).Trigger = 0
            End If
            
            'Erase NPCs
            If mapData(x, y).CharIndex > 0 Then
                Engine.Character_Destroy mapData(x, y).CharIndex
            End If
            
            'Erase OBJs
            mapData(x, y).ObjGrh.GrhIndex = 0
        Next x
    Next y
    
    Close Handle

    Map_Load = True
    Exit Function
ErrHandle:
    Map_Load = False
    Debug.Print "Error al cargar el mapa " & Err.Number; Err.Description
End Function

Public Function Map_Change(ByVal UserMap As Integer, ByVal UserMapName As String) As Boolean

On Error GoTo ErrHandle
    
    Engine.UserMap = UserMap

    Engine.UserMapName = UserMapName
    
    If FileExist(App.path & "\MAPAS\Mapa" & Engine.UserMap & ".map", vbNormal) Then
        
        'Load Map
        
        If Map_Load(UserMap) = True Then
            If mapRain(UserMap) = 0 Then
                If bRain Then
                    Call Sound_Stop(Engine.RainBufferIndex)
                    RainBufferIndex = 0
                    IsPlaying = PlayLoop.plNone
                End If
            End If
        Else
            Map_Change = False
        End If
        
    Else
        Map_Change = False
    End If

    Map_Change = True
    Exit Function
    
ErrHandle:
    Map_Change = False
End Function

Private Sub Map_LoadRainArray()

    Dim N As Integer: Dim numMap As Integer
    
    N = FreeFile()
    Open App.path & "\init\fk.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    Get #N, , numMap
    
    'Resize array
    ReDim mapRain(1 To numMap) As Byte
    
    Dim i As Long
    
    For i = 1 To numMap
        Get #N, , mapRain(i)
    Next i
    
    Close #N
    
End Sub

Public Sub Map_RainToggle()
    
    If Not Map_InBounds(UserPos.x, UserPos.y) Then Exit Sub
    
    bTecho = (mapData(UserPos.x, UserPos.y).Trigger = 1 Or _
            mapData(UserPos.x, UserPos.y).Trigger = 2 Or _
            mapData(UserPos.x, UserPos.y).Trigger = 4)
    If bRain Then
        If mapRain(UserMap) Then
            'Stop playing the rain sound
            Call Engine.Sound_Stop(RainBufferIndex)
            RainBufferIndex = 0
            If bTecho Then
                Call Engine.Sound_Play("lluviainend.wav", 0, 0, Sound_LoopStyle.Disabled)
            Else
                Call Engine.Sound_Play("lluviaoutend.wav", 0, 0, Sound_LoopStyle.Disabled)
            End If
            IsPlaying = PlayLoop.plNone
        End If
    End If
    
    bRain = Not bRain

End Sub

Sub Map_Render(ByRef PixelOffsetX As Single, ByRef PixelOffsetY As Single)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 04/03/2009
'Last modified by: Parra
'Renders everything to the viewport
'**************************************************************
    
    Dim y           As Long     'Keeps track of where on map we are
    Dim x           As Long     'Keeps track of where on map we are
    Dim screenminY  As Integer  'Start Y pos on current screen
    Dim screenmaxY  As Integer  'End Y pos on current screen
    Dim screenminX  As Integer  'Start X pos on current screen
    Dim screenmaxX  As Integer  'End X pos on current screen
    Dim minY        As Integer  'Start Y pos on current map
    Dim maxY        As Integer  'End Y pos on current map
    Dim minX        As Integer  'Start X pos on current map
    Dim maxX        As Integer  'End X pos on current map
    Dim ScreenX     As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY     As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset  As Integer
    Dim minYOffset  As Integer
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    
    
    'Figure out Ends and Starts of screen
    screenminY = UserPos.y - AddtoUserPos.y - HalfWindowTileHeight
    screenmaxY = UserPos.y - AddtoUserPos.y + HalfWindowTileHeight
    screenminX = UserPos.x - AddtoUserPos.x - HalfWindowTileWidth
    screenmaxX = UserPos.x - AddtoUserPos.x + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1
    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1
    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1

   'Draw floor layer
    For y = screenminY To screenmaxY
        For x = screenminX To screenmaxX

            'Layer 1 **********************************
            If mapData(x, y).Graphic(1).GrhIndex <> 0 Then
                Graphics_Draw_GrhToSurface mapData(x, y).Graphic(1), _
                    (ScreenX - 9) * TilePixelWidth + PixelOffsetX + TileBufferPixelOffsetX, _
                    (ScreenY - 5) * TilePixelHeight + PixelOffsetY + TileBufferPixelOffsetY, 0, 1
            End If
            '******************************************
            
            'Layer 2 **********************************
            If mapData(x, y).Graphic(2).GrhIndex <> 0 Then
                Graphics_Draw_GrhToSurface mapData(x, y).Graphic(2), _
                    (ScreenX - 9) * TilePixelWidth + PixelOffsetX + TileBufferPixelOffsetX, _
                    (ScreenY - 5) * TilePixelHeight + PixelOffsetY + TileBufferPixelOffsetY, 0, 1
            End If
            '******************************************
            
            ScreenX = ScreenX + 1
        Next x

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - x + screenminX
        ScreenY = ScreenY + 1
    Next y
    
    'Draw Transparent Layers
    ScreenY = minYOffset
    For y = minY To maxY
        ScreenX = minXOffset
        For x = minX To maxX
            PixelOffsetXTemp = (ScreenX - 9) * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = (ScreenY - 5) * TilePixelHeight + PixelOffsetY
                
            With mapData(x, y)
            
                'Object Layer **********************************
                If .ObjGrh.GrhIndex <> 0 Then
                    Graphics_Draw_GrhToSurface .ObjGrh, _
                        PixelOffsetXTemp, PixelOffsetYTemp, 1, 1
                End If
                '***********************************************
    
                'Char layer ************************************
                If .CharIndex <> 0 Then
                    Character_Render .CharIndex, PixelOffsetXTemp, PixelOffsetYTemp
                End If
                '*************************************************
    
    
                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then
                    'Draw
                    Graphics_Draw_GrhToSurface .Graphic(3), _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1
                End If
                '************************************************
            End With
    
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next y
    
    If Not bTecho Then
        'Draw blocked tiles and grid
        ScreenY = minYOffset
        For y = minY To maxY
            ScreenX = minXOffset
            For x = minX To maxX
    
                'Layer 4 **********************************
                If mapData(x, y).Graphic(4).GrhIndex Then
                    'Draw
                    Graphics_Draw_GrhToSurface mapData(x, y).Graphic(4), _
                        (ScreenX - 9) * TilePixelWidth + PixelOffsetX, _
                        (ScreenY - 5) * TilePixelHeight + PixelOffsetY, _
                        1, 1
                End If
                '**********************************
    
                ScreenX = ScreenX + 1
            Next x
            ScreenY = ScreenY + 1
        Next y
    End If
    
    'Lluvia
    If mapRain(UserMap) = 1 Then
        If bRain Then
    
            ' Rain..
        
        End If
    End If
    
End Sub
Public Sub Map_RemoveOldUser()

    If mapData(UserPos.x, UserPos.y).CharIndex = UserCharIndex Then
        mapData(UserPos.x, UserPos.y).CharIndex = 0
    End If

End Sub
Public Sub Map_CreateObject(GrhIndex As Integer, x As Byte, y As Byte)

    mapData(x, y).ObjGrh.GrhIndex = GrhIndex
    
    Call Grh_Initialize(mapData(x, y).ObjGrh, mapData(x, y).ObjGrh.GrhIndex)
    
End Sub
Public Sub Map_DestroyObject(x As Byte, y As Byte)

    mapData(x, y).ObjGrh.GrhIndex = 0
    
End Sub
Public Sub Map_SetBlocked(x As Byte, y As Byte, Block As Byte)

    mapData(x, y).Blocked = Block

End Sub
Function Map_LegalPos(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************
    'Limites del mapa
    If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If mapData(x, y).Blocked = 1 Then
        Exit Function
    End If
    
    '¿Hay un personaje?
    If mapData(x, y).CharIndex > 0 Then
        Exit Function
    End If
   
    If UserNavegando <> Map_CheckWater(x, y) Then
        Exit Function
    End If
    
    LegalPos = True
End Function

Function Map_InBounds(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If x < XMinMapSize Or x > XMaxMapSize Or y < YMinMapSize Or y > YMaxMapSize Then
        Exit Function
    End If
    
    Map_InBounds = True
End Function

Function Map_CheckWater(ByVal x As Integer, ByVal y As Integer) As Boolean
    Map_CheckWater = ((mapData(x, y).Graphic(1).GrhIndex >= 1505 And mapData(x, y).Graphic(1).GrhIndex <= 1520) Or _
            (mapData(x, y).Graphic(1).GrhIndex >= 5665 And mapData(x, y).Graphic(1).GrhIndex <= 5680) Or _
            (mapData(x, y).Graphic(1).GrhIndex >= 13547 And mapData(x, y).Graphic(1).GrhIndex <= 13562)) And _
                mapData(x, y).Graphic(2).GrhIndex = 0
                
End Function

Private Function Map_CheckBonfire(ByRef Location As Point) As Boolean
    Dim j As Long
    Dim k As Long
    
    For j = UserPos.x - 8 To UserPos.x + 8
        For k = UserPos.y - 6 To UserPos.y + 6
            If Map_InBounds(j, k) Then
                If mapData(j, k).ObjGrh.GrhIndex = 1521 Then
                    Location.x = j
                    Location.y = k
                    
                    Map_CheckBonfire = True
                    Exit Function
                End If
            End If
        Next k
    Next j
End Function

Public Sub Map_ChangeArea(ByVal x As Byte, ByVal y As Byte)
    Dim loopX As Long, loopY As Long
    
    MinLimit.x = (x \ 9 - 1) * 9
    MaxLimit.x = MinLimit.x + 26
    
    MinLimit.y = (y \ 9 - 1) * 9
    MaxLimit.y = MinLimit.y + 26
    
    For loopX = 1 To 100
        For loopY = 1 To 100
        
            If (loopY < MinLimit.y) Or (loopY > MaxLimit.y) Or (loopX < MinLimit.x) Or (loopX > MaxLimit.x) Then
                'Erase NPCs
                
                If mapData(loopX, loopY).CharIndex > 0 Then
                    If mapData(loopX, loopY).CharIndex <> UserCharIndex Then
                        Character_Destroy mapData(loopX, loopY).CharIndex
                    End If
                End If
                
                'Erase OBJs
                mapData(loopX, loopY).ObjGrh.GrhIndex = 0
            End If
            
        Next loopY
    Next loopX

    Character_RefreshAll

End Sub

'***************************************************
' Character Functions
'***************************************************


Public Sub Character_Create(ByVal index As Integer, ByVal Body As Integer, ByVal Head As Integer, _
                             ByVal Heading As Byte, ByVal x As Integer, ByVal y As Integer, _
                             ByVal Weapon As Integer, ByVal Shield As Integer, ByVal Helmet As Integer, _
                             ByVal Name As String, ByVal Crim As Byte, ByVal privs As Byte)

    Dim i As Long
    
    
    ' Search the index on the list
    With charList(index)
                  
'===========================

        .Name = Name
        .Criminal = Crim
        
        If privs <> 0 Then
            'If the player belongs to a council AND is an admin, only whos as an admin
            If (privs And PlayerType.ChaosCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.ChaosCouncil
            End If
            
            If (privs And PlayerType.RoyalCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.RoyalCouncil
            End If
            
            'If the player is a RM, ignore other flags
            If privs And PlayerType.RoleMaster Then
                privs = PlayerType.RoleMaster
            End If
            
            'Log2 of the bit flags sent by the server gives our numbers ^^
            .Priv = Log(privs) / Log(2)
        Else
            .Priv = 0
        End If

'===========================
                    
        If Weapon = 0 Then Weapon = 2 ' Sources cutres de AO...
        If Shield = 0 Then Shield = 2
        If Helmet = 0 Then Helmet = 2
                    
                    
        .Head = HeadData(Head): .iHead = Head
        .Body = BodyData(Body): .iBody = Body
        .Arma = WeaponAnimData(Weapon)
        .Escudo = ShieldAnimData(Shield)
        .Helmet = HelmetData(Helmet)
                    
        .Heading = Heading
                
        'Reset moving stats
        .Moving = 0: .MoveOffsetX = 0: .MoveOffsetY = 0
                
        'Update position
        .Pos.x = x: .Pos.y = y

        'Make Active Char
        .Active = 1
        
    End With
    
    'Plot on map-> Destroy this
    mapData(x, y).CharIndex = index
End Sub

Public Sub Character_Change(ByVal CharIndex As Integer, ByVal TempInt As Integer, ByVal HeadIndex As Integer, ByVal Heading As Byte, _
                            ByVal Weapon As Integer, ByVal Shield As Integer, ByVal Helmet As Integer, _
                            ByVal Fx As Integer, ByVal Loops As Integer)

    With charList(CharIndex)
        
        If TempInt < LBound(BodyData()) Or TempInt > UBound(BodyData()) Then
            .Body = BodyData(0)
            .iBody = 0
        Else
            .Body = BodyData(TempInt)
            .iBody = TempInt
        End If
        
        If TempInt < LBound(HeadData()) Or TempInt > UBound(HeadData()) Then
            .Head = HeadData(0)
            .iHead = 0
        Else
            .Head = HeadData(HeadIndex)
            .iHead = HeadIndex
        End If
        
        .Muerto = (HeadIndex = CASPER_HEAD)
        
        .Heading = Heading
        
        If Weapon <> 0 Then .Arma = WeaponAnimData(Weapon)
        If Shield <> 0 Then .Escudo = ShieldAnimData(Shield)
        If Helmet <> 0 Then .Helmet = HelmetData(Helmet)
        
        Call Character_SetFx(CharIndex, Fx, Loops)
    End With

End Sub

Private Sub Character_Refresh(ByRef Character As CharacterAparence, ByRef index As Long)
    
    'Refresh all chars....
    mapData(Character.Pos.x, Character.Pos.y).CharIndex = index
    
End Sub


Public Function Character_GetName(ByRef index As Integer) As String

    Character_GetName = charList(index).Name

End Function


Public Function Character_IsActive(ByRef index As Integer) As Boolean

  
    Character_IsActive = charList(index).Active = 1

End Function

Public Sub Character_ChangeVisibility(ByVal index As Integer, ByVal Invi As Boolean)

    charList(index).Invisible = Invi

End Sub

Public Sub Character_ChangeTagAndStatus(ByVal index As Integer, ByVal Criminal As Boolean, ByVal userTag As String)

    'Update char status and tag!
    With charList(index)
        If Criminal Then
            .Criminal = 1
        Else
            .Criminal = 0
        End If
        
        .Name = userTag
    End With

End Sub

Public Sub Character_Destroy(ByVal index As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
    
    
    With charList(index)
        .Active = 0
        
        'Update lastchar
        If index = lastChar Then
            Do Until charList(lastChar).Active = 1
                lastChar = lastChar - 1
                If lastChar = 0 Then Exit Do
            Loop
        End If
        
        'Remove char's dialog
        Dialogos.RemoveDialog index
        

        If .Pos.x > 0 And .Pos.y > 0 Then
            If mapData(.Pos.x, .Pos.y).CharIndex = index Then
                mapData(.Pos.x, .Pos.y).CharIndex = 0
            End If
        End If
        .Pos.x = 0
        .Pos.y = 0
    End With

End Sub

Public Sub Character_RefreshAll()
    
    Dim i As Long

    For i = 1 To lastChar
        If charList(i).Active > 0 Then
            Character_Refresh charList(i), i
        End If
    Next i
    
End Sub

Public Sub Character_DestroyAll()
'**************************************************************
'Clean up
'**************************************************************
    Dim i As Long
    
    For i = 1 To lastChar
        Character_Destroy i
    Next i
    
End Sub

Public Sub User_MoveByHead(ByVal nHeading As Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim addX As Integer
    Dim addY As Integer
    Dim x As Integer
    Dim y As Integer
    
    Dim Character As CharacterAparence
    
    With charList(UserCharIndex)
        x = .Pos.x
        y = .Pos.y
        
        'Figure out which way to move
        Select Case nHeading
            Case Heading.North: addY = -1
            Case Heading.East:  addX = 1
            Case Heading.South: addY = 1
            Case Heading.West:  addX = -1
        End Select
        
        mapData(x + addX, y + addY).CharIndex = UserCharIndex
        .Pos.x = x + addX
        .Pos.y = y + addY
        mapData(x, y).CharIndex = 0
        
        .MoveOffsetX = -1 * (TilePixelWidth * addX)
        .MoveOffsetY = -1 * (TilePixelHeight * addY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = addX
        .scrollDirectionY = addY
    End With
    
    If UserEstado = 0 Then Call Character_WalkFx(UserCharIndex)
    
End Sub

Public Sub User_UpdatePos(ByVal x As Byte, ByVal y As Byte)

    'Set new pos
    UserPos.x = x: UserPos.y = y
    mapData(UserPos.x, UserPos.y).CharIndex = UserCharIndex
End Sub

Public Sub User_CheckPosition()

    'mapData(UserPos.x, UserPos.y).CharIndex = UserCharIndex
    UserPos = charList(Engine.UserCharIndex).Pos
        
    'Are we under a roof?
    bTecho = IIf(mapData(UserPos.x, UserPos.y).Trigger = 1 Or _
            mapData(UserPos.x, UserPos.y).Trigger = 2 Or _
            mapData(UserPos.x, UserPos.y).Trigger = 4, True, False)

    frmMain.Coord.Caption = "(" & UserMap & "[" & UserMapName & "]" & "," & UserPos.x & "," & UserPos.y & ")"

End Sub
Private Function Character_CheckArea(ByVal CharIndex As Integer) As Boolean
    
    With charList(CharIndex).Pos
        Character_CheckArea = .x > UserPos.x - MinXBorder And .x < UserPos.x + MinXBorder And .y > UserPos.y - MinYBorder And .y < UserPos.y + MinYBorder
    End With

End Function
Private Sub Character_WalkFx(ByVal CharIndex As Integer)
    
    With charList(CharIndex)
    
        If Not UserNavegando Then
            If Not .Muerto And Character_CheckArea(CharIndex) Then
                .Pie = Not .Pie
    
                If .Pie Then
                    Call Engine.Sound_Play(SND_PASOS1, .Pos.x, .Pos.y)
                Else
                    Call Engine.Sound_Play(SND_PASOS2, .Pos.x, .Pos.y)
                End If
            End If
        Else
            Call Sound_Play(SND_NAVEGANDO, .Pos.x, .Pos.y)
        End If
    
    End With
    
End Sub
Public Sub Character_MoveByPos(ByVal index As Integer, ByVal nX As Integer, ByVal nY As Integer)

On Error GoTo ErrHandle

    Dim x As Integer
    Dim y As Integer
    Dim addX As Integer
    Dim addY As Integer
    Dim nHeading As Heading
    
    
    With charList(index)
    
        If .FxIndex >= 40 And .FxIndex <= 49 Then   'If it's meditating, we remove the FX
            .FxIndex = 0
        End If
        
        ' Play steps sounds if the user is not an admin of any kind
        If .Priv <> 1 And .Priv <> 2 And .Priv <> 3 And .Priv <> 5 And .Priv <> 25 Then
            Character_WalkFx index
        End If
    
        x = .Pos.x
        y = .Pos.y
        
        mapData(x, y).CharIndex = 0
        
        addX = nX - x
        addY = nY - y
        
        If Sgn(addX) = 1 Then
            nHeading = Heading.East
        ElseIf Sgn(addX) = -1 Then
            nHeading = Heading.West
        ElseIf Sgn(addY) = -1 Then
            nHeading = Heading.North
        ElseIf Sgn(addY) = 1 Then
            nHeading = Heading.South
        End If
        
        mapData(nX, nY).CharIndex = index
        
        .Pos.x = nX
        .Pos.y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * addX)
        .MoveOffsetY = -1 * (TilePixelHeight * addY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addX)
        .scrollDirectionY = Sgn(addY)
        
        'parche para que no medite cuando camina
        If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.GRANDE Or .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XGRANDE Or .FxIndex = FxMeditar.XXGRANDE Then
            .FxIndex = 0
        End If
        
    End With
    
    If Not Character_CheckArea(index) Then Call Dialogos.RemoveDialog(index)
    
    If (nY < MinLimit.y) Or (nY > MinLimit.y) Or (nX < MinLimit.x) Or (nX > MinLimit.x) Then
        Character_Destroy index
    End If
    
ErrHandle:


End Sub

Public Sub User_MoveScreen(ByVal nHeading As Heading) ' PARRA: Esto no es del renderscreen???
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim x As Integer
    Dim y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case Heading.North
            y = -1
        
        Case Heading.East
            x = 1
        
        Case Heading.South
            y = 1
        
        Case Heading.West
            x = -1
    End Select
    
    'Fill temp pos
    tX = UserPos.x + x
    tY = UserPos.y + y
    
    'Check to see if its out of bounds
    If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.x = x
        UserPos.x = tX
        AddtoUserPos.y = y
        UserPos.y = tY
        UserMoving = 1
        
        bTecho = IIf(mapData(UserPos.x, UserPos.y).Trigger = 1 Or _
                mapData(UserPos.x, UserPos.y).Trigger = 2 Or _
                mapData(UserPos.x, UserPos.y).Trigger = 4, True, False)
    End If
End Sub
Public Sub User_Move(ByVal Direccion As Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/28/2008
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
' 06/28/2008: NicoNZ - Saqué lo que impedía que si el usuario estaba paralizado se ejecute el sub.
'***************************************************
    Dim LegalOk As Boolean
    
    If Cartel Then Cartel = False
    
    Select Case Direccion
        Case Heading.North
            LegalOk = User_MoveToLegalPos(UserPos.x, UserPos.y - 1)
        Case Heading.East
            LegalOk = User_MoveToLegalPos(UserPos.x + 1, UserPos.y)
        Case Heading.South
            LegalOk = User_MoveToLegalPos(UserPos.x, UserPos.y + 1)
        Case Heading.West
            LegalOk = User_MoveToLegalPos(UserPos.x - 1, UserPos.y)
    End Select
    
    If LegalOk And Not UserParalizado Then
        Call WriteWalk(Direccion)
        If Not UserDescansar And Not UserMeditar Then
            User_MoveByHead Direccion
            User_MoveScreen Direccion
        End If
    Else
        If charList(UserCharIndex).Heading <> Direccion Then
            Call WriteChangeHeading(Direccion)
        End If
    End If
    
    ' Update 3D sounds!
    Call Engine.SoundEffects_MoveListener(UserPos.x, UserPos.y)
End Sub

Function User_MoveToLegalPos(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 01/08/2009
'Checks to see if a tile position is legal, including if there is a casper in the tile
'10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
'01/08/2009: ZaMa - Now invisible admins can't change position with caspers.
'*****************************************************************
    Dim CharIndex As Integer
    
    'Limites del mapa
    If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If mapData(x, y).Blocked = 1 Then
        Exit Function
    End If
    
    CharIndex = mapData(x, y).CharIndex
    '¿Hay un personaje?
    If CharIndex > 0 Then
    
        If mapData(UserPos.x, UserPos.y).Blocked = 1 Then
            Exit Function
        End If
        
        With charList(CharIndex)
            ' Si no es casper, no puede pasar
            If .iHead <> CASPER_HEAD And .iBody <> FRAGATA_FANTASMAL Then
                Exit Function
            Else
                ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
                If Map_CheckWater(UserPos.x, UserPos.y) Then
                    If Not Map_CheckWater(x, y) Then Exit Function
                Else
                    ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
                    If Map_CheckWater(x, y) Then Exit Function
                End If
                
                Dim UserChar As CharacterAparence: UserChar = charList(UserCharIndex)
                
                ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles
                If UserChar.Priv > 0 And UserChar.Priv < 6 Then
                    If UserChar.Invisible = True Then Exit Function
                End If
                
            End If
        End With
    End If
   
    If UserNavegando <> Map_CheckWater(x, y) Then
        Exit Function
    End If
    
    User_MoveToLegalPos = True
End Function

Private Sub Character_Render(ByVal CharIndex As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 03/09/09
'Last modified by: •Parra
'Draw char's to screen without offcentering them
'***************************************************
    Dim moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim color As Long
    
    With charList(CharIndex)
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                        (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
            
            .Arma.WeaponWalk(.Heading).Started = 0
            .Arma.WeaponWalk(.Heading).FrameCounter = 1
            
            .Escudo.ShieldWalk(.Heading).Started = 0
            .Escudo.ShieldWalk(.Heading).FrameCounter = 1
            
            .Moving = False
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        If .Head.Head(.Heading).GrhIndex Then
            If Not .Invisible Then
                'Draw Body
                If .Body.Walk(.Heading).GrhIndex Then _
                     Graphics_Draw_GrhToSurface .Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1
            
                'Draw Head
                If .Head.Head(.Heading).GrhIndex Then
                     Graphics_Draw_GrhToSurface .Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0
                    
                    'Draw Helmet
                    If .Helmet.Head(.Heading).GrhIndex Then _
                        Graphics_Draw_GrhToSurface .Helmet.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0
                    
                    'Draw Weapon
                    If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                        Graphics_Draw_GrhToSurface .Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1
                    
                    'Draw Shield
                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                        Graphics_Draw_GrhToSurface .Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1
                
                
                    'Draw name over head
                    If LenB(.Name) > 0 Then
                        If Nombres And Abs(MouseTileX - .Pos.x) < 2 And (Abs(MouseTileY - .Pos.y)) < 2 Then
                            Pos = InStr(.Name, "<")
                            If Pos = 0 Then Pos = Len(.Name) + 2
                            
                            If .Priv = 0 Then
                                If .Criminal Then
                                    color = RGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).B)
                                Else
                                    color = RGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).B)
                                End If
                            Else
                                color = RGB(ColoresPJ(.Priv).r, ColoresPJ(.Priv).g, ColoresPJ(.Priv).B)
                            End If
                            
                            'Nick
                            line = Left$(.Name, Pos - 2)
                            Engine.Fonts_Render 1, PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 30, PixelOffsetX + TilePixelWidth \ 2 + 105, PixelOffsetY + 45, line, color
                            
                            'Clan
                            line = mid$(.Name, Pos)
                            Engine.Fonts_Render 1, PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 50, PixelOffsetX + TilePixelWidth \ 2 + 105, PixelOffsetY + 65, line, color
                        End If
                    End If
                End If
            End If
        Else
            'Draw Body
            If .Body.Walk(.Heading).GrhIndex Then _
                Graphics_Draw_GrhToSurface .Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1
        End If

        
        'Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, CharIndex)
        
        'Draw FX
        If .FxIndex <> 0 Then
            Graphics_Draw_GrhToSurface .Fx, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY, 1, 1
            
            'Check if animation is over
            If .Fx.Started = 0 Then _
                .FxIndex = 0
        End If
    End With
End Sub

Public Sub Character_SetFx(ByVal CharIndex As Integer, ByVal Fx As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************
    With charList(CharIndex)
        .FxIndex = Fx
        
        If .FxIndex > 0 Then
            Call Grh_Initialize(.Fx, FxData(Fx).Animation)
        
            .Fx.Loops = Loops
        End If
    End With
End Sub

' *********************************************************

Public Function InitTileEngine(ByVal Frm As Form, ByVal setMainViewTop As Integer, ByVal setMainViewLeft As Integer, ByVal setMainViewBottom As Integer, ByVal setMainViewRight As Integer, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal setWindowTileHeight As Integer, ByVal setWindowTileWidth As Integer, ByVal setTileBufferSize As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer, ByVal engineSpeed As Single, ByVal BitsPerPixel As Byte, ByVal Windowed As Boolean) As Boolean
    
    With regRenderDX
        .X1 = setMainViewLeft
        .X2 = setMainViewRight
        .Y1 = setMainViewTop
        .Y2 = setMainViewBottom
    End With
    
    'Fill startup variables
    MainViewTop = setMainViewTop
    MainViewLeft = setMainViewLeft
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize
    
    HalfWindowTileHeight = setWindowTileHeight \ 2
    HalfWindowTileWidth = setWindowTileWidth \ 2
    
    'Compute offset in pixels when rendering tile buffer.
    'We diminish by one to get the top-left corner of the tile for rendering.
    TileBufferPixelOffsetX = ((TileBufferSize - 1) * TilePixelWidth)
    TileBufferPixelOffsetY = ((TileBufferSize - 1) * TilePixelHeight)
    
    engineBaseSpeed = engineSpeed
    
    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
    
    
    'Resize mapdata array
    ReDim mapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As mapBlock
    
    'Set intial user position
    UserPos.x = MinXBorder
    UserPos.y = MinYBorder
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY
    

    Grh_Load_Data 3
    Grh_Load_Bodys
    Grh_Load_Heads
    Grh_Load_Helmets
    Grh_Load_Fxs
    Grh_Load_WeaponsAnim
    Grh_Load_ShieldsAnim
    
    DirectX_Initialize Frm
    
    Map_LoadRainArray
    
    InitTileEngine = True
End Function
