VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Apariencia del personaje

Private Type CharacterAparence
     Heading     As Heading
     Pos         As Point
     
     Body        As BodyData: iBody       As Integer
     Head        As HeadData: iHead       As Integer
     Helmet      As HeadData
     Arma        As WeaponAnimData
     Escudo      As ShieldAnimData
    
     FX          As Grh
     FxIndex     As Integer
    
     Criminal    As Byte
    
     Nombre      As String
    
     scrollDirectionX As Integer
     scrollDirectionY As Integer
    
     Moving      As Byte
     MoveOffsetX As Single
     MoveOffsetY As Single
    
     Pie         As Boolean
     Muerto      As Boolean
     Invisible   As Boolean
     Priv        As Byte

End Type: Private CharacterAparence As CharacterAparence

Private Sub Class_Initialize(ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal x As Integer, ByVal y As Integer, ByVal Weapon As Integer, ByVal Shield As Integer, ByVal Helmet As Integer)
'On Error Resume Next
    'Apuntamos al ultimo Char
    'If CharIndex > LastChar Then LastChar = CharIndex
    
    If Weapon = 0 Then Weapon = 2
    If Shield = 0 Then Shield = 2
    If Helmet = 0 Then Helmet = 2
    
    With CharacterAparence
    
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
        
        'Make active
        .Active = 1
    
    End With
            
    'Plot on map
    mapData(x, y).CharIndex = CharIndex
End Sub

Public Sub Refresh(ByVal Index As Integer)
    
    mapData(CharacterAparence.Pos.x, CharacterAparence.Pos.y).CharIndex = Index
    
End Sub

Public Sub Reset()

    With CharacterAparence

        .Criminal = 0
        .FxIndex = 0
        .Invisible = False: .Muerto = False
        .Moving = 0: .Pie = False
        .Nombre = vbNullString ' GSZ
        .Pos.x = 0: .Pos.y = 0

    End With
    
End Sub

'TODO : Never was sure this is really necessary....
'Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    'Dim loopc As Long
    
    'For loopc = 1 To LastChar
    '    If Charlist(loopc).Active = 1 Then
    '        mapData(Charlist(loopc).Pos.x, Charlist(loopc).Pos.y).CharIndex = loopc
    '    End If
    'Next loopc
'End Sub


'PARRA: Separar ChararcterManager - MapChararcterManager

'Public Sub Destroy(ByVal CharIndex As Integer)
''*****************************************************************
''Erases a character from CharList and map
''*****************************************************************
'On Error Resume Next
'    Charlist(CharIndex).Active = 0
'
'    mapData(Charlist(CharIndex).Pos.x, Charlist(CharIndex).Pos.y).CharIndex = 0
'
'    'Remove char's dialog
'    Call Dialogos.RemoveDialog(CharIndex)
'
'    Call ResetCharInfo(CharIndex)
'
'End Sub



Public Sub MoveByHead(ByVal nHeading As Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim addX As Integer
    Dim addY As Integer
    Dim x As Integer
    Dim y As Integer
    
    With CharacterAparence
        x = .Pos.x
        y = .Pos.y
        
        'Figure out which way to move
        Select Case nHeading
            Case Heading.North: addY = -1
            Case Heading.East:  addX = 1
            Case Heading.South: addY = 1
            Case Heading.West:  addX = -1
        End Select
        
        mapData(x + addX, y + addY).CharIndex = CharIndex
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
    
    If UserEstado = 0 Then Call DoPasosFx(CharIndex)
    
    'areas viejos
    If (y + addY < MinLimit.y) Or (y + addY > MinLimit.y) Or (x + addX < MinLimit.x) Or (x + addX > MinLimit.x) Then
        If CharIndex <> UserCharIndex Then
            Call EraseChar(CharIndex)
        End If
    End If
End Sub

Private Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With Charlist(CharIndex).Pos
        EstaPCarea = .x > UserPos.x - MinXBorder And .x < UserPos.x + MinXBorder And .y > UserPos.y - MinYBorder And .y < UserPos.y + MinYBorder
    End With
End Function

Sub DoPasosFx(ByVal CharIndex As Integer)
    If Not UserNavegando Then
        With Charlist(CharIndex)
            If Not .Muerto And EstaPCarea(CharIndex) Then
                .Pie = Not .Pie
                
                If .Pie Then
                    Call Engine.Sound_Play(SND_PASOS1, .Pos.x, .Pos.y)
                Else
                    Call Engine.Sound_Play(SND_PASOS2, .Pos.x, .Pos.y)
                End If
            End If
        End With
    Else
' TODO : Actually we would have to check if the CharIndex char is in the water or not....
        Call Engine.Sound_Play(SND_NAVEGANDO, Charlist(CharIndex).Pos.x, Charlist(CharIndex).Pos.y)
    End If
End Sub

Public Sub MoveByPos(ByVal nX As Integer, ByVal nY As Integer)
'On Error Resume Next

On Error GoTo ErrHandle

    Dim x As Integer
    Dim y As Integer
    Dim addX As Integer
    Dim addY As Integer
    Dim nHeading As Heading
    
    With Charlist(CharIndex)
        x = .Pos.x
        y = .Pos.y
        
        mapData(x, y).CharIndex = 0
        
        addX = nX - x
        addY = nY - y
        
        If Sgn(addX) = 1 Then: nHeading = Heading.East
        ElseIf Sgn(addX) = -1 Then: nHeading = Heading.West
        ElseIf Sgn(addY) = -1 Then: nHeading = Heading.North
        ElseIf Sgn(addY) = 1 Then:  nHeading = Heading.South
        End If
        
        mapData(nX, nY).CharIndex = CharIndex
        
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
    
    If Not EstaPCarea(CharIndex) Then Call Dialogos.RemoveDialog(CharIndex)
    
    If (nY < MinLimit.y) Or (nY > MinLimit.y) Or (nX < MinLimit.x) Or (nX > MinLimit.x) Then
        Call EraseChar(CharIndex)
    End If
    
ErrHandle:


End Sub

Sub MoveScreen(ByVal nHeading As Heading) ' PARRA: Esto no es del renderscreen???
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



'Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
'    Dim loopc As Long
'    Dim Dale As Boolean
'
'    loopc = 1
'    Do While Charlist(loopc).Active And Dale
'        loopc = loopc + 1
'        Dale = (loopc <= UBound(Charlist))
'    Loop
'
'    NextOpenChar = loopc
'End Function

Sub MoveTo(ByVal Direccion As Heading)
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
            LegalOk = MoveToLegalPos(UserPos.x, UserPos.y - 1)
        Case Heading.East
            LegalOk = MoveToLegalPos(UserPos.x + 1, UserPos.y)
        Case Heading.South
            LegalOk = MoveToLegalPos(UserPos.x, UserPos.y + 1)
        Case Heading.West
            LegalOk = MoveToLegalPos(UserPos.x - 1, UserPos.y)
    End Select
    
    If LegalOk And Not UserParalizado Then
        Call WriteWalk(Direccion)
        If Not UserDescansar And Not UserMeditar Then
            MoveCharbyHead UserCharIndex, Direccion
            MoveScreen Direccion
        End If
    Else
        If Charlist(UserCharIndex).Heading <> Direccion Then
            Call WriteChangeHeading(Direccion)
        End If
    End If
    
    ' Update 3D sounds!
    Call Engine.SoundEffects_MoveListener(UserPos.x, UserPos.y)
End Sub

Function MoveToLegalPos(ByVal x As Integer, ByVal y As Integer) As Boolean
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
        
        With Charlist(CharIndex)
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
                
                ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles
                If Charlist(UserCharIndex).Priv > 0 And Charlist(UserCharIndex).Priv < 6 Then
                    If Charlist(UserCharIndex).Invisible = True Then Exit Function
                End If
            End If
        End With
    End If
   
    If UserNavegando <> Map_CheckWater(x, y) Then
        Exit Function
    End If
    
    MoveToLegalPos = True
End Function

Private Sub CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 03/09/09
'Last modified by: •Parra
'Draw char's to screen without offcentering them
'***************************************************
    Dim moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim Color As Long
    
    With Charlist(CharIndex)
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
                     DrawGrhToSurface .Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1
            
                'Draw Head
                If .Head.Head(.Heading).GrhIndex Then
                     DrawGrhToSurface .Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0
                    
                    'Draw Helmet
                    If .Helmet.Head(.Heading).GrhIndex Then _
                        DrawGrhToSurface .Helmet.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0
                    
                    'Draw Weapon
                    If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                        DrawGrhToSurface .Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1
                    
                    'Draw Shield
                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                        DrawGrhToSurface .Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1
                
                
                    'Draw name over head
                    If LenB(.Nombre) > 0 Then
                        If Nombres And Abs(MouseTileX - .Pos.x) < 2 And (Abs(MouseTileY - .Pos.y)) < 2 Then
                            Pos = InStr(.Nombre, "<")
                            If Pos = 0 Then Pos = Len(.Nombre) + 2
                            
                            If .Priv = 0 Then
                                If .Criminal Then
                                    Color = RGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).B)
                                Else
                                    Color = RGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).B)
                                End If
                            Else
                                Color = RGB(ColoresPJ(.Priv).r, ColoresPJ(.Priv).g, ColoresPJ(.Priv).B)
                            End If
                            
                            'Nick
                            line = Left$(.Nombre, Pos - 2)
                            Call RenderText(1, PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 30, PixelOffsetX + TilePixelWidth \ 2 + 105, PixelOffsetY + 45, line, Color)
                            
                            'Clan
                            line = mid$(.Nombre, Pos)
                            Call RenderText(1, PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 50, PixelOffsetX + TilePixelWidth \ 2 + 105, PixelOffsetY + 65, line, Color)
                        End If
                    End If
                End If
            End If
        Else
            'Draw Body
            If .Body.Walk(.Heading).GrhIndex Then _
                DrawGrhToSurface .Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1
        End If

        
        'Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, CharIndex)
        
        'Draw FX
        If .FxIndex <> 0 Then
            DrawGrhToSurface .FX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY, 1, 1
            
            'Check if animation is over
            If .FX.Started = 0 Then _
                .FxIndex = 0
        End If
    End With
End Sub

Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal FX As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************
    With Charlist(CharIndex)
        .FxIndex = FX
        
        If .FxIndex > 0 Then
            Call Grh_Initialize(.FX, FxData(FX).Animation)
        
            .FX.Loops = Loops
        End If
    End With
End Sub
