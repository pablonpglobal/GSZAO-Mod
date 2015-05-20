Attribute VB_Name = "modLoaderGraphics"
Option Explicit

'Manage Textures
Private Const lMaxEntrys As Integer = 32000

Private Type structGraphic
        FileName   As Long
        D3DTexture As Direct3DTexture8
        Used       As Long
        Available  As Boolean
        Width      As Integer
        Height     As Integer
End Type: Public oGraphic() As structGraphic: Public lKeys() As Long

Private lSurfaceSize As Long
Private TexInfo      As D3DXIMAGE_INFO

Private Type tCache
    Number        As Long
    SrcHeight     As Integer
    SrcWidth      As Integer
End Type: Private Cache As tCache

Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, ByVal numBytes As Long)

Public Function texInitialize() As Boolean
On Error GoTo errHandle

    ReDim oGraphic(lMaxEntrys)
    ReDim lKeys(1 To lMaxEntrys)
        
    texInitialize = True
    Exit Function
errHandle:
    texInitialize = False
End Function
Private Function texDelete(ByRef FileNumber As Long) As Boolean
    
    lKeys(FileNumber) = 0

    ZeroMemory oGraphic(FileNumber), Len(oGraphic(FileNumber))
    lSurfaceSize = lSurfaceSize + 1
    
End Function
Public Function textureLoad(grhIndex As Long) As Boolean
    
    If (Cache.Number <> grhIndex) Then

        If oGraphic(lKeys(grhIndex)).D3DTexture Is Nothing Then
        
            With Cache
                Set oGraphic(lKeys(grhIndex)).D3DTexture = texLoad(grhIndex)
        
                .SrcHeight = oGraphic(lKeys(grhIndex)).Height + 1
                .SrcWidth = oGraphic(lKeys(grhIndex)).Width + 1
                
                Cache.Number = grhIndex
            End With
        
        End If
    
    End If
    
    If oGraphic(lKeys(grhIndex)).D3DTexture Is Nothing Then
        textureLoad = False
    Else
        textureLoad = True
    End If

End Function
Private Function texLoad(ByRef FileNumber As Long) As Direct3DTexture8

    oGraphic(lKeys(FileNumber)).Used = oGraphic(lKeys(FileNumber)).Used + 1

    If (oGraphic(lKeys(FileNumber)).Available = False) Then
        If (texCreateFromFile(FileNumber) = False) Then
            Set texLoad = Nothing: Exit Function
        Else
            lSurfaceSize = lSurfaceSize - 1
        End If
    End If

    Set texLoad = oGraphic(lKeys(FileNumber)).D3DTexture

End Function
Private Function texCreateFromFile(ByRef FileNumber As Long) As Boolean
    Dim i As Long
    Dim TexNum As Long
    Dim DelTex As Long
    Dim TexInfo As D3DXIMAGE_INFO
    
        TexNum = 0
    
        For i = 1 To lMaxEntrys
            If (oGraphic(i).Available = False) Then
                TexNum = i
                oGraphic(i).Available = True
                Exit For
            Else
                If (oGraphic(i).Used < 0) Then oGraphic(i).Used = 0: DelTex = i
            End If
        Next i
    
        If (TexNum = 0) Then
            If (texDelete(DelTex) = False) Then
                texCreateFromFile = False: Exit Function
            Else
                lKeys(FileNumber) = DelTex
            End If
        Else
            lKeys(FileNumber) = TexNum
        End If
        
        
    Dim Handle As Integer

    Handle = FreeFile() ' Get a free file number

    Open App.path & "\Graficos\" & CStr(FileNumber) & ".bmp" For Binary Access Read As #Handle
        Dim FileData() As Byte

        ReDim FileData(LOF(Handle) - 1) As Byte  ' Create an array just big enough to hold the whole file

        Get #Handle, , FileData() ' Read the file into that array

        Set oGraphic(lKeys(FileNumber)).D3DTexture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, FileData(0), UBound(FileData()) + 1, D3DX_DEFAULT, _
                                                D3DX_DEFAULT, 6, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                                D3DX_FILTER_DITHER Or D3DX_FILTER_TRIANGLE, D3DX_FILTER_BOX, _
                                                D3DColorXRGB(0, 0, 0), TexInfo, ByVal 0)

    Close #Handle
    
    
    ' Create the Texture using the filedata() array
    'Set oGraphic(lKeys(FileNumber)).D3DTexture = Loader.CreateTextureFromFileInMemory(D3DDevice, FileData(0), lBitmap)
    
    'Set oGraphic(lKeys(FileNumber)).D3DTexture = Loader.CreateTextureFromFileEx(D3DDevice, App.Path & "\Graphics\" & CStr(FileNumber) & ".png", D3DX_DEFAULT, _
                                                D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                                D3DX_FILTER_POINT, D3DX_FILTER_POINT, D3DColorXRGB(0, 0, 0), TexInfo, ByVal 0)
    
    
    With oGraphic(lKeys(FileNumber))
        .Width = TexInfo.Width
        .Height = TexInfo.Height
    End With

    texCreateFromFile = True

End Function
Public Sub texReloadAll()
    Dim i As Long

    For i = 1 To lSurfaceSize
        With oGraphic(i)
            If (.Available = True) Then
                Set .D3DTexture = Nothing
                    .Available = False
            End If
        End With
    Next i
    
    lSurfaceSize = lMaxEntrys

    ReDim oGraphic(lMaxEntrys)
    ReDim lKeys(1 To lMaxEntrys)
    
End Sub
Public Sub texDestroyAll()
    Dim i As Long

    For i = 1 To lSurfaceSize
        With oGraphic(i)
            If (.Available = True) Then
                Set .D3DTexture = Nothing
                    .Available = False
            End If
        End With
    Next i
    
    Erase oGraphic
    Erase lKeys
End Sub


