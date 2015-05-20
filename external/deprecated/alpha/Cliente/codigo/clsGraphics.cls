VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const BYTES_PER_MB As Long = 1048576                        '1Mb = 1024 Kb = 1024 * 1024 bytes = 1048576 bytes
Private Const MIN_MEMORY_TO_USE As Long = 4 * BYTES_PER_MB          '4 Mb
Private Const DEFAULT_MEMORY_TO_USE As Long = 16 * BYTES_PER_MB     '16 Mb

'Number of buckets in our hash table. Must be a nice prime number.
Const HASH_TABLE_SIZE As Long = 337

Private Type SURFACE_ENTRY_DYN
    fileIndex As Long
    lastAccess As Long
    Surface As DxVBLibA.Direct3DTexture8
End Type

Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type

Private surfaceList(HASH_TABLE_SIZE - 1) As HashNode

Private maxBytesToUse As Long
Private usedBytes As Long

Private ResourcePath As String

Private D3DLoader As D3DX8
Private Device8 As Direct3DDevice8

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Sub Class_Initialize()
    usedBytes = 0
    maxBytesToUse = MIN_MEMORY_TO_USE
End Sub
Private Sub Class_Terminate()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Clean up
'**************************************************************
    Dim i As Long
    Dim j As Long
    
    'Destroy every surface in memory
    For i = 0 To HASH_TABLE_SIZE - 1
        With surfaceList(i)
            For j = 1 To .surfaceCount
                Set .SurfaceEntry(j).Surface = Nothing
            Next j
            
            'Destroy the arrays
            Erase .SurfaceEntry
        End With
    Next i
End Sub
Public Sub GraphicsInitialize(ByVal graphicPath As String, ByRef D3DX As D3DX8, ByRef Device As Direct3DDevice8, Optional ByVal maxMemoryUsageInMb As Long = -1)
    ResourcePath = graphicPath
    
    Set D3DLoader = D3DX
    Set Device8 = Device
    
    If maxMemoryUsageInMb = -1 Then
        maxBytesToUse = DEFAULT_MEMORY_TO_USE   ' 16 Mb by default
    ElseIf maxMemoryUsageInMb * BYTES_PER_MB < MIN_MEMORY_TO_USE Then
        maxBytesToUse = MIN_MEMORY_TO_USE       ' 4 Mb is the minimum allowed
    Else
        maxBytesToUse = maxMemoryUsageInMb * BYTES_PER_MB
    End If
End Sub
Public Function GetSurface(ByVal fileIndex As Long, ByRef surfaceWidth As Integer, ByRef surfaceHeight As Integer) As Direct3DTexture8
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Retrieves the requested texture
'**************************************************************
    Dim i As Long
    
    ' Search the index on the list
    With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
        For i = 1 To .surfaceCount
            If .SurfaceEntry(i).fileIndex = fileIndex Then
                .SurfaceEntry(i).lastAccess = GetTickCount
                Set GetSurface = .SurfaceEntry(i).Surface
                Exit Function
            End If
        Next i
    End With
    
    'Not in memory, load it!
    Set GetSurface = LoadSurface(fileIndex, surfaceWidth, surfaceHeight)
End Function

Private Function LoadSurface(ByRef fileIndex As Long, ByRef surfaceWidth As Integer, ByRef surfaceHeight As Integer) As Direct3DTexture8
    Dim newSurface As SURFACE_ENTRY_DYN
    Dim TexInfo As D3DXIMAGE_INFO
    Dim surfDesc As D3DSURFACE_DESC
        
    With newSurface
    
        .fileIndex = fileIndex
        
        'Set last access time (if we didn't we would reckon this texture as the one lru)
        .lastAccess = GetTickCount
        

            'Set .Surface = D3DLoader.CreateTextureFromFileInMemoryEx(Device8, data(0), UBound(data()) + 1, D3DX_DEFAULT, _
                                                            D3DX_DEFAULT, 6, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                                            D3DX_FILTER_DITHER Or D3DX_FILTER_TRIANGLE, D3DX_FILTER_BOX, _
                                                            D3DColorXRGB(0, 0, 0), TexInfo, ByVal 0)
            
            'Set .Surface = D3DLoader.CreateTextureFromFileInMemory(Device8, data(0), lBitMap)
                
        
                
            Set .Surface = D3DLoader.CreateTextureFromFileEx(Device8, ResourcePath & CStr(fileIndex) & ".bmp", D3DX_DEFAULT, _
                                                        D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                                        D3DX_FILTER_POINT, D3DX_FILTER_POINT, D3DColorXRGB(0, 0, 0), TexInfo, ByVal 0)
                
        .Surface.GetLevelDesc 0, surfDesc
                
        surfaceWidth = TexInfo.Width
        surfaceHeight = TexInfo.Height

    End With

    'Insert surface to the list
    With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
        .surfaceCount = .surfaceCount + 1
        
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
        
        .SurfaceEntry(.surfaceCount) = newSurface
        
        Set LoadSurface = newSurface.Surface
    End With
    
    'Update used bytes
    usedBytes = usedBytes + surfDesc.Size
    
    'Check if we have exceeded our allowed share of memory usage
    Do While usedBytes > maxBytesToUse
        'Remove a file. If no file could be removed we continue, if the file was previous to our surface we update the index
        If Not RemoveLRU() Then
            Exit Do
        End If
    Loop
Exit Function

ErrHandler:
    MsgBox "Un error inesperado ocurrió al intentar cargar el gráfico " & CStr(fileIndex) & ".png" & ". " & vbCrLf & _
                "El código de error es " & CStr(Err.Number) & " - " & Err.Description & vbCrLf & vbCrLf & "Copia este mensaje y notifica a los administradores.", _
                vbOKOnly Or vbCritical Or vbExclamation, "Error"
    End
End Function
Private Function RemoveLRU() As Boolean
'**************************************************************
'Removes the Least Recently Used surface to make some room for new ones
'**************************************************************

    Dim LRUi As Long
    Dim LRUj As Long
    Dim LRUtime As Long
    Dim i As Long
    Dim j As Long
    Dim surfDesc As D3DSURFACE_DESC
    
    LRUtime = GetTickCount
    
    'Check out through the whole list for the least recently used
    For i = 0 To HASH_TABLE_SIZE - 1
        With surfaceList(i)
            For j = 1 To .surfaceCount
                If LRUtime > .SurfaceEntry(j).lastAccess Then
                    LRUi = i
                    LRUj = j
                    LRUtime = .SurfaceEntry(j).lastAccess
                End If
            Next j
        End With
    Next i
    
    'Retrieve the surface desc
    surfaceList(LRUi).SurfaceEntry(LRUj).Surface.GetLevelDesc 0, surfDesc
       
    'Remove it
    Set surfaceList(LRUi).SurfaceEntry(LRUj).Surface = Nothing
    surfaceList(LRUi).SurfaceEntry(LRUj).fileIndex = 0
    
    'Move back the list (if necessary)
    If LRUj Then
        RemoveLRU = True
        
        With surfaceList(LRUi)
            For j = LRUj To .surfaceCount - 1
                .SurfaceEntry(j) = .SurfaceEntry(j + 1)
            Next j
            
            .surfaceCount = .surfaceCount - 1
            If .surfaceCount Then
                ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
            Else
                Erase .SurfaceEntry
            End If
        End With
    End If
    
    'Update the used bytes
    usedBytes = usedBytes - surfDesc.Size
End Function


