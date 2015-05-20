VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSurfaceManDyn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsSurfaceManDyn.cls - Inherits from clsSurfaceManager. Is designed to load
'surfaces dynamically without using more than an arbitrary amount of Mb.
'For removale it uses LRU, attempting to just keep in memory those surfaces
'that are actually usefull.
'
' Developed by Maraxus (Juan Martín Sotuyo Dodero - juansotuyo@hotmail.com)
' Last Modify Date: 3/06/2006
'**************************************************************

'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

Option Explicit

'Inherit from the surface manager
Implements clsSurfaceManager

Private Const BYTES_PER_MB As Long = 1048576                        '1Mb = 1024 Kb = 1024 * 1024 bytes = 1048576 bytes
Private Const MIN_MEMORY_TO_USE As Long = 4 * BYTES_PER_MB          '4 Mb
Private Const DEFAULT_MEMORY_TO_USE As Long = 16 * BYTES_PER_MB     '16 Mb

'Number of buckets in our hash table. Must be a nice prime number.
Const HASH_TABLE_SIZE As Long = 337

Private Type SURFACE_ENTRY_DYN
    fileIndex As Long
    lastAccess As Long
    Surface As Direct3DTexture8
End Type

Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type

Private surfaceList(HASH_TABLE_SIZE - 1) As HashNode

Private D3DX As D3DX8
Private Device8 As Direct3DDevice8

Private maxBytesToUse As Long
Private usedBytes As Long

Private ResourcePath As String

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Class_Initialize()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'
'**************************************************************
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

Private Sub clsSurfaceManager_Initialize(ByRef Device As Direct3DDevice8, ByRef D3D As D3DX8, ByVal graphicPath As String, Optional ByVal maxMemoryUsageInMb As Long = -1)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 04/09/09
'Last Modify Author: Parra
'Initializes the manager
'**************************************************************
    Set Device8 = Device
    Set D3DX = D3D
        
    ResourcePath = graphicPath
    
    If maxMemoryUsageInMb = -1 Then
        maxBytesToUse = DEFAULT_MEMORY_TO_USE   ' 16 Mb by default
    ElseIf maxMemoryUsageInMb * BYTES_PER_MB < MIN_MEMORY_TO_USE Then
        maxBytesToUse = MIN_MEMORY_TO_USE       ' 4 Mb is the minimum allowed
    Else
        maxBytesToUse = maxMemoryUsageInMb * BYTES_PER_MB
    End If
End Sub

Private Function clsSurfaceManager_Surface(ByVal fileIndex As Long) As Direct3DTexture8
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 04/09/09
'Last Modify Author: Parra
'Retrieves the requested texture
'**************************************************************
    Dim i As Long
    
    ' Search the index on the list
    With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
        For i = 1 To .surfaceCount
            If .SurfaceEntry(i).fileIndex = fileIndex Then
                .SurfaceEntry(i).lastAccess = GetTickCount
                Set clsSurfaceManager_Surface = .SurfaceEntry(i).Surface
                Exit Function
            End If
        Next i
    End With
    
    'Not in memory, load it!
    Set clsSurfaceManager_Surface = LoadSurface(fileIndex)
End Function

Private Function LoadSurface(ByVal fileIndex As Long) As Direct3DTexture8
'**************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 04/09/09
'Last Modify Author: Parra
'Loads the surface named fileIndex + ".bmp" and inserts it to the
'surface list in the listIndex position
'**************************************************************
'On Error GoTo ErrHandler

    Dim newSurface As SURFACE_ENTRY_DYN
    Dim data() As Byte, lBitMap As Long
    Dim surfDesc As D3DSURFACE_DESC
    Dim TexInfo As D3DXIMAGE_INFO
    
    'get Bitmap
    'GetBitmap data(), lBitMap, fileIndex
        
    With newSurface
        .fileIndex = fileIndex
        
        'Set last access time (if we didn't we would reckon this texture as the one lru)
        .lastAccess = GetTickCount
        
        'Set .Surface = D3DX.CreateTextureFromFileInMemory(Device8, Data(0), lBitMap)
        
        'Set .Surface = D3DX.CreateTextureFromFileInMemoryEx(Device8, data(0), lBitMap, D3DX_DEFAULT, _
                                             D3DX_DEFAULT, 0, 0, 0, D3DPOOL_MANAGED, _
                                             D3DX_FILTER_POINT, D3DX_FILTER_NONE, D3DColorXRGB(0, 0, 0), _
                                             TexInfo, ByVal 0)
                                             
        Set .Surface = D3DX.CreateTextureFromFileEx(Device8, ResourcePath & CStr(fileIndex) & ".bmp", _
                                        D3DX_DEFAULT, D3DX_DEFAULT, 3, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                                        D3DX_FILTER_NONE, &HFF000000, TexInfo, ByVal 0)
        
        'Retrieve the updated surface desc
        .Surface.GetLevelDesc 0, surfDesc
        
    End With
    
    'Insert surface to the list
    With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
        .surfaceCount = .surfaceCount + 1
        
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
        
        .SurfaceEntry(.surfaceCount) = newSurface
        
        Set LoadSurface = newSurface.Surface
    End With
    
    'Update used bytes
    usedBytes = usedBytes + surfDesc.size
    
    'Check if we have exceeded our allowed share of memory usage
    Do While usedBytes > maxBytesToUse
        'Remove a file. If no file could be removed we continue, if the file was previous to our surface we update the index
        If Not RemoveLRU() Then
            Exit Do
        End If
    Loop
'Exit Function

'ErrHandler:
    'Remove a surface and try again
    'If RemoveLRU() Then
    '    Resume
    'Else
    '    MsgBox "No hay memoria disponible! El programa abortará. Cierra algunos programas e intenta de nuevo"
    '    End
    'End If

'    Debug.Print Err.number; Err.Description
End Function
Private Function RemoveLRU() As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 04/09/09
'Last Modify Author: Parra
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
    
    If LRUj Then
        RemoveLRU = True
        'Retrieve the surface desc
        surfaceList(LRUi).SurfaceEntry(LRUj).Surface.GetLevelDesc 0, surfDesc
        
        'Remove it
        Set surfaceList(LRUi).SurfaceEntry(LRUj).Surface = Nothing
        surfaceList(LRUi).SurfaceEntry(LRUj).fileIndex = 0
        
        'Move back the list (if necessary)
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
        
        'Update the used bytes
        usedBytes = usedBytes - surfDesc.size
    End If
End Function
Private Function GetBitmap(ByRef data() As Byte, ByRef lBitMap As Long, ByVal fileIndex As Long)
'**************************************************************
'Author: Parra
'Last Modify Date: 04/09/09
'Load graphic in array
'**************************************************************

    Dim Handle As Integer
    
    Handle = FreeFile() ' Get a free file number
    
        Open ResourcePath & CStr(fileIndex) & ".bmp" For Binary Access Read As #Handle
            
            lBitMap = LOF(Handle)
            
            ReDim data(lBitMap - 1) As Byte  ' Create an array just big enough to hold the whole file
            
            Get #Handle, , data() ' Read the file into that array
   
        Close #Handle

End Function

Public Sub clsSurfaceManager_CleanSurface()
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
