VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSurfaceManStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsSurfaceManStatic.cls - Inherits from clsSurfaceManager. Is designed to load
' surfaces at startup, and never deallocating them.
' This grants high performance can use a lot of RAM.
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

'Number of buckets in our hash table. Must be a nice prime number.
Const HASH_TABLE_SIZE As Long = 541

Private Type SURFACE_ENTRY_STATIC
    fileIndex As Long
    Surface As Direct3DTexture8
End Type

Private Type HashBucket
    sufraceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_STATIC
End Type

Private surfaceList(HASH_TABLE_SIZE - 1) As HashBucket

Private D3DX As D3DX8
Private Device8 As Direct3DDevice8
Private Sub Class_Initialize()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'
'**************************************************************
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
        For j = 1 To surfaceList(i).sufraceCount
            Set surfaceList(i).SurfaceEntry(j).Surface = Nothing
        Next j
        
        'Destroy the array
        Erase surfaceList(i).SurfaceEntry
    Next i
End Sub

Private Sub clsSurfaceManager_Initialize(ByRef Device As Direct3DDevice8, ByRef D3D As D3DX8, ByVal graphicPath As String, Optional ByVal maxMemoryUsageInMb As Long = -1)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 04/09/09
'Last Modify Author: Parra
'
'**************************************************************
    Set Device8 = Device
    Set D3DX = D3D
    
    Call LoadSurfaces(graphicPath)
End Sub

Private Function clsSurfaceManager_Surface(ByVal fileIndex As Long) As Direct3DTexture8
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 04/09/09
'Last Modify Author: Parra
'Retrieves the requested texture
'**************************************************************
On Error GoTo ErrHandler:
    Dim i As Long
    
    ' Search for it and return it
    With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
        For i = 1 To .sufraceCount
            If .SurfaceEntry(i).fileIndex = fileIndex Then
                Set clsSurfaceManager_Surface = .SurfaceEntry(i).Surface
                Exit Function
            End If
        Next i
    End With
    
    'We have a missing bmp!
    MsgBox "No se encuentra el archivo " & CStr(fileIndex) & ".bmp. Reinstale el juego, " _
        & "y si el problema persiste contactese con los adminsitradores", vbOKOnly Or vbCritical Or vbExclamation, "Error"
    
    End
Exit Function

ErrHandler:
    MsgBox "Un error inesperado ocurrió a intentar cargar el archivo " & CStr(fileIndex) & ".bmp. & vbcrlf" _
        & "Error: " & CStr(Err.number), vbOKOnly Or vbCritical Or vbExclamation, "Error"
    
    End
End Function

Private Sub LoadSurfaces(ByVal ResourcePath As String)
'**************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 04/09/2009
'Last Modify Author: Parra
'Loads all surfaces in random order and then sorts them
'**************************************************************
On Error Resume Next
    Dim data() As Byte
    Dim i As Long
    
    For i = 1 To 20207 'Ultimo grafico
        If FileExist(ResourcePath & CStr(i) & ".BMP", vbArchive) Then
            GetBitmap data(), ResourcePath & CStr(i) & ".BMP"
            LoadSurface data(), i
        End If
    Next i
            
End Sub

Private Sub LoadSurface(ByRef data() As Byte, ByVal fileIndex As Long)
'**************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 04/09/09
'Last Modify Author: Parra
'Loads the surface named fileIndex + ".bmp" and inserts it to the
'surface list in the listIndex position
'**************************************************************
'On Error GoTo ErrHandler

    Dim newSurface As SURFACE_ENTRY_STATIC
    Dim TexInfo As D3DXIMAGE_INFO
    Dim surfDesc As D3DSURFACE_DESC

    With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
    
        .sufraceCount = .sufraceCount + 1
        ReDim Preserve .SurfaceEntry(1 To .sufraceCount) As SURFACE_ENTRY_STATIC
        
        With .SurfaceEntry(.sufraceCount)
            .fileIndex = fileIndex
            
            Set .Surface = D3DX.CreateTextureFromFileInMemoryEx(Device8, data(0), UBound(data), D3DX_DEFAULT, _
                                                 D3DX_DEFAULT, 0, 0, 0, D3DPOOL_MANAGED, _
                                                 D3DX_FILTER_POINT, D3DX_FILTER_NONE, D3DColorXRGB(0, 0, 0), _
                                                 TexInfo, ByVal 0)
            
            'Retrieve the updated surface desc
            .Surface.GetLevelDesc 0, surfDesc
        End With
        
    End With
    
'Exit Sub

'ErrHandler:
    'If Err.number = DDERR_OUTOFVIDEOMEMORY Then
    '    ' No video memory? Try system memory! We want all surfaces loaded!
    '    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    '    Resume
    'End If
    
    'MsgBox "Un error inesperado ocurrió al intentar cargar el gráfico " & CStr(fileIndex) & ".BMP" & ". " & vbCrLf & _
    '        "El código de error es " & CStr(Err.number) & " - " & Err.Description & vbCrLf & vbCrLf & "Copia este mensaje y notifica a los administradores.", _
    '        vbOKOnly Or vbCritical Or vbExclamation, "Error"
    '    End
End Sub
Private Function GetBitmap(ByRef data() As Byte, ByVal path As String)
'**************************************************************
'Author: Parra
'Last Modify Date: 04/09/09
'Load graphic in array
'**************************************************************

    Dim Handle As Integer
    
    Handle = FreeFile() ' Get a free file number
    
        Open path For Binary Access Read As #Handle
            
            ReDim data(LOF(Handle) - 1) As Byte  ' Create an array just big enough to hold the whole file
            
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
        For j = 1 To surfaceList(i).sufraceCount
            Set surfaceList(i).SurfaceEntry(j).Surface = Nothing
        Next j
        
        'Destroy the array
        Erase surfaceList(i).SurfaceEntry
    Next i
End Sub
