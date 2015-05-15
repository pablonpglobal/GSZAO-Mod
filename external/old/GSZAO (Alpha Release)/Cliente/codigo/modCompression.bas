Attribute VB_Name = "modCompression"
Option Explicit

Public Const GRH_SOURCE_FILE_EXT As String = ".bmp"
Public Const GRH_RESOURCE_FILE As String = "Graphics.AO"
Public Const GRH_PATCH_FILE As String = "Graphics.PATCH"

'This structure will describe our binary file's
'size, number and version of contained files
Public Type FILEHEADER
    lngNumFiles As Long                 'How many files are inside?
    lngFileSize As Long                 'How big is this file? (Used to check integrity)
    lngFileVersion As Long              'The resource version (Used to patch)
End Type

'This structure will describe each file contained
'in our binary file
Public Type INFOHEADER
    lngFileSize As Long             'How big is this chunk of stored data?
    lngFileStart As Long            'Where does the chunk start?
    strFileName As String * 16      'What's the name of the file this data came from?
    lngFileSizeUncompressed As Long 'How big is the file compressed
End Type

Private Enum PatchInstruction
    Delete_File
    Create_File
    Modify_File
End Enum

Private Declare Function compress Lib "zlib.dll" (dest As Any, destlen As Any, Src As Any, ByVal srclen As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destlen As Any, Src As Any, ByVal srclen As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef source As Any, ByVal byteCount As Long)

'BitMaps Strucures
Public Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type
Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD
End Type

Private Const BI_RGB As Long = 0
Private Const BI_RLE8 As Long = 1
Private Const BI_RLE4 As Long = 2
Private Const BI_BITFIELDS As Long = 3
Private Const BI_JPG As Long = 4
Private Const BI_PNG As Long = 5


'To get free bytes in drive
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, FreeBytesToCaller As Currency, bytesTotal As Currency, FreeBytesTotal As Currency) As Long

Private Function General_Drive_Get_Free_Bytes(ByVal DriveName As String) As Currency
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 6/07/2004
'
'**************************************************************
    Dim retval As Long
    Dim FB As Currency
    Dim BT As Currency
    Dim FBT As Currency
    
    retval = GetDiskFreeSpace(Left$(DriveName, 2), FB, BT, FBT)
    
    General_Drive_Get_Free_Bytes = FB * 10000 'convert result to actual size in bytes
End Function

''
' Searches for the specified InfoHeader.
'
' @param    ResourceFile A handler to the data file.
' @param    InfoHead The header searched.
' @param    FirstHead The first head to look.
' @param    LastHead The last head to look.
' @param    FileHeaderSize The bytes size of a FileHeader.
' @param    InfoHeaderSize The bytes size of a InfoHeader.
'
' @return   True if found.
'
' @remark   File must be already open.
' @remark   InfoHead must have set its file name to perform the search.

Private Function BinarySearch(ByRef ResourceFile As Integer, ByRef InfoHead As INFOHEADER, ByVal FirstHead As Long, ByVal LastHead As Long, ByVal FileHeaderSize As Long, ByVal InfoHeaderSize As Long) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 08/21/2007
'Searches for the specified InfoHeader
'*****************************************************************
    Dim ReadingHead As Long
    Dim ReadInfoHead As INFOHEADER
    
    Do Until FirstHead > LastHead
        ReadingHead = (FirstHead + LastHead) \ 2

        Get ResourceFile, FileHeaderSize + InfoHeaderSize * (ReadingHead - 1) + 1, ReadInfoHead

        If InfoHead.strFileName = ReadInfoHead.strFileName Then
            InfoHead = ReadInfoHead
            BinarySearch = True
            Exit Function
        Else
            If InfoHead.strFileName < ReadInfoHead.strFileName Then
                LastHead = ReadingHead - 1
            Else
                FirstHead = ReadingHead + 1
            End If
        End If
    Loop
End Function

''
' Retrieves the InfoHead of the specified graphic file.
'
' @param    ResourcePath The resource file folder.
' @param    FileName The graphic file name.
' @param    InfoHead The InfoHead where data is returned.
'
' @return   True if found.

Private Function GetInfoHeader(ByRef ResourcePath As String, ByRef FileName As String, ByRef InfoHead As INFOHEADER) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 08/21/2007
'Retrieves the InfoHead of the specified graphic file
'*****************************************************************
    Dim ResourceFile As Integer
    Dim ResourceFilePath As String
    Dim FileHead As FILEHEADER
    
On Local Error GoTo ErrHandler

    ResourceFilePath = ResourcePath & GRH_RESOURCE_FILE
    
    'Set InfoHeader we are looking for
    InfoHead.strFileName = UCase$(FileName)
    
    'Open the binary file
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
        'Extract the FILEHEADER
        Get ResourceFile, 1, FileHead
        
        'Check the file for validity
        If LOF(ResourceFile) <> FileHead.lngFileSize Then
            MsgBox "Archivo de recursos dañado. " & ResourceFilePath, , "Error"
            Close ResourceFile
            Exit Function
        End If
        
        'Search for it!
        If BinarySearch(ResourceFile, InfoHead, 1, FileHead.lngNumFiles, Len(FileHead), Len(InfoHead)) Then
            GetInfoHeader = True
        End If
        
    Close ResourceFile
Exit Function

ErrHandler:
    Close ResourceFile
    
    Call MsgBox("Error al intentar leer el archivo " & ResourceFilePath & ". Razón: " & Err.number & " : " & Err.Description, vbOKOnly, "Error")
End Function

''
' Decompresses binary data.
'
' @param    data() The data array.
' @param    OrigSize The original data size.

Private Sub DecompressData(ByRef Data() As Byte, ByVal OrigSize As Long)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Decompresses binary data
'*****************************************************************
    Dim BufTemp() As Byte
    
    ReDim BufTemp(OrigSize - 1)
    
    Call uncompress(BufTemp(0), OrigSize, Data(0), UBound(Data) + 1)
    
    ReDim Data(OrigSize - 1)
    
    Data = BufTemp
    
    Erase BufTemp
End Sub

''
' Retrieves a byte array with the compressed data from the specified file.
'
' @param    ResourcePath The resource file folder.
' @param    InfoHead The header specifiing the graphic file info.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.
'
' @remark   InfoHead must not be encrypted.
' @remark   Data is not desencrypted.

Public Function GetFileRawData(ByRef ResourcePath As String, ByRef InfoHead As INFOHEADER, ByRef Data() As Byte) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 08/24/2007
'Retrieves a byte array with the compressed data from the specified file
'*****************************************************************
    Dim ResourceFilePath As String
    Dim ResourceFile As Integer
    
On Local Error GoTo ErrHandler
    ResourceFilePath = ResourcePath & GRH_RESOURCE_FILE
    
    'Size the Data array
    ReDim Data(InfoHead.lngFileSize - 1)
    
    'Open the binary file
    ResourceFile = FreeFile
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
        'Get the data
        Get ResourceFile, InfoHead.lngFileStart, Data
    'Close the binary file
    Close ResourceFile
    
    GetFileRawData = True
Exit Function

ErrHandler:
    Close ResourceFile
End Function

''
' Extract the specific file from a resource file.
'
' @param    ResourcePath The resource file folder.
' @param    InfoHead The header specifiing the graphic file info.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.
'
' @remark   Data is desencrypted.

Public Function ExtractFile(ByRef ResourcePath As String, ByRef InfoHead As INFOHEADER, ByRef Data() As Byte) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 08/20/2007
'Extract the specific file from a resource file
'*****************************************************************
On Local Error GoTo ErrHandler
    
    If GetFileRawData(ResourcePath, InfoHead, Data) Then
        'Decompress all data
        If InfoHead.lngFileSize < InfoHead.lngFileSizeUncompressed Then
            Call DecompressData(Data, InfoHead.lngFileSizeUncompressed)
        End If
        
        ExtractFile = True
    End If
Exit Function

ErrHandler:
    Call MsgBox("Error al intentar decodificar recursos. Razon: " & Err.number & " : " & Err.Description, vbOKOnly, "Error")
End Function

''
' Extracts all files from a resource file.
'
' @param    ResourcePath The resource file folder.
' @param    OutputPath The folder where graphic files will be extracted.
' @param    PrgBar The control that shows the process state.
'
' @return   True if no error occurred.

Public Function ExtractFiles(ByRef ResourcePath As String, ByRef OutputPath As String, ByRef prgBar As ProgressBar) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 08/20/2007
'Extracts all files from a resource file
'*****************************************************************
    Dim loopc As Long
    Dim ResourceFile As Integer
    Dim ResourceFilePath As String
    Dim OutputFile As Integer
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim RequiredSpace As Currency
    
On Local Error GoTo ErrHandler
    ResourceFilePath = ResourcePath & GRH_RESOURCE_FILE
    
    'Open the binary file
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
        'Extract the FILEHEADER
        Get ResourceFile, 1, FileHead
        
        'Check the file for validity
        If LOF(ResourceFile) <> FileHead.lngFileSize Then
            Call MsgBox("Archivo de recursos dañado. " & ResourceFilePath, , "Error")
            Close ResourceFile
            Exit Function
        End If
        
        'Size the InfoHead array
        ReDim InfoHead(FileHead.lngNumFiles - 1)
        
        'Extract the INFOHEADER
        Get ResourceFile, , InfoHead
        
        'Check if there is enough hard drive space to extract all files
        For loopc = 0 To UBound(InfoHead)
            RequiredSpace = RequiredSpace + InfoHead(loopc).lngFileSizeUncompressed
        Next loopc
        
        If RequiredSpace >= General_Drive_Get_Free_Bytes(Left$(App.path, 3)) Then
            Erase InfoHead
            Close ResourceFile
            Call MsgBox("No hay suficiente espacio en el disco para extraer los archivos.", , "Error")
            Exit Function
        End If
    Close ResourceFile
    
    'Update progress bar
    If Not prgBar Is Nothing Then
        prgBar.max = FileHead.lngNumFiles
        prgBar.value = 0
    End If
    
    'Extract all of the files from the binary file
    For loopc = 0 To UBound(InfoHead)
        'Extract this file
        If ExtractFile(ResourcePath, InfoHead(loopc), SourceData) Then
            'Destroy file if it previuosly existed
            If FileExist(OutputPath & InfoHead(loopc).strFileName, vbNormal) Then
                Call Kill(OutputPath & InfoHead(loopc).strFileName)
            End If
            
            'Save it!
            OutputFile = FreeFile()
            Open OutputPath & InfoHead(loopc).strFileName For Binary As OutputFile
                Put OutputFile, , SourceData
            Close OutputFile
            
            Erase SourceData
        Else
            Erase SourceData
            Erase InfoHead
            
            Call MsgBox("No se pudo extraer el archivo " & InfoHead(loopc).strFileName, vbOKOnly, "Error")
            Exit Function
        End If
            
        'Update progress bar
        If Not prgBar Is Nothing Then prgBar.value = prgBar.value + 1
        DoEvents
    Next loopc
    
    Erase InfoHead
    ExtractFiles = True
Exit Function

ErrHandler:
    Close ResourceFile
    Erase SourceData
    Erase InfoHead
    
    Call MsgBox("No se pudo extraer el archivo binario correctamente. Razon: " & Err.number & " : " & Err.Description, vbOKOnly, "Error")
End Function

''
' Retrieves a byte array with the specified file data.
'
' @param    ResourcePath The resource file folder.
' @param    FileName The graphic file name.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.
'
' @remark   Data is desencrypted.

Public Function GetFileData(ByRef ResourcePath As String, ByRef FileName As String, ByRef Data() As Byte) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 08/21/2007
'Retrieves a byte array with the specified file data
'*****************************************************************
    Dim InfoHead As INFOHEADER
    
    If GetInfoHeader(ResourcePath, FileName, InfoHead) Then
        'Extract!
        GetFileData = ExtractFile(ResourcePath, InfoHead, Data)
    Else
        Call MsgBox("No se se encontro el recurso " & FileName)
    End If
End Function

''
' Retrieves bitmap file data.
'
' @param    ResourcePath The resource file folder.
' @param    FileName The graphic file name.
' @param    bmpInfo The bitmap info structure.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.

Public Function GetBitmap(ByRef ResourcePath As String, ByRef FileName As String, ByRef bmpInfo As BITMAPINFO, ByRef Data() As Byte) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 11/30/2007
'Retrieves bitmap file data
'*****************************************************************
    Dim InfoHead As INFOHEADER
    Dim rawData() As Byte
    Dim offBits As Long
    Dim bitmapSize As Long
    Dim colorCount As Long
    
    If GetInfoHeader(ResourcePath, FileName, InfoHead) Then
        'Extract the file and create the bitmap data from it.
        If ExtractFile(ResourcePath, InfoHead, rawData) Then
            Call CopyMemory(offBits, rawData(10), 4)
            Call CopyMemory(bmpInfo.bmiHeader, rawData(14), 40)
            
            With bmpInfo.bmiHeader
                bitmapSize = AlignScan(.biWidth, .biBitCount) * Abs(.biHeight)
                
                If .biBitCount < 24 Or .biCompression = BI_BITFIELDS Or (.biCompression <> BI_RGB And .biBitCount = 32) Then
                    If .biClrUsed < 1 Then
                        colorCount = 2 ^ .biBitCount
                    Else
                        colorCount = .biClrUsed
                    End If
                    
                    ' When using bitfields on 16 or 32 bits images, bmiColors has a 3-longs mask.
                    If .biBitCount >= 16 And .biCompression = BI_BITFIELDS Then colorCount = 3
                    
                    Call CopyMemory(bmpInfo.bmiColors(0), rawData(54), colorCount * 4)
                End If
            End With
            
            ReDim Data(bitmapSize - 1) As Byte
            Call CopyMemory(Data(0), rawData(offBits), bitmapSize)
            
            GetBitmap = True
        End If
    Else
        Call MsgBox("No se encontro el recurso " & FileName)
    End If
End Function

''
' Compare two byte arrays to detect any difference.
'
' @param    data1() Byte array.
' @param    data2() Byte array.
'
' @return   True if are equals.

Private Function CompareData(ByRef data1() As Byte, ByRef data2() As Byte) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 02/11/2007
'Compare two byte arrays to detect any difference
'*****************************************************************
    Dim length As Long
    Dim act As Long
    
    length = UBound(data1) + 1
    
    If (UBound(data2) + 1) = length Then
        While act < length
            If data1(act) Xor data2(act) Then Exit Function
            
            act = act + 1
        Wend
        
        CompareData = True
    End If
End Function

''
' Retrieves the next InfoHeader.
'
' @param    ResourceFile A handler to the resource file.
' @param    FileHead The reource file header.
' @param    InfoHead The returned header.
' @param    ReadFiles The number of headers that have already been read.
'
' @return   False if there are no more headers tu read.
'
' @remark   File must be already open.
' @remark   Used to walk through the resource file info headers.
' @remark   The number of read files will increase although there is nothing else to read.
' @remark   InfoHead is encrypted.

Private Function ReadNextInfoHead(ByRef ResourceFile As Integer, ByRef FileHead As FILEHEADER, ByRef InfoHead As INFOHEADER, ByRef ReadFiles As Long) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 08/24/2007
'Reads the next InfoHeader
'*****************************************************************

    If ReadFiles < FileHead.lngNumFiles Then
        'Read header
        Get ResourceFile, Len(FileHead) + Len(InfoHead) * ReadFiles + 1, InfoHead
        
        'Update
        ReadNextInfoHead = True
    End If
    
    ReadFiles = ReadFiles + 1
End Function

''
' Retrieves the next bitmap.
'
' @param    ResourcePath The resource file folder.
' @param    ReadFiles The number of bitmaps that have already been read.
' @param    bmpInfo The bitmap info structure.
' @param    data() The byte array to return data.
'
' @return   False if there are no more bitmaps tu get.
'
' @remark   Used to walk through the resource file bitmaps.

Public Function GetNextBitmap(ByRef ResourcePath As String, ByRef ReadFiles As Long, ByRef bmpInfo As BITMAPINFO, ByRef Data() As Byte, ByRef fileIndex As Long) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 12/02/2007
'Reads the next InfoHeader
'*****************************************************************
On Error Resume Next

    Dim ResourceFile As Integer
    Dim FileHead As FILEHEADER
    Dim InfoHead As INFOHEADER
    Dim FileName As String
    
    ResourceFile = FreeFile
    Open ResourcePath & GRH_RESOURCE_FILE For Binary Access Read Lock Write As ResourceFile
    Get ResourceFile, 1, FileHead
    
    If ReadNextInfoHead(ResourceFile, FileHead, InfoHead, ReadFiles) Then
        Call GetBitmap(ResourcePath, InfoHead.strFileName, bmpInfo, Data())
        FileName = Trim$(InfoHead.strFileName)
        fileIndex = CLng(Left$(FileName, Len(FileName) - 4))
        
        GetNextBitmap = True
    End If
    
    Close ResourceFile
End Function

''
' Compares two resource versions and makes a patch file.
'
' @param    NewResourcePath The actual reource file folder.
' @param    OldResourcePath The previous reource file folder.
' @param    OutputPath The patchs file folder.
' @param    PrgBar The control that shows the process state.
'
' @return   True if no error occurred.

Public Function MakePatch(ByRef NewResourcePath As String, ByRef OldResourcePath As String, ByRef OutputPath As String, ByRef prgBar As ProgressBar) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 08/24/2007
'Compares two resource versions and make a patch file
'*****************************************************************
    Dim NewResourceFile As Integer
    Dim NewResourceFilePath As String
    Dim NewFileHead As FILEHEADER
    Dim NewInfoHead As INFOHEADER
    Dim NewReadFiles As Long
    Dim NewReadNext As Boolean
    
    Dim OldResourceFile As Integer
    Dim OldResourceFilePath As String
    Dim OldFileHead As FILEHEADER
    Dim OldInfoHead As INFOHEADER
    Dim OldReadFiles As Long
    Dim OldReadNext As Boolean
    
    Dim OutputFile As Integer
    Dim OutputFilePath As String
    Dim Data() As Byte
    Dim auxData() As Byte
    Dim Instruction As Byte
    
'Set up the error handler
'On Local Error GoTo ErrHandler

    NewResourceFilePath = NewResourcePath & GRH_RESOURCE_FILE
    OldResourceFilePath = OldResourcePath & GRH_RESOURCE_FILE
    OutputFilePath = OutputPath & GRH_PATCH_FILE
    
    'Open the old binary file
    OldResourceFile = FreeFile
    Open OldResourceFilePath For Binary Access Read Lock Write As OldResourceFile
        
        'Get the old FileHeader
        Get OldResourceFile, 1, OldFileHead
        
        'Check the file for validity
        If LOF(OldResourceFile) <> OldFileHead.lngFileSize Then
            Call MsgBox("Archivo de recursos anterior dañado. " & OldResourceFilePath, , "Error")
            Close OldResourceFile
            Exit Function
        End If
        
        'Open the new binary file
        NewResourceFile = FreeFile()
        Open NewResourceFilePath For Binary Access Read Lock Write As NewResourceFile
            
            'Get the new FileHeader
            Get NewResourceFile, 1, NewFileHead
            
            'Check the file for validity
            If LOF(NewResourceFile) <> NewFileHead.lngFileSize Then
                Call MsgBox("Archivo de recursos anterior dañado. " & NewResourceFilePath, , "Error")
                Close NewResourceFile
                Close OldResourceFile
                Exit Function
            End If
            
            'Destroy file if it previuosly existed
            If Dir(OutputFilePath, vbNormal) <> vbNullString Then Kill OutputFilePath ' GSZ
            
            'Open the patch file
            OutputFile = FreeFile()
            Open OutputFilePath For Binary Access Read Write As OutputFile
                
                If Not prgBar Is Nothing Then
                    prgBar.max = OldFileHead.lngNumFiles + NewFileHead.lngNumFiles
                    prgBar.value = 0
                End If
                
                'put previous file version (unencrypted)
                Put OutputFile, , OldFileHead.lngFileVersion
                
                'Put the new file header
                Put OutputFile, , NewFileHead
                
                'Try to read old and new first files
                If ReadNextInfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles) _
                  And ReadNextInfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles) Then
                    
                    'Update
                    prgBar.value = prgBar.value + 2
                    
                    Do 'Main loop
                        'Comparisons are between encrypted names, for ordering issues
                        If OldInfoHead.strFileName = NewInfoHead.strFileName Then

                            'Get old file data
                            Call GetFileRawData(OldResourcePath, OldInfoHead, auxData)
                            
                            'Get new file data
                            Call GetFileRawData(NewResourcePath, NewInfoHead, Data)
                            
                            If Not CompareData(Data, auxData) Then
                                'File was modified
                                Instruction = PatchInstruction.Modify_File
                                Put OutputFile, , Instruction
                                
                                'Write header
                                Put OutputFile, , NewInfoHead
                                
                                'Write data
                                Put OutputFile, , Data
                            End If
                            
                            'Read next OldResource
                            If Not ReadNextInfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles) Then
                                Exit Do
                            End If
                            
                            'Read next NewResource
                            If Not ReadNextInfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles) Then
                                'Reread last OldInfoHead
                                OldReadFiles = OldReadFiles - 1
                                Exit Do
                            End If
                            
                            'Update
                            If Not prgBar Is Nothing Then prgBar.value = prgBar.value + 2
                        
                        ElseIf OldInfoHead.strFileName < NewInfoHead.strFileName Then
                            
                            'File was deleted
                            Instruction = PatchInstruction.Delete_File
                            Put OutputFile, , Instruction
                            Put OutputFile, , OldInfoHead
                            
                            'Read next OldResource
                            If Not ReadNextInfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles) Then
                                'Reread last NewInfoHead
                                NewReadFiles = NewReadFiles - 1
                                Exit Do
                            End If
                            
                            'Update
                            If Not prgBar Is Nothing Then prgBar.value = prgBar.value + 1
                        
                        Else
                            
                            'New file
                            Instruction = PatchInstruction.Create_File
                            Put OutputFile, , Instruction
                            Put OutputFile, , NewInfoHead
                            
                            'Get file data
                            Call GetFileRawData(NewResourcePath, NewInfoHead, Data)
                            
                            'Write data
                            Put OutputFile, , Data
                            
                            'Read next NewResource
                            If Not ReadNextInfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles) Then
                                'Reread last OldInfoHead
                                OldReadFiles = OldReadFiles - 1
                                Exit Do
                            End If
                            
                            'Update
                            If Not prgBar Is Nothing Then prgBar.value = prgBar.value + 1
                        End If
                        
                        DoEvents
                    Loop
                
                Else
                    'if at least one is empty
                    OldReadFiles = 0
                    NewReadFiles = 0
                End If
                
                'Read everything?
                While ReadNextInfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles)
                    'Delete file
                    Instruction = PatchInstruction.Delete_File
                    Put OutputFile, , Instruction
                    Put OutputFile, , OldInfoHead
                    
                    'Update
                    If Not prgBar Is Nothing Then prgBar.value = prgBar.value + 1
                    DoEvents
                Wend
                
                'Read everything?
                While ReadNextInfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles)
                    'Create file
                    Instruction = PatchInstruction.Create_File
                    Put OutputFile, , Instruction
                    Put OutputFile, , NewInfoHead
                    
                    'Get file data
                    Call GetFileRawData(NewResourcePath, NewInfoHead, Data)
                    'Write data
                    Put OutputFile, , Data
                    
                    'Update
                    If Not prgBar Is Nothing Then prgBar.value = prgBar.value + 1
                    DoEvents
                Wend
            
            'Close the patch file
            Close OutputFile
        
        'Close the new binary file
        Close NewResourceFile
    
    'Close the old binary file
    Close OldResourceFile
    
    MakePatch = True
Exit Function

ErrHandler:
    Close OutputFile
    Close NewResourceFile
    Close OldResourceFile
    
    Call MsgBox("No se pudo terminar de crear el parche. Razon: " & Err.number & " : " & Err.Description, vbOKOnly, "Error")
End Function

''
' Follows patches instructions to update a resource file.
'
' @param    ResourcePath The reource file folder.
' @param    PatchPath The patch file folder.
' @param    PrgBar The control that shows the process state.
'
' @return   True if no error occurred.
Public Function Apply_Patch(ByRef ResourcePath As String, ByRef PatchPath As String, ByRef prgBar As ProgressBar) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 08/24/2007
'Follows patches instructions to update a resource file
'*****************************************************************
    Dim ResourceFile As Integer
    Dim ResourceFilePath As String
    Dim FileHead As FILEHEADER
    Dim InfoHead As INFOHEADER
    Dim ResourceReadFiles As Long
    Dim EOResource As Boolean

    Dim PatchFile As Integer
    Dim PatchFilePath As String
    Dim PatchFileHead As FILEHEADER
    Dim PatchInfoHead As INFOHEADER
    Dim Instruction As Byte
    Dim OldResourceVersion As Long

    Dim OutputFile As Integer
    Dim OutputFilePath As String
    Dim Data() As Byte
    Dim WrittenFiles As Long
    Dim DataOutputPos As Long

On Local Error GoTo ErrHandler

    ResourceFilePath = ResourcePath & GRH_RESOURCE_FILE
    PatchFilePath = PatchPath & GRH_PATCH_FILE
    OutputFilePath = ResourcePath & GRH_RESOURCE_FILE & "tmp"
    
    'Open the old binary file
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
        
        'Read the old FileHeader
        Get ResourceFile, , FileHead
        
        'Check the file for validity
        If LOF(ResourceFile) <> FileHead.lngFileSize Then
            Call MsgBox("Archivo de recursos anterior dañado. " & ResourceFilePath, , "Error")
            Close ResourceFile
            Exit Function
        End If
        
        'Open the patch file
        PatchFile = FreeFile()
        Open PatchFilePath For Binary Access Read Lock Write As PatchFile
            
            'Get previous file version
            Get PatchFile, , OldResourceVersion
            
            'Check the file version
            If OldResourceVersion <> FileHead.lngFileVersion Then
                Call MsgBox("Incongruencia en versiones.", , "Error")
                Close ResourceFile
                Close PatchFile
                Exit Function
            End If
            
            'Read the new FileHeader
            Get PatchFile, , PatchFileHead
            
            'Destroy file if it previuosly existed
            If FileExist(OutputFilePath, vbNormal) Then Call Kill(OutputFilePath)
            
            'Open the patch file
            OutputFile = FreeFile()
            Open OutputFilePath For Binary Access Read Write As OutputFile
                
                'Save the file header
                Put OutputFile, , PatchFileHead
                
                If Not prgBar Is Nothing Then
                    prgBar.max = PatchFileHead.lngNumFiles
                    prgBar.value = 0
                End If
                
                'Update
                DataOutputPos = Len(FileHead) + Len(InfoHead) * PatchFileHead.lngNumFiles + 1
                
                'Process loop
                While Loc(PatchFile) < LOF(PatchFile)
                    
                    'Get the instruction
                    Get PatchFile, , Instruction
                    'Get the InfoHead
                    Get PatchFile, , PatchInfoHead
                    
                    Do
                        EOResource = Not ReadNextInfoHead(ResourceFile, FileHead, InfoHead, ResourceReadFiles)
                        
                        'Comparison is performed among encrypted names for ordering issues
                        If Not EOResource And InfoHead.strFileName < PatchInfoHead.strFileName Then

                            'GetData and update InfoHead
                            Call GetFileRawData(ResourcePath, InfoHead, Data)
                            InfoHead.lngFileStart = DataOutputPos
                            
                            'Save file!
                            Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, InfoHead
                            Put OutputFile, DataOutputPos, Data
                            
                            'Update
                            DataOutputPos = DataOutputPos + UBound(Data) + 1
                            WrittenFiles = WrittenFiles + 1
                            If Not prgBar Is Nothing Then prgBar.value = WrittenFiles
                        Else
                            Exit Do
                        End If
                    Loop
                    
                    Select Case Instruction
                        'Delete
                        Case PatchInstruction.Delete_File
                            If InfoHead.strFileName <> PatchInfoHead.strFileName Then
                                Err.Description = "Incongruencia en archivos de recurso"
                                GoTo ErrHandler
                            End If
                        
                        'Create
                        Case PatchInstruction.Create_File
                            If (InfoHead.strFileName > PatchInfoHead.strFileName) Or EOResource Then

                                'Get file data
                                ReDim Data(PatchInfoHead.lngFileSize - 1)
                                Get PatchFile, , Data
                                
                                'Save it
                                Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, PatchInfoHead
                                Put OutputFile, DataOutputPos, Data
                                
                                'Reanalize last Resource InfoHead
                                EOResource = False
                                ResourceReadFiles = ResourceReadFiles - 1
                                
                                'Update
                                DataOutputPos = DataOutputPos + UBound(Data) + 1
                                WrittenFiles = WrittenFiles + 1
                                If Not prgBar Is Nothing Then prgBar.value = WrittenFiles
                            Else
                                Err.Description = "Incongruencia en archivos de recurso"
                                GoTo ErrHandler
                            End If
                        
                        'Modify
                        Case PatchInstruction.Modify_File
                            If InfoHead.strFileName = PatchInfoHead.strFileName Then

                                'Get file data
                                ReDim Data(PatchInfoHead.lngFileSize - 1)
                                Get PatchFile, , Data
                                
                                'Save it
                                Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, PatchInfoHead
                                Put OutputFile, DataOutputPos, Data
                                
                                'Update
                                DataOutputPos = DataOutputPos + UBound(Data) + 1
                                WrittenFiles = WrittenFiles + 1
                                If Not prgBar Is Nothing Then prgBar.value = WrittenFiles
                            Else
                                Err.Description = "Incongruencia en archivos de recurso"
                                GoTo ErrHandler
                            End If
                    End Select
                    
                    DoEvents
                Wend
                
                'Read everything?
                While ReadNextInfoHead(ResourceFile, FileHead, InfoHead, ResourceReadFiles)
                    'GetData and update InfoHeader
                    Call GetFileRawData(ResourcePath, InfoHead, Data)
                    InfoHead.lngFileStart = DataOutputPos
          
                    'Save file!
                    Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, InfoHead
                    Put OutputFile, DataOutputPos, Data
                    
                    'Update
                    DataOutputPos = DataOutputPos + UBound(Data) + 1
                    WrittenFiles = WrittenFiles + 1
                    If Not prgBar Is Nothing Then prgBar.value = WrittenFiles
                    DoEvents
                Wend
            
            'Close the patch file
            Close OutputFile
        
        'Close the new binary file
        Close PatchFile
    
    'Close the old binary file
    Close ResourceFile
    
    'Check integrity
    If (PatchFileHead.lngNumFiles = WrittenFiles) Then
            'Replace File
            Call Kill(ResourceFilePath)
            Name OutputFilePath As ResourceFilePath
    Else
        Err.Description = "Falla al procesar parche"
        GoTo ErrHandler
    End If
    
    Apply_Patch = True
Exit Function

ErrHandler:
    Close OutputFile
    Close PatchFile
    Close ResourceFile
    'Destroy file if created
    If FileExist(OutputFilePath, vbNormal) Then Call Kill(OutputFilePath)
    
    Call MsgBox("No se pudo parchear. Razon: " & Err.number & " : " & Err.Description, vbOKOnly, "Error")
End Function

Private Function AlignScan(ByVal inWidth As Long, ByVal inDepth As Integer) As Long
'*****************************************************************
'Author: Unknown
'Last Modify Date: Unknown
'*****************************************************************
    AlignScan = (((inWidth * inDepth) + &H1F) And Not &H1F&) \ &H8
End Function

''
' Retrieves the version number of a given resource file.
'
' @param    ResourceFilePath The resource file complete path.
'
' @return   The version number of the given file.

Public Function GetVersion(ByVal ResourceFilePath As String) As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/23/2008
'
'*****************************************************************
    Dim ResourceFile As Integer
    Dim FileHead As FILEHEADER
    
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
        'Extract the FILEHEADER
        Get ResourceFile, 1, FileHead

    Close ResourceFile
    
    GetVersion = FileHead.lngFileVersion
End Function
