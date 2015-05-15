Attribute VB_Name = "modCarga"
Option Explicit

'Parra: No me gusta nada este modulo, hay que pensar algo.. (esto.. provisional)

Public Sub InicializarNombres()
'**************************************************************
'Author: •Parra
'Last Modify Date: 03/09/2009
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc... desde un archivo.
'**************************************************************
    Dim handle As Long
    
    handle = FreeFile()
    
    Open App.path & "\Init\Nombres.cnfig" For Binary Access Read As #handle
    
        Get #handle, , Ciudades()
        Get #handle, , ListaRazas()
        Get #handle, , ListaClases()
        Get #handle, , SkillsNames()
        Get #handle, , AtributosNames()
        
    Close #handle

End Sub

Sub CargarVersiones()
On Error GoTo errorH:
'Parra: ¡¡¡QUE PORQUERIA DE SUB!!!! -__-U
    
    Versiones(1) = Val(GetVar(App.path & "\init\" & "versiones.ini", "Graficos", "Val"))
    Versiones(2) = Val(GetVar(App.path & "\init\" & "versiones.ini", "Wavs", "Val"))
    Versiones(3) = Val(GetVar(App.path & "\init\" & "versiones.ini", "Midis", "Val"))
    Versiones(4) = Val(GetVar(App.path & "\init\" & "versiones.ini", "Init", "Val"))
    Versiones(5) = Val(GetVar(App.path & "\init\" & "versiones.ini", "Mapas", "Val"))
    Versiones(6) = Val(GetVar(App.path & "\init\" & "versiones.ini", "E", "Val"))
    Versiones(7) = Val(GetVar(App.path & "\init\" & "versiones.ini", "O", "Val"))
Exit Sub

errorH:
    Call MsgBox("Error cargando versiones")
End Sub

Sub CargarColores()
On Error Resume Next
    Dim archivoC As String
    
    archivoC = App.path & "\init\colores.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim i As Long
    
    For i = 0 To 48 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(i).r = CByte(GetVar(archivoC, CStr(i), "R"))
        ColoresPJ(i).g = CByte(GetVar(archivoC, CStr(i), "G"))
        ColoresPJ(i).B = CByte(GetVar(archivoC, CStr(i), "B"))
    Next i
    
    ColoresPJ(50).r = CByte(GetVar(archivoC, "CR", "R"))
    ColoresPJ(50).g = CByte(GetVar(archivoC, "CR", "G"))
    ColoresPJ(50).B = CByte(GetVar(archivoC, "CR", "B"))
    ColoresPJ(49).r = CByte(GetVar(archivoC, "CI", "R"))
    ColoresPJ(49).g = CByte(GetVar(archivoC, "CI", "G"))
    ColoresPJ(49).B = CByte(GetVar(archivoC, "CI", "B"))
End Sub


