Attribute VB_Name = "General"
'@Folder "GilmanAddinVB6.Modules"
Option Explicit

Public Sub RemoveSpaces(ByRef Cadena As String)
    
    Dim l As Long
'    Dim oRep As clsReplace
    On Error GoTo Err_RemoveSpaces
'    Set oRep = New clsReplace
    l = Len(Cadena)
    Cadena = clsReplace.Replace(Cadena, "  ", " ")
    While l <> Len(Cadena)
        Cadena = clsReplace.Replace(Cadena, "  ", " ")
        l = Len(Cadena)
    Wend

    Exit Sub
Err_RemoveSpaces:
    If Erl = 0 Then
        Err.Raise Err.Number, "frmAddIn.RemoveSpaces" & vbCrLf & Err.Source, Err.Description
    Else
        Err.Raise Err.Number, "frmAddIn.RemoveSpaces Linea " & Erl & vbCrLf & Err.Source, Err.Description
    End If
End Sub


Private Function ModuleLine(ByVal RealLine As Long, ByVal ModuleFile As String) As Long
    Dim oFS As Scripting.FileSystemObject: Set oFS = New Scripting.FileSystemObject
    Dim oTS As Scripting.TextStream
    Dim Line As String
    On Error GoTo Err_ModuleLine
    Set oTS = oFS.OpenTextFile(ModuleFile, ForReading)
    
    While Left(Line, 10) <> "Attribute "
        Line = oTS.ReadLine
        RealLine = RealLine - 1
    Wend
    'Estamos en la primera linea de atributos del archivo
    While Left(Line, 10) = "Attribute "
        Line = oTS.ReadLine
        RealLine = RealLine - 1
    Wend
    'Estamos en la primera linea del modulo
    ModuleLine = 1
    While RealLine > 0
        Line = oTS.ReadLine
        If Left(Line, 10) <> "Attribute " Then
            'es una linea de código
            ModuleLine = ModuleLine + 1
        End If
        RealLine = RealLine - 1
    Wend
    'hemos llegado a la linea de módulo
    oTS.Close
    

    Exit Function
Err_ModuleLine:
    'El único error posible es que el archivo no tenga suficientes líneas.
    Dim FileLines As Long
    FileLines = oTS.Line
    oTS.Close
    If Err.Number = 62 Then
        Err.Raise -1, "Class1.ModuleLine" & vbCrLf & Err.Source, "El fichero del módulo solo tiene " & FileLines & " líneas."
    ElseIf Erl = 0 Then
        Err.Raise Err.Number, "Class1.ModuleLine" & vbCrLf & Err.Source, Err.Description
    Else
        Err.Raise Err.Number, "Class1.ModuleLine Linea " & Erl & vbCrLf & Err.Source, Err.Description
    End If
End Function

