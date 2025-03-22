Attribute VB_Name = "Main"
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: Printf
' Purpose: Formats a string by replacing the fields with the corresponding values
' Procedure Kind: Function
' Procedure Access: Public
' Parameter Message (String): String to format, the field have to match with #i# i being the order of the value with the first element as 0 pos
'                             the char # must be s
' Parameter vars (Variant): values to put in the Message string.
'                           This function is used internaly, so the vars is not normal param array, it has a unique value containing an array with the original
'                           Param array
' Return Type: String
' Author: alber
' Date: 22/03/2025
' ----------------------------------------------------------------
Public Function Printf(ByRef Message As String, ParamArray vars() As Variant) As String
    
    Dim I As Long
    Dim t() As String
    On Error GoTo Err_Printf
    ' Check if the input format string is empty; if so, return an empty string
    If Len(Message) = 0 Then
        Printf = vbNullString
    Else
        Printf = Replace(Message, "\\", Chr$(1))
        Printf = Replace(Printf, "\t", vbTab)
        Printf = Replace(Printf, "\n", vbNewLine)
        Printf = Replace(Printf, "\T", vbTab)
        Printf = Replace(Printf, "\N", vbNewLine)
        Printf = Replace(Printf, "\%", Chr$(2))
    
        ' Check if no arguments were provided (empty ParamArray); if so, return the format string unchanged
        If UBound(vars) >= LBound(vars) Then
            'first replace \\
            If InStr(Printf, "\") > 0 Then
                'there should be no more \
                Err.Raise -1, "printf", "Bad formatted string"
            End If
    
            If Left(Printf, 1) = "%" Then
                'preppend a char if the first word in the message is an argument
                Printf = Chr$(3) & Printf
            End If
            t = Split(Printf, "%")
            For I = 1 To UBound(t) Step 2
                t(I) = vars(0)(t(I))
            Next I
            Printf = Join(t, vbNullString)
        End If
    
        Printf = Replace(Printf, Chr$(1), "\")
        Printf = Replace(Printf, Chr$(2), "%")
        Printf = Replace(Printf, Chr$(3), vbNullString)
    End If

    Exit Function
Err_Printf:
    If Erl = 0 Then
        Err.Raise Err.Number, "Main.Printf" & vbCrLf & Err.Source, Err.Description
    Else
        Err.Raise Err.Number, "Main.Printf Linea " & Erl & vbCrLf & Err.Source, Err.Description
    End If
End Function


