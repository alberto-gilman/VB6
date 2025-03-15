VERSION 5.00
Begin VB.Form frmGotoLine 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Goto Line"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin GilmansAddinVB6.NumTextBox ntbLine 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkLineaReal 
      Caption         =   "Línea Real"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   0
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Ir a Linea"
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmGotoLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "GilmanAddinVB6.Forms"

Option Explicit

Public vbInstance As VBIDE.VBE
Public oConnect As Connect

Private Sub cmdCancel_Click()
    Me.Hide
End Sub


Private Sub cmdOK_Click()
    Dim vbComp As VBComponent
    Dim LineaEnModulo As Long
    Dim Respuesta As VbMsgBoxResult
    On Error GoTo Err_OKButton_Click
    If Not IsNumeric(Me.ntbLine.value) Then
        MsgBox "Introduzca un número de línea correcto."
    Else
        If Me.chkLineaReal.value = vbChecked Then
            Set vbComp = vbInstance.ActiveCodePane.CodeModule.Parent
            If vbComp.IsDirty Then
                If vbComp.FileNames(1) <> "" Then

                    Respuesta = MsgBox("Para localizar la línea es necesario guardar primero el archivo. ¿Desea guardarlo ahora?", vbYesNo + vbQuestion)
                    If Respuesta = vbYes Then
                        vbComp.SaveAs vbComp.FileNames(1)
                    Else
                        Exit Sub
                    End If
                Else
                    MsgBox "Para localizar la línea es necesario guardar primero el archivo.", vbOKOnly + vbInformation
                End If
            End If
            LineaEnModulo = ModuleLine(ntbLine.value, vbComp.FileNames(1))
            vbInstance.ActiveCodePane.SetSelection LineaEnModulo, 1, LineaEnModulo, Len(vbInstance.ActiveCodePane.CodeModule.Lines(LineaEnModulo, 1)) + 1
            Me.Hide
        Else
       
            If CLng(ntbLine.value) < 1 Or CLng(ntbLine.value) > vbInstance.ActiveCodePane.CodeModule.CountOfLines Then
                MsgBox "Introduzca un número entre 1 y " & vbInstance.ActiveCodePane.CodeModule.CountOfLines
            Else
                vbInstance.ActiveCodePane.SetSelection CLng(ntbLine.value), 1, _
                    CLng(ntbLine.value), Len(vbInstance.ActiveCodePane.CodeModule.Lines(CLng(ntbLine.value), 1)) + 1
                Me.Hide
            End If
        End If
    End If

    Exit Sub
Err_OKButton_Click:
    Screen.MousePointer = vbDefault
    If Erl = 0 Then
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Sub 'cmdOK_Click' del Formulario frmAddIn" & vbCrLf & "[" & Err.Source & "]", vbCritical
    Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Sub 'cmdOK_Click' del Formulario frmAddIn en linea: " & Erl & vbCrLf & "[" & Err.Source & "]", vbCritical
    End If
End Sub




Public Function ModuleLine(ByVal RealLine As Long, ByVal ModuleFile As String) As Long
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


