VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MsComCtl.ocx"
Begin VB.Form frmEncapsulateField 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Encapsulate fields"
   ClientHeight    =   8070
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Begin GilmansAddinVB6.NumTextBox ntbSpacesInTab 
      Height          =   285
      Left            =   6360
      TabIndex        =   15
      Top             =   3480
      Width           =   495
      _ExtentX        =   873
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
      Precision       =   2
      Scal            =   0
      Text            =   "4"
   End
   Begin VB.TextBox txtPrivateVarName 
      Height          =   375
      Left            =   3240
      TabIndex        =   14
      Text            =   "this"
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox txtPrivateTypeName 
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Text            =   "InternalType"
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   7800
      TabIndex        =   11
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox txtPreview 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   4440
      Width           =   8655
   End
   Begin VB.CommandButton cmdDeselectAll 
      Caption         =   "Deselect All"
      Height          =   255
      Left            =   7920
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame fraProperty 
      Caption         =   "Property"
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   4695
      Begin VB.OptionButton optLet_Set 
         Caption         =   "Let/Set"
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optLet_Set 
         Caption         =   "Set"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optLet_Set 
         Caption         =   "Let"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox chkReadOnly 
         Caption         =   "Read Only"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select All"
      Height          =   255
      Left            =   6600
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin MSComctlLib.ListView lvwFields 
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "VAR"
         Text            =   "Nombre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "TIPO"
         Text            =   "Tipo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "READONLY"
         Text            =   "Read Only"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "LETSET"
         Text            =   "Let/Set"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label lblSpacesInTab 
      Caption         =   "Spaces In tab"
      Height          =   255
      Left            =   5160
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
End
Attribute VB_Name = "frmEncapsulateField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "GilmanAddinVB6.Forms"

Option Explicit

Private vbInstance As VBIDE.VBE
Private oConnect As Connect


Private oCodeModule As VBIDE.CodeModule
Private oProject As VBProject

Private Lines As Collection
'Contiene la declarción del tipo sin el End Type
Private CurrentPrivateType As String
Private CurrentPrivateTypeStartLine As Long
Private CurrentPrivateTypeEndLine As Long

Private thisLine As Long
Private Sub LoadVars()
    Dim var As clsMyVar
    Dim Line As clsLine
    
    Dim lItem As MSComctlLib.ListItem
    On Error GoTo Err_LoadVars
    For Each Line In Lines
        For Each var In Line
            Set lItem = Me.lvwFields.ListItems.Add(, var.Nombre, var.Nombre)
            lItem.SubItems(1) = var.Tipo
            lItem.SubItems(2) = False
            Select Case var.Tipo
                Case "Object"
                    lItem.SubItems(3) = "Set"
                Case "Variant"
                    lItem.SubItems(3) = "Let/Set"
                Case Else
                    lItem.SubItems(3) = "Let"
            End Select
            lItem.Tag = "LIN" & Line.CodeLocation
        Next var
    Next Line

    Exit Sub
Err_LoadVars:
    If Erl = 0 Then
        Err.Raise Err.Number, "frmEncapsulateField.LoadVars" & vbCrLf & Err.Source, Err.Description
    Else
        Err.Raise Err.Number, "frmEncapsulateField.LoadVars Linea " & Erl & vbCrLf & Err.Source, Err.Description
    End If
End Sub


Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub chkReadOnly_Click()
    lvwFields.SelectedItem.SubItems(2) = chkReadOnly.value = vbChecked
End Sub

Private Sub cmdDeselectAll_Click()
    SelectItems False
    
End Sub


Private Sub cmdRefresh_Click()
    Dim DeclaresLines As String
    Dim Text As cStringBuilder
    On Error GoTo Err_cmdRefresh_Click
    DeclaresLines = Declares
    
    
    If DeclaresLines <> "" Then
        Set Text = New cStringBuilder
        Text.Append DeclaresLines
        
        Text.Append NewProperties
        
        txtPreview.Text = Text.ToString
    Else
        txtPreview.Text = ""
    End If


    Exit Sub
Err_cmdRefresh_Click:
    Screen.MousePointer = vbDefault
    If Erl = 0 Then
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Sub 'cmdRefresh_Click' del Formulario frmEncapsulateField" & vbCrLf & "[" & Err.Source & "]", vbCritical
    Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Sub 'cmdRefresh_Click' del Formulario frmEncapsulateField en linea: " & Erl & vbCrLf & "[" & Err.Source & "]", vbCritical
    End If
End Sub

Private Sub cmdSelectAll_Click()
    SelectItems True
End Sub


Private Sub SelectItems(ByVal All As Boolean)
    Dim lItem As MSComctlLib.ListItem
    For Each lItem In lvwFields.ListItems
        lItem.Checked = All
        Lines(lItem.Tag)(lItem.Text).Removed = All
    Next lItem
End Sub

Private Sub lvwFields_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Lines(Item.Tag)(Item.Text).Removed = Item.Checked
End Sub

Private Sub lvwFields_ItemClick(ByVal Item As MSComctlLib.ListItem)
    chkReadOnly.value = IIf(Item.SubItems(2), vbChecked, vbUnchecked)
    Select Case Item.SubItems(3)
        Case "Let"
            optLet_Set(0).value = True
        Case "Set"
            optLet_Set(1).value = True
        Case "Let/Set"
            optLet_Set(2).value = True
    End Select
    optLet_Set(2).Enabled = Item.SubItems(1) = "Variant"
End Sub

Private Sub cmdOK_Click()
    Dim DeclaresLines As String
    On Error GoTo Err_cmdOK_Click
    DeclaresLines = Declares
    
    
    If DeclaresLines <> "" Then
        BorrarLineas
        oCodeModule.InsertLines oCodeModule.CountOfDeclarationLines, DeclaresLines
        oCodeModule.InsertLines oCodeModule.CountOfLines, NewProperties
        oCodeModule.CodePane.SetSelection 1, 1, 1, 1
        Me.Hide
    End If
    Exit Sub
Err_cmdOK_Click:
    Screen.MousePointer = vbDefault
    If Erl = 0 Then
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Sub 'cmdOK_Click' del Formulario frmEncapsulateField" & vbCrLf & "[" & Err.Source & "]", vbCritical
    Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Sub 'cmdOK_Click' del Formulario frmEncapsulateField en linea: " & Erl & vbCrLf & "[" & Err.Source & "]", vbCritical
    End If
End Sub


Private Sub optLet_Set_Click(Index As Integer)
    lvwFields.SelectedItem.SubItems(3) = optLet_Set(Index).Caption
End Sub





Public Sub LoadFields(ByVal Instance As VBIDE.VBE, ByVal Conn As Connect)

    Dim x As Member
    Dim Vars() As String
    Dim Comentario As String
    Dim var As clsMyVar
    Dim CommentPos As Long
    Dim OldLine As Long
    Dim Data() As String
    Dim Line As clsLine
    Dim n As Long
    
    
    
    On Error GoTo Err_LoadFields
    
    Set vbInstance = Instance
    Set oConnect = Conn
    Set oCodeModule = vbInstance.ActiveCodePane.CodeModule
    Set oProject = oCodeModule.Parent.Collection.Parent
    getCurrentPrivateType
    If CurrentPrivateType = "" Then
        MsgBox "El tipo privado está mal declarado."
        Exit Sub
    End If
    Set Lines = New Collection
    
    For Each x In oCodeModule.Members
        If x.Type = vbext_mt_Variable And x.Scope = vbext_Public Then
            If OldLine <> x.CodeLocation Then
                Set Line = New clsLine
                Line.Code = vbInstance.ActiveCodePane.CodeModule.Lines(x.CodeLocation, 1)
                CommentPos = InStr(1, Line.Code, "'")
                If CommentPos > 0 Then
                    Line.Comentario = Mid$(Line.Code, CommentPos + 1)
                    Line.Code = Left(Line.Code, CommentPos - 1)
                End If
                
                Line.CodeLocation = x.CodeLocation
                RemoveSpaces Line.Code
                
                Vars = Split(Trim(Line.Code), ",")
                
                For n = 0 To UBound(Vars)
                    Set var = New clsMyVar
                    Data = Split(Trim(Vars(n)))
                
                    If n = 0 Then
                       
                        
                        var.Nombre = Data(1)
                        If UBound(Data) = 1 Then
                            var.Tipo = "Variant"
                        Else
                            var.Tipo = Data(3)
                        End If
                    Else
                        var.Nombre = Data(0)
                        If UBound(Data) = 0 Then
                            var.Tipo = "Variant"
                        Else
                            var.Tipo = Data(2)
                        End If
                    End If
                    Line.Add var, var.Nombre
                Next n
                OldLine = Line.CodeLocation
                Lines.Add Line, "LIN" & Line.CodeLocation
            End If
        End If
    Next x
    If Lines.Count = 0 Then
        MsgBox "No se a encontrado ningún campo"
    Else
        LoadVars
        Me.Show vbModal
    End If

    Exit Sub
Err_LoadFields:
    If Erl = 0 Then
        Err.Raise Err.Number, "frmEncapsulateField.LoadFields" & vbCrLf & Err.Source, Err.Description
    Else
        Err.Raise Err.Number, "frmEncapsulateField.LoadFields Linea " & Erl & vbCrLf & Err.Source, Err.Description
    End If
End Sub
'CurrentPrivateType contendrá la definición original del tipo
'Si no está definido contedrá:
'Private Type InternalType
'End Type
'
'Si hay una sentencia Private Type InternalType, pero no una End Type posterior contendrá una cadena vacía
'TODO: [-] No está permitida la declaración del tipo en la cual la sentencia End Type esté en la misma línea que la declaración de un miembro del tipo
'La declaración
'Private Type InternalType
'   A As Long
'B As Double: End Type
'No está permitida
'Si está permitida una declaración así
'Private Type InternalType:   A As Long:    B As Double
'End Type
'también está permitida
Private Sub getCurrentPrivateType()
    Dim n As Long
    Dim NL As Long
    CurrentPrivateType = ""
    CurrentPrivateTypeStartLine = 0
    CurrentPrivateTypeEndLine = 0
    thisLine = 0
    If oCodeModule.Find("Private Type " & txtPrivateTypeName.Text & "", 1, 0, oCodeModule.CountOfDeclarationLines, -1) Then
        n = 1
        While InStr(1, oCodeModule.Lines(n, 1), "Private Type " & txtPrivateTypeName.Text) = 0
            n = n + 1
        Wend
        CurrentPrivateTypeStartLine = n
        If oCodeModule.Find("End Type", n, 0, oCodeModule.CountOfDeclarationLines + 1, -1) Then
            While InStr(1, oCodeModule.Lines(n, 1), "End Type") = 0
                n = n + 1
            Wend
            CurrentPrivateTypeEndLine = n
            CurrentPrivateType = oCodeModule.Lines(CurrentPrivateTypeStartLine, CurrentPrivateTypeEndLine - CurrentPrivateTypeStartLine)
            
        End If
        'thisLine
        'Private this As InternalType
        If oCodeModule.Find("Private " & txtPrivateVarName.Text & " As " & txtPrivateTypeName.Text & "", n, 0, oCodeModule.CountOfDeclarationLines + 1, -1) Then
            While InStr(1, oCodeModule.Lines(n, 1), "Private " & txtPrivateVarName.Text & " As " & txtPrivateTypeName.Text) = 0
                n = n + 1
            Wend
            thisLine = n
        End If

    Else
        CurrentPrivateType = "Private Type " & txtPrivateTypeName.Text & ""
    End If
End Sub

' ----------------------------------------------------------------
' Procedure Name: BorrarLineas
' Purpose: Borra las lineas que sobran tras insertar los cambios
' Procedure Kind: Sub
' Procedure Access: Private
' Author: alber
' Date: 01/03/2025
'
' ----------------------------------------------------------------
Private Sub BorrarLineas()
    'Se deben borrar las lineas en orden inverso al que aparecen en el código
    'En la colección Lines están las lineas cargadas en orden inverso
    Dim oL As clsLine
    Dim CPTBorrado As Boolean
    
    CPTBorrado = CurrentPrivateTypeEndLine = 0
    For Each oL In Lines
        If oL.CodeLocation > CurrentPrivateTypeEndLine Or CPTBorrado Then
            oCodeModule.DeleteLines oL.CodeLocation
        Else
            'Si la definición del tipo es posterior, primero borrar el tipo, así la linea estará en la misma posición
            'Primero borrar la declaración de this
            
            oCodeModule.DeleteLines thisLine
            
            oCodeModule.DeleteLines CurrentPrivateTypeStartLine, CurrentPrivateTypeEndLine - CurrentPrivateTypeStartLine + 1
            CPTBorrado = True
            oCodeModule.DeleteLines oL.CodeLocation
        End If
    Next oL
End Sub


Private Function NewPrivateType(ByVal TabSpace As String) As String
    Dim lItem As ListItem
    Dim Text As cStringBuilder: Set Text = New cStringBuilder
    
    Dim NewVarAdded As Boolean
    getCurrentPrivateType
    If CurrentPrivateType = "" Then
        Text.Append "Private Type " & Me.txtPrivateTypeName.Text & " & vbCrLf"
    Else
        Text.Append CurrentPrivateType & vbCrLf
    End If
    For Each lItem In lvwFields.ListItems
        If lItem.Checked Then
            Text.Append TabSpace & lItem.Text & " As " & lItem.SubItems(1) & vbCrLf
            NewVarAdded = True
        End If
    Next lItem
    If NewVarAdded Then
        Text.Append "End Type" & vbCrLf
        Text.Append "Private " & txtPrivateVarName.Text & " AS " & txtPrivateTypeName.Text & "" & vbCrLf

        NewPrivateType = Text.ToString
    End If
End Function


Private Function Declares() As String
    Dim Text As cStringBuilder: Set Text = New cStringBuilder
    Dim oL As clsLine
    Dim NewCode As String
    Dim TabSpace As String
    Dim NPT As String
    Dim lItem As ListItem
    
    TabSpace = Space(Me.ntbSpacesInTab.value)
    'Primero los campos publicos que no han sido asignados al tipo Privado
    For Each oL In Lines
        NewCode = oL.NewCode
        If NewCode <> "" Then
            Text.Append NewCode & vbCrLf
        End If
    Next oL
    'Luego el tipo Privado.
    NPT = NewPrivateType(TabSpace)
    If NPT <> vbNullString Then
        Text.Append NPT & vbCrLf
        Declares = Text.ToString
    End If
End Function


Private Function NewProperties()
    Dim Text As cStringBuilder: Set Text = New cStringBuilder
    Dim oL As clsLine
    Dim NewCode As String
    Dim TabSpace As String
    Dim NPT As String
    Dim lItem As ListItem
    
    TabSpace = Space(Me.ntbSpacesInTab.value)
    For Each lItem In lvwFields.ListItems
        Set oL = Lines(lItem.Tag)
        If lItem.Checked Then
            Text.Append "Public Property Get " & lItem.Text & "() As " & lItem.SubItems(1) & vbCrLf
            Select Case lItem.SubItems(3)
                Case "Let/Set"
                    Text.Append TabSpace & "If IsObject(" & txtPrivateVarName.Text & "." & lItem.Text & ") Then" & vbCrLf
                    Text.Append TabSpace & TabSpace & "Set " & lItem.Text & " = " & txtPrivateVarName.Text & "." & lItem.Text & vbCrLf
                    Text.Append TabSpace & "Else" & vbCrLf
                    Text.Append TabSpace & TabSpace & lItem.Text & " = " & txtPrivateVarName.Text & "." & lItem.Text & vbCrLf
                    Text.Append TabSpace & "End If" & vbCrLf
                Case "Let"
                    Text.Append TabSpace & lItem.Text & " = " & txtPrivateVarName.Text & "." & lItem.Text & vbCrLf
                Case "Set"
                    Text.Append TabSpace & "Set " & lItem.Text & " = " & txtPrivateVarName.Text & "." & lItem.Text & vbCrLf
            End Select
            Text.Append "End Property" & vbCrLf
                
            If Not CBool(lItem.SubItems(2)) Then
                If InStr(1, lItem.SubItems(3), "Let") > 0 Then
                    Text.Append "Public Property Let " & lItem.Text & "(ByVal RHC As " & lItem.SubItems(1) & ")" & vbCrLf
                    Text.Append TabSpace & txtPrivateVarName.Text & "." & lItem.Text & " = RHC" & vbCrLf
                    Text.Append "End Property" & vbCrLf
                End If
                If InStr(1, lItem.SubItems(3), "Set") > 0 Then
                    Text.Append "Public Property Set " & lItem.Text & "(ByVal RHC As " & lItem.SubItems(1) & ")" & vbCrLf
                    Text.Append TabSpace & "Set " & txtPrivateVarName.Text & "." & lItem.Text & " = RHC" & vbCrLf
                    Text.Append "End Property" & vbCrLf
                End If
            End If
        End If
    Next lItem
    NewProperties = Text.ToString
End Function
