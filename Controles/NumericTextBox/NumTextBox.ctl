VERSION 5.00
Begin VB.UserControl NumTextBox 
   ClientHeight    =   1005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2970
   ScaleHeight     =   1005
   ScaleWidth      =   2970
   ToolboxBitmap   =   "NumTextBox.ctx":0000
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Width           =   2805
   End
End
Attribute VB_Name = "NumTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Utilidades.Controles"
Option Explicit
'Default Property Values:
Const m_def_Enabled = 0
Const m_def_LimitedDecimals = True
Private Const m_def_Alignment = AlignmentConstants.vbRightJustify
Private Const m_def_Value = Null
Private Const m_def_BackStyle = 0
Private Const m_def_BorderStyle = 0
Private Const m_def_Precision = 10
Private Const m_def_Scale = 2
'Property Variables:
Dim m_Enabled As Boolean
Dim m_LimitedDecimals As Boolean
Private m_Value As Variant
Private m_BackStyle As Integer
Private m_BorderStyle As Integer
Private m_Precision As Long
Private m_Scale As Long


Private SepDec As String

Private blnSepDec As Boolean

Public Event Change()
Public Event EnterPresed()


Private Type tLastValue
    Text As String
    SelStart As Long
    SelLength As Long
End Type

Private LastValue As tLastValue

Public Property Let Alignment(ByVal value As AlignmentConstants)
    txtValor.Alignment = value
    PropertyChanged "Alignment"
End Property

Public Property Get Alignment() As AlignmentConstants
    Alignment = txtValor.Alignment
End Property

Private Sub txtValor_Change()
    Select Case Trim$(txtValor.Text)
        Case "", SepDec, "-", "-" & SepDec, "+", "+" & SepDec
        Case Else
            If Not IsNumeric(txtValor.Text) Then
                With LastValue
                    txtValor.Text = .Text
                    txtValor.SelStart = .SelStart
                    txtValor.SelLength = .SelLength
                End With
            Else
                If ExceedCapacity(txtValor.Text) Then
                    With LastValue
                        txtValor.Text = .Text
                        txtValor.SelStart = .SelStart
                        txtValor.SelLength = .SelLength
                    End With
                End If
            End If
    End Select
    If LastValue.Text <> txtValor.Text Then
        RaiseEvent Change
    End If
End Sub


Private Sub txtValor_KeyDown(KeyCode As Integer, Shift As Integer)
    blnSepDec = KeyCode = vbKeyDecimal
End Sub


Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If blnSepDec Then
        Debug.Print KeyAscii
        KeyAscii = Asc(SepDec)
        Debug.Print KeyAscii
        Debug.Print SepDec
    End If
    If InStr(1, "ED", UCase$(Chr$(KeyAscii))) > 0 Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        RaiseEvent EnterPresed
    ElseIf InStr(1, "+-", Chr$(KeyAscii)) > 0 And txtValor.SelStart <> 0 Then
        KeyAscii = 0
    Else
        If blnSepDec Then
            KeyAscii = Asc(SepDec)
        End If
        With LastValue
            .Text = txtValor.Text
            .SelStart = txtValor.SelStart
            .SelLength = txtValor.SelLength
        End With
    End If
    
End Sub


Private Sub UserControl_Initialize()
    SepDec = Format(0, ".")
    With LastValue
        .Text = txtValor.Text
    End With
End Sub

Private Sub UserControl_Resize()
    txtValor.Move 0, 0, UserControl.Width, UserControl.Height
    If UserControl.Height < txtValor.Height Then
        UserControl.Height = txtValor.Height
    End If
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtValor,txtValor,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = txtValor.BackColor
End Property


Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtValor.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtValor,txtValor,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = txtValor.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtValor.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=txtLastValue,txtLastValue,-1,Enabled
'Public Property Get Enabled() As Boolean
'    Enabled = txtValor.Enabled
'End Property
'
'Public Property Let Enabled(ByVal New_Enabled As Boolean)
'    txtValor.Enabled() = New_Enabled
'    PropertyChanged "Enabled"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtValor,txtValor,-1,Font
Public Property Get Font() As Font
    Set Font = txtValor.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtValor.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get BackStyle() As Integer
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Public Property Get BorderStyle() As Integer
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Sub Refresh()
     
End Sub


Public Property Get Precision() As Long
    Precision = m_Precision
End Property

Public Property Let Precision(ByVal New_Precision As Long)
    m_Precision = New_Precision
    PropertyChanged "Precision"
End Property

Public Property Get Scal() As Long
    Scal = m_Scale
End Property

Public Property Let Scal(ByVal New_Scale As Long)
    m_Scale = New_Scale
    PropertyChanged "Scal"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    txtValor.Alignment = m_def_Alignment
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_Precision = m_def_Precision
    m_Scale = m_def_Scale
    
    txtValor.Text = ""
    m_Value = m_def_Value
    m_Enabled = m_def_Enabled
    m_LimitedDecimals = m_def_LimitedDecimals
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    txtValor.Alignment = PropBag.ReadProperty("Alignment", vbRightJustify)

    txtValor.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtValor.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtValor.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_Precision = PropBag.ReadProperty("Precision", m_def_Precision)
    m_Scale = PropBag.ReadProperty("Scal", m_def_Scale)
    txtValor.Text = PropBag.ReadProperty("Text", "")
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_LimitedDecimals = PropBag.ReadProperty("LimitedDecimals", m_def_LimitedDecimals)
    txtValor.SelText = PropBag.ReadProperty("SelText", "")
    txtValor.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtValor.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtValor.Locked = PropBag.ReadProperty("Locked", False)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Alignment", txtValor.Alignment, vbRightJustify)

    Call PropBag.WriteProperty("BackColor", txtValor.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", txtValor.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", txtValor.Enabled, True)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Precision", m_Precision, m_def_Precision)
    Call PropBag.WriteProperty("Scal", m_Scale, m_def_Scale)
    Call PropBag.WriteProperty("Text", txtValor.Text, "")
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("LimitedDecimals", m_LimitedDecimals, m_def_LimitedDecimals)
    Call PropBag.WriteProperty("SelText", txtValor.SelText, "")
    Call PropBag.WriteProperty("SelStart", txtValor.SelStart, 0)
    Call PropBag.WriteProperty("SelLength", txtValor.SelLength, 0)
    Call PropBag.WriteProperty("Locked", txtValor.Locked, False)
End Sub


Private Function ExceedCapacity(ByRef Num As String) As Boolean

    Dim l() As String
    'If in the number there is an  E or D the number is in scientific notation
    'At the moment we do not accept numbers in scientific notation
    ExceedCapacity = InStr(1, Replace$(UCase$(Num), "D", "E"), "E") > 0
    
    If Not ExceedCapacity Then
        'Ignore the left zeroes  of the integer part
        l = Split(Replace$(Abs(Num), "0", " "), SepDec)
        l(0) = LTrim$(l(0))
        If UBound(l) = 1 Then
            l(1) = RTrim$(l(1))
        End If
        If m_LimitedDecimals Then
            If UBound(l) >= 0 Then
                ExceedCapacity = Len(l(0)) > m_Precision - m_Scale
                If UBound(l) = 1 And Not ExceedCapacity Then
                    'The zeros of the right side of the number in l(0) are spaces, ignore then
                    ExceedCapacity = Len(RTrim$(l(1))) > m_Scale
                End If
            End If
        Else
            If UBound(l) = 1 Then
                ExceedCapacity = Len(l(0)) + Len(l(1)) > m_Precision
            Else
                ExceedCapacity = Len(l(0)) > m_Precision
            End If
        End If
    End If

End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtValor,txtValor,-1,Text
Public Property Get Text() As String
    Text = txtValor.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtValor.Text() = New_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,2,Null
Public Property Get value() As Variant
    If txtValor.Text <> "" Then
        value = CCur(txtValor.Text)
    Else
        value = Null
    End If
End Property

Public Property Let value(ByVal New_Value As Variant)
    If Ambient.UserMode = False Then Err.Raise 387
    txtValor.Text = New_Value & ""
    PropertyChanged "Value"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get LimitedDecimals() As Boolean
    LimitedDecimals = m_LimitedDecimals
End Property

Public Property Let LimitedDecimals(ByVal New_LimitedDecimals As Boolean)
    m_LimitedDecimals = New_LimitedDecimals
    PropertyChanged "LimitedDecimals"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtValor,txtValor,-1,SelText
Public Property Get SelText() As String
    SelText = txtValor.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    txtValor.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtValor,txtValor,-1,SelStart
Public Property Get SelStart() As Long
    SelStart = txtValor.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtValor.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtValor,txtValor,-1,SelLength
Public Property Get SelLength() As Long
    SelLength = txtValor.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtValor.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtValor,txtValor,-1,Locked
Public Property Get Locked() As Boolean
    Locked = txtValor.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtValor.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

