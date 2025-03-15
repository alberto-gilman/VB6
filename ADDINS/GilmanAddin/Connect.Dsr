VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   7440
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   20715
   _ExtentX        =   36539
   _ExtentY        =   13123
   _Version        =   393216
   Description     =   "Utilities from Gilman"
   DisplayName     =   "Gilmans Addin VB6"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 98 (ver 6.0)"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder "GilmanAddinVB6"
Option Explicit

Public FormDisplayed          As Boolean
Public vbInstance             As VBIDE.VBE
'Dim mcbMenuCommandBar         As Office.CommandBarControl

Private cwcmGotoLine     As Office.CommandBarControl
Public WithEvents mhGotoLine As CommandBarEvents          'command bar event handler
Attribute mhGotoLine.VB_VarHelpID = -1


Private cwcmEncapsulateFields     As Office.CommandBarControl
Public WithEvents mhEncapsulateFields As CommandBarEvents          'command bar event handler
Attribute mhEncapsulateFields.VB_VarHelpID = -1

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    Set vbInstance = Application
    
    Set cwcmGotoLine = AddItemToMenu("Code Window", "Goto Line", True)   ' And "Code Window (Break)" ???
    Set mhGotoLine = vbInstance.Events.CommandBarEvents(cwcmGotoLine)
    
    Set cwcmEncapsulateFields = AddItemToMenu("Code Window", "Encapsulate Fields", False)   ' And "Code Window (Break)" ???
    Set mhEncapsulateFields = vbInstance.Events.CommandBarEvents(cwcmEncapsulateFields)
    
'    'save the vb instance
'    Set VBInstance = Application
'
'    'this is a good place to set a breakpoint and
'    'test various addin objects, properties and methods
'    Debug.Print VBInstance.FullName
'
'    If ConnectMode = ext_cm_External Then
'        'Used by the wizard toolbar to start this wizard
'        Me.Show
'    Else
'        Set mcbMenuCommandBar = AddToAddInCommandBar("Gilman Addin")
'
'        'sink the event
'        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
'    End If
'
'    If ConnectMode = ext_cm_AfterStartup Then
'        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
'            'set this to display the form on connect
'            Me.Show
'        End If
'    End If
  
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'delete the command bar entry
'    mcbMenuCommandBar.Delete
    
    'shut down the Add-In
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    

End Sub

Private Sub mhEncapsulateFields_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error GoTo Err_mhEncapsulateFields_Click
    ShowEncapsulateFields

    Exit Sub
Err_mhEncapsulateFields_Click:
    Screen.MousePointer = vbDefault
    If Erl = 0 Then
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Sub 'mhEncapsulateFields_Click' del Diseñador Connect" & vbCrLf & "[" & Err.Source & "]", vbCritical
    Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Sub 'mhEncapsulateFields_Click' del Diseñador Connect en linea: " & Erl & vbCrLf & "[" & Err.Source & "]", vbCritical
    End If
End Sub

'this event fires when the menu is clicked in the IDE
Private Sub mhGotoLine_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error GoTo Err_mhGotoLine_Click
    ShowGotoLine

    Exit Sub
Err_mhGotoLine_Click:
    Screen.MousePointer = vbDefault
    If Erl = 0 Then
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Sub 'mhGotoLine_Click' del Diseñador Connect" & vbCrLf & "[" & Err.Source & "]", vbCritical
    Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Sub 'mhGotoLine_Click' del Diseñador Connect en linea: " & Erl & vbCrLf & "[" & Err.Source & "]", vbCritical
    End If
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = vbInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    
    'set the caption
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function


Private Sub ShowGotoLine()
    Dim ofrmGotoLine As frmGotoLine

    
    Dim CurrentLine As Long, x As Long, Y As Long, z As Long
    On Error GoTo Err_ShowGotoLine
    
    If vbInstance.ActiveWindow Is Nothing Then
        MsgBox "No hay una ventana activa"
    ElseIf vbInstance.ActiveWindow.Type <> vbext_wt_CodeWindow Then
        MsgBox "La ventana activa no es una ventana de código"
    Else
        Set ofrmGotoLine = New frmGotoLine
    
        Set ofrmGotoLine.vbInstance = vbInstance
        Set ofrmGotoLine.oConnect = Me
        vbInstance.ActiveCodePane.GetSelection CurrentLine, x, Y, z
        ofrmGotoLine.ntbLine.Text = CurrentLine
        FormDisplayed = True
        ofrmGotoLine.Show vbModal
    End If
    
   

    Exit Sub
Err_ShowGotoLine:
    If Erl = 0 Then
        Err.Raise Err.Number, "Connect.ShowGotoLine" & vbCrLf & Err.Source, Err.Description
    Else
        Err.Raise Err.Number, "Connect.ShowGotoLine Linea " & Erl & vbCrLf & Err.Source, Err.Description
    End If
End Sub

Private Sub ShowEncapsulateFields()
    On Error GoTo Err_ShowEncapsulateFields
    If vbInstance.ActiveWindow Is Nothing Then
        MsgBox "No hay una ventana activa"
    ElseIf vbInstance.ActiveWindow.Type <> vbext_wt_CodeWindow Then
        MsgBox "La ventana activa no es una ventana de código"
    Else
        LoadFields
    End If


    Exit Sub
Err_ShowEncapsulateFields:
    If Erl = 0 Then
        Err.Raise Err.Number, "Connect.ShowEncapsulateFields" & vbCrLf & Err.Source, Err.Description
    Else
        Err.Raise Err.Number, "Connect.ShowEncapsulateFields Linea " & Erl & vbCrLf & Err.Source, Err.Description
    End If
End Sub
Private Sub LoadFields()

    Dim f As frmEncapsulateField
    
    On Error GoTo Err_LoadFields
    Set f = New frmEncapsulateField
    f.LoadFields vbInstance, Me
    Exit Sub
Err_LoadFields:
    If Erl = 0 Then
        Err.Raise Err.Number, "Connect.LoadFields" & vbCrLf & Err.Source, Err.Description
    Else
        Err.Raise Err.Number, "Connect.LoadFields Linea " & Erl & vbCrLf & Err.Source, Err.Description
    End If
End Sub




Private Function AddItemToMenu(sCommandBarName As String, sCaption As String, Optional bStartNewSection As Boolean) As Office.CommandBarControl
    '
    ' Returns NOTHING object if command bar not found.
    '
    ' See if we can find the menu.
    ' These don't iterate, as the Office stuff is weird like that.
    Dim oMenu As Office.CommandBar
    On Error Resume Next
    Set oMenu = vbInstance.CommandBars(sCommandBarName)
    On Error GoTo 0
    If oMenu Is Nothing Then Exit Function
    '
    ' Add new item to the menu.
    
    Const msoControlButton As Long = 1&
    Set AddItemToMenu = oMenu.Controls.Add(msoControlButton)
    If bStartNewSection Then AddItemToMenu.BeginGroup = True
    '
    ' Set the caption.
    AddItemToMenu.Caption = sCaption
    
End Function

