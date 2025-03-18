VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGestionIdioma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MFT_STRING = &H0
Private Const MFT_RADIOCHECK = &H200&
Private Const MIIM_TYPE = &H10
Private Const MIIM_SUBMENU = &H4
Private Type MENUITEMINFO
  cbSize As Long
  fMask As Long
  fType As Long
  fState As Long
  wID As Long
  hSubMenu As Long
  hbmpChecked As Long
  hbmpUnchecked As Long
  dwItemData As Long
  dwTypeData As String
  cch As Long
End Type
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpmii As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long


Private XMLDoc As MSXML2.DOMDocument60
Private RootNode As String
Private mDefaultLanguage As String
Private mDir As String
Private mProductName As String


Private Sub Class_Initialize()
    'I usually works with Spanish Language, so by default set the default Language to it
    mDefaultLanguage = "ESP"
    Set XMLDoc = New MSXML2.DOMDocument60
    
End Sub



Public Sub LoadLanguage(ByVal dir As String, ByVal ProductName As String, ByVal of As Object, ByVal Language As String)
    Dim objNode As MSXML2.IXMLDOMNode
    Dim nodoHijo As MSXML2.IXMLDOMNode
    Dim nodoNieto As MSXML2.IXMLDOMNode
    Dim oControl As Control
    Dim oAtt As MSXML2.IXMLDOMAttribute
    Dim oAtt2 As MSXML2.IXMLDOMAttribute
    Dim hMenu As Long


    
    RootNode = of.Name
    mDir = dir
    mProductName = ProductName
    XMLDoc.Load dir & "\" & mProductName & "." & RootNode & "." & Language & ".xml"
    
    
    Set objNode = XMLDoc.selectSingleNode("/" & RootNode)
    For Each oAtt In objNode.Attributes
        CallByName of, oAtt.nodeName, VbLet, oAtt.Text
    Next oAtt
    
    For Each nodoHijo In objNode.childNodes
        If nodoHijo.nodeName = "Messages" Then
            'Los mensajes no se procesan
        ElseIf nodoHijo.nodeName = "Menu" Then
    
            hMenu = GetMenu(of.hwnd)
            SetMenuInfo hMenu, nodoHijo
        Else
            If nodoHijo.Attributes.getNamedItem("Index").Text = "" Then
                Set oControl = of.Controls(nodoHijo.nodeName)
            Else
                Set oControl = of.Controls(nodoHijo.nodeName)(nodoHijo.Attributes.getNamedItem("Index").Text)
            End If
            For Each oAtt In nodoHijo.Attributes
                If oAtt.nodeName <> "Index" Then
                    CallByName oControl, oAtt.nodeName, VbLet, oAtt.Text
                End If
            Next oAtt
            Select Case TypeName(oControl)
                Case "SSDBGrid"
                    For Each nodoNieto In nodoHijo.childNodes
                        For Each oAtt2 In nodoNieto.Attributes
                            CallByName oControl.Columns(nodoNieto.nodeName), oAtt2.nodeName, VbLet, oAtt2.Text
                        Next oAtt2
                    Next nodoNieto
                Case "Toolbar"
                    Dim BTN As Object
                    Dim nodoBisNieto As MSXML2.IXMLDOMNode
                   
                    For Each nodoNieto In nodoHijo.childNodes
                        Set BTN = oControl.Buttons(CInt(Mid(nodoNieto.nodeName, 4)))
                        For Each oAtt2 In nodoNieto.Attributes
                            CallByName BTN, oAtt2.nodeName, VbLet, oAtt2.Text
                        Next oAtt2
                        If BTN.Style = 5 Then
                            Dim MBTN As Object
                            
                            For Each nodoBisNieto In nodoNieto.childNodes
                                Set MBTN = BTN.ButtonMenus(CInt(Mid(nodoBisNieto.nodeName, 4)))
                                    
                                    
                                For Each oAtt2 In nodoBisNieto.Attributes
                                    CallByName MBTN, oAtt2.nodeName, VbLet, oAtt2.Text
                                Next oAtt2
                            Next nodoBisNieto
                        End If
                    Next nodoNieto
                Case "ListView"
                    Dim CH As Object
                    For Each nodoNieto In nodoHijo.childNodes
                        Set CH = oControl.ColumnHeaders(CInt(Mid(nodoNieto.nodeName, 4)))
                        For Each oAtt2 In nodoNieto.Attributes
                            CallByName CH, oAtt2.nodeName, VbLet, oAtt2.Text
                        Next oAtt2
                    Next nodoNieto
                Case "SSTab"
                    For Each nodoNieto In nodoHijo.childNodes
                        oControl.TabCaption(CInt(Mid(nodoNieto.nodeName, 4))) = nodoNieto.Attributes(0).Text
                    Next nodoNieto
                Case "StatusBar"
                    For Each nodoNieto In nodoHijo.childNodes
                        Set CH = oControl.Panels(CInt(Mid(nodoNieto.nodeName, 4)))
                        For Each oAtt2 In nodoNieto.Attributes
                            CallByName CH, oAtt2.nodeName, VbLet, oAtt2.Text
                        Next oAtt2
                    Next nodoNieto
                Case "MSFlexGrid"
                    For Each nodoNieto In nodoHijo.childNodes
                        If nodoNieto.nodeName = "TextArray" Then
                            For Each oAtt2 In nodoNieto.Attributes
                                oControl.TextArray(Mid(oAtt2.Name, 4)) = oAtt2.Text
                            Next oAtt2
                        End If
                    Next nodoNieto
            End Select
        End If
    Next nodoHijo
    
    
End Sub

'
Public Sub SaveLanguage(ByVal dir As String, ByVal ProductName As String, ByVal of As Object, ByVal Language As String)
Dim objNode As MSXML2.IXMLDOMNode
Dim nodoHijo As MSXML2.IXMLDOMNode
Dim oAtt As MSXML2.IXMLDOMAttribute
Dim oControl As Control
Dim hMenu As Long
Dim i As Long

    Set objNode = XMLDoc.createNode(1, of.Name, XMLDoc.namespaceURI)
    Set oAtt = XMLDoc.createNode(2, "Caption", XMLDoc.namespaceURI)
    oAtt.Value = of.Caption
    objNode.Attributes.setNamedItem oAtt
    
    For Each oControl In of.Controls
        If HaveToolTip(oControl) Then
            Set nodoHijo = Nothing
            Set nodoHijo = XMLDoc.createNode(1, oControl.Name, XMLDoc.namespaceURI)
            Set oAtt = XMLDoc.createNode(2, "ToolTipText", XMLDoc.namespaceURI)
            oAtt.Value = oControl.ToolTipText
            nodoHijo.Attributes.setNamedItem oAtt
            Set oAtt = XMLDoc.createNode(2, "Index", XMLDoc.namespaceURI)
            On Error Resume Next
            oAtt.Value = oControl.Index
            On Error GoTo 0
            nodoHijo.Attributes.setNamedItem oAtt
            If HaveCaption(oControl) Then
                Set oAtt = XMLDoc.createNode(2, "Caption", XMLDoc.namespaceURI)
                oAtt.Value = oControl.Caption
                nodoHijo.Attributes.setNamedItem oAtt
            End If
            If HaveText(oControl) Then
                Set oAtt = XMLDoc.createNode(2, "Text", XMLDoc.namespaceURI)
                oAtt.Value = oControl.Text
                nodoHijo.Attributes.setNamedItem oAtt
            End If
            Select Case TypeName(oControl)
                Case "SSDBGrid"
                    Set oAtt = XMLDoc.createNode(2, "Caption", XMLDoc.namespaceURI)
                    oAtt.Value = oControl.Caption
                    nodoHijo.Attributes.setNamedItem oAtt
                    Dim oCol As Object
                    Dim ncol As MSXML2.IXMLDOMNode
                    For Each oCol In oControl.Columns
                        Set ncol = XMLDoc.createNode(1, oCol.Name, XMLDoc.namespaceURI)
                        Set oAtt = XMLDoc.createNode(2, "Caption", XMLDoc.namespaceURI)
                        oAtt.Value = oCol.Caption
                        ncol.Attributes.setNamedItem oAtt
                        nodoHijo.appendChild ncol
                    Next oCol
                Case "Toolbar"
                    Dim BTN As Object
                    For Each BTN In oControl.Buttons
                        Set ncol = XMLDoc.createNode(1, "BTN" & BTN.Index, XMLDoc.namespaceURI)
                        Set oAtt = XMLDoc.createNode(2, "Caption", XMLDoc.namespaceURI)
                        oAtt.Value = BTN.Caption
                        ncol.Attributes.setNamedItem oAtt
                        Set oAtt = XMLDoc.createNode(2, "ToolTipText", XMLDoc.namespaceURI)
                        oAtt.Value = BTN.ToolTipText
                        ncol.Attributes.setNamedItem oAtt
                        If BTN.Style = 5 Then
                            'dropdown button
                            Dim bm As Object
                            Dim nbm  As MSXML2.IXMLDOMNode
                            For Each bm In BTN.ButtonMenus
                                Set nbm = XMLDoc.createNode(1, "BTM" & bm.Index, XMLDoc.namespaceURI)
                                Set oAtt = XMLDoc.createNode(2, "Text", XMLDoc.namespaceURI)
                                oAtt.Value = bm.Text
                                nbm.Attributes.setNamedItem oAtt
                                ncol.appendChild nbm
                            Next bm
                        End If
                        nodoHijo.appendChild ncol
                    Next BTN
                Case "ListView"
                    Dim CH As Object
                    For Each CH In oControl.ColumnHeaders
                        Set ncol = XMLDoc.createNode(1, "CHE" & CH.Index, XMLDoc.namespaceURI)
                        Set oAtt = XMLDoc.createNode(2, "Text", XMLDoc.namespaceURI)
                        oAtt.Value = CH.Text
                        ncol.Attributes.setNamedItem oAtt
                        nodoHijo.appendChild ncol
                    Next CH
                Case "SSTab"
                    For i = 0 To oControl.Tabs - 1
                        Set ncol = XMLDoc.createNode(1, "TAB" & i, XMLDoc.namespaceURI)
                        Set oAtt = XMLDoc.createNode(2, "Caption", XMLDoc.namespaceURI)
                        oAtt.Value = oControl.TabCaption(i)
                        ncol.Attributes.setNamedItem oAtt
                        nodoHijo.appendChild ncol
                    Next i
                Case "StatusBar"
                    Dim Pan As Object
                    For Each Pan In oControl.Panels
                        Set ncol = XMLDoc.createNode(1, "PAN" & Pan.Index, XMLDoc.namespaceURI)
                        Set oAtt = XMLDoc.createNode(2, "Text", XMLDoc.namespaceURI)
                        oAtt.Value = Pan.Text
                        ncol.Attributes.setNamedItem oAtt
                        Set oAtt = XMLDoc.createNode(2, "ToolTipText", XMLDoc.namespaceURI)
                        oAtt.Value = Pan.ToolTipText
                        ncol.Attributes.setNamedItem oAtt
                        nodoHijo.appendChild ncol
                    Next Pan
                Case "MSFlexGrid"
                    Dim F As Long

                    Set ncol = XMLDoc.createNode(1, "TextArray", XMLDoc.namespaceURI)
                    For F = 0 To oControl.Rows * oControl.Cols - 1
                        If oControl.TextArray(F) <> "" Then
                            Set oAtt = XMLDoc.createNode(2, "TXA" & F, XMLDoc.namespaceURI)
                            oAtt.Value = oControl.TextArray(F)
                            ncol.Attributes.setNamedItem oAtt
                        End If
                        nodoHijo.appendChild ncol
                    Next F
                
            End Select
            objNode.appendChild nodoHijo
        End If
    Next oControl
    
    Set nodoHijo = XMLDoc.createNode(1, "Menu", XMLDoc.namespaceURI)
    
    hMenu = GetMenu(of.hwnd)

    MenuInfo hMenu, nodoHijo
    objNode.appendChild nodoHijo
    XMLDoc.appendChild objNode
    
    XMLDoc.save dir & "\" & ProductName & "." & of.Name & "." & Language & ".xml"
    
    
End Sub


Private Function HaveToolTip(ByVal o As Object) As Boolean
    Dim x As String

    On Error GoTo Err_HaveToolTip

    x = o.ToolTipText
    HaveToolTip = True
    
    On Error GoTo 0
    Exit Function

Err_HaveToolTip:
    HaveToolTip = False
    
End Function

Private Function HaveCaption(ByVal o As Object) As Boolean
    Dim x As String

    On Error GoTo Err_HaveCaption

    x = o.Caption
    HaveCaption = True
    
    Exit Function

Err_HaveCaption:
    HaveCaption = False
    
End Function


Private Function HaveText(ByVal o As Object) As Boolean
    On Error GoTo HaveText_Error
    Dim x As String


    x = o.Text
    HaveText = True
    
    Exit Function

HaveText_Error:


End Function



Private Function MenuInfo(ByVal vhMenu As Long, ByVal oNodo As MSXML2.IXMLDOMNode)
    Dim MII As MENUITEMINFO
    Dim Cont As Integer
    Dim i As Integer
    Dim nodoHijo As MSXML2.IXMLDOMNode
    Dim oAtt As MSXML2.IXMLDOMAttribute
    
    'initialize the structure
    MII.cbSize = Len(MII)
    MII.fMask = MIIM_TYPE Or MIIM_SUBMENU
    MII.fType = MFT_STRING
    Cont = GetMenuItemCount(vhMenu)
    
    For i = 0 To Cont - 1
        MII.cch = 256
        MII.dwTypeData = String(256, Chr(0))
        GetMenuItemInfo vhMenu, i, True, MII
        Set nodoHijo = oNodo.ownerDocument.createNode(1, "MenuItem", oNodo.namespaceURI)
        Set oAtt = oNodo.ownerDocument.createNode(2, "Text", oNodo.namespaceURI)
        oAtt.Value = MII.dwTypeData
        nodoHijo.Attributes.setNamedItem oAtt
        oNodo.appendChild nodoHijo
        'retrieve information about the menu item
        If MII.hSubMenu <> 0 Then
            MenuInfo MII.hSubMenu, nodoHijo
        End If
    Next i
    
End Function


Private Function SetMenuInfo(ByVal vhMenu As Long, ByVal oNodo As MSXML2.IXMLDOMNode)
    Dim MII As MENUITEMINFO
    Dim Cont As Integer
    Dim i As Integer
    Dim nodoHijo As MSXML2.IXMLDOMNode
    
    'initialize the structure
    MII.cbSize = Len(MII)
    MII.fMask = MIIM_TYPE Or MIIM_SUBMENU
    MII.fType = MFT_STRING
    
    Cont = oNodo.childNodes.length
    For i = 0 To Cont - 1
        Set nodoHijo = oNodo.childNodes(i)
        GetMenuItemInfo vhMenu, i, True, MII
        MII.cch = Len(nodoHijo.Attributes.getNamedItem("Text").Text) + 1
        MII.dwTypeData = nodoHijo.Attributes.getNamedItem("Text").Text & Chr(0)
        SetMenuItemInfo vhMenu, i, -1, MII
        If MII.hSubMenu <> 0 Then
            SetMenuInfo MII.hSubMenu, nodoHijo
        End If
    Next i
    
End Function


Public Function MessageStr(ByVal Message As String, ParamArray vars()) As String
    Dim i As Long
    Dim objNode As MSXML2.IXMLDOMNode
    Dim t() As String
    Dim ret As String
    Dim msg As String
    Set objNode = XMLDoc.selectSingleNode("/" & RootNode & "/Messages/" & Message)
    If objNode Is Nothing Then
        'if failed to load the message in current language try to load in the default language
        Dim oDoc As MSXML2.DOMDocument60
        Set oDoc = New MSXML2.DOMDocument60
        oDoc.Load mDir & "\" & mProductName & "." & RootNode & "." & mDefaultLanguage & ".xml"
        Set objNode = oDoc.selectSingleNode("/" & RootNode & "/Messages/" & Message)
        If objNode Is Nothing Then
            'If failed to load in default language too then
            MessageStr = "Failed to load message " & Message
            Exit Function
        End If
    End If
    On Error GoTo 0
    
    msg = objNode.Text
    msg = Replace(msg, Chr(0), Empty)
    msg = Replace(msg, Chr(1), Empty)
    msg = Replace(msg, "\\", Chr(1))
    msg = Replace(msg, "\t", vbTab)
    msg = Replace(msg, "\n", vbCrLf) 'simplified
    msg = Replace(msg, "%%", Chr(0))
    
    t = Split(msg, "%")
    
    ret = t(0)
    For i = 1 To UBound(t) Step 2
        ret = ret & vars(t(i)) & t(i + 1)
    Next i
    
    ret = Replace(ret, Chr(0), "%")
    ret = Replace(ret, Chr(1), "\")
    MessageStr = ret
    
End Function




Public Property Get DefaultLanguage() As String

    DefaultLanguage = mDefaultLanguage

End Property

Public Property Let DefaultLanguage(ByVal sNewValue As String)

    mDefaultLanguage = sNewValue

End Property
