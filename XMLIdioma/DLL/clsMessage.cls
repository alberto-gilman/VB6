VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private XMLDoc As MSXML2.DOMDocument

Public Function MessageStr(ByVal Message As String, ParamArray vars()) As String
    Dim i As Long
    Dim objNode As MSXML2.IXMLDOMNode
    Set objNode = XMLDoc.selectSingleNode("/Messages/" & Message)
    Dim t() As String
    Dim ret As String
    Dim msg As String
    
    msg = objNode.Text
    msg = Replace(msg, Chr(0), Empty)
    msg = Replace(msg, "\\", Chr(1))
    msg = Replace(msg, "\t", vbTab)
    msg = Replace(msg, "\n", vbCrLf) 'simplified
    msg = Replace(msg, "%%", Chr(0))
    msg = Replace(msg, Chr(1), Empty)
    
    t = Split(msg, "%")
    
    ret = t(0)
    For i = 1 To UBound(t) Step 2
        ret = ret & vars(t(i)) & t(i + 1)
    Next i
    
    ret = Replace(ret, Chr(0), "%")
    ret = Replace(ret, Chr(1), "\")
    MessageStr = ret
    
End Function



Public Sub LoadMessages(ByVal dir As String, ByVal ProductName As String, ByVal Lenguage As String)

    XMLDoc.Load dir & "\" & ProductName & ".Messages." & Lenguage & ".xml"

End Sub

Private Sub Class_Initialize()

    Set XMLDoc = New MSXML2.DOMDocument
End Sub


