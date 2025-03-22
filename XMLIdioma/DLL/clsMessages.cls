VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type InternalType
    Language As String
    DefaultLanguage As String
    Application As String
End Type
Private this As InternalType


Private XMLDoc As MSXML2.DOMDocument60
Public Property Get Language() As String
    Language = this.Language
End Property
Public Property Let Language(ByVal RHC As String)
    this.Language = RHC
End Property


Private Sub Class_Initialize()
    mDefaultLanguage = "ESP"
    Set XMLDoc = New MSXML2.DOMDocument60

End Sub

Public Property Get DefaultLanguage() As String
    DefaultLanguage = this.DefaultLanguage
End Property
Public Property Let DefaultLanguage(ByVal RHC As String)
    this.DefaultLanguage = RHC
End Property
Public Property Get Application() As String
    Application = this.Application
End Property
Public Property Let Application(ByVal RHC As String)
    this.Application = RHC
End Property


