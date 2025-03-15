VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@Folder "GilmanAddinVB6.Clases"
Option Explicit



Private Type TClsMyVar
    Linea As Long
    Nombre As String
    Tipo As String
    Removed As Boolean
End Type
Private this  As TClsMyVar

Public Property Get Removed() As Boolean
    Removed = this.Removed
End Property

Public Property Let Removed(ByVal RHS As Boolean)
    this.Removed = RHS
End Property

Public Property Get Linea() As Long
    Linea = this.Linea
End Property

Public Property Let Linea(ByVal RHS As Long)
    this.Linea = RHS
End Property

Public Property Get Nombre() As String
    Nombre = this.Nombre
End Property

Public Property Let Nombre(ByVal RHS As String)
    this.Nombre = RHS
End Property

Public Property Get Tipo() As String
    Tipo = this.Tipo
End Property

Public Property Let Tipo(ByVal RHS As String)
    this.Tipo = RHS
End Property
