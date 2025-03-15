VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsMyVar"
Attribute VB_Ext_KEY = "Member0" ,"clsMyVar"
'@Folder "GilmanAddinVB6.Clases"
Option Explicit


Private Type TClsLine
    LineNumber As Long
    Code As String
    Comentario As String
    Vars As Collection
End Type
Private this As TClsLine

Public Property Get Comentario() As String
    Comentario = this.Comentario
End Property

Public Property Let Comentario(ByVal RHS As String)
    this.Comentario = RHS
End Property

Public Property Get CodeLocation() As Long
    CodeLocation = this.LineNumber
End Property

Public Property Let CodeLocation(ByVal RHS As Long)
    this.LineNumber = RHS
End Property

Public Property Get Code() As String
    Code = this.Code
End Property

Public Property Let Code(ByVal RHS As String)
    this.Code = RHS
End Property

Public Sub Add(ByVal objNewMember As clsMyVar, Optional sKey As String)
    If Len(sKey) = 0 Then
        this.Vars.Add objNewMember
    Else
        this.Vars.Add objNewMember, sKey
    End If

End Sub

Public Property Get Item(vntIndexKey As Variant) As clsMyVar
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = this.Vars(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = this.Vars.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    this.Vars.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = this.Vars.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set this.Vars = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set this.Vars = Nothing
End Sub



Public Function NewCode() As String
    Dim var As clsMyVar
    For Each var In this.Vars
        If Not var.Removed Then
            NewCode = NewCode & var.Nombre & " As " & var.Tipo & ","
        End If
    Next var
    If NewCode <> "" Then
        NewCode = "Public " & Left(NewCode, Len(NewCode) - 1)
    End If
End Function
