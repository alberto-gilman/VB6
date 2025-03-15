VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Forms"

'TODO:[-] The next declaration is valid in VB6 but not in the addin
'only gets the first var

'TODO:[+] The next declaration is valid in VB6 but not in the addin
'And since it is a sentence with continuation, if there are more statements afterwards, the Addin will not work correctly.
'
''Public R As Long, _
''       s As String, _
''       t As Variant
'TODO:[-] The next declaration is valid in VB6 but not in the addin
''Private Type elTipoInterno
''a As Long: End Type



Option Explicit

Public Codigo As Long 'campo código
Public Nombre As String
Public Variable As Variant
Public Objeto As Object

Public X As Double, Y As Double
'De momento no está permitida una dec
Private Type InternalType
    a As Long
    B As Double
End Type

Private this As InternalType

'TODO:[-] The next declaration is valid in VB6 but not in the addin
'only gets the first var

'TODO:[+] The next declaration is valid in VB6 but not in the addin
'And since it is a sentence with continuation, if there are more statements afterwards, the Addin will not work correctly.
'
''Public R As Long, _
''       s As String, _
''       t As Variant
'TODO:[-] The next declaration is valid in VB6 but not in the addin
''Private Type elTipoInterno
''a As Long: End Type

Public Sub unmetodo()
    'algún código
    'tiene que haber
End Sub


