VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MsComCtl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frmFormulario 
   Caption         =   "Formulario1"
   ClientHeight    =   9030
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMsg 
      Caption         =   "Saludos"
      Height          =   735
      Left            =   8640
      TabIndex        =   13
      Top             =   5520
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid 
      Height          =   855
      Left            =   6000
      TabIndex        =   12
      Top             =   6960
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1508
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid 
      Height          =   1215
      Left            =   6000
      TabIndex        =   11
      Top             =   5520
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2143
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Timer Timer 
      Left            =   8325
      Top             =   4980
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2355
      Left            =   240
      TabIndex        =   10
      ToolTipText     =   "Son las pestañas"
      Top             =   5985
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   4154
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Primera Pestaña"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Segunda pestaña"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   990
      Left            =   270
      TabIndex        =   9
      ToolTipText     =   "ToolTip del FlexGrid"
      Top             =   4890
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   1746
      _Version        =   393216
      Rows            =   3
      FixedRows       =   2
      AllowUserResizing=   1
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   8
      ToolTipText     =   "Es la barra de estado"
      Top             =   8700
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   582
      SimpleText      =   "TEXTO SIMPLE"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "PANEL IZDA"
            TextSave        =   "PANEL IZDA"
            Object.ToolTipText     =   "Es el panel de la Izda"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "OTRO PANEL"
            TextSave        =   "OTRO PANEL"
            Object.ToolTipText     =   "Es otro Panel"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            TextSave        =   "22/03/2025"
            Object.ToolTipText     =   "Es panel de fecha"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   3750
      Left            =   4710
      TabIndex        =   7
      ToolTipText     =   "Otro Control"
      Top             =   885
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6615
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cabecera 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cabecera2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbarx 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "barra de herramientas"
      Top             =   0
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   1164
      ButtonWidth     =   953
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ver"
            Key             =   "VER"
            Object.ToolTipText     =   "Ver archivo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "oir"
            Object.ToolTipText     =   "Oir archivo"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "abajo"
            Object.ToolTipText     =   "desplegar abajo"
            Style           =   5
            Object.Width           =   1e-4
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "primera opcion"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "segunda opcion"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Index           =   1
      Left            =   1320
      TabIndex        =   5
      Text            =   "Texto 1"
      ToolTipText     =   "Un texto para pruebas 2"
      Top             =   1335
      Width           =   3090
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Index           =   0
      Left            =   1290
      TabIndex        =   4
      Text            =   "Texto 0"
      ToolTipText     =   "Un texto para pruebas"
      Top             =   930
      Width           =   3090
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   630
      Left            =   5580
      TabIndex        =   3
      ToolTipText     =   "Cerrar la aplicación"
      Top             =   7950
      Width           =   5100
   End
   Begin VB.Frame Frame1 
      Caption         =   "Otro Control"
      Height          =   2745
      Left            =   330
      TabIndex        =   2
      Top             =   1770
      Width           =   4350
   End
   Begin VB.Label Label2 
      Caption         =   "Etiqueta 2"
      Height          =   345
      Left            =   360
      TabIndex        =   1
      Top             =   1395
      Width           =   2490
   End
   Begin VB.Label Label1 
      Caption         =   "Etiqueta 1"
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   915
      Width           =   900
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuGuardar 
         Caption         =   "Guardar"
      End
      Begin VB.Menu mnuSalvar 
         Caption         =   "Cargar"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "Edición"
   End
End
Attribute VB_Name = "frmFormulario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Private Const MFT_STRING = &H0
'Private Const MFT_RADIOCHECK = &H200&
'Private Const MIIM_TYPE = &H10
'Private Const MIIM_SUBMENU = &H4
'Private Type MENUITEMINFO
'  cbSize As Long
'  fMask As Long
'  fType As Long
'  fState As Long
'  wID As Long
'  hSubMenu As Long
'  hbmpChecked As Long
'  hbmpUnchecked As Long
'  dwItemData As Long
'  dwTypeData As String
'  cch As Long
'End Type
'Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpmii As MENUITEMINFO) As Long
'Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
'Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long


Private oIdioma As New clsGestionIdioma

Private Function F()

    On Error GoTo Err_F

    Exit Function
Err_F:
    Screen.MousePointer = vbDefault
    If Erl = 0 Then
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Function 'F' del Formulario Formulario" & vbCrLf & "[" & Err.Source & "]", vbCritical
    Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Function 'F' del Formulario Formulario en linea: " & Erl & vbCrLf & "[" & Err.Source & "]", vbCritical
    End If
End Function

Private Sub cmdMsg_Click()
10        On Error GoTo Err_cmdMsg_Click
20        MsgBox oIdioma.MessageStr("SALUDOS", "Gilman", "Bilbao")
          
30        Err.Raise 7, "hello world"

40        Exit Sub
Err_cmdMsg_Click:
50        Screen.MousePointer = vbDefault
60        If Erl = 0 Then
70            MsgBox oIdioma.MessageStr("ErrMsg", Err.Number, Err.Description, "Sub", "'cmdMsg_Click'", "Formulario", "frmFormulario", Err.Source), vbCritical
80        Else
90            MsgBox oIdioma.MessageStr("ErrErlMsg", Err.Number, Err.Description, "Sub", "'cmdMsg_Click'", "Formulario", "frmFormulario", Err.Source, Erl), vbCritical
100       End If
End Sub

Private Sub Command1_Click()
    Dim answer As VbMsgBoxResult
    On Error GoTo Err_Command1_Click
    'The message ASKFOREXIT is not defined in ENG Language, so it is showed in the default language
    answer = MsgBox(oIdioma.MessageStr("ASKFOREXIT"), vbQuestion + vbYesNo + vbDefaultButton2)
    If answer = vbYes Then
        Unload Me
    End If
    

    Exit Sub
Err_Command1_Click:
    Screen.MousePointer = vbDefault
    If Erl = 0 Then
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Sub 'Command1_Click' del Formulario Formulario" & vbCrLf & "[" & Err.Source & "]", vbCritical
    Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Sub 'Command1_Click' del Formulario Formulario en linea: " & Erl & vbCrLf & "[" & Err.Source & "]", vbCritical
    End If
End Sub

Private Sub Form_Load()
  
    Dim F As Long
    Dim C As Long
    
    On Error GoTo Err_Form_Load
    Set oIdioma = New DllCargarIdioma.clsGestionIdioma
    'Establish your default language
    'if a message is not defined in the selected language it is showed in the default language
    oIdioma.DefaultLanguage = "ESP"
    
    For F = 0 To Me.MSFlexGrid.FixedRows - 1
        For C = 0 To Me.MSFlexGrid.Cols - 1
            MSFlexGrid.TextMatrix(F, C) = "Fila " & F & " Columna " & C
        Next C
    Next F
    For F = Me.MSFlexGrid.FixedRows To Me.MSFlexGrid.Rows - 1
        For C = 0 To Me.MSFlexGrid.FixedCols - 1
            MSFlexGrid.TextMatrix(F, C) = "Fila " & F & " Columna " & C
        Next C
    Next F
    
    'La llamada a Guardar idioma solo es necesaría cuando se modifica el Diseño de un form
    'Está llamada se hará despues de la configuración del formulario
    'Si algún control tiene una configuración condicional, por ejemplo que cargue más o menos botones
    'en un toolbar según desde donde se llame, lo normal sería hacer la llamada antes.
    'Hay que tener en cuenta que borrarará la sección de mensajes
'        oIdioma.SaveLanguage App.Path & "\LANGUAGES", App.ProductName, Me, "ESP"


    'Si algún control tiene una configuración condicional, por ejemplo que cargue más o menos botones
    'en un toolbar según desde donde se llame, esta llamada deberá hacerse antes de dicha carga
    'y en el procedimiento de carga tener en cuenta el multi idioma
    
    oIdioma.LoadLanguage App.Path & "\LANGUAGES", App.ProductName, Me, "ENG"

    Exit Sub
Err_Form_Load:
    Screen.MousePointer = vbDefault
    If Erl = 0 Then
        MsgBox oIdioma.MessageStr("ErrMsg", Err.Number, Err.Description, "Sub", "'Form_Load'", "Formulario", "frmFormulario", Err.Source), vbCritical
    Else
        MsgBox oIdioma.MessageStr("ErrErlMsg", Err.Number, Err.Description, "Sub", "'Form_Load'", "Formulario", "frmFormulario", Err.Source, Erl), vbCritical
    End If
End Sub


