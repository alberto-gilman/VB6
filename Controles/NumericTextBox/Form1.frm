VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin Project1.NumTextBox ntbExp 
      Height          =   400
      Left            =   2760
      TabIndex        =   1
      Top             =   600
      Width           =   2640
      _extentx        =   4657
      _extenty        =   714
      backcolor       =   0
      forecolor       =   49152
      font            =   "Form1.frx":0000
      precision       =   2
   End
   Begin Project1.NumTextBox NumTextBox 
      Height          =   400
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   2640
      _extentx        =   4657
      _extenty        =   714
      backcolor       =   0
      forecolor       =   49152
      font            =   "Form1.frx":0024
   End
   Begin Project1.NumTextBox NumTextBox1 
      Height          =   400
      Left            =   2760
      TabIndex        =   4
      Top             =   1200
      Width           =   2640
      _extentx        =   4657
      _extenty        =   714
      backcolor       =   0
      forecolor       =   49152
      font            =   "Form1.frx":0048
      precision       =   2
      scal            =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Precision 2, Scale 0"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Precision 2, Scale 2"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lbl 
      Caption         =   "Precision 10, Scale 2"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

