VERSION 5.00
Begin VB.Form frmValoresDelChar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Valores del char"
   ClientHeight    =   1695
   ClientLeft      =   840
   ClientTop       =   4800
   ClientWidth     =   7995
   Icon            =   "frmValoresDelChar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   6000
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtValorPorDefecto 
      Height          =   615
      Left            =   3240
      TabIndex        =   2
      Text            =   "0"
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txtTamaño 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Text            =   "50"
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lbValorPorDefecto 
      Caption         =   "Valor por defecto"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      ToolTipText     =   "Datos por defecto"
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lbTamaño 
      Caption         =   "Tamaño del texto"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Por ejemplo: 50"
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "frmValoresDelChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Autor Rafael Angel Montero Fernández
'Correo: Sharkyc12@gmail.com
Private vValor As Variant

Private Sub cmdAceptar_Click()
    vValor = "(" & Me.txtTamaño.Text & ") default " & COMILLAS_DOBLES & Me.txtValorPorDefecto.Text & COMILLAS_DOBLES
    Unload Me
End Sub

Public Property Get Valor()
    Valor = vValor
End Property
