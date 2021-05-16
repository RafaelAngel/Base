VERSION 5.00
Begin VB.Form frmInsert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert"
   ClientHeight    =   6525
   ClientLeft      =   4050
   ClientTop       =   1185
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   7140
   Begin VB.TextBox txtConsultaHecha 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   3360
      Width           =   6855
   End
   Begin VB.CommandButton cmdListo 
      Caption         =   "Listo"
      Height          =   375
      Left            =   3023
      TabIndex        =   1
      Top             =   6000
      Width           =   1095
   End
   Begin AsistenteMySQL.ConsultaInsert ConsultaInsert1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5530
   End
End
Attribute VB_Name = "frmInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdListo_Click()
    Unload Me
End Sub

Private Sub ConsultaInsert1_ClickGenerarConsulta()
    txtConsultaHecha.Text = ConsultaInsert1.fOptenerConsulta
End Sub
