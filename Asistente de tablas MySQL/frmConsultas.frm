VERSION 5.00
Begin VB.Form frmConsultaSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultas"
   ClientHeight    =   5925
   ClientLeft      =   915
   ClientTop       =   1545
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   6825
   Begin VB.CommandButton cmdListo 
      Caption         =   "Listo"
      Height          =   375
      Left            =   2865
      TabIndex        =   2
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txtConsultaHecha 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   2640
      Width           =   6495
   End
   Begin AsistenteMySQL.ConsultaSelect Consulta1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4471
   End
End
Attribute VB_Name = "frmConsultaSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdListo_Click()
    Unload Me
End Sub

Private Sub Consulta1_ClickGenerarConsulta()
    txtConsultaHecha.Text = Consulta1.fOptenerConsulta()
End Sub
