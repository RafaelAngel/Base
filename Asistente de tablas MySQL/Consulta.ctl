VERSION 5.00
Begin VB.UserControl ConsultaSelect 
   ClientHeight    =   2610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7725
   ScaleHeight     =   2610
   ScaleWidth      =   7725
   Begin VB.CommandButton cmdGenerarConsulta 
      Caption         =   "Generar consulta"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   1800
      Width           =   2775
   End
   Begin VB.CommandButton cmdWhere 
      Caption         =   "Agregar Where"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      ToolTipText     =   "Las condiciones Where permiten crear consultas más potentes"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtTabla 
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdAddCampo 
      Caption         =   "Add"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtNombreDeLosCampos 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin VB.ListBox lstCampos 
      Height          =   1425
      ItemData        =   "Consulta.ctx":0000
      Left            =   0
      List            =   "Consulta.ctx":0002
      TabIndex        =   0
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label lbTabla 
      Caption         =   "nombre de la tabla"
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lbNombreDeLosCampos 
      Caption         =   "Escriba el nombre de cada campo de la tabla"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "ConsultaSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event ClickGenerarConsulta()

Private vListaDeCampos As Variant
Private fWhere As New frmWhere

Private Sub cmdAddCampo_Click()
    lstCampos.AddItem txtNombreDeLosCampos.Text
    
    txtNombreDeLosCampos.Text = ""
End Sub


Public Property Get prListaDeCampos() As Variant
    CargarListaDeCampos
    prListaDeCampos = vListaDeCampos
End Property

Public Property Get prTabla() As Variant
    prTabla = txtTabla.Text
End Property

Private Sub cmdGenerarConsulta_Click()
    RaiseEvent ClickGenerarConsulta
    lstCampos.Clear
    vListaDeCampos = ""
End Sub

Private Sub cmdWhere_Click()
    fWhere.Show
    fWhere.Clear
End Sub


Public Function fOptenerConsulta()
    fOptenerConsulta = " Select " & Me.prListaDeCampos & " from " & Me.prTabla & fWhere.prCondicionesWhere & ";"
End Function

Private Sub CargarListaDeCampos()
    vListaDeCampos = ""
    Dim i As Integer
    For i = 0 To lstCampos.ListCount - 1
        lstCampos.ListIndex = i
        If vListaDeCampos = "" Then
            vListaDeCampos = lstCampos.Text
        Else
            vListaDeCampos = vListaDeCampos & ", " & lstCampos.Text
        End If
    Next
End Sub
