VERSION 5.00
Begin VB.UserControl ConsultaInsert 
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7215
   ScaleHeight     =   3495
   ScaleWidth      =   7215
   Begin VB.TextBox txtValues 
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Text            =   "0"
      Top             =   480
      Width           =   3015
   End
   Begin VB.ListBox LstValues 
      Height          =   1425
      ItemData        =   "ConsultaInsert.ctx":0000
      Left            =   3120
      List            =   "ConsultaInsert.ctx":0002
      TabIndex        =   7
      Top             =   840
      Width           =   3015
   End
   Begin VB.CommandButton cmdGenerarConsulta 
      Caption         =   "Generar consulta"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   2520
      Width           =   2775
   End
   Begin VB.ListBox LstCampos 
      Height          =   1425
      ItemData        =   "ConsultaInsert.ctx":0004
      Left            =   0
      List            =   "ConsultaInsert.ctx":0006
      TabIndex        =   3
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox txtNombreDeLosCampos 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton cmdAddCampo 
      Caption         =   "Add"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtTabla 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label LbValues 
      Caption         =   "Escriba el texto que va a insertar en la tabla."
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label lbNombreDeLosCampos 
      Caption         =   "Escriba el nombre de cada campo de la tabla"
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label lbTabla 
      Caption         =   "Nombre de la tabla"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2400
      Width           =   2055
   End
End
Attribute VB_Name = "ConsultaInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event ClickGenerarConsulta()
Private vListaDeCampos As Variant
Private vListaDeDatos As Variant

Public Function fOptenerConsulta()
    fOptenerConsulta = "Insert into " & Me.prTabla & " (" & prListaDeCampos & ")Values(" & prListaDeDatos & ");"
End Function

Private Sub cmdAddCampo_Click()
    LstCampos.AddItem txtNombreDeLosCampos.Text
    
    txtNombreDeLosCampos.Text = ""
    LstValues.AddItem txtValues.Text
    txtValues.Text = 0
End Sub

Private Sub CargarListaDeDatos()
    vListaDeDatos = ""
    Dim i As Integer
    For i = 0 To LstValues.ListCount - 1
        LstValues.ListIndex = i
        If vListaDeDatos = "" Then
            vListaDeDatos = "'" & LstValues.Text & "'"
        Else
            vListaDeDatos = vListaDeDatos & ", '" & LstValues.Text & "'"
        End If
    Next
End Sub

Private Sub CargarListaDeCampos()
    vListaDeCampos = ""
    Dim i As Integer
    For i = 0 To LstCampos.ListCount - 1
        LstCampos.ListIndex = i
        If vListaDeCampos = "" Then
            vListaDeCampos = LstCampos.Text
        Else
            vListaDeCampos = vListaDeCampos & ", " & LstCampos.Text
        End If
    Next
End Sub
Public Property Get prListaDeDatos() As Variant
'Carga las 2 listas al mismo tiempo para mantener la sincronia.
    CargarListaDeCampos
    CargarListaDeDatos
    prListaDeDatos = vListaDeDatos
End Property

Public Property Get prListaDeCampos() As Variant
'Carga las 2 listas al mismo tiempo para mantener la sincronia.
    CargarListaDeCampos
    CargarListaDeDatos
    prListaDeCampos = vListaDeCampos
End Property

Public Property Get prTabla() As Variant
    prTabla = txtTabla.Text
End Property

Private Sub cmdGenerarConsulta_Click()
    RaiseEvent ClickGenerarConsulta
End Sub
