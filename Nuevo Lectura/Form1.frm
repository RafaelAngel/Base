VERSION 5.00
Begin VB.Form FrmNuevoLectura 
   Caption         =   "Nuevo lectura"
   ClientHeight    =   6105
   ClientLeft      =   1110
   ClientTop       =   1710
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   10140
   Begin VB.OptionButton OptParar 
      Caption         =   "Parar"
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   4800
      Width           =   1455
   End
   Begin VB.OptionButton OptPausa 
      Caption         =   "Pausa"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   4680
      Width           =   1575
   End
   Begin VB.OptionButton OptContinuar 
      Caption         =   "Continuar"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   4680
      Width           =   1335
   End
   Begin Rnn.Libro Libro1 
      Height          =   4335
      Left            =   3000
      TabIndex        =   1
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7646
   End
   Begin Rnn.Hablar Hablar1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   8070
   End
End
Attribute VB_Name = "FrmNuevoLectura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private conteo As Integer



'Private Sub Command1_Click()
''Codigo para recrear el recorido del texto.
'Dim Sel As Variant
'    With Text1
'        .SetFocus
'        .SelStart = conteo
'        .SelLength = 5
'        Sel = .SelText
'        If conteo > Len(.Text) Then
'            conteo = Len(.Text)
'        End If
'        conteo = conteo + 5
'        '
'    End With
'
'Dim i As Integer
'For i = 0 To Len(Sel)
'DoEvents
'
'Next
'
'End Sub






Private Sub Hablar1_EveAudioStop()
    Hablar1.Leer Libro1.fContinuarLeyendo
End Sub

Private Sub Libro1_ClickContinuar(Habilitar As Boolean)
    
    'Hablar1.Leer Datos
End Sub

Private Sub Libro1_ClickIniciarLectura(Datos As Variant)
    Hablar1.Leer Datos
End Sub

Private Sub Libro1_EveDeshabilitarControles(Habilitados As Boolean)
    'No por el momento.
End Sub

Private Sub Libro1_EveSePuedeContinuar(Se_puede_continuar As Boolean)
    
    Se_puede_continuar = OptContinuar.Value
End Sub

Private Sub OptContinuar_Click()
    Hablar1.Continuar
End Sub

Private Sub OptParar_Click()
    Hablar1.Parar
End Sub

Private Sub OptPausa_Click()
    Hablar1.Pausa
End Sub
