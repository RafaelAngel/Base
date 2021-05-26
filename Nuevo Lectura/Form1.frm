VERSION 5.00
Begin VB.Form FrmNuevoLectura 
   Caption         =   "Nuevo lectura"
   ClientHeight    =   5625
   ClientLeft      =   1110
   ClientTop       =   2010
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   9120
   Begin VB.CommandButton CMDEnd 
      Height          =   495
      Left            =   8400
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cerrar este formulario"
      Top             =   4920
      Width           =   495
   End
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
      Value           =   -1  'True
      Width           =   1335
   End
   Begin Rnn.Libro Libro1 
      Height          =   4695
      Left            =   3000
      TabIndex        =   1
      Top             =   0
      Width           =   6015
      _extentx        =   10610
      _extenty        =   7646
   End
   Begin Rnn.Hablar Hablar1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      _extentx        =   5106
      _extenty        =   8070
   End
   Begin VB.Label LbStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   5400
      Width           =   7815
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuAbrirLibro 
         Caption         =   "Abrir libro"
      End
      Begin VB.Menu mnu_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu MnuVer 
      Caption         =   "Ver"
      Begin VB.Menu MnuLibros 
         Caption         =   "Libros"
      End
   End
End
Attribute VB_Name = "FrmNuevoLectura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private conteo As Integer
'Private ClsLibro As New clsLecturaEnProgreso

Private Sub CMDEnd_Click()
    Unload Me
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Libro1.GuardarPropiedades
End Sub

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

Private Sub Libro1_EveSegundosLeyendo(Segundos_leyendo As Double)
'    Me.Caption = Segundos_leyendo
End Sub

Private Sub Libro1_EveSePuedeContinuar(Se_puede_continuar As Boolean)
    
    Se_puede_continuar = OptContinuar.Value
End Sub

Private Sub Libro1_TiempoDeLectura(Horas As Integer, Minutos As Integer, Segundos As Integer)
    LbStatus.Caption = "Horas (" & Horas & "), minutos(" & Minutos & "), segundos(" & Segundos & ")"
End Sub

Private Sub mnuAbrirLibro_Click()
    Libro1.sAbrirLectura
End Sub

Private Sub MnuLibros_Click()
    FrmLibrosLeidos.Show 'vbModal, Me
End Sub

Private Sub mnuSalir_Click()
    CMDEnd_Click
End Sub

Private Sub OptContinuar_Click()
    Hablar1.Continuar
    Libro1.prHabilitarAnexar = OptParar.Value
    Libro1.prHabilitarPegar = OptParar.Value
  
End Sub

Private Sub OptParar_Click()
    Hablar1.Parar
    Libro1.prHabilitarAnexar = OptParar.Value
    Libro1.prHabilitarPegar = OptParar.Value
  
End Sub

Private Sub OptPausa_Click()
    Hablar1.Pausa
    Libro1.prHabilitarAnexar = OptPausa.Value
    Libro1.prHabilitarPegar = OptParar.Value
    
End Sub
