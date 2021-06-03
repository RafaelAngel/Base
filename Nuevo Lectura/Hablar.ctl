VERSION 5.00
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "XVoice.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Hablar 
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3510
   ScaleHeight     =   4410
   ScaleWidth      =   3510
   ToolboxBitmap   =   "Hablar.ctx":0000
   Begin ACTIVEVOICEPROJECTLibCtl.DirectSS DirectSS1 
      Height          =   975
      Left            =   0
      OleObjectBlob   =   "Hablar.ctx":0312
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin MSComctlLib.Slider SliderVolumen 
      Height          =   675
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1191
      _Version        =   393216
      Min             =   100
      Max             =   65535
      SelStart        =   100
      Value           =   100
   End
   Begin MSComctlLib.Slider SliderVelocidad 
      Height          =   555
      Left            =   0
      TabIndex        =   3
      Top             =   2520
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   979
      _Version        =   393216
      Min             =   136
      Max             =   225
      SelStart        =   136
      Value           =   136
   End
   Begin MSComctlLib.Slider SliderNivelDeVoz 
      Height          =   555
      Left            =   0
      TabIndex        =   5
      Top             =   3600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   979
      _Version        =   393216
      Max             =   82
   End
   Begin VB.Label LbNivelDeVoz 
      Caption         =   "Nivel de voz"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label LbVelocidad 
      Caption         =   "Velocidad"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label LbVolumen 
      Caption         =   "Volumen"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2655
   End
End
Attribute VB_Name = "Hablar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Event EveAudioStart()
Public Event EveAudioStop()
Public Event EveParadaDefinitiva()

Private vParadaDefinitiva As Boolean

Public Property Get PrParadaDefinitiva() As Boolean
    PrParadaDefinitiva = vParadaDefinitiva
End Property

Public Property Let PrParadaDefinitiva(Nuevo As Boolean)
    vParadaDefinitiva = Nuevo
End Property


Private Sub DirectSS1_AudioStart(ByVal hi As Long, ByVal lo As Long)
    RaiseEvent EveAudioStart
End Sub

Private Sub DirectSS1_AudioStop(ByVal hi As Long, ByVal lo As Long)
    If PrParadaDefinitiva = False Then
        RaiseEvent EveAudioStop
    Else
        RaiseEvent EveParadaDefinitiva
    End If
End Sub

Public Sub Continuar()
    DirectSS1.AudioResume
End Sub

Public Sub Parar()
    vParadaDefinitiva = True
    DirectSS1.AudioReset
End Sub

Public Sub Pausa()
    vParadaDefinitiva = False
    DirectSS1.AudioPause
End Sub

Private Sub SliderNivelDeVoz_Click()
    DirectSS1.Pitch = DirectSS1.MinPitch + SliderNivelDeVoz.Value
End Sub

Private Sub SliderVelocidad_Click()
    DirectSS1.Speed = SliderVelocidad.Value 'DirectSS1.MinSpeed + SliderVelocidad.Value
    LbVelocidad.Caption = "Velocidad=" & DirectSS1.Speed
End Sub

Private Sub SliderVolumen_Click()
    DirectSS1.VolumeLeft = SliderVolumen.Value
    DirectSS1.VolumeRight = SliderVolumen.Value
End Sub

Public Sub Leer(Texto As Variant)
'El texto debe ser menor a 8 mil letras; los espacios valen como letras.
    vParadaDefinitiva = False
    DirectSS1.Speak Texto
End Sub

