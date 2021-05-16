VERSION 5.00
Begin VB.Form frmWhere 
   Caption         =   "Condiciones Where"
   ClientHeight    =   4815
   ClientLeft      =   1710
   ClientTop       =   1875
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   6180
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4800
      Top             =   2040
   End
   Begin VB.CheckBox chkActivarCeparadores 
      Caption         =   "Usar ceparadores logicos. Si va a usar varias condicoines entonces debe elegir un ceparador por cada condicion."
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Width           =   5295
   End
   Begin VB.OptionButton optOR 
      Caption         =   "OR"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      ToolTipText     =   "Ceparador OR"
      Top             =   1800
      Width           =   855
   End
   Begin VB.OptionButton optAnd 
      Caption         =   "AND"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   10
      ToolTipText     =   "Ceparador AND"
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3263
      TabIndex        =   9
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   1823
      TabIndex        =   8
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txtOtroCampo 
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtUnCampo 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.ListBox lstCondicionesWhere 
      Height          =   1230
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   5655
   End
   Begin VB.ComboBox cmbOperadoresLogicos 
      Height          =   315
      ItemData        =   "frmWhere.frx":0000
      Left            =   1920
      List            =   "frmWhere.frx":0016
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lbOtroCampor 
      Caption         =   "Otro campo o datos"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      ToolTipText     =   "Este otro campo puede ser el mismo campo"
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lbUnCampo 
      Caption         =   "Un campo"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lbOperadoresLogicos 
      Caption         =   "Operadores logicos"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmWhere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private vCondicionesWhere As Variant
Private vCeparadorLogico As Variant


Private Sub cmdAceptar_Click()
    fObtenerCondicionesWhere
    Me.Hide
End Sub

Private Sub cmdAdd_Click()
'Las condiciones Where pueden tener los dos campos iguales.
'Ademas, puede que el segundo campo no sea otro campo sino datos.
    If lstCondicionesWhere.ListCount >= 1 Then
        If vCeparadorLogico = "" Then
            vCeparadorLogico = " Or "
        End If
        lstCondicionesWhere.AddItem vCeparadorLogico & fCondicionWhere
    End If
    
    If lstCondicionesWhere.ListCount <= 0 Then
        lstCondicionesWhere.AddItem fCondicionWhere
    End If
End Sub


Private Function fCondicionWhere()
    Dim vCondicion As Variant
    vCondicion = " " & Me.txtUnCampo.Text & " " & Me.cmbOperadoresLogicos.Text & " " & Me.txtOtroCampo.Text & " "
    txtUnCampo.Text = ""
    cmbOperadoresLogicos.Text = ""
    txtOtroCampo.Text = ""
    fCondicionWhere = vCondicion
End Function

Friend Function fObtenerCondicionesWhere()
    Dim i As Integer
    
    For i = 0 To Me.lstCondicionesWhere.ListCount - 1
        lstCondicionesWhere.ListIndex = i
        If i < Me.lstCondicionesWhere.ListCount - 1 Then
            vCondicionesWhere = vCondicionesWhere & lstCondicionesWhere.Text
        ElseIf i = Me.lstCondicionesWhere.ListCount - 1 Then
            vCondicionesWhere = vCondicionesWhere & lstCondicionesWhere.Text
        End If
    Next
    Me.lstCondicionesWhere.Clear
End Function

Public Property Get prCondicionesWhere() As Variant
    prCondicionesWhere = vCondicionesWhere
End Property

Private Sub LuzVerde()
    If lstCondicionesWhere.ListCount = 1 Or lstCondicionesWhere.ListCount > 1 Then
        If optAnd.Value = True Or optOR.Value = True Then
            cmdAdd.Enabled = True
            Timer1.Enabled = False
            Me.chkActivarCeparadores.BackColor = &H8000000F
        Else
            Timer1.Enabled = True
            cmdAdd.Enabled = False
        End If
    End If
    
End Sub


Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Public Sub Clear()
    Me.lstCondicionesWhere.Clear
    vCondicionesWhere = ""
End Sub

Private Sub chkActivarCeparadores_Click()
    optAnd.Enabled = chkActivarCeparadores.Value
    optOR.Enabled = chkActivarCeparadores.Value
    Timer1.Enabled = False
End Sub

Private Sub Form_Load()
    vCondicionesWhere = ""
End Sub

Private Sub optAnd_Click()
    vCeparadorLogico = optAnd.Caption & " "
    cmdAdd.Enabled = True
End Sub

Private Sub optOR_Click()
    vCeparadorLogico = optOR.Caption & " "
    cmdAdd.Enabled = True
End Sub

Private Sub Timer1_Timer()
Static v As Boolean
    If v = False Then
        Me.chkActivarCeparadores.BackColor = &H8000000F
        v = True
    Else
        Me.chkActivarCeparadores.BackColor = &HFF00&
        v = False
    End If
End Sub


Private Sub txtOtroCampo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LuzVerde
End Sub

Private Sub txtUnCampo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LuzVerde
End Sub
