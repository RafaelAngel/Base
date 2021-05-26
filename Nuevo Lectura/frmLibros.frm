VERSION 5.00
Begin VB.Form frmLibros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Libros"
   ClientHeight    =   3165
   ClientLeft      =   1110
   ClientTop       =   1755
   ClientWidth     =   5745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5745
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5745
      TabIndex        =   22
      Top             =   2565
      Width           =   5745
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   300
         Left            =   1213
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "A&ctualizar"
         Height          =   300
         Left            =   59
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Cerrar"
         Height          =   300
         Left            =   4675
         TabIndex        =   27
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Reno&var"
         Height          =   300
         Left            =   3521
         TabIndex        =   26
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Eliminar"
         Height          =   300
         Left            =   2367
         TabIndex        =   25
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edición"
         Height          =   300
         Left            =   1213
         TabIndex        =   24
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Ag&regar"
         Height          =   300
         Left            =   59
         TabIndex        =   23
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5745
      TabIndex        =   16
      Top             =   2865
      Width           =   5745
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frmLibros.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frmLibros.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frmLibros.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frmLibros.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   21
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TemaDelLibro"
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   15
      Top             =   2300
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PaginasLen"
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   13
      Top             =   1980
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Notas"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   1660
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NombreDelLibro"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   1340
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Id"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1020
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DireccionURL"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   700
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Autor"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   380
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "AñoDePublicasion"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   60
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "TemaDelLibro:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   2300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PaginasLen:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Notas:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "NombreDelLibro:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Id:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DireccionURL:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Autor:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "AñoDePublicasion:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmLibros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents PrimaryCLS As clsLibros
Attribute PrimaryCLS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub Form_Load()
  Set PrimaryCLS = New clsLibros

  Dim oText As TextBox
  'Enlaza los cuadros de texto con el proveedor de datos
  For Each oText In Me.txtFields
    oText.DataMember = "Primary"
    Set oText.DataSource = PrimaryCLS
  Next
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub PrimaryCLS_MoveComplete()
  'Esto mostrará la posición de registro actual para este Recordset
  lblStatus.Caption = "Record: " & CStr(PrimaryCLS.AbsolutePosition)
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  PrimaryCLS.AddNew
  lblStatus.Caption = "Agregar registro"
  mbAddNewFlag = True
  SetButtons False

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  PrimaryCLS.Delete
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Esto sólo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  PrimaryCLS.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr

  lblStatus.Caption = "Modificar registro"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  PrimaryCLS.Cancel
  SetButtons True
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  PrimaryCLS.Update
  SetButtons True
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  PrimaryCLS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  PrimaryCLS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  PrimaryCLS.MoveNext
  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  PrimaryCLS.MovePrevious
  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

