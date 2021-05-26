VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmRelacion 
   Caption         =   "Relacion"
   ClientHeight    =   4605
   ClientLeft      =   1110
   ClientTop       =   1755
   ClientWidth     =   6615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   6615
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   5340
      TabIndex        =   0
      Top             =   3960
      Width           =   1080
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3840
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   6773
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Rows            =   4
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      GridColor       =   12632256
      GridColorFixed  =   16777215
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      FormatString    =   "IdRelacion|IdLibro|IdTema|IdAño|IdAutor"
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLineWidthBand=   1
      _Band(0).TextStyleBand=   0
   End
End
Attribute VB_Name = "frmRelacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MARGIN_SIZE = 60      ' en twips
' variables para enlace de datos
Private datPrimaryRS As ADODB.Recordset

Private Sub Form_Load()

    Dim sConnect As String
    Dim sSQL As String
    Dim dfwConn As ADODB.Connection
    Dim i As Integer
    Dim j As Integer
    Dim m_iMaxCol As Integer

    ' establecer cadenas
    sConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Password='';User ID=Admin;Data Source=C:\Documents and Settings\Administrador\Escritorio\Nuevo Lectura\DbLectura2021.mdb;Mode=Share Deny None;Extended Properties='';Jet OLEDB:System database='';Jet OLEDB:Registry Path='';Jet OLEDB:Database Password='';Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password='';Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False"
    sSQL = "select IdAño,IdAutor,IdLibro,IdRelacion,IdTema from Relacion"

    ' abrir conexión
    Set dfwConn = New Connection
    dfwConn.Open sConnect

    ' crear un conjunto de registros con la colección proporcionada
    Set datPrimaryRS = New Recordset
    datPrimaryRS.CursorLocation = adUseClient
    datPrimaryRS.Open sSQL, dfwConn, adOpenForwardOnly, adLockReadOnly

    Set MSHFlexGrid1.DataSource = datPrimaryRS

    With MSHFlexGrid1

        .Redraw = False
        ' colocar las columnas en el orden apropiado
        .ColData(3) = 0
        .ColData(2) = 1
        .ColData(4) = 2
        .ColData(0) = 3
        .ColData(1) = 4

        ' repetir para reordenar las columnas
        For i = 0 To .Cols - 1
            m_iMaxCol = i                   ' busca el valor más alto partiendo de esta columna
            For j = i To .Cols - 1
                If .ColData(j) > .ColData(m_iMaxCol) Then m_iMaxCol = j
            Next j
            .ColPosition(m_iMaxCol) = 0     ' mueve la columna con el valor máximo a la izquierda
        Next i

        ' establecer anchos de columna de cuadrícula
        .ColWidth(0) = -1
        .ColWidth(1) = -1
        .ColWidth(2) = -1
        .ColWidth(3) = -1
        .ColWidth(4) = -1

        ' establecer tipo de cuadrícula
        .AllowBigSelection = True
        .FillStyle = flexFillRepeat

        ' encabezado en negrita
        .Row = 0
        .Col = 0
        .RowSel = .FixedRows - 1
        .ColSel = .Cols - 1
        .CellFontBold = True

        ' atenuar las demás filas
        For i = .FixedRows + 1 To .Rows - 1 Step 2
            .Row = i
            .Col = .FixedCols
            .ColSel = .Cols() - .FixedCols - 1
            .CellBackColor = &HC0C0C0   ' gris claro
        Next i

        .AllowBigSelection = False
        .FillStyle = flexFillSingle
        .Redraw = True

    End With

End Sub

Private Sub Form_Resize()

    Dim sngButtonTop As Single
    Dim sngScaleWidth As Single
    Dim sngScaleHeight As Single

    On Error GoTo Form_Resize_Error
    With Me
        sngScaleWidth = .ScaleWidth
        sngScaleHeight = .ScaleHeight

        ' mueve el botón Cerrar a la esquina superior derecha
        With .cmdClose
                sngButtonTop = sngScaleHeight - (.Height + MARGIN_SIZE)
                .Move sngScaleWidth - (.Width + MARGIN_SIZE), sngButtonTop
        End With

        .MSHFlexGrid1.Move MARGIN_SIZE, _
            MARGIN_SIZE, _
            sngScaleWidth - (2 * MARGIN_SIZE), _
            sngButtonTop - (2 * MARGIN_SIZE)

    End With
    Exit Sub

Form_Resize_Error:
    ' evita errores en valores negativos
    Resume Next

End Sub
Private Sub cmdClose_Click()

    Unload Me

End Sub


