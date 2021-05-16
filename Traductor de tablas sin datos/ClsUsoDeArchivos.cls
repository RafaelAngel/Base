VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsUsoDeArchivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vDireccion_del_archivo As String
Public vNombreDelArchivo As String
Public Event EveLinea(Datos As Variant)

Public Property Get prDireccionParaElArchivo() As String
    prDireccionParaElArchivo = vDireccion_del_archivo
End Property

Public Property Let prDireccionParaElArchivo(Nuevo As String)
    vDireccion_del_archivo = Nuevo
End Property

Public Sub EscribirDatos(Datos As Variant)
    On Error Resume Next
    Dim vTexto As Variant
    vTexto = Datos
    'MsgBox (App.Path & "\Script\")
    Dim i As Integer
    'For i = 1 To 4
        Open vDireccion_del_archivo & "\Script\" & vNombreDelArchivo For Append As 1
        Print #1, vTexto
        Close #1
    'Next
End Sub

Public Sub LeerTexto(Nombre As String)
    On Error Resume Next
    Dim vTexto As Variant
    Open vDireccion_del_archivo & "\Script\" & Nombre For Input As 1
    'MsgBox (App.Path & "\Script\")
    Dim i As Integer
    Do Until EOF(1)
        
        Line Input #1, vTexto
        Debug.Print vTexto
        RaiseEvent EveLinea(vTexto)
        'vTexto = ""
        DoEvents
        Loop
        Close #1
        
    'Next
End Sub
