VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCrearMySQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Autor: Rafael Angel Montero Fernández.
'Fecha de creacion: Jueves 11 de junio del 2020.
'Esta clase debe ser implementada con WithEvents cuando se cree una instancia de ella.

'Evento para el progressbar
Public Event eveTrabajandoEnCampos(conteo As Integer, maximo_de_tareas As Integer)
Public Event eveTrabajandoEnTablas(conteo As Integer, maximo_de_tareas As Integer)


Private vConteo As Integer

'Evento para crear codigo MySQL designado a la creacion de tablas.
Public Event eveCrearTablaMySQL(linea_sql_de_la_tabla As Variant)

'Evento para la creacion de datos.
Public Event eveCrearInsertDatos(linea_sql_insert As Variant)

'Incluye el nombre y el tipo de datos. Listo para usarse en la creacion de tablas.
Private vCamposFormateados As Variant

Private vDireccionDeLaBase As Variant 'variable para guardar la direccion de la base para ser consultada en ocasiones.
Private mCampos() As Variant  'Matris dinamica para obtener los campos.
Private vCampos As String 'Variable para listar los campos en modo plano.
Private vTabla As Variant 'Variable para conservar la linea de la tabla.
Private vCamposParaDatos As Variant 'Variable para listar los campos para usarse en consultas insert.


Private Type tyTodasLosColumnas
'Typo de datos personalizado para guardar datos de las filas.
    vNombre_de_la_columna As Variant 'Es opcional.
    mDatosDeUnaColumna() As String
    
End Type


Private Type tyDatos
  vDatosDeTodosLosCampos As Variant 'Variable para listar los datos de todos los campos
  mDatosDeTodosLosCampos() As String ' Matriz para listar los datos de todos los campos pero ceparados por campos.
  mDatosDeTodasLasColumnas() As tyTodasLosColumnas 'Matriz de objeto para almacenar los datos de un solo campo(columna).
  mDatosDeUnaColumnaTemp() As String 'Datos de todas las filas de una columna en forma temporal.
End Type
Private vDatosDeLosCampos As tyDatos 'Variable del objeto (tipo) personalizado de datos.

Public WithEvents clsTraductor As clsTraductorDB 'Objeto class que permite obtener los nombres de las tablas, los nombres de los campos y los datos de cada campo.
Attribute clsTraductor.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set clsTraductor = New clsTraductorDB
End Sub

Public Sub crearMySQLBase(direccion_de_la_base As Variant)
    'direccion_de_la_base incluye el nombre de la base junto con su extencion de archivo.
    vDireccionDeLaBase = direccion_de_la_base
    clsTraductor.ListandoTablas direccion_de_la_base
End Sub

Private Sub clsTraductor_eveCrearTabla(Lineas_de_creacion As Variant)
    RaiseEvent eveCrearTablaMySQL(Lineas_de_creacion)
End Sub

Private Sub clsTraductor_eveInsert(clausula_insert As Variant)
    'MsgBox clausula_insert
    RaiseEvent eveCrearInsertDatos(clausula_insert)
End Sub

Private Sub clsTraductor_eveListandoCampos(nombre_del_campo As Variant)
vCampos = vCampos & nombre_del_campo & "."
'clsTraductor.irListandoDatos nombre_del_campo
'vDatosDeLosCampos.vDatosDeTodosLosCampos = vDatosDeLosCampos.vDatosDeTodosLosCampos & "°"
End Sub

Private Sub clsTraductor_eveListandoDatos(datos_de_una_fila_de_la_columna As Variant, nombre_de_la_columna As Variant)
    vDatosDeLosCampos.vDatosDeTodosLosCampos = vDatosDeLosCampos.vDatosDeTodosLosCampos & datos_de_una_fila_de_la_columna & "|"
End Sub

Private Sub clsTraductor_eveListandoTablas(nombre_de_la_tabla As Variant)
    'vTabla = "drop table if exists " & nombre_de_la_tabla & ";" & rtc & "create table " & nombre_de_la_tabla
    clsTraductor.ListandoCampos nombre_de_la_tabla, vDireccionDeLaBase
End Sub

Private Sub clsTraductor_eveListandoTiposDeDatos(tipo_de_datos As Variant, valor_del_tipo_de_datos As Variant, nombre_del_campo As Variant, cantidad_de_campos As Integer, nombre_de_la_tabla As Variant, datos_de_este_campo As Variant)
    On Error Resume Next
    Dim vTipoDeDatosFormateado As Variant
    
    Select Case tipo_de_datos
    Case "int"
        vTipoDeDatosFormateado = nombre_del_campo & " int "
    Case "Memo"
        vTipoDeDatosFormateado = nombre_del_campo & " LongText default " & COMILLAS_DOBLES & "0" & COMILLAS_DOBLES & " "
    Case "Boolean"
        vTipoDeDatosFormateado = nombre_del_campo & " boolean default false"
    Case "Double"
        vTipoDeDatosFormateado = nombre_del_campo & " double "
    Case "DateTime"
        vTipoDeDatosFormateado = nombre_del_campo & " DateTime default " & COMILLAS_DOBLES & "1978/1/25" & COMILLAS_DOBLES & " "
    Case "LongBlob not null"
        vTipoDeDatosFormateado = nombre_del_campo & " LongBlob not null "
    Case "int not null auto_numeric"
        vTipoDeDatosFormateado = nombre_del_campo & " int not null auto_numeric, primafry key( " & nombre_del_campo & ") "
    Case "Char"
        vTipoDeDatosFormateado = nombre_del_campo & " char(50) "
    End Select
    If vConteo >= cantidad_de_campos Then
        vCamposFormateados = vCamposFormateados & vTipoDeDatosFormateado
        vConteo = 0
    Else
        vCamposFormateados = vCamposFormateados & vTipoDeDatosFormateado & ","
    End If
    
    Debug.Print "Tabla " & nombre_de_la_tabla & " campo " & nombre_del_campo & ", " & tipo_de_datos & " valor " & valor_del_tipo_de_datos
    vConteo = vConteo + 1
End Sub

Private Sub clsTraductor_eveTrabajandoEnCampos(conteo As Integer, maximo_de_tareas As Integer)
    RaiseEvent eveTrabajandoEnCampos(conteo, maximo_de_tareas)
End Sub

Private Sub clsTraductor_eveTrabajandoEnTablas(conteo As Integer, maximo_de_tareas As Integer)
    RaiseEvent eveTrabajandoEnTablas(conteo, maximo_de_tareas)
End Sub
