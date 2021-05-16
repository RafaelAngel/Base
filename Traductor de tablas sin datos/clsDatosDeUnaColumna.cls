VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatosDeUnaColumna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Autor: Rafael Angel Montero Fernández.
'Fecha de creacion: Jueves 11 de junio del 2020.
'Clase para usarse en la recopilacion de datos en el class clsCrearMySQL.
'Su instancia sera a modo de matriz dinamica.


Private vDatos As Variant

Public Property Let Datos(nuevos_datos As Variant)
    vDatos = nuevos_datos
End Property

Public Property Get Datos() As Variant
   Datos = vDatos
End Property
