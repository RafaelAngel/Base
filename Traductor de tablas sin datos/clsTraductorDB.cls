VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTraductorDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Autor: Rafael Angel Montero Fernández.
'Fecha de creacion: Jueves 11 de junio del 2020.
'Esta clase debe ser implementada con WithEvents cuando se cree una instancia de ella.


'Evento para el progressbarr
Public Event eveTrabajandoEnCampos(conteo As Integer, maximo_de_tareas As Integer)
Public Event eveTrabajandoEnTablas(conteo As Integer, maximo_de_tareas As Integer)



Private Type tyColumnasConDatos
    mRegistros() As Variant
End Type


Private Type tyTablasConDatos
    Nombre As Variant
    mColumnas() As Variant
    mColumnasConDatos() As tyColumnasConDatos  'Por cada matriz tyColumnasConDatos hay una matriz de datos.
End Type
Private vTablasConDatos As tyTablasConDatos

Private mFilas() As String
'Este evento se usa en ListandoTablas.
'Para pasar los datos al cliente.
Public Event eveListandoTablas(nombre_de_la_tabla As Variant)
'Evento para obteenr los campos se usa en ListandoCampos.
Public Event eveListandoCampos(nombre_del_campo As Variant)
'Permite obtener los registros de las tablas.
Private clsDatosDeLasTablas As New clsDatos
Attribute clsDatosDeLasTablas.VB_VarHelpID = -1
'Evento para agarrar los datos de una columna en un recorrido.
'El nombre de la columna es solo para validar que se esta en la columna correcta.
Public Event eveListandoDatos(datos_de_una_fila_de_la_columna As Variant, nombre_de_la_columna As Variant)

'Evento para crear la tabla.
Public Event eveCrearTabla(Lineas_de_creacion As Variant)

'Evento para los campos y los tipos de datos.
Public Event eveListandoTiposDeDatos(tipo_de_datos As Variant, valor_del_tipo_de_datos As Variant, nombre_del_campo As Variant, cantidad_de_campos As Integer, nombre_de_la_tabla As Variant, datos_de_este_campo As Variant)
Private vNombreDeLaBase As Variant

'Evento de insercion
Public Event eveInsert(clausula_insert As Variant)

Public Function fObtenerDatosDeUnCampo(nombre_del_campo As Variant)
'Funcion para obtener datos de una fila de una sola columna.
'Se necesita usar en combinacion con un procedimiento de avanzar.
fObtenerDatosDeUnCampo = clsDatosDeLasTablas.GetF(nombre_del_campo)
End Function


Public Sub irAdelante()
'Para ir hacia adelante obteniendo los campos.
clsDatosDeLasTablas.mSiguiente
End Sub
Public Function irListandoDatos(nombre_del_campo As Variant, Optional id_de_columna As Integer = 0)
    'Permite obtener registro a registro hasta el final de una sola columna.
    On Error GoTo n
    Dim mCampos() As String
    Dim i As Integer
    Dim vRegistrosDeUnCampo As Variant
    clsDatosDeLasTablas.mPrimero
    ReDim mFilas(clsDatosDeLasTablas.Cont)
    For i = 0 To clsDatosDeLasTablas.Cont
        clsDatosDeLasTablas.mSiguiente
        If i = clsDatosDeLasTablas.Cont - 1 Then
            vTablasConDatos.mColumnasConDatos(id_de_columna).mRegistros(i) = COMILLAS_DOBLES & clsDatosDeLasTablas.GetF(nombre_del_campo) & COMILLAS_DOBLES
            vRegistrosDeUnCampo = vRegistrosDeUnCampo & COMILLAS_DOBLES & clsDatosDeLasTablas.GetF(nombre_del_campo) & COMILLAS_DOBLES
        ElseIf i < clsDatosDeLasTablas.Cont Then
            vTablasConDatos.mColumnasConDatos(id_de_columna).mRegistros(i) = COMILLAS_DOBLES & clsDatosDeLasTablas.GetF(nombre_del_campo) & COMILLAS_DOBLES & ", "
            vRegistrosDeUnCampo = vRegistrosDeUnCampo & COMILLAS_DOBLES & clsDatosDeLasTablas.GetF(nombre_del_campo) & COMILLAS_DOBLES & "|"
        End If
        
        RaiseEvent eveListandoDatos(clsDatosDeLasTablas.GetF(nombre_del_campo), nombre_del_campo)
        'Debug.Print clsDatosDeLasTablas.GetF(nombre_del_campo)
    Next
    'Set clsDatosDeLasTablas = Nothing
    
    irListandoDatos = vRegistrosDeUnCampo
    Exit Function
n:
    irListandoDatos = vRegistrosDeUnCampo
    Debug.Print "Problemas en irListandoDatos " & Err.Description
End Function

Public Sub irPrimero()
'En caso de usarse la obtencion de datos dentro de un bucle, siempre es bueno usar regresar al primero.
clsDatosDeLasTablas.mPrimero
End Sub

Private Sub crearScriptDeInsercion(nombre_de_la_tabla As Variant, nombre_de_todos_los_campos As Variant, todos_los_datos_de_todos_los_campos As Variant)
    'mCampos = Split(todos_los_datos_de_todos_los_campos, "|")
    
    RaiseEvent eveInsert("insert into " & nombre_de_la_tabla & "(" & nombre_de_todos_los_campos & ")values(")
End Sub

Public Sub ListandoTablas(direccion_de_la_base As Variant)

On Error Resume Next
   'Se usarán eventos para pasar los datos.
    'On Error GoTo error_handler
    'clsDatosDeLasTablas.mPrimero
    'On Error Resume Next
    Set clsDatosDeLasTablas = New clsDatos
    
   Dim vListaDeCampos As Variant
   Dim vListaDeDatos As Variant
   
     ' -- Conexión
    Dim cnn As New ADODB.Connection
   
     ' -- Variables ADOX
    Dim oCatalog  As New ADOX.Catalog
    Dim Tablas   As ADOX.Tables
    Dim Tabla     As ADOX.Table
    Dim campo As ADOX.Column
    Dim fd As Field
   
      ' Abrir la base de datos
     cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & direccion_de_la_base
     cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & direccion_de_la_base
        
     ' Asignar la conexión activa al objeto catalog
     Set oCatalog.ActiveConnection = cnn
       
     ' Colección con las tablas
    Set Tablas = oCatalog.Tables
     'Recorre las tablas y agregar al control Listbox
     Dim formulario() As New frmCrearScript
     
     'For Each Tabla In Tablas
         'clsDatosDeLasTablas.LoadDBs Tabla.Name, direccion_de_la_base
         'Set campo = Tablas.Item(id).Columns.Item(0).Type
         'campo.Name = "Numero"
         'On Error GoTo eN
         Dim cid As Integer
         '__________________________________
         'Codigo para obtener el tipo de datos del campo y el nombre del campo.
         ReDim vTablasConDatos.mColumnas(Tablas.Count) As Variant
         For cid = 0 To Tablas.Count - 1
            'Set clsDatosDeLasTablas = New clsDatos
            On Error Resume Next
            clsDatosDeLasTablas.Cerrar
            clsDatosDeLasTablas.LoadDBs Tablas.Item(cid).Name, direccion_de_la_base
            DoEvents
            Dim fid As Integer
            
            'Todas las enumeraciones de fechas se toman como DateTime.
            'Todos los que tengan int o integer se toman como int.
            'adVarNumeric se toma como int.
            'Todos los que tengan char son char simplemente.
            
            ReDim formulario(Tablas.Count)
            'f(cid).direccion = direccion_de_la_base
            formulario(cid).Caption = Tablas.Item(cid).Name
            For fid = 0 To Tablas.Item(cid).Columns.Count - 1
                ReDim vTablasConDatos.mColumnasConDatos(Tablas.Item(cid).Columns.Count) As tyColumnasConDatos
                
                DoEvents
                Dim vCamposParaRegistros As Variant
                Dim vNombre_del_campo As Variant
                Dim vValorDelTipoDeDatos As DataTypeEnum
                Dim vNombreDelTipoDeDatos As Variant
                Dim vMax As Integer
                vMax = Tablas.Item(cid).Columns.Count '- 1
                vNombre_del_campo = Tablas.Item(cid).Columns.Item(fid).Name
                vValorDelTipoDeDatos = Tablas.Item(cid).Columns.Item(fid).Type
                Select Case vValorDelTipoDeDatos
                Case DataTypeEnum.adBigInt
                    vNombreDelTipoDeDatos = " int "
                Case DataTypeEnum.adBinary
                    vNombreDelTipoDeDatos = " Boolean "
                Case DataTypeEnum.adBoolean 'adBoolean valor 11
                    vNombreDelTipoDeDatos = " Boolean "
                Case DataTypeEnum.adCurrency 'El valor moneda se toma como un doble.
                    vNombreDelTipoDeDatos = " Double "
                Case DataTypeEnum.adChar
                    vNombreDelTipoDeDatos = "char(50)"
                Case DataTypeEnum.adDate
                    vNombreDelTipoDeDatos = " DateTime default " & COMILLAS_DOBLES & "1978/1/25 00:00:00" & COMILLAS_DOBLES
                Case DataTypeEnum.adDBDate
                    vNombreDelTipoDeDatos = " DateTime default " & COMILLAS_DOBLES & "1978/1/25 00:00:00" & COMILLAS_DOBLES
                Case DataTypeEnum.adDBTime
                    vNombreDelTipoDeDatos = " DateTime default " & COMILLAS_DOBLES & "1978/1/25 00:00:00" & COMILLAS_DOBLES
                Case DataTypeEnum.adDBTimeStamp
                    vNombreDelTipoDeDatos = " DateTime default " & COMILLAS_DOBLES & "1978/1/25 00:00:00" & COMILLAS_DOBLES
                Case DataTypeEnum.adDecimal
                    vNombreDelTipoDeDatos = "adDecimal"
                Case DataTypeEnum.adDouble
                    vNombreDelTipoDeDatos = " Double "
                Case DataTypeEnum.adLongVarBinary
                    vNombreDelTipoDeDatos = " LongBlob not null " 'Campo OLE, adLongVarBinary valor 205
                Case DataTypeEnum.adLongVarChar
                    vNombreDelTipoDeDatos = " LongText default " & COMILLAS_DOBLES & COMILLAS_DOBLES '"Memo"
                Case DataTypeEnum.adNumeric
                    vNombreDelTipoDeDatos = "adNumeric"
                Case DataTypeEnum.adSingle
                    vNombreDelTipoDeDatos = " int "
                Case DataTypeEnum.adVarBinary 'Campo de datos binarios (ceros y unos), puede ser guardado como un texto largo.
                'Viene en las tablas automaticas de Access
                    vNombreDelTipoDeDatos = " VarChar(500) "
                Case DataTypeEnum.adVarChar
                    vNombreDelTipoDeDatos = " char(50)"
                Case DataTypeEnum.adVarNumeric
                    vNombreDelTipoDeDatos = " int "
                Case DataTypeEnum.adBSTR
                    vNombreDelTipoDeDatos = "adBSTR"
                Case DataTypeEnum.adChapter
                    vNombreDelTipoDeDatos = " char(50) "
                Case DataTypeEnum.adVarWChar
                    vNombreDelTipoDeDatos = " Char(50) "
                Case DataTypeEnum.adArray
                    vNombreDelTipoDeDatos = "adArray"
                Case DataTypeEnum.adWChar 'texto pequeño adVarWChar valor 202
                    vNombreDelTipoDeDatos = " Char(char(50) "
                Case DataTypeEnum.adEmpty
                    vNombreDelTipoDeDatos = "adEmpty"
                Case DataTypeEnum.adFileTime
                    vNombreDelTipoDeDatos = "adFileTime"
                Case DataTypeEnum.adVariant
                    vNombreDelTipoDeDatos = "adVariant"
                Case DataTypeEnum.adUnsignedTinyInt
                    vNombreDelTipoDeDatos = " int "
                Case DataTypeEnum.adUnsignedSmallInt
                    vNombreDelTipoDeDatos = " int "
                Case DataTypeEnum.adUnsignedInt
                    vNombreDelTipoDeDatos = " int "
                Case DataTypeEnum.adUnsignedBigInt
                    vNombreDelTipoDeDatos = " int "
                Case DataTypeEnum.adTinyInt
                    vNombreDelTipoDeDatos = "adTinyInt"
                Case DataTypeEnum.adSmallInt 'numerico int; adSmallInt valor 2
                    vNombreDelTipoDeDatos = " int "
                Case DataTypeEnum.adIDispatch
                    vNombreDelTipoDeDatos = "adIDispatch"
                Case DataTypeEnum.adGUID
                    vNombreDelTipoDeDatos = " int not null auto_increment, primary key( " & Tablas.Item(cid).Columns.Item(fid).Name & ") " 'primarykey adGUID valor 72
                Case DataTypeEnum.adInteger 'Valor 3
                    If UCase(Tablas.Item(cid).Columns.Item(fid).Name) = UCase("id") Then
                      vNombreDelTipoDeDatos = " int not null auto_increment, primary key( " & Tablas.Item(cid).Columns.Item(fid).Name & ") " 'primarykey adGUID valor 72
                    Else
                        vNombreDelTipoDeDatos = " int "
                    End If
                Case DataTypeEnum.adLongVarWChar
                    vNombreDelTipoDeDatos = " LongText default " & COMILLAS_DOBLES & COMILLAS_DOBLES ' "Memo"
                Case Else
                    vNombreDelTipoDeDatos = "Alguno de los que no fueron anotados."
                End Select
               formulario(cid).add Tablas.Item(cid).Columns.Item(fid).Name
                    If fid = Tablas.Item(cid).Columns.Count - 1 Then
                        
                        vCamposParaRegistros = vCamposParaRegistros & " " & Tablas.Item(cid).Columns.Item(fid).Name & " "
                        vListaDeCampos = vListaDeCampos & " " & Tablas.Item(cid).Columns.Item(fid).Name & vNombreDelTipoDeDatos 'Tablas.Item(cid).Columns.Item(fid).Name
                        irListandoDatos Tablas.Item(cid).Columns.Item(fid).Name, fid
                    Else
                        irListandoDatos Tablas.Item(cid).Columns.Item(fid).Name, fid
                        vCamposParaRegistros = vCamposParaRegistros & Tablas.Item(cid).Columns.Item(fid).Name & ", "
                        vListaDeCampos = vListaDeCampos & " " & Tablas.Item(cid).Columns.Item(fid).Name & vNombreDelTipoDeDatos & ", "  'Tablas.Item(cid).Columns.Item(fid).Name & ","
                    End If
                'MsgBox vMax
                ReDim vTablasConDatos.mColumnasConDatos(fid).mRegistros(clsDatosDeLasTablas.Cont) As Variant
                vListaDeDatos = vListaDeDatos & "°" & Me.irListandoDatos(Tablas.Item(cid).Columns.Item(fid).Name, fid)
                RaiseEvent eveListandoTiposDeDatos(vNombreDelTipoDeDatos, vValorDelTipoDeDatos, Tablas.Item(cid).Columns.Item(fid).Name, vMax, Tablas.Item(cid).Name, irListandoDatos(Tablas.Item(cid).Columns.Item(fid).Name))
                RaiseEvent eveTrabajandoEnCampos(fid, Tablas.Item(cid).Columns.Count - 1)
            Next
            direccion = direccion_de_la_base
            formulario(cid).Show
            formulario(cid).SetFocus
            formulario(cid).activar = True
            
            'Debug.Print direccion_de_la_base
            'clsDatosDeLasTablas.mSiguiente
            Select Case Tablas.Item(cid).Name
            Case "MSysRelationships", "MSysQueries", "MSysACEs", "MSysAccessObjects", "MSysObjects", "MSysQueries"
                Debug.Print "Tablas del sistema access"
            Case Else
                RaiseEvent eveCrearTabla("Drop table if exists " & Tablas.Item(cid).Name & ";" & rtc & "Create table " & Tablas.Item(cid).Name & "(" & vListaDeCampos & ");" & rtc & rtc)
            End Select
            
            RaiseEvent eveListandoTablas(Tablas.Item(cid).Name)
'            Dim xid As Integer
'            Dim xaid As Integer
'            Dim datos_a_insertar As Variant
'            Dim idCelda As Integer
'            For xaid = 0 To Tablas.Item(cid).Columns.Count - 1
'            DoEvents
'            For idCelda = 0 To Len(vTablasConDatos.mColumnasConDatos(xaid)) - 1
'            DoEvents
'                For xid = 0 To Tablas.Item(cid).Columns.Count - 1
'                DoEvents
'                'MsgBox "len de columna " & Len(vTablasConDatos.mColumnasConDatos(xaid) & " y " & idCelda)
'                    On Error Resume Next
'                    'datos_a_insertar = datos_a_insertar & " " & vTablasConDatos.mColumnasConDatos(xid).mRegistros(idCelda)
'
'                Next
'                'RaiseEvent eveInsert("insert into " & Tablas.Item(cid).Name & "(" & vCamposParaRegistros & ")values(" & datos_a_insertar & "); " & rtc & rtc)
'                datos_a_insertar = ""
'                vCamposParaRegistros = ""
'            Next
'            Next
            RaiseEvent eveTrabajandoEnTablas(cid, Tablas.Count - 1)
            vListaDeCampos = ""
            'RaiseEvent eveInsert(formulario(cid).txtInsert & rtc)
         Next

Dim ra As Integer
For ra = 0 To Tablas.Count - 1
RaiseEvent eveInsert(formulario(ra).rtxtInsert.Text & rtc)
Next
         '__________________________________
         'RaiseEvent eveListandoTablas(Tabla.Name)
     'Next
eN:
         'Debug.Print "Error al tratar de obtener los tipos de datos. " & rtc & Err.Description
   
 ' -- Error
'error_handler:
'MsgBox Err.Description
 On Error Resume Next
   
 ' --Cierra la conexión
    If cnn.State <> 0 Then
         cnn.Close
     End If
  
    ' -- DEscargar las referencias
     Set cnn = Nothing
      Set oCatalog = Nothing
      Set Tabla = Nothing
      Set Tablas = Nothing
   
 End Sub



Public Sub ListandoCampos(La_Tabla As Variant, direccion_de_la_base As Variant)
     'Permite obtener los campos de la tabla.
     On Error GoTo Err_Sub

    Const cadena As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= "

       ' Nuevo objeto catalog
     Dim Obj_catalog As ADOX.Catalog
     Set Obj_catalog = New ADOX.Catalog

    ' Abre la base de datos
    Obj_catalog.ActiveConnection = cadena & direccion_de_la_base

     ' Nuevo objeto Table para hacer referencia a la tabla _
       que contiene los campos para agregar al listox
   Dim Obj_Tabla As ADOX.Table
  Set Obj_Tabla = New ADOX.Table


     ' crea la referencia a la tabla pasandole el nombre
    Set Obj_Tabla = Obj_catalog.Tables(La_Tabla)

     Dim i As Integer
n1:
    ' Recorre los campos con la colección Columns

     For i = 0 To Obj_Tabla.Columns.Count - 1
        RaiseEvent eveListandoCampos(Obj_Tabla.Columns(i).Name)
     Next

Eliminar_Objetos:
On Error Resume Next
     ' Elimina las referencias a Adox
    Set Obj_catalog = Nothing
     Set Obj_Tabla = Nothing

     Exit Sub
 
Err_Sub:
      ' Elimina las referencias a Adox
      MsgBox Err.Description, vbCritical
      GoTo Eliminar_Objetos
 End Sub

Private Sub Class_Initialize()
    Set clsDatosDeLasTablas = New clsDatos
End Sub

Private Sub clsDatosDeLasTablas_Conectado(Estado As Boolean)
'
End Sub


