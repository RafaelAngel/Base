VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObtenerBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Obtiene la base con las tablas.

'Atentamente Rafael Angel Montero Fernández.
'Usted tiene toda la livertad de hacer lo mismo con este proyecto.
'Yo visualicé la posibilidad de facilitarme algunas
'tareas tediosas con bases de datos.

Private WithEvents K As clsDatos
Attribute K.VB_VarHelpID = -1
Private WithEvents Tr As TreeView
Attribute Tr.VB_VarHelpID = -1
Dim L As ListView
Public ITM As ListItem 'Variable para los SubItems

   '***************************************************************************
    '*  Name         : Ejemplo para obtener los campos de una tabla

    '*  Referencias:   Microsoft Ado Ext for dll and security ( Adox)
    '***************************************************************************

 Public Sub Obtener_Campos(La_Tabla As String, Lista As Object, Optional PathBD As Variant = "D:\Ralfy\Notas\Proyecto\Notas.mdb")

     On Error GoTo Err_Sub

    Const cadena As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= "

       ' Nuevo objeto catalog
     Dim Obj_catalog As ADOX.Catalog
     Set Obj_catalog = New ADOX.Catalog

    ' Abre la base de datos
    Obj_catalog.ActiveConnection = cadena & PathBD

     ' Nuevo objeto Table para hacer referencia a la tabla _
       que contiene los campos para agregar al listox
   Dim Obj_Tabla As ADOX.Table
  Set Obj_Tabla = New ADOX.Table


     ' crea la referencia a la tabla pasandole el nombre
    Set Obj_Tabla = Obj_catalog.Tables(La_Tabla)

     Dim i As Integer

    ' Recorre los campos con la colección Columns
     Lista.clear
     For i = 0 To Obj_Tabla.Columns.Count - 1
          Lista.AddItem Obj_Tabla.Columns(i).Name
      Next

Eliminar_Objetos:

     ' Elimina las referencias a Adox
    Set Obj_catalog = Nothing
     Set Obj_Tabla = Nothing

     Exit Sub
 
Err_Sub:
      ' Elimina las referencias a Adox
      MsgBox Err.Descriptio, vbCritical, "Obtener_Campos"
      GoTo Eliminar_Objetos
 End Sub


    
Public Sub AddFieldsToListView(La_Tabla As String, ListView_control As Object, Optional PathBD As Variant = "D:\Ralfy\Notas\Proyecto\Notas.mdb")

     On Error GoTo Err_Sub

    Const cadena As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= "

       ' Nuevo objeto catalog
     Dim Obj_catalog As ADOX.Catalog
     Set Obj_catalog = New ADOX.Catalog

    ' Abre la base de datos
    Obj_catalog.ActiveConnection = cadena & PathBD

     ' Nuevo objeto Table para hacer referencia a la tabla _
       que contiene los campos para agregar al listox
   Dim Obj_Tabla As ADOX.Table
  Set Obj_Tabla = New ADOX.Table


     ' crea la referencia a la tabla pasandole el nombre
    Set Obj_Tabla = Obj_catalog.Tables(La_Tabla)

     Dim i As Integer
     Static E As Long
     
     E = ListView_control.ColumnHeaders.Count
     Do
     On Error GoTo n1
     DoEvents
     ListView_control.ColumnHeaders.Remove (E)
     E = E - Val(1)
     Loop Until E = 0
n1:
    ' Recorre los campos con la colección Columns

     For i = 0 To Obj_Tabla.Columns.Count - 1
          ListView_control.ColumnHeaders.add , , Obj_Tabla.Columns(i).Name, 2500
      Next

Eliminar_Objetos:

     ' Elimina las referencias a Adox
    Set Obj_catalog = Nothing
     Set Obj_Tabla = Nothing

     Exit Sub
 
Err_Sub:
      ' Elimina las referencias a Adox
      MsgBox Err.Description, vbCritical, "AddFieldsToListView"
      GoTo Eliminar_Objetos
 End Sub


Public Sub AddFielsToList(PathBD As Variant, List As Object)
 On Error Resume Next
 Dim ret As Boolean
      Dim Tabla As String
      'Dim Path_Bd As String
     Tabla = InputBox(" Nombre de la tabla para obtener los campos ", " Nombre de tabla ")

     If Tabla = vbNullString Then
         Exit Sub
     End If

    ' Me.Caption = " Tabla selecionada: " & tabla

     ' Path de la base de datos
    'Path_Bd = "D:\Proyectos\Alternativo\Notas.mdb"

     ' Le pasa a la rutina el path de la base de datos y el nombre de la tabla
    Call Obtener_Campos(Tabla, List, PathBD)
End Sub


Public Sub Obtener_Tablas(ConnectionString As String, Lista As Object)
   
 On Error GoTo error_handler
   
   
     ' -- Conexión
    Dim cnn             As New ADODB.Connection
   
     ' -- Variables ADOX
     Dim oCatalog        As New ADOX.Catalog
    Dim Tablas          As ADOX.Tables
  Dim Tabla           As ADOX.Table
   
   
      ' -- Abrir la base de datos
     cnn.ConnectionString = ConnectionString
     cnn.Open ConnectionString
        
     ' -- Asignar la conexión activa al objeto catalog
     Set oCatalog.ActiveConnection = cnn
       
     ' -- Colección con las tablas
    Set Tablas = oCatalog.Tables
   Lista.clear
     ' -- Recorre las tablas y agregar al control Listbox
     For Each Tabla In Tablas
         Lista.AddItem Tabla.Name
     Next
   
 ' -- Error
error_handler:
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

Public Sub AddTablasToTree(ConnectionString As String, Tree As Object, _
Optional RaizIcon As Long = 1, Optional Icono As Long = 1)
   
 On Error GoTo error_handler
   
   
     ' -- Conexión
    Dim cnn             As New ADODB.Connection
   
     ' -- Variables ADOX
     Dim oCatalog        As New ADOX.Catalog
    Dim Tablas          As ADOX.Tables
  Dim Tabla           As ADOX.Table
   
   
      ' -- Abrir la base de datos
     cnn.ConnectionString = ConnectionString
     cnn.Open ConnectionString
        
     ' -- Asignar la conexión activa al objeto catalog
     Set oCatalog.ActiveConnection = cnn
       
     ' -- Colección con las tablas
    Set Tablas = oCatalog.Tables
   'Lista.Clear
     ' -- Recorre las tablas y agregar al control Listbox
     
     For Each Tabla In Tablas
     'Xz = Xz + Val(1)
         'Lista.AddItem Tabla.Name
         LoadTree "Base", Tabla.Name, Tabla.Name, Tree, RaizIcon, Icono
     Next
   
 ' -- Error
error_handler:
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

Public Function Base(Optional Nombre_de_la_base_de_datos As Variant = "D:\Ralfy\Notas\Proyecto\Notas.mdb") As Variant
Base = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                   Nombre_de_la_base_de_datos & ";"
End Function


Public Function NotasDB() 'Cambien esta direccion("D:\Ralfy\Notas\Proyecto\Notas.mdb") por la base de datos suya.
NotasDB = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
"D:\Ralfy\Notas\Proyecto\Notas.mdb" & ";"
End Function

Public Function NombreDB(Nombre_de_la_base_de_datos As Variant)
NombreDB = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
Nombre_de_la_base_de_datos & ";"
End Function


Public Sub LoadTree(Raiz As Variant, Clave As Variant, Datos As Variant, Tree As Object, Optional RaizIcon As Long = 1, Optional Icono As Long = 1)
Dim nodX As Node
On Error GoTo n
Set nodX = Tree.Nodes.add(, , "R", Raiz, RaizIcon)
n:
Set nodX = Tree.Nodes.add("R", tvwChild, Clave, Datos, Icono)
End Sub

Private Sub Class_Initialize()
Set K = New clsDatos
End Sub

Public Sub RegTree(Tree As Object, Optional Base As Variant = "D:\Proyectos\Alternativo\Notas/Notas.mdb")
Set Tr = Tree
K.Cerrar
K.LoadDBs Tr.Nodes.Item(1).Text, Base  ', "D:\Proyectos\Alternativo\Notas/Notas.mdb"

End Sub

Private Sub K_Conectado(Estado As Boolean)
'
End Sub


Private Sub Tr_BeforeLabelEdit(Cancel As Integer)
'
End Sub






