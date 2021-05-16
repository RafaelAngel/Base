VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.UserControl Coneccion 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10890
   ScaleHeight     =   6000
   ScaleWidth      =   10890
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9120
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFerst 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2460
      TabIndex        =   10
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4980
      TabIndex        =   9
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton cmdSiguiente 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4140
      TabIndex        =   8
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3300
      TabIndex        =   7
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   1335
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3480
      Width           =   7815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      ToolTipText     =   "Campo seleccionado"
      Top             =   2760
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "Tabla seleccionada"
      Top             =   2760
      Width           =   3735
   End
   Begin VB.CommandButton cmdCargar 
      Caption         =   "Cargar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      ToolTipText     =   "Carga la coneccion establecida mediante propiedades."
      Top             =   5040
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   2010
      ItemData        =   "Coneccion.ctx":0000
      Left            =   4440
      List            =   "Coneccion.ctx":0002
      TabIndex        =   2
      Top             =   600
      Width           =   3735
   End
   Begin VB.ListBox List1 
      Height          =   2010
      ItemData        =   "Coneccion.ctx":0004
      Left            =   360
      List            =   "Coneccion.ctx":0006
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.CommandButton cmdBuscarBase 
      Height          =   495
      Left            =   1785
      Picture         =   "Coneccion.ctx":0008
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Carga una base de datos que usted seleccione en tiempo de ejecucion."
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Lista de campos"
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Lista de tablas"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label lbDatosDelCampoSeleccionado 
      Caption         =   "Datos del campo seleccionado"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   3240
      Width           =   3615
   End
End
Attribute VB_Name = "Coneccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private vFieldsMax As Integer 'Variable para uso de la propiedad prFieldsMax
Private vSQLParaBusquedas As Variant 'Variable para uso de la propiedad prSQLParaBusquedas
'Private vAboudMe As Variant 'Variable para uso de la propiedad prAboudMe
Private vTodosLosFields As Variant 'Variable para uso de la propiedad prTodosLosFields

Dim Mx As New clsObtenerBase
Private vBaseNombre As String
'Dim Dx As New clsCommon
Private WithEvents K As clsDatos
Attribute K.VB_VarHelpID = -1
Private WithEvents Tr As TreeView
Attribute Tr.VB_VarHelpID = -1
Dim L As ListView
Public Event evListFieldsClick(ByRef Field_seleccionado As Variant)
Public Event evListTablasClick(ByRef Field_seleccionado As Variant)
Public Event evDblClickCargandoTodosLosFields(ByRef Nombre_del_field As Variant, ByVal Id_del_field_en_la_lista As Integer, ByRef Todos_los_Fields As Variant) 'Sucede al hacer doble Click en la lista de Fields.
Public Event evAllFieldsMatriz(ByRef Matriz_de_campos() As Variant, ByRef MaxFields As Integer) 'Devuelve una matriz con todos los campos de una tabla, sucede al hacer doble click en la lista de campos (Fields).



Public Function rtc()
'Crea un salto de linea.
rtc = Chr(13) + Chr(10)
End Function


Public Property Get prTodosLosFields() As Variant 'Contiene la lista de todos los campos de una tabla base en una base de datos.

On Error GoTo n

prTodosLosFields = vTodosLosFields
Exit Property
n:
MsgBox "Tengo problemas con prTodosLosFields"
End Property
Public Property Let prTodosLosFields(vNuevosDatos As Variant)  'Contiene la lista de todos los campos de una tabla base en una base de datos.
On Error GoTo n

vTodosLosFields = vNuevosDatos
PropertyChanged "prTodosLosFields"
Exit Property
n:
MsgBox "Tengo problemas con prTodosLosFields"
End Property


'Public ITM As ListItem 'Variable para los SubItems

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

List2_DblClick 'A ver que sucede

Eliminar_Objetos:

     ' Elimina las referencias a Adox
    Set Obj_catalog = Nothing
     Set Obj_Tabla = Nothing

     Exit Sub
 
Err_Sub:
      ' Elimina las referencias a Adox
      'MsgBox Err.Description, vbCritical
      GoTo Eliminar_Objetos
 End Sub


    
Public Sub AddFieldsToListView(La_Tabla As String, Lista As Object, Optional PathBD As Variant = "D:\Ralfy\Notas\Proyecto\Notas.mdb")

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
     
     E = Lista.ColumnHeaders.Count
     Do
     On Error GoTo n1
     DoEvents
     Lista.ColumnHeaders.Remove (E)
     E = E - Val(1)
     Loop Until E = 0
n1:
    ' Recorre los campos con la colección Columns

     For i = 0 To Obj_Tabla.Columns.Count - 1
          Lista.ColumnHeaders.add , , Obj_Tabla.Columns(i).Name, 2500
      Next

Eliminar_Objetos:

     ' Elimina las referencias a Adox
    Set Obj_catalog = Nothing
     Set Obj_Tabla = Nothing

     Exit Sub
 
Err_Sub:
      ' Elimina las referencias a Adox
      MsgBox Err.Description, vbCritical
      GoTo Eliminar_Objetos
 End Sub


Public Sub AddFielsToList(PathBD As Variant, List As Object)
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

Public Function Base(Optional Nombre_de_la_base_de_datos As String = "D:\Ralfy\Notas\Proyecto\Notas.mdb") As Variant

Base = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                Nombre_de_la_base_de_datos & ";"
                
End Function

Public Property Get prBaseNombre() As String

On Error GoTo n

prBaseNombre = vBaseNombre
Exit Property
n:
MsgBox "Tengo problemas con vBaseNombre"
End Property
Public Property Let prBaseNombre(vNuevosDatos As String)
On Error GoTo n
Dim vString As String
vString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                vNuevosDatos & ";"

vBaseNombre = vNuevosDatos
PropertyChanged "prBaseNombre"
Obtener_Tablas vString, List1
Exit Property
n:
MsgBox "Tengo problemas con vBaseNombre"
End Property




Public Function NotasDB()
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

Public Sub RegTree(Tree As Object)
Set Tr = Tree
K.Cerrar
K.LoadDBs Tr.Nodes.Item(1).Text, vBaseNombre  ', "D:\Proyectos\Alternativo\Notas/Notas.mdb"

End Sub

Private Sub CMDAnterior_Click()
K.mAnterior
End Sub

Private Sub cmdBuscarBase_Click()
'Dim V As String
' Dim vString As String
CommonDialog1.ShowOpen 'Dx.fSeleccionarArchivo
' V = CommonDialog1.FileName
' prBaseNombre = V
'vString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
'                V & ";"
'
'Obtener_Tablas vString, List1
 fConectandoBase CommonDialog1.FileName
 prBaseNombre = CommonDialog1.FileName
End Sub

Public Function fConectandoBase(ByRef Direccion_de_la_base_de_datos As Variant) 'Activa la coneccion con la base de datos y devuelve unos comandos de coneccion.
On Error GoTo n 'Este procedimiento permitirá cargar la ultima base de datos llamada por el usuario.
Dim v As String
 Dim vString As String
 v = Direccion_de_la_base_de_datos
 'prBaseNombre = V'Porque sucederia un desvordamiento de pila...
vString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                v & ";"

Obtener_Tablas vString, List1

 fConectandoBase = vString

Exit Function
n: 'No molestar al usuario.
'MsgBox "Tengo problemas con fConectandoBase"
End Function


Private Sub cmdCargar_Click()
 Obtener_Tablas prBaseNombre, List1
End Sub

Private Sub CMDEnd_Click()
K.mUltimo
End Sub

Private Sub cmdFerst_Click()
K.mPrimero
End Sub

Private Sub CMDSiguiente_Click()
K.mSiguiente
End Sub

Private Sub K_Actualizando()
Dim H As Long
Text3 = ""
For H = 0 To List2.ListCount
'List2.ListIndex = H
DoEvents
If H > List2.ListCount Then Exit Sub
Text3.SelText = Chr(13) + Chr(10) & "////////////////////////" & Chr(13) + Chr(10) & K.GetF(H)
Next
End Sub

Private Sub K_Conectado(Estado As Boolean)
'
End Sub


Private Sub List1_Click()
 Obtener_Campos List1.Text, List2, Me.NotasDB
Text1 = List1.Text
Obtener_Campos List1.Text, List2, vBaseNombre
K.Cerrar
  
K.LoadDBs List1.Text, prBaseNombre       ', "D:\Proyectos\Alternativo\Notas"
RaiseEvent evListTablasClick(List1.Text)
End Sub

Private Sub List2_Click()
Text2 = List2.Text
RaiseEvent evListFieldsClick(List2.Text) 'Cuando se pasan argumentos a un evento se debe crear una base local, cargarla con los datos y pasarla al argumento; esto para impedir que el argumento devuelva desde otras aplicasiones datos que puedan causar una distorcion en el funcionamiento del control ActiveX.
End Sub

Private Sub List2_DblClick()
Dim mList() As Variant 'Matris local para cargar los items (Fields).
Dim vCont As Integer, mSqlFomrat() As Variant
For vCont = 0 To (List2.ListCount - 1)
DoEvents
    If vCont = 0 Then
    ReDim mList(0) 'Le da el primer id y ademas borra una posible anterior llamada.
    ReDim mSqlFomrat(0)
    Else
    ReDim Preserve mList(vCont)
    ReDim Preserve mSqlFomrat(vCont)
    End If
mSqlFomrat(vCont) = List2.List(vCont) & " As " & List2.List(vCont)
mList(vCont) = List2.List(vCont)
Next vCont
prTodosLosFields = Join(mList, rtc) 'Cargando lista de campos.
Me.prSQLParaBusquedas = Join(mSqlFomrat, ", ")
RaiseEvent evDblClickCargandoTodosLosFields(List2.Text, List2.ListIndex, prTodosLosFields)
prFieldsMax = UBound(mList)
RaiseEvent evAllFieldsMatriz(mList, prFieldsMax) 'No importa si el argumento es manipulado ya que, la matriz esta por morir despues de esta linea.

End Sub

Private Sub Tr_BeforeLabelEdit(Cancel As Integer)
'
End Sub


Private Sub UserControl_Initialize()
Set K = New clsDatos
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
prBaseNombre = .ReadProperty("prBaseNombre", "")
'prConectarYa = .ReadProperty("prConectarYa", False)
End With
prTodosLosFields = PropBag.ReadProperty("prTodosLosFields", 0)

'sScroolBarHorizontal List1
'sScroolBarHorizontal List2
prSQLParaBusquedas = PropBag.ReadProperty("prSQLParaBusquedas", 0)
prFieldsMax = PropBag.ReadProperty("prFieldsMax", 0)
fConectandoBase prBaseNombre 'Ya que se usa prBaseNombre para autocargarse.
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "prSQLParaBusquedas", prSQLParaBusquedas
With PropBag
.WriteProperty "prBaseNombre", prBaseNombre
'.WriteProperty "prConectarYa", prConectarYa
End With
PropBag.WriteProperty "prTodosLosFields", prTodosLosFields
PropBag.WriteProperty "prFieldsMax", prFieldsMax

End Sub


Public Property Get prtxtTabla() As Variant

On Error GoTo n

prtxtTabla = Text1
Exit Property
n:
MsgBox "Tengo problemas con vtxtTabla"
End Property
Public Property Let prtxtTabla(vNuevosDatos As Variant)
On Error GoTo n
MsgBox "Es de solo lectura."
'vtxtTabla = vNuevosDatos
'PropertyChanged "prtxtTabla"
Exit Property
n:
MsgBox "Tengo problemas con vtxtTabla"
End Property

Public Property Get prtxtField() As Variant

On Error GoTo n

prtxtField = Text2
Exit Property
n:
MsgBox "Tengo problemas con vtxtField"
End Property
Public Property Let prtxtField(vNuevosDatos As Variant)
On Error GoTo n
MsgBox "Es de solo lectura."
'vtxtField = vNuevosDatos
'PropertyChanged "prtxtField"
Exit Property
n:
MsgBox "Tengo problemas con vtxtField"
End Property


Public Property Get prAboudMe() As Variant 'No devuelve nada solo muestra una pagina de propiedad a modo de Aboud...

On Error GoTo n

'prAboudMe = vAboudMe
Exit Property
n:
MsgBox "Tengo problemas con prAboudMe"
End Property
Public Property Let prAboudMe(vNuevosDatos As Variant)  'No devuelve nada solo muestra una pagina de propiedad a modo de Aboud...
On Error GoTo n

'vAboudMe = vNuevosDatos
Exit Property
n:
MsgBox "Tengo problemas con prAboudMe"
End Property

Public Property Get prSQLParaBusquedas() As Variant 'Para insertar una secuencia: Field as Field, Field1 as Field1... que se usará en busquedas.

On Error GoTo n

prSQLParaBusquedas = vSQLParaBusquedas
Exit Property
n:
MsgBox "Tengo problemas con prSQLParaBusquedas"
End Property
Public Property Let prSQLParaBusquedas(vNuevosDatos As Variant)  'Para insertar una secuencia: Field as Field, Field1 as Field1... que se usará en busquedas.
On Error GoTo n

vSQLParaBusquedas = vNuevosDatos
PropertyChanged "prSQLParaBusquedas"
Exit Property
n:
MsgBox "Tengo problemas con prSQLParaBusquedas"
End Property

Public Property Get prFieldsMax() As Integer 'Maxima cantidad de campos en una tabla. Util para el manejo de matrices.

On Error GoTo n

prFieldsMax = vFieldsMax
Exit Property
n:
MsgBox "Tengo problemas con prFieldsMax"
End Property
Public Property Let prFieldsMax(vNuevosDatos As Integer)  'Maxima cantidad de campos en una tabla. Util para el manejo de matrices.
On Error GoTo n

vFieldsMax = vNuevosDatos
PropertyChanged "prFieldsMax"
Exit Property
n:
MsgBox "Tengo problemas con prFieldsMax"
End Property

