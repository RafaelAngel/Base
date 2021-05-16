VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Declaracion para activar controles
Private Declare Function SendMessage Lib "user32" Alias _
 "SendMessageA" (ByVal hwnd As Long, ByVal wMsg _
 As Long, ByVal wParam As Long, lParam As Long) _
 As Long
 'Variables de apoyo para las declaraciones
Private Const CB_SHOWDROPDOWN = &H14F
   Private Tmp As Variant



Public ITM As ListItem 'Variable para los SubItems


'Private R As New ClsServer

Private WithEvents vComboBox As ComboBox 'Para cargarlo desde aqui.
Attribute vComboBox.VB_VarHelpID = -1

'Los eventos de la clase se pueden llamar o usar _
de la siguiente manera: _
Usando WithEvents NombreClass
'----------------------------------------------
'Luego la clase pasa una plantilla con todos los eventos que posee _
para que se puedan usar como eventos estandar.
'----------------------------------------------
Private vTabla As Variant
Private vQCampo As Single
'Declaración del objeto CONNECTION, el
'cual me permite acceso a una base de datos
Private Conect As New ADODB.Connection
'Declaración de un objeto RECORDSET, el cual
'representa los registros de una tabla
Private Records As New ADODB.Recordset
'----------------------------------------------
'Debe volverse a activar en caso de error.
Public Fields As New clsCampos
Public Event Conectado(Estado As Boolean)
Public Event ControlesHabilitados() 'Para habilitar los controles.
Public Event NuevoRegistro()
Public Event ControlesDeshabilitados() 'Deshabilita los controles.

Public Event Actualizando() 'Para actualizar datos y campos.

Public Event RsErr(Mensaje As Variant, ProcedimientoNombre As Variant) 'Para devolver informacion de los errores.

Public Event SentFields() 'Para cargar los campos de la tabla.


Public Event LimpiandoParaAgregarNuevoRegistro() 'Sucede al agregar un nuevo registro.
'variables locales para almacenar los valores de las propiedades
Private mvarConecting As Boolean 'copia local

Public Event EVECargandoListView(Estado As Variant) 'Cuandos e carga el ListView
Public Property Let Conecting(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Conecting = 5
    mvarConecting = vData
End Property


Public Property Get Conecting() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Conecting
    Conecting = mvarConecting
End Property


Public Sub Limpiar(Nombre_del_form As Object)
'Este procedimiento limpia los campos de texto.
Dim Objeto As Object
On Error GoTo n
For Each Objeto In Nombre_del_form
If TypeOf Objeto Is TextBox Then
 Objeto.Text = ""
End If
Next Objeto
Exit Sub
n:
'RaiseEvent RsErr(RST(115), "Limpiar en clsevents")
Exit Sub
End Sub

Public Sub LoadDBs(nombre_de_la_tabla As Variant, Base As Variant)
    'Conecta con una base de datos

    On Error GoTo n
    Cerrar 'Por si esta abierta con otra tabla.
    Conect.Provider = "Microsoft.Jet.OLEDB.4.0"
    'Aqui establesco la coneccion con la base de datos
    Conect.ConnectionString = Base '"D:\Proyectos\Alternativo\Notas.mdb"
    'Habre la conexion
    Conect.Open
    'Aqui se obtiene una serie de registros: en este
    'caso de la
    'tabla
    Records.Open "select * from " & nombre_de_la_tabla, Conect, , adLockOptimistic
    'Siguiente
    'Anterior
    RaiseEvent Actualizando
    RaiseEvent Conectado(True)
    'MsgBox "La tabla " & Nombre_de_la_tabla & " esta Conectada.", vbInformation
    Exit Sub
n:
    Debug.Print "Hay problemas en LoadDBs" & rtc & Err.Description
    RaiseEvent Conectado(False)
    RaiseEvent RsErr(Err.Description, " LoadDBs en clsevents")   '"No puedo conectar con la tabla.")
    'LoadDBsAlternate Nombre_de_la_tabla
    'Err.Raise 4000, "LoadDBs", Err.Description
End Sub


Public Function Cont() As Long
On Error GoTo n
Static v As Long
v = 0
mPrimero
 While Not Records.EOF
      v = v + Val(1)
      Records.MoveNext  ' Pasa al registro siguiente.
      
   Wend
   Records.MoveFirst
   Cont = v
   v = 0
   Exit Function
n:
   RaiseEvent RsErr("Hay un valor nulo en la tabla, por favor eliminelo.", "Cont en clsEvents")
  'R.msgB App.EXEName, "Cont en clsEvents", Err.Description, enClass
  Cont = v
  Exit Function
End Function

Public Sub Nuevo()
'Agrega un nuevo registro pero tambien debe usarse el _
procedimiento limpiar.
On Error GoTo n
Records.AddNew
RaiseEvent ControlesDeshabilitados
RaiseEvent Actualizando
RaiseEvent LimpiandoParaAgregarNuevoRegistro
RaiseEvent NuevoRegistro
Exit Sub
n:
RaiseEvent RsErr("No puedo crear un nuevo registro. ", "Nuevo en clsEvents")
Exit Sub
End Sub

Private Sub MsgErr(Informacion_del_error As Variant, _
ProcedimientoNombre As Variant)
'Activa el evento de error y devuelve la informacion.
RaiseEvent RsErr(Informacion_del_error, ProcedimientoNombre)
'R.msgB App.EXEName, ProcedimientoNombre, Informacion_del_error, enClass
End Sub

Public Sub Guardar()
'Guarda los datos editados...
On Error GoTo n
 RaiseEvent SentFields
RaiseEvent Actualizando
RaiseEvent ControlesHabilitados

Exit Sub
n:
'RaiseEvent RsErr(RST(110) & " no puedo guardar los datos en la tabla.", " Guardar en clsEvents")
Exit Sub
End Sub

Public Sub mPrimero()
'Va al primer registro.
On Error GoTo n:
Records.MoveFirst
RaiseEvent Actualizando
Exit Sub
n:

Exit Sub
End Sub

Public Sub mUltimo()
'Va al ultimo registro
On Error GoTo n:
Records.MoveLast
RaiseEvent Actualizando
Exit Sub
n:

Exit Sub
End Sub

Public Sub mAnterior()
'Va al anterior registro
On Error GoTo n
Records.MovePrevious
If Records.BOF = True Then
Records.MoveFirst
Else
RaiseEvent Actualizando   'Actualiza los datos en el form
End If
Exit Sub
n:
'No hacer nada
Exit Sub
End Sub
Public Sub mSiguiente()
'Va al siguiente registro
On Error GoTo n
Records.MoveNext
If Records.EOF = True Then
Records.MoveLast
Else
RaiseEvent Actualizando
End If
Exit Sub
n:
RaiseEvent RsErr(Err.Description, "mSiguiente")
Exit Sub
End Sub


Public Sub Eliminar()
'Elimina datos de la tabla o registros, eso cada _
vez que se ejecute este comando.
Dim vSiNo As Long
On Error GoTo Cancel
vSiNo = MsgBox("¿En realidad quiere eliminar el registro?", vbYesNo)
Select Case vSiNo
Case 6
 Records.Delete 'Elimina el registro celeccionado
 ElP
 RaiseEvent Actualizando
Case 7
'No se hace nada.
'R.msgB App.EXEName, "Eliminar en clsEvents", "Se cancelo el comando de eliminar el registro, buena decision.", enClass
End Select
Exit Sub
Cancel:
'R.msgB App.EXEName, "Eliminar en clsEvents", Err.Description, enClass
'RaiseEvent RsErr(RST(106) & " No puedo eliminar el registro.", "Eliminar en ClsEvents")
End Sub

Private Sub ElP()
'Al eliminar un registro se mueve al registro anterior.
On Error GoTo n
Records.MovePrevious
Exit Sub
n:
ElN
Exit Sub
End Sub

Private Sub ElN()
'Al eliminar un registro se mueve al siguiente registro.
On Error GoTo n
Records.MoveNext
Exit Sub
n:
'RaiseEvent RsErr(RST(105) & " Un procedimiento private.", " eln en clsevents")
Exit Sub
End Sub

Public Function Enter_()
'Es lo mismo que RetornoDeCarro
Enter_ = Chr(13) & Chr(10)
End Function

Public Function Arroba()
'Devuelve un arroba.
Arroba = "@"
End Function

Public Function Colon()
'Devuelve el simbolo de la moneda tica _
aunque se parece mucho al simbolo de centavos.
Colon = "¢"
End Function


Public Function Sharp()
'Devuelve el simbolo de numeral o Sharp
Sharp = "#"
End Function

Public Sub AcercaDe_()
'Muestra el dialogo acerca de...
'FrmAcercaDe.Show
End Sub

Public Sub Buscando(Nombre_del_field As Variant, Datos_del_campo As Variant, Informacion_buscada As Variant)
'Este procedimiento busca datos en la tabla en la que se este conectado.
Dim criteriob As String
criteriob = "[" & Nombre_del_field & "]=" & Datos_del_campo
Records.MoveFirst
criteriob = "[" & Nombre_del_field & "] like '*" & Informacion_buscada & "*'"
Records.Find criteriob
If Informacion_buscada <> "" Then
Records.MoveFirst
Records.Find criteriob
RaiseEvent Actualizando
End If
End Sub

Public Sub CargandoCombo(Nombre_del_ComboBox As Object, Nombre_del_field As Variant)
'Esto carga un ComboBox
Dim F As Object
Static v As Long
On Error GoTo n
Nombre_del_ComboBox.clear
Fields.FieldDeBusqueda = Nombre_del_field
F.Show
F.Progres1.Max = Me.Cont
Records.MoveFirst
 While Not Records.EOF
      Nombre_del_ComboBox.AddItem CStr(Records.Fields(Nombre_del_field))
      Records.MoveNext  ' Pasa al registro siguiente.
   v = v + Val(1)
   F.Progres1.Avanzando v
   Wend
   Records.MoveFirst
   Unload F
   v = 0
   Set F = Nothing
   Set vComboBox = Nombre_del_ComboBox 'Para asociarlo.
   Exit Sub
n:
 Unload F
   v = 0
   Set F = Nothing
  'RaiseEvent RsErr(RST(103) & "CargandoCombo", "CargandoCombo en clsEvents")
'R.msgB App.EXEName, "CargandoCombo en clsEvents", Err.Description, enClass
End Sub


Public Sub CargandoListView(Nombre_del_ListView As Object, Optional Icono As Long = 1, Optional SmallIcono As Long = 1)
'Esto carga listview con datos de una tabla base.
'Tambien llevo dias tratando de crear algo nunca cre4ado.
'Ya esta perfeccionado.
Const Cambiar_a_reporte = "Cambie a vista reporte o no se cargaran los datos."
Static v As Long
Static Nf As Variant
Dim Nc
On Error GoTo n
x1:
Select Case Nombre_del_ListView.View

Case lvwIcon
Nombre_del_ListView.View = lvwReport
GoTo x1
'MsgBox Cambiar_a_reporte, vbInformation
Case lvwList
Nombre_del_ListView.View = lvwReport
GoTo x1
'MsgBox Cambiar_a_reporte, vbInformation
Case lvwSmallIcon
Nombre_del_ListView.View = lvwReport
GoTo x1
'MsgBox Cambiar_a_reporte, vbInformation
Case lvwReport
Nombre_del_ListView.ListItems.clear
'Fields.FieldDeBusqueda = Nombre_del_Field
'F.Show
'F.Progres1.Max = Me.Cont
Nc = ListViewColumnasCount(Nombre_del_ListView)

For Nf = 2 To Nombre_del_ListView.ColumnHeaders.Count
Records.MoveFirst
 While Not Records.EOF
      Records.MoveNext  ' Pasa al registro siguiente.
   v = v + Val(1)
   AddToListView Nc, Nombre_del_ListView, CStr(Records.Fields(Nombre_del_ListView.ColumnHeaders.Item(1).Text)), Icono, SmallIcono
   'F.Progres1.Avanzando V
   RaiseEvent EVECargandoListView("Cargando")
   Wend
   Records.MoveFirst
   'Unload F
   v = 0
 
   Next
   Nf = 0
   'Set F = Nothing
  ' Set vComboBox = Nombre_del_ComboBox 'Para asociarlo.
  End Select

   Exit Sub
n:
 'Unload F
   v = 0
   Nf = 0
   'Set F = Nothing
  RaiseEvent RsErr("CargandoListView", "CargandoListView en clsX")
'R.msgB App.EXEName, "CargandoCombo en clsEvents", Err.Description, enClass
RaiseEvent EVECargandoListView("Cargado")
End Sub

Public Sub CargandoLista(Nombre_del_ListBox As Object, Nombre_del_field As Variant)
'Esto carga un ComboBox
On Error GoTo n
Dim F As Object
Static v As Long
F.Show
F.Progres1.Max = Me.Cont
mPrimero
 While Not Records.EOF
      Nombre_del_ListBox.AddItem CStr(Records.Fields(Nombre_del_field))
      Records.MoveNext  ' Pasa al registro siguiente.
    v = v + Val(1)
   F.Progres1.Avanzando v
   Wend
   Records.MoveFirst
    Unload F
   v = 0
   Set F = Nothing
  Exit Sub
n:
 Unload F
   v = 0
   Set F = Nothing
   ' RaiseEvent RsErr(RST(103) & " Cargando el List", "CargandoLista en clsEvents")
   'R.msgB App.EXEName, "CargandoLista en clsEvents", Err.Description, enClass
   Exit Sub
End Sub

Public Sub FieldRecuperar(Campo_TextBox_para_el_Field As Object, _
Nombre_del_Field_en_la_tabla As Variant)
'Recupera datos de un campo Field que esta en la tabla especificada en la coneccion. _
Y los pasa al TextBox. No es necesario el nombre de la tabla porque se supone que ya la coneccion se ha efectuado.
Campo_TextBox_para_el_Field = Records.Fields(Nombre_del_Field_en_la_tabla)
End Sub
Public Sub GetField(Campo_TextBox_para_el_Field As Object, _
Nombre_del_Field_en_la_tabla As Variant)
On Error GoTo n
'Recupera datos de un campo Field que esta en la tabla especificada en la coneccion. _
Y los pasa al TextBox. No es necesario el nombre de la tabla porque se supone que ya la coneccion se ha efectuado.
Campo_TextBox_para_el_Field = Records.Fields(Nombre_del_Field_en_la_tabla)
Exit Sub
n:
Debug.Print "Hay problemas en GetF" & rtc & Err.Description
'RaiseEvent RsErr("No puedo hacer sub getfield.", "GetField en clsEvents")
End Sub
Public Sub GetFldComboBox(Combo As Object, _
Nombre_del_Field_en_la_tabla As Variant)
On Error GoTo n
'Recupera datos de un campo Field que esta en la tabla especificada en la coneccion. _
Y los pasa al TextBox. No es necesario el nombre de la tabla porque se supone que ya la coneccion se ha efectuado.
Combo.Text = Records.Fields(Nombre_del_Field_en_la_tabla)
Exit Sub
n:
'RaiseEvent RsErr("No puedo hacer sub getfield.", "GetField en clsEvents")
End Sub
Public Function GetF(FieldNombre As Variant)
On Error GoTo n
'Recupera datos de un campo Field que esta en la tabla especificada en la coneccion. _
Y los pasa al TextBox. No es necesario el nombre de la tabla porque se supone que ya la coneccion se ha efectuado.
Debug.Print "Datos de " & FieldNombre & "= " & CStr(Records.Fields(FieldNombre)) & ""
GetF = CStr(Records.Fields(FieldNombre)) & ""
Exit Function
n:
Debug.Print "Hay problemas en GetF" & rtc & Err.Description

'RaiseEvent RsErr("No puedo recuperar los datos de fuction GetF.", "GetF en clsEvents")
'R.msgB App.EXEName, "GetF en clsEvents", Err.Description, enClass
End Function
Public Sub LetField(Campo_TextBox_para_el_Field As Object, _
Nombre_del_Field_en_la_tabla As Variant)
On Error GoTo n
'Se supone que esto es lo mismo que guardar pero guarda field por field.
'debe usarse el evento guardar o sendfields.
Records.Update
Records.Fields(Nombre_del_Field_en_la_tabla) = Campo_TextBox_para_el_Field
Exit Sub
n:
'RaiseEvent RsErr("No puedo realizar LetField. LetField(Guardar)", "Let field en clsevents")
End Sub

Public Sub LetFieldNuevoRegistro(Datos As Variant, _
Nombre_del_Field_en_la_tabla As Variant)
On Error GoTo n
'Se supone que esto es lo mismo que guardar pero guarda field por field.
'debe usarse el evento guardar o sendfields.
Records.Update
Records.Fields(Nombre_del_Field_en_la_tabla) = Datos
Exit Sub
n:
'RaiseEvent RsErr("No puedo realizar LetField. LetField(Guardar)", "Let field en clsevents")
End Sub

Public Sub LetFieldComboBox(Combo As Object, _
Nombre_del_Field_en_la_tabla As Variant)
On Error GoTo n
'Se supone que esto es lo mismo que guardar pero guarda field por field.
Records.Update
Records.Fields(Nombre_del_Field_en_la_tabla) = Combo.Text
Exit Sub
n:
'RaiseEvent RsErr("No puedo realizar LetField. LetField(Guardar)", "Let field en clsevents")
End Sub

Public Sub FieldActualizar(Campo_TextBox_para_el_Field As Object, _
Nombre_del_Field_en_la_tabla As Variant)
'Actualiza los datos en la tabla, basicamente es lo mismo que recuperar datos _
pero con una pequeña diferencia, en este caso se envian _
datos a la tabla.
Records.Fields(Nombre_del_Field_en_la_tabla) = Campo_TextBox_para_el_Field
End Sub

Public Sub GuardandoFieldPorField(Campo_TextBox_para_el_Field As Object, _
Nombre_del_Field_en_la_tabla As Variant)
'Se supone que esto es lo mismo que guardar pero guarda field por field.
Records.Update
Records.Fields(Nombre_del_Field_en_la_tabla) = Campo_TextBox_para_el_Field
End Sub

Public Sub CmbActivate(Combo As Object)
'Activa un ComboBox
Tmp = SendMessage(Combo.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Public Sub Cerrar()
On Error GoTo n
    Records.Close
    Conect.Close
Exit Sub
n:
'RaiseEvent RsErr(Err.Description, "Cerrar en clsEvents")
'R.msgB App.EXEName, "Cerrar en clsEvents", Err.Description, enClass
End Sub

Public Function InterrogativSimbol()
InterrogativSimbol = "¿?"
End Function

Public Function ExclamativSimbol()
ExclamativSimbol = "¡!"
End Function

Public Property Get Tabla() As Variant
Tabla = vTabla
End Property
Public Property Let Tabla(ByVal vNewValue As Variant)
vTabla = vNewValue
End Property

Public Sub GetBase(BaseNombre As Variant)
LoadDBs Tabla, BaseNombre
End Sub

Public Property Get QCampo() As Single
QCampo = vQCampo
End Property
Public Property Let QCampo(ByVal vNewValue As Single)
vQCampo = vNewValue
End Property

Private Sub A(Datos As Variant)
End Sub

Private Sub Class_Initialize()
'R.AppNombre = App.EXEName
End Sub

Private Sub Class_Terminate()
'Set R = Nothing
Set vComboBox = Nothing
Set Fields = Nothing
End Sub

Public Sub ActualizandoComboBox()
Me.CargandoCombo vComboBox, Me.Fields.FieldDeBusqueda
End Sub

Private Sub vComboBox_Click()
Buscando Fields.FieldDeBusqueda, vComboBox.Text, vComboBox.Text
End Sub

Private Sub vComboBox_GotFocus()
CmbActivate vComboBox
End Sub

Private Sub vComboBox_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CmbActivate vComboBox
End If
End Sub

Public Sub LoadDBsAlternate(nombre_de_la_tabla)
'Conecta con una base de datos
Dim x As Variant
On Error GoTo n
x = GetSetting(App.Path, "db", "base")
If x = "" Then
x = InputBox("Escriba o pegue la direccion de la base de datos Notas.mdb.", , "Notas" & ".mdb")
SaveSetting App.Path, "db", "base", x
End If
Conect.Provider = "Microsoft.Jet.OLEDB.4.0"
'Aqui establesco la coneccion con la base de datos
Conect.ConnectionString = x '"D:\Proyectos\Alternativo/Notas.mdb"
'Habre la conexion
Conect.Open
'Aqui se obtiene una serie de registros: en este
'caso de la
'tabla
Records.Open "select * from " & nombre_de_la_tabla, Conect, , adLockOptimistic
'Siguiente
'Anterior
RaiseEvent Actualizando
'MsgBox "La tabla " & Nombre_de_la_tabla & " esta Conectada.", vbInformation
Exit Sub
n:
x = InputBox("Escriba o pegue la direccion de la base de datos Notas.mdb.", , "Notas" & ".mdb")
SaveSetting App.Path, "db", "base", x
MsgBox "Reinicie la base de datos si proporciono la direccion correcta.", vbInformation
End Sub

Public Sub AddToListView(Columnas As Variant, ListViewNombre As Object, Datos As Variant, Optional Icono As Long = 1, Optional SmallIcono As Long = 1)
'He durado muchos dias perfeccionando este codigo, ya lo logré.
'Ademas se puede amplificar y basado en esto se puede usar de diferente manera.
'Lo mío era para bases de datos.
On Error Resume Next

Static v
Set ITM = ListViewNombre.ListItems.add(, , Datos)
For v = 1 To Columnas '+ Val(1)
DoEvents
With ITM
    .Icon = Icono   ' Establece un icono de ImageList1.
    .SmallIcon = SmallIcono
    .SubItems(v - Val(1)) = GetF(ListViewNombre.ColumnHeaders.Item(v).Text) 'K.GetF(ListViewNombre.ColumnHeaders.Item(V).Text)
    End With
Next
v = 0
Set ITM = Nothing
Exit Sub
n:
Set ITM = Nothing
MsgBox Err.Description, , "addtoListView"
End Sub

Public Function ListViewColumnasCount(ListViewNombre As Object)
ListViewColumnasCount = ListViewNombre.ColumnHeaders.Count
End Function

Public Function DBDir(Optional direccion As String = "D:\Ralfy\Notas\Proyecto\Notas.mdb") As String
DBDir = direccion
End Function



