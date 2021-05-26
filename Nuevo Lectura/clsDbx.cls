VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDbx"
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

Public Event evContando(ByRef Realizando_conteo As Long)


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
Private vprQueCampo As Single
'Declaración del objeto CONNECTION, el
'cual me permite acceso a una base de datos
Public Conect As New ADODB.Connection
'Declaración de un objeto RECORDSET, el cual
'representa los registros de una tabla
Public Records As New ADODB.Recordset
'----------------------------------------------
Public Fields As New ClsFields
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
Public Event MinRs()

Public Event MaxRs()

Public Event Eliminando(Eliminado As Boolean)



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
On Error GoTo N

For Each Objeto In Nombre_del_form

    If TypeOf Objeto Is TextBox Then
        Objeto.Text = ""
    End If
    
Next Objeto

Exit Sub
N:
'RaiseEvent RsErr(RST(115), "Limpiar en clsevents")
Exit Sub
End Sub

Public Sub LoadDBs(Nombre_de_la_tabla As Variant)
    'Conecta con una base de datos
    On Error GoTo N
    Cerrar 'Por si esta abierta con otra tabla.
    Conect.Provider = "Microsoft.Jet.OLEDB.4.0"
    'Aqui establesco la coneccion con la base de datos
    Conect.ConnectionString = DBDir(App.Path & "/DbLectura2021.mdb") ' "D:\Proyectos\Alternativo\Notas.mdb"
    'Habre la conexion
    Conect.Open
    'Aqui se obtiene una serie de registros: en este
    'caso de la
    'tabla
    Records.Open "select * from " & Nombre_de_la_tabla, Conect, , adLockOptimistic
    'Siguiente
    'Anterior
    RaiseEvent Actualizando
    RaiseEvent Conectado(True)
    'MsgBox "La tabla " & Nombre_de_la_tabla & " esta Conectada.", vbInformation
    Exit Sub
N:
    MsgBox Err.Description, vbExclamation, "LoadDBs"
    RaiseEvent Conectado(False)
    RaiseEvent RsErr(Err.Description, " LoadDBs en clsevents")   '"No puedo conectar con la tabla.")
    'LoadDBsAlternate Nombre_de_la_tabla
    'Err.Raise 4000, "LoadDBs", Err.Description
End Sub

Public Sub LoadDBConSQL(Consulta_sql As Variant)
    'Conecta con una base de datos
    On Error GoTo N
    Cerrar 'Por si esta abierta con otra tabla.
    Conect.Provider = "Microsoft.Jet.OLEDB.4.0"
    'Aqui establesco la coneccion con la base de datos
    Conect.ConnectionString = DBDir(App.Path & "/DbLectura2021.mdb") ' "D:\Proyectos\Alternativo\Notas.mdb"
    'Habre la conexion
    Conect.Open
    'Aqui se obtiene una serie de registros: en este
    'caso de la
    'tabla
    Records.Open Consulta_sql, Conect, , adLockOptimistic
    'Siguiente
    'Anterior
    RaiseEvent Actualizando
    RaiseEvent Conectado(True)
    'MsgBox "La tabla " & Nombre_de_la_tabla & " esta Conectada.", vbInformation
    Exit Sub
N:
    MsgBox Err.Description, vbExclamation, "LoadDBsConSQL"
    RaiseEvent Conectado(False)
    RaiseEvent RsErr(Err.Description & RTC & RTC & Consulta_sql, " LoadDBConSQL en clsevents")   '"No puedo conectar con la tabla.")
    'LoadDBsAlternate Nombre_de_la_tabla
    'Err.Raise 4000, "LoadDBs", Err.Description
End Sub


Public Function Cont() As Long
On Error GoTo N
Static V As Long
V = 0
mPrimero

While Not Records.EOF
      V = V + Val(1)
      RaiseEvent evContando(V)
      Records.MoveNext  ' Pasa al registro siguiente.
Wend

   Records.MoveFirst
   Cont = V
   V = 0
   Exit Function
N:
   RaiseEvent RsErr("Hay un valor nulo en la tabla, por favor eliminelo.", "Cont en clsEvents")
  'R.msgB App.EXEName, "Cont en clsEvents", Err.Description, enClass
  Exit Function
End Function

Public Sub Nuevo()
    'Agrega un nuevo registro pero tambien debe usarse el _
    procedimiento limpiar.
On Error GoTo N
    Records.AddNew
    RaiseEvent ControlesDeshabilitados
    RaiseEvent LimpiandoParaAgregarNuevoRegistro
    RaiseEvent NuevoRegistro
    
    'No se debe llamar a este evento.
    'RaiseEvent Actualizando 'El ultimo evento en ser llamado es Actualizando.
    Exit Sub
N:
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
On Error GoTo N
RaiseEvent SentFields
RaiseEvent Actualizando
RaiseEvent ControlesHabilitados

Exit Sub
N:
'RaiseEvent RsErr(RST(110) & " no puedo guardar los datos en la tabla.", " Guardar en clsEvents")
Exit Sub
End Sub

Public Sub mPrimero()
'Va al primer registro.
On Error GoTo N:
Records.MoveFirst
RaiseEvent Actualizando
RaiseEvent MinRs
Exit Sub
N:

Exit Sub
End Sub

Public Sub mUltimo()
'Va al ultimo registro
On Error GoTo N:
Records.MoveLast
RaiseEvent Actualizando
RaiseEvent MaxRs
Exit Sub
N:

Exit Sub
End Sub

Public Sub mAnterior()
'Va al anterior registro
On Error GoTo N
Records.MovePrevious

If Records.BOF = True Then
    Records.MoveFirst
Else
    'RaiseEvent Actualizando   'Actualiza los datos en el form
End If

RaiseEvent Actualizando
Exit Sub
N:
RaiseEvent MinRs
'No hacer nada
Exit Sub
End Sub
Public Sub mSiguiente()
'Va al siguiente registro
On Error GoTo N
Records.MoveNext

If Records.EOF = True Then
    Records.MoveLast
Else
    'RaiseEvent Actualizando
End If

RaiseEvent Actualizando
Exit Sub
N:
RaiseEvent MaxRs
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
         RaiseEvent Eliminando(True)
         Records.Delete 'Elimina el registro celeccionado
         sEliminarMoverseAlAnterior
         RaiseEvent Actualizando
        ' RaiseEvent Actualizando
    Case 7
        RaiseEvent Eliminando(False)
        'No se hace nada.
        'R.msgB App.EXEName, "Eliminar en clsEvents", "Se cancelo el comando de eliminar el registro, buena decision.", enClass
    End Select
    
     RaiseEvent Actualizando
    Exit Sub
Cancel:
    RaiseEvent Actualizando
    'R.msgB App.EXEName, "Eliminar en clsEvents", Err.Description, enClass
    'RaiseEvent RsErr(RST(106) & " No puedo eliminar el registro.", "Eliminar en ClsEvents")
End Sub

Private Sub sEliminarMoverseAlAnterior()
    'Al eliminar un registro se mueve al registro anterior.
    On Error GoTo N
    Records.MovePrevious
    Exit Sub
N:
    sEliminarMoverseAlSiguiente
Exit Sub
End Sub

Private Sub sEliminarMoverseAlSiguiente()
    'Al eliminar un registro se mueve al siguiente registro.
    On Error GoTo N
    Records.MoveNext
    Exit Sub
N:
    'RaiseEvent RsErr(RST(105) & " Un procedimiento private.", " sEliminarMoverseAlSiguiente en clsevents")
Exit Sub
End Sub

Public Function fRTC()
    'Es lo mismo que RetornoDeCarro
    fRTC = Chr(13) & Chr(10)
End Function

Public Function fArroba()
    'Devuelve un arroba.
    fArroba = "@"
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

Public Sub Buscando(Nombre_del_Field As Variant, Datos_del_campo As Variant, Informacion_buscada As Variant)
    'Este procedimiento busca datos en la tabla en la que se este conectado.
    On Error Resume Next
    Dim criteriob As String
    criteriob = "[" & Nombre_del_Field & "]=" & Datos_del_campo
    Records.MoveFirst
    criteriob = "[" & Nombre_del_Field & "] like '*" & Informacion_buscada & "*'"
    Records.Find criteriob
    
    If Informacion_buscada <> "" Then
        Records.MoveFirst
        Records.Find criteriob
        RaiseEvent Actualizando
    End If

End Sub

Public Sub CargandoCombo(Nombre_del_ComboBox As Object, Nombre_del_Field As Variant)
'Esto carga un ComboBox
    'Dim F As New FrmProgress
Static V As Long
On Error GoTo N
Nombre_del_ComboBox.Clear
Fields.FieldDeBusqueda = Nombre_del_Field
    'F.Show
    'F.Progres1.Max = Me.Cont
Records.MoveFirst

While Not Records.EOF
      Nombre_del_ComboBox.AddItem CStr(Records.Fields(Nombre_del_Field))
      Records.MoveNext  ' Pasa al registro siguiente.
   V = V + Val(1)
    'F.Progres1.Avanzando V
Wend

   Records.MoveFirst
    'Unload F
   V = 0
    'Set F = Nothing
   Set vComboBox = Nombre_del_ComboBox 'Para asociarlo.
   Exit Sub
N:
    'Unload F
   V = 0
    'Set F = Nothing
   MsgBox Err.Description & RTC & "Ha ocurrido un error al cargar los datos del ComboBox.", vbExclamation, "CargandoCombo"
  'RaiseEvent RsErr(RST(103) & "CargandoCombo", "CargandoCombo en clsEvents")
'R.msgB App.EXEName, "CargandoCombo en clsEvents", Err.Description, enClass
End Sub


Public Sub CargandoListView(Nombre_del_ListView As ListView, Optional Icono As Long = 1, Optional SmallIcono As Long = 1)
'Esto carga listview con datos de una tabla base.
'Tambien llevo dias tratando de crear algo nunca cre4ado.
'Ya esta perfeccionado.
Const Cambiar_a_reporte = "Cambie a vista reporte o no se cargaran los datos."
Static V As Long
Static Nf As Variant
Dim Nc
On Error GoTo N
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
    Nombre_del_ListView.ListItems.Clear
    'Fields.FieldDeBusqueda = Nombre_del_Field
    'F.Show
    'F.Progres1.Max = Me.Cont
    Nc = ListViewColumnasCount(Nombre_del_ListView)
    
    For Nf = 2 To Nombre_del_ListView.ColumnHeaders.Count
           Records.MoveFirst
           
           While Not Records.EOF
                Records.MoveNext  ' Pasa al registro siguiente.
                V = V + Val(1)
                AddToListView Nc, Nombre_del_ListView, CStr(Records.Fields(Nombre_del_ListView.ColumnHeaders.Item(1).Text)), Icono, SmallIcono
                'F.Progres1.Avanzando V
           Wend
           
           Records.MoveFirst
           'Unload F
           V = 0
    Next
       
       Nf = 0
       'Set F = Nothing
      ' Set vComboBox = Nombre_del_ComboBox 'Para asociarlo.
  End Select
   Exit Sub
N:
 'Unload F
   V = 0
   Nf = 0
   'Set F = Nothing
  RaiseEvent RsErr("CargandoListView", "CargandoListView en clsX")
'R.msgB App.EXEName, "CargandoCombo en clsEvents", Err.Description, enClass
End Sub

Public Sub CargandoLista(Nombre_del_ListBox As Object, Nombre_del_Field As Variant)
    'Esto carga un ComboBox
    On Error GoTo N
    Dim F As Object
    Static V As Long
    F.Show
    F.Progres1.Max = Me.Cont
    mPrimero
    
     While Not Records.EOF
          Nombre_del_ListBox.AddItem CStr(Records.Fields(Nombre_del_Field))
          Records.MoveNext  ' Pasa al registro siguiente.
          V = V + Val(1)
          F.Progres1.Avanzando V
    Wend
    
       Records.MoveFirst
        Unload F
       V = 0
       Set F = Nothing
      Exit Sub
N:
     Unload F
       V = 0
       Set F = Nothing
       ' RaiseEvent RsErr(RST(103) & " Cargando el List", "CargandoLista en clsEvents")
       'R.msgB App.EXEName, "CargandoLista en clsEvents", Err.Description, enClass
       Exit Sub
End Sub

Public Sub FieldRecuperar(Campo_TextBox_para_el_Field As TextBox, _
    Nombre_del_Field_en_la_tabla As Variant)
    'Recupera datos de un campo Field que esta en la tabla especificada en la coneccion. _
    Y los pasa al TextBox. No es necesario el nombre de la tabla porque se supone que ya la coneccion se ha efectuado.
    Campo_TextBox_para_el_Field = Records.Fields(Nombre_del_Field_en_la_tabla)
End Sub
Public Sub fGetField(Campo_TextBox_para_el_Field As TextBox, _
    Nombre_del_Field_en_la_tabla As Variant)
    On Error GoTo N
    'Recupera datos de un campo Field que esta en la tabla especificada en la coneccion. _
    Y los pasa al TextBox. No es necesario el nombre de la tabla porque se supone que ya la coneccion se ha efectuado.
    Campo_TextBox_para_el_Field = Records.Fields(Nombre_del_Field_en_la_tabla)
    Exit Sub
N:
    'RaiseEvent RsErr("No puedo hacer sub fGetField.", "fGetField en clsEvents")
End Sub
Public Sub fGetFldComboBox(Combo As ComboBox, _
    Nombre_del_Field_en_la_tabla As Variant)
    On Error GoTo N
    'Recupera datos de un campo Field que esta en la tabla especificada en la coneccion. _
    Y los pasa al TextBox. No es necesario el nombre de la tabla porque se supone que ya la coneccion se ha efectuado.
    Combo.Text = Records.Fields(Nombre_del_Field_en_la_tabla)
    Exit Sub
N:
    'RaiseEvent RsErr("No puedo hacer sub fGetField.", "fGetField en clsEvents")
End Sub
Public Function fGetF(FieldNombre As Variant)
    On Error GoTo N
    'Recupera datos de un campo Field que esta en la tabla especificada en la coneccion. _
    Y los pasa al TextBox. No es necesario el nombre de la tabla porque se supone que ya la coneccion se ha efectuado.
    fGetF = Records.Fields(FieldNombre)
    Exit Function
N:
    fGetF = ""
    'RaiseEvent RsErr("No puedo recuperar los datos de fuction fGetF.", "fGetF en clsEvents")
    'R.msgB App.EXEName, "fGetF en clsEvents", Err.Description, enClass
End Function
Public Sub LetField(Campo_TextBox_para_el_Field As TextBox, _
    Nombre_del_Field_en_la_tabla As Variant)
    On Error GoTo N
    'Se supone que esto es lo mismo que guardar pero guarda field por field.
    'debe usarse el evento guardar o sendfields.
    Records.Update
    Records.Fields(Nombre_del_Field_en_la_tabla) = Campo_TextBox_para_el_Field
    RaiseEvent Actualizando
    Exit Sub
N:
    'RaiseEvent RsErr("No puedo realizar LetField. LetField(Guardar)", "Let field en clsevents")
End Sub


Public Sub LetF(Datos As Variant, _
    Nombre_del_Field_en_la_tabla As Variant)
    On Error GoTo N
    'Se supone que esto es lo mismo que guardar pero guarda field por field.
    'debe usarse el evento guardar o sendfields.
    Records.Update Nombre_del_Field_en_la_tabla, Datos
    'Si no funciona habilitar este otro.
    Records.Fields(Nombre_del_Field_en_la_tabla) = Datos
    Exit Sub
N:
    MsgBox Err.Description, , "LetF"
    'RaiseEvent RsErr("No puedo realizar LetField. LetField(Guardar)", "Let field en clsevents")
End Sub
Public Sub LetFieldNuevoRegistro(Nombre_del_Field_en_la_tabla As Variant, Optional Datos As Variant = 0)
    On Error GoTo N
    'Se supone que esto es lo mismo que guardar pero guarda field por field.
    'debe usarse el evento guardar o sendfields.
    Records.Update
    Records.Fields(Nombre_del_Field_en_la_tabla) = Datos
    Exit Sub
N:
    'RaiseEvent RsErr("No puedo realizar LetField. LetField(Guardar)", "Let field en clsevents")
End Sub

Public Sub LetFieldComboBox(Combo As ComboBox, _
    Nombre_del_Field_en_la_tabla As Variant)
    On Error GoTo N
    'Se supone que esto es lo mismo que guardar pero guarda field por field.
    Records.Update
    Records.Fields(Nombre_del_Field_en_la_tabla) = Combo.Text
    Exit Sub
N:
    'RaiseEvent RsErr("No puedo realizar LetField. LetField(Guardar)", "Let field en clsevents")
End Sub

Public Sub FieldActualizar(Campo_TextBox_para_el_Field As TextBox, _
    Nombre_del_Field_en_la_tabla As Variant)
    'Actualiza los datos en la tabla, basicamente es lo mismo que recuperar datos _
    pero con una pequeña diferencia, en este caso se envian _
    datos a la tabla.
    Records.Fields(Nombre_del_Field_en_la_tabla) = Campo_TextBox_para_el_Field
End Sub

Public Sub GuardandoFieldPorField(Campo_TextBox_para_el_Field As TextBox, _
    Nombre_del_Field_en_la_tabla As Variant)
    'Se supone que esto es lo mismo que guardar pero guarda field por field.
    Records.Update
    Records.Fields(Nombre_del_Field_en_la_tabla) = Campo_TextBox_para_el_Field
End Sub

Public Sub CmbActivate(Combo As ComboBox)
    'Activa un ComboBox
    Tmp = SendMessage(Combo.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Public Sub Cerrar()
    On Error GoTo N
    Records.Close
    Conect.Close
    'Set Records = Nothing
    'Set Conect = Nothing
    Exit Sub
N:
    'RaiseEvent RsErr(Err.Description, "Cerrar en clsEvents")
    'R.msgB App.EXEName, "Cerrar en clsEvents", Err.Description, enClass
End Sub

Public Function fInterrogacion()
    fInterrogacion = "¿?"
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

Public Sub GetBase()
    LoadDBs Tabla
End Sub

Public Property Get prQueCampo() As Single
    prQueCampo = vprQueCampo
End Property
Public Property Let prQueCampo(ByVal vNewValue As Single)
    vprQueCampo = vNewValue
End Property



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

Public Sub LoadDBsAlternate(Nombre_de_la_tabla)
    'Conecta con una base de datos
    Dim X As Variant
    On Error GoTo N
    X = GetSetting(App.Path, "db", "base")
    If X = "" Then
        X = InputBox("Escriba o pegue la direccion de la base de datos Notas.mdb.", , "Notas" & ".mdb")
        SaveSetting App.Path, "db", "base", X
    End If
    Conect.Provider = "Microsoft.Jet.OLEDB.4.0"
    'Aqui establesco la coneccion con la base de datos
    Conect.ConnectionString = X '"D:\Proyectos\Alternativo/Notas.mdb"
    'Habre la conexion
    Conect.Open
    'Aqui se obtiene una serie de registros: en este
    'caso de la
    'tabla
    Records.Open "select * from " & Nombre_de_la_tabla, Conect, , adLockOptimistic
    'Siguiente
    'Anterior
    RaiseEvent Actualizando
    'MsgBox "La tabla " & Nombre_de_la_tabla & " esta Conectada.", vbInformation
    Exit Sub
N:
    X = InputBox("Escriba o pegue la direccion de la base de datos Notas.mdb.", , "Notas" & ".mdb")
    SaveSetting App.Path, "db", "base", X
    MsgBox "Reinicie la base de datos si proporciono la direccion correcta.", vbInformation
End Sub

Public Sub AddToListView(Columnas As Variant, ListViewNombre As ListView, Datos As Variant, Optional Icono As Long = 1, Optional SmallIcono As Long = 1)
    'He durado muchos dias perfeccionando este codigo, ya lo logré.
    'Ademas se puede amplificar y basado en esto se puede usar de diferente manera.
    'Lo mío era para bases de datos.
    On Error Resume Next
    
    Static V
    Set ITM = ListViewNombre.ListItems.Add(, , Datos)
    
    For V = 1 To Columnas '+ Val(1)
        DoEvents
        
        With ITM
            .Icon = Icono   ' Establece un icono de ImageList1.
            .SmallIcon = SmallIcono
            .SubItems(V - Val(1)) = fGetF(ListViewNombre.ColumnHeaders.Item(V).Text) 'K.fGetF(ListViewNombre.ColumnHeaders.Item(V).Text)
        End With
        
    Next V
    
    V = 0
    Set ITM = Nothing
    Exit Sub
N:
    Set ITM = Nothing
    MsgBox Err.Description, , "addtoListView"
End Sub

Public Function ListViewColumnasCount(ListViewNombre As ListView)
    ListViewColumnasCount = ListViewNombre.ColumnHeaders.Count
End Function

Public Function DBDir(Direccion As String) As String
    DBDir = Direccion
End Function

Public Sub sSentNewDatos(ParamArray Lista_de_Fields() As Variant) 'Se crea un nuevo registro por medio de una lista de campos, sin necesidad de llamar una y otra vez al Sub para cada campo.
    On Error GoTo AccionesCorrectivas
    Dim vField As Variant
    
    For Each vField In Lista_de_Fields
        LetFieldNuevoRegistro vField
    Next vField
    
    Exit Sub
AccionesCorrectivas:
    MsgBox "Tengo problemas con SentNewDatos"
End Sub



