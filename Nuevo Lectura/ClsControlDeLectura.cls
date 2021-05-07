VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsControlDeLectura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vBackUpGuardado As Boolean 'Variable para uso de la propiedad prBackUpGuardado


Private vNombreAntiguo As Variant 'Variable para uso de la propiedad prNombreAntiguo


Private mDiezCargas(1 To 10) As Variant 'Copia local de las 10 cargas de texto.


Private vCreandoNuevoRegistro As Boolean 'Variable para uso de la propiedad prCreandoNuevoRegistro


Private vDepurandoTexto_ultimo_texto As Variant 'Variable para uso de la propiedad prDepurandoTextoUltimoTexto


Private vLen_incremental_del_texto As Double 'Variable para uso de la propiedad prLenIncrementalDelTexto


Public Enum enBackColor 'Se usa en sBackColor.
Verde_claro = &H80FF80   '&H8000000E
Normal = &H8000000F
Amarillo = &HFFFF&
End Enum

'Creado a finales de julio del 2015 por Rafael Angel Montero Fernández.
'La intencion de este Class es reducir la cantidad de lineas de codigos de otros controles y objetos.
'Se pretende reflejar en cada objeto o control solo lo más expecifico y lo que se pueda generalizar se coloca en este Class.



Public Event evUpMenu(ByRef Menu As Object) 'Sucede al llamar al procedimiento sUpMenu.

'Para los colores de los cuadros de texto
Public Enum Colores_Txt
    entxt_Sin_enfoque = &H80000005
    entxt_Enfocado = 9034953 '&H80000013
End Enum

Public Sub InformacionDesconocida(Control As Object)

If Control.Text = "" Then
    Control.Text = 0
End If

End Sub


Public Sub sArrastrarUnArchivo(Data As Object, TextBox_para_la_direccion As Object, Control_para_el_nombre_del_archivo As Object) 'As Variant 'Para recuperar la direccion del archivo.
On Error GoTo AccionesCorrectivas
TextBox_para_la_direccion.Text = Data.Files.Item(1)
Control_para_el_nombre_del_archivo.Text = fNombreDelArchivo(Data.Files.Item(1))
Exit Sub
AccionesCorrectivas:
'MsgBox "Tengo problemas con sArrastrarUnArchivo"
End Sub


Public Function fNombreDelArchivo(Direccion_de_la_carpeta As Variant) As Variant  'Nombre del video que se esta reproducioendo. Es de solo lectura.
'On Error GoTo AccionesCorrectivas
Const cntSlash = "\"

If Direccion_de_la_carpeta = "" Then
    Exit Function
End If

Dim vNombre As Variant, mNombre() As String

mNombre = Split(Direccion_de_la_carpeta, cntSlash) 'Se obtiene el nombre con su extencion de archivo.
vNombre = mNombre(UBound(mNombre)) 'Se carga en la variable.

mNombre = Split(vNombre, ".") 'Despues se obtien el nombre sin su extencion de archivo.
vNombre = mNombre(LBound(mNombre)) 'Se carga la nueva informacion en la variable.
fNombreDelArchivo = vNombre 'Se retorna el nombre sin su extencion.

'Exit Function
'AccionesCorrectivas:
'MsgBox Err.Description
'MsgBox "Tengo problemas con prNombreDelVideo"
End Function

Public Function fToolTip(ByVal Numero_del_1_al_100 As Double, ByRef Max_del_valor_real As Double, ByRef Valor_real_obtenido As Double)   'sprgVolumen.ToolTipText = "El " & Valor & "% de " & cPorcentajes.prMaximoDelValorReal & " es " & cPorcentajes.prValorReal
On Error GoTo AccionesCorrectivas
fToolTip = "El " & Numero_del_1_al_100 & "% de " & Max_del_valor_real & " es " & Valor_real_obtenido
'Slider1.ToolTipText = ProgressBar1.ToolTipText
Exit Function
AccionesCorrectivas:
MsgBox "Tengo problemas con sToolTip"
End Function

Public Sub PonerIcono(vControl As Object, Optional id_del_icono As Integer = 104) 'Pone un icono en cada control que tenga la propiedad Icon.
vControl.MouseIcon = LoadResPicture(id_del_icono, vbResIcon)
End Sub

Public Sub sNormalizarColor(ParamArray Controles() As Variant)
'Hace que el color de los labels regrese a ser como el del control ActiveX
On Error GoTo N
Dim vControl As Object 'Label  'Puntero al id del array, para poderlo manipular.
Dim i As Integer

For i = LBound(Controles) To UBound(Controles)
    DoEvents
    Set vControl = Controles(i)
    vControl.BackColor = &H8000000F
    Set vControl = Nothing
Next i

'sBackColor Normal, Controles'No funciona
Exit Sub
N:

End Sub


Private Function fVarTipo(Variable_nombre As Variant)
'Comprueva la variable.
fVarTipo = TypeName(Variable_nombre)
End Function

Public Sub sHabilitarControles(ParamArray Controles() As Variant) 'Habilita una lista de botones y controles.
On Error GoTo AccionesCorrectivas

Dim i As Integer
Dim vBoton As Object 'Esto funciona como una especie de puntero al control, ya que no se puede hacer directamente desde el array.

For i = LBound(Controles) To UBound(Controles)
    DoEvents
    Set vBoton = Controles(i)
    vBoton.Enabled = True
    Set vBoton = Nothing
Next i

Exit Sub
AccionesCorrectivas:

MsgBox Err.Description '"Tengo problemas con sHabilitarBotones"
End Sub


Public Sub sDeshabilitarControles(ParamArray Controles() As Variant) 'Deshabilita una lista de botones y controles.
On Error GoTo AccionesCorrectivas

Dim i As Integer
Dim vBoton As Object 'Esto funciona como una especie de puntero al control, ya que no se puede hacer directamente desde el array.

For i = LBound(Controles) To UBound(Controles)
    DoEvents
    Set vBoton = Controles(i)
    vBoton.Enabled = False
    Set vBoton = Nothing
Next i

Exit Sub
AccionesCorrectivas:

MsgBox Err.Description '"Tengo problemas con sHabilitarBotones"
End Sub

Public Sub sProgressBarProgresoDeLaLectura(ByRef Nombre_del_ProgressBar As Object, ByRef Nombre_del_segundo_ProgressBar As Object, ByRef Nombre_del_RichTextBox_o_TextBox As Object) 'Mide el progreso de la barra de progreso o ProgressBar.
'
End Sub



Public Sub sEstadoDelProgressBar(ByVal Len_de_la_totalidad_del_texto As Double, ByRef Nombre_del_ProgressBar As Object, ByRef Nombre_del_segundo_ProgressBar As Object) 'Valora el estado del ProgressBar.
'
End Sub

Public Sub sBackColorDelTextBox(Nombre_del_TextBox As Object, Color As Colores_Txt)
'Para cambiar el color de los campos de texto.
On Error GoTo N
 Nombre_del_TextBox.BackColor = Color
Exit Sub
N:
MsgBox Err.Description, vbExclamation
End Sub

Public Function fGetTextoDelPortapapeles(ByRef Control_con_una_propiedad_Text As Object) As Double 'Carga al RichTextBox con nuevos datos sin borrar los anteriores. Ademas devuelve el Len de todo el texto.
On Error GoTo AccionesCorrectivas

Dim vLen As Double 'El uso de variables locales evita la sobrecarga. Y así Len(RichTextBox1.Text) solo se carga una vez.
vLen = Len(Control_con_una_propiedad_Text.Text) 'Se mide al principio y al final despues de cargar los datos.

On Error GoTo N

If Control_con_una_propiedad_Text.Text = "" Or Control_con_una_propiedad_Text.Text = "0" Then
    
    Control_con_una_propiedad_Text.Text = Clipboard.GetText
Else

N:
    On Error Resume Next 'No es muy correcto usar On Error Resume Next, es mejor usar On Error GoTo N; pero por el momento usaré esta linea aquí.
    
        
        Control_con_una_propiedad_Text.SelStart = vLen 'Va al final del texto.
        Control_con_una_propiedad_Text.SelText = RTC & RTC & Clipboard.GetText  'Inserta la informacion que esta en el portapapeles.
        
        sSeltxtColor Control_con_una_propiedad_Text 'Ahora si se puede usar un TextBox o un RichTextBox, el error lo agarra el procedimiento sSeltxtColor sin afectar a esta funcion (fGetTextoDelPortapapeles.
 
End If 'Control_con_una_propiedad_Text.Text

Clipboard.Clear 'Se limpia el portapapeles.

vLen = Len(Control_con_una_propiedad_Text.Text) 'Se mide al principio y al final despues de cargar los datos.


 fGetTextoDelPortapapeles = vLen
prLenIncrementalDelTexto = prLenIncrementalDelTexto + vLen
vLen = 0
Exit Function
AccionesCorrectivas:
MsgBox "Tengo problemas con fGetTextoDelPortapapeles"
End Function


Public Function fDuracionDeLaLectura(ByVal Len_del_texto As Double, ByRef Tiempo_almacenado As Double) 'Permite calcular la cantidad de tiempo que podria durar la lectura. No es preciso pero se acerca bastante.
On Error GoTo AccionesCorrectivas

'Esto permite saber los segundos, minutos y horas que tardará el libro en ser leido.
    Dim vSegundos As Integer, vMinutos As Double, vHoras As Double, vMenosDeUnaHora As Variant
    Dim vRespuesta As Variant


If Len_del_texto = 0 Or Tiempo_almacenado = 0 Then 'No se deven permitir diviciones por cero.
    MsgBox "Divicion por cero, algo no permitido.", vbExclamation
    Exit Function
End If

    vSegundos = Len_del_texto / Tiempo_almacenado
    
    vMinutos = vSegundos / 60 'Cantidad de minutos
    
    If vMinutos > 60 Then
        vHoras = vMinutos / 60 'Cantidad de horas
        vRespuesta = "Tiempo de lectura transcurido en segundos: " & Tiempo_almacenado & RTC & "Cantidad de minutos que podria durar la lectura: " & vMinutos & RTC & "Cantidad de horas que podria durar la lectura: " & vHoras

    Else
        vMenosDeUnaHora = "Menos de una hora."
        vRespuesta = "Tiempo de lectura transcurido en segundos: " & Tiempo_almacenado & RTC & "Cantidad de minutos que podria durar la lectura: " & vMinutos & RTC & "Cantidad de horas que podria durar la lectura: " & vMenosDeUnaHora

    End If
    
MsgBox vRespuesta, vbInformation

fDuracionDeLaLectura = vRespuesta

Exit Function
AccionesCorrectivas:
MsgBox Err.Description, vbExclamation, "fDuracionDeLaLectura" '"Tengo problemas con fDuracionDeLaLectura"
End Function


Public Function fMaxDeLosProgressBars(ByVal Len_del_texto As Double, ParamArray controles_ProgressBar() As Variant)  'Permite obtener el maximo de cada ProgressBar y ademas devuelve el maximo.
On Error GoTo N
Dim vPrg As Object, i As Long

For i = LBound(controles_ProgressBar) To UBound(controles_ProgressBar)
    DoEvents
    Set vPrg = controles_ProgressBar(i)
    vPrg.Max = Len_del_texto
    Set vPrg = Nothing
Next i

N:
MsgBox Err.Description, vbExclamation, "Problemas con fMaxDeLosProgressBars"
End Function


Public Sub sPegarTextoNumeroDeVeces(ByRef Control_RichTextBox As Object) ', ByRef Control_ProgressBar As Object) 'Pega un texto un cierto numero de veces; el maximo es 38.
On Error GoTo AccionesCorrectivas

Dim NVeces As Variant
NVeces = InputBox("¿Cuantas veces desea pegar el mismo texto?" & RTC & "Esto es por si tiene que leer una sierta cantidad de veces el texto por ejemplo para aplicar PNL (Programacion Neuroligüistica." & RTC & "Recuerde que el texto se copia automaticamente desde el portapapeles.", , 3)

If IsNumeric(NVeces) = True Then

 If NVeces > 38 Then MsgBox "El maximo es 38 veces, se ajustará a 38.", vbInformation: NVeces = 38
    
    If NVeces = 0 Then
NoPuedoPegarElTexto:
        MsgBox "No puedo pegar el texto porque no anotó una cantidad de veces o puso cero (0).", vbInformation
    Else
        Dim Nv As Long
        
        For Nv = 1 To NVeces
            DoEvents
            Control_RichTextBox.SelText = Clipboard.GetText & RTC & RTC
        Next
        
    End If
    
    'Dim vActualizarMax As Long
    
   ' vActualizarMax = Len(Control_RichTextBox.Text)
    
    'If vActualizarMax = 0 Then vActualizarMax = 1
    
    'Control_ProgressBar.Max = vActualizarMax 'Para actualizar el maximo.
    'Control_ProgressBar.Value = Control_ProgressBar.Max 'Para darle el maximo valor al progress.
Else
    GoTo NoPuedoPegarElTexto
End If


Exit Sub
AccionesCorrectivas:

MsgBox "Tengo problemas con sPegarTextoNumeroDeVeces"
End Sub


Public Sub sUpMenu(ByRef Menu As Object, ByVal Boton_clickeado As Integer) 'Levanta un menu emergente.
On Error GoTo AccionesCorrectivas

If Boton_clickeado = 2 Then
    RaiseEvent evUpMenu(Menu)
End If

Exit Sub
AccionesCorrectivas:

MsgBox "Tengo problemas con sUpMenu"
End Sub


Public Sub sCopiarTextoSeleccionado(ByRef Control_RichTextBox As Object, Nombre_del_campo As Variant) 'Copia el texto que se ha seleccionado en la ventana de texto de un control RichTextBox.
On Error GoTo AccionesCorrectivas

If Control_RichTextBox.SelText = "" Then
    MsgBox "Seleccione algun texto del campo " & Nombre_del_campo
Else
    Clipboard.SetText Control_RichTextBox.SelText
End If

Exit Sub
AccionesCorrectivas:

MsgBox "Tengo problemas con sCopiarTextoSeleccionado"
End Sub


Public Sub sQuitarEspaciosAlSelText(ByRef Control_con_una_propiedad_SelText As Object) 'Quita todos los espacios del texto. Esto es util ocasionalmente...
On Error GoTo AccionesCorrectivas

Dim SiNo As Single
Dim R As New ClsRecursos
SiNo = MsgBox("¿Desea quitar los espacios inecesarios entre letras y palabras?", vbQuestion + vbYesNo)

If SiNo = vbYes Then
    Control_con_una_propiedad_SelText.SelText = R.fReemplazarTexto(Control_con_una_propiedad_SelText.SelText, " ", "")  'Reemplaza los espacios.
End If

Exit Sub
AccionesCorrectivas:

MsgBox "Tengo problemas con sQuitarEspaciosAlSelText"
End Sub


Public Sub sTextoInicial(ByRef Nombre_del_primer_RichTextBox As Object, ByRef Nombre_del_segundo_RichTextBox As Object, ByRef Nombre_del_libro As Variant) 'Carga el texto inicial en el segundo control.
On Error GoTo AccionesCorrectivas

Nombre_del_segundo_RichTextBox.Text = Nombre_del_primer_RichTextBox.Text 'RichTextBoxInicial toma el valor de RichTextBox1.

If Nombre_del_segundo_RichTextBox.Text = "" Then Nombre_del_segundo_RichTextBox.Text = 0

    Nombre_del_segundo_RichTextBox.SaveFile App.Path & "\" & Nombre_del_libro & " todo el texto.TXTLeido"

Exit Sub
AccionesCorrectivas:

MsgBox "Tengo problemas con sTextoInicial"
End Sub


Public Function fPreparandoElProgressBarParaNuevaLectura(ByRef Control_RichTextBox As Object, ByRef Nombre_del_ProgressBar As Object) As Integer 'Realiza algunas preparaciones al iniciarse la lectura, como ajustar el progressbar vertical...
'
End Function


Public Sub sEliminarLibro(ByRef Direccion_del_libro As Variant, ByRef Titulo_del_libro As Variant) 'Borra un libro o un archivo...


On Error GoTo N
   Kill App.Path & "/" & Titulo_del_libro & ".Lectura"
Exit Sub
N:
sEliminarLibro2 Direccion_del_libro, Titulo_del_libro
On Error GoTo AccionesCorrectivas
   Kill Direccion_del_libro & ".Lectura"

Exit Sub
AccionesCorrectivas:

MsgBox "Tengo problemas con sEliminarLibro"
End Sub

Public Sub sEliminarLibro2(ByRef Direccion_del_libro As Variant, ByRef Titulo_del_libro As Variant) 'Borra un libro o un archivo...

On Error Resume Next ' A veces si es valioso usar On Error Resume Next.
   Kill Direccion_del_libro & ".Lectura"
End Sub

Public Function fPreparandoProgressBarParaActualizarDatos(ByRef ProgressBar_1 As Object, ByRef ProgressBar_2 As Object, ByVal Len_de_todo_el_texto As Double, ByVal Len_de_la_lectura_en_progreso As Double) As Double 'Prepara los dos ProgressBar durante la actualizacion de datos.
'
End Function


Public Sub sPoniendoCeroAlText(ParamArray Control_con_una_propiedad_Text() As Variant) 'Le pone un cero a los controles que tengan una propiedad Text.
On Error GoTo AccionesCorrectivas

Dim vControl As Object 'Label  'Puntero al id del array, para poderlo manipular.
Dim i As Integer

For i = LBound(Control_con_una_propiedad_Text) To UBound(Control_con_una_propiedad_Text)
    DoEvents
    Set vControl = Control_con_una_propiedad_Text(i)
    If vControl.Text = "" Then vControl.Text = 0
    Set vControl = Nothing
Next i



Exit Sub
AccionesCorrectivas:

MsgBox "Tengo problemas con sPoniendoCeroAlText"
End Sub


'Public Sub sComprobarSiHayDatosEnElPortapapeles(ByRef Control_que_mostrará_el_icono As Object) 'Comprueva si hay datos en el portapapeles y de acuerdo a eso, carga un cursor verde o rojo.
'On Error GoTo AccionesCorrectivas
'
'Dim vDatos As Variant 'Para comprobar si hay o no hay datos.
'vDatos = Clipboard.GetText 'Agarro los datos del portapapeles.
'
'If vDatos = "" Then
'    Control_que_mostrará_el_icono.MouseIcon = LoadIcon(Bola_roja) 'Cargo un icono color rojo.
'    Control_que_mostrará_el_icono.Tag = ""
'Else
'    Control_que_mostrará_el_icono.MouseIcon = LoadIcon(Bola_verde) 'Cargo un icono verde.
'    Control_que_mostrará_el_icono.Tag = "Datos"
'End If
'
'Exit Sub
'AccionesCorrectivas:
'
'MsgBox "Tengo problemas con sComprobarSiHayDatosEnElPortapapeles"
'End Sub


Public Function fProgressBarParaLosUltimosDatosLeidos(ByRef PorgressBar As Object, ByRef PorgressBar_2 As Object, ByRef RichTextBox_control As Object, Len_de_la_lectura_en_progreso As Double) As Double 'Realiza tareas durante la activacion de la lectura de los ultimos datos, ademas, devuelve el maximo de todo el texto o el Len.
'
End Function


Public Sub sProgressBarToolTipText(ByRef ProgressBar_nombre As Object, ByRef Caption_de_la_lectura As Variant, Len_de_todo_el_texto As Double) 'Muestra un ToolTipText con informacion del estado del progressbar y el texto de algun campo que acepte texto.
ProgressBar_nombre.ToolTip = Caption_de_la_lectura & "=" & Len_de_todo_el_texto
End Sub


Public Sub sMaxLenDeUnTextBox(ByRef TextBox_nombre As Object, ByRef ProgressBar_nombre As Object) 'Hace tareas relacionadas con el maximo de datos permitido en un TextBox.
On Error GoTo AccionesCorrectivas
'Se usa en Text1_Change.

 ' If TextBox_nombre = "" Then
    'Se realiza alguna tarea relacionada con datos en blanco, aunque por ahora no es necesario.
  'End If
  
If Len(TextBox_nombre.Text) > ProgressBar_nombre.Max Then 'Este debe pasarse a clsaControles
    Dim SiNo As Variant
    TextBox_nombre.SetFocus
    SiNo = MsgBox("No puedo leer más de 65000 caracteres. Los datos serán recortados a 500 caracteres. ¿Desea recortarlos o solo seleccionarlos?", vbYesNo)
    
    TextBox_nombre.SelStart = 0
    TextBox_nombre.SelLength = ProgressBar_nombre.Max
    
    Select Case SiNo
    Case vbYes
        TextBox_nombre.Text = TextBox_nombre.SelText
    Case vbNo
        'No se hace nada.
    End Select
    
Else

    If TextBox_nombre.Text = "" Then
        ProgressBar_nombre.Value = Val(0)
    Else
        ProgressBar_nombre.Value = Len(TextBox_nombre.Text)
    End If
    
End If

Exit Sub
AccionesCorrectivas:

MsgBox "Tengo problemas con sMaxLenDeUnTextBox"
End Sub


Public Sub sBackColor(Elija_un_color As enBackColor, ParamArray Controles() As Variant) 'Para cambiar el color del o de los controles.
On Error GoTo AccionesCorrectivas
Dim vControl As Object, i As Integer

For i = LBound(Controles) To UBound(Controles)
    DoEvents
    Set vControl = Controles(i)
    vControl.BackColor = Elija_un_color
    Set vControl = Nothing
Next i

Exit Sub
AccionesCorrectivas:
MsgBox Err.Description '"Tengo problemas con BackColor"
End Sub


Public Function fGuardarLibro(ByRef Lectura_en_progreso As Object, BackUp_del_libro As Object, Nombre_del_libro As Variant) 'Guarda la lectura en progreso y el texto BackUp
On Error GoTo AccionesCorrectivas
Lectura_en_progreso.SaveFile App.Path & "\" & Nombre_del_libro & ".Lectura"
BackUp_del_libro.SaveFile App.Path & "\" & Nombre_del_libro & ".txtLeido"
fGuardarLibro = App.Path & "\" & Nombre_del_libro & ".Lectura"
Exit Function
AccionesCorrectivas:
MsgBox "Tengo problemas con GuardarLibro"
End Function


Public Sub sAbrirLecturaEnProgreso(ByRef RichTextBox_nombre As Object, Nombre_del_libro As Variant) 'Abre el libro que se esta leyendo.
On Error GoTo AccionesCorrectivas
RichTextBox_nombre.LoadFile App.Path & "\" & Nombre_del_libro & ".Lectura"   'Abre el archivo
Exit Sub
AccionesCorrectivas:
MsgBox "Tengo problemas con AbrirLecturaEnProgreso"
End Sub


Public Sub sAbrirBackUp(ByRef RichTextBox_nombre As Object, Nombre_del_libro As Variant) 'Abre el BackUp del libro que se esta leyendo.
On Error GoTo AccionesCorrectivas
RichTextBox_nombre.Text = ""
RichTextBox_nombre.LoadFile App.Path & "\" & Nombre_del_libro & ".txtLeido"
Exit Sub
AccionesCorrectivas:
MsgBox "Tengo problemas con AbrirBackUp"
End Sub


Public Property Get prLenIncrementalDelTexto() As Double 'Se carga con el Len del texto copiado del portapapeles con el Function fGetTextoDelPortapapeles.

On Error GoTo AccionesCorrectivas

prLenIncrementalDelTexto = vLen_incremental_del_texto
Exit Property
AccionesCorrectivas:
MsgBox "Tengo problemas con prLenIncrementalDelTextoprLenIncrementalDelTexto"
End Property
Public Property Let prLenIncrementalDelTexto(vNuevosDatos As Double)  'Se carga con el Len del texto copiado del portapapeles con el Function fGetTextoDelPortapapeles.
On Error GoTo AccionesCorrectivas

vLen_incremental_del_texto = vNuevosDatos
Exit Property
AccionesCorrectivas:
MsgBox "Tengo problemas con prLenIncrementalDelTextoprLenIncrementalDelTexto"
End Property


Public Function fDepurandoTexto(ByRef Datos As Variant, Libro As Object) 'Depura el texto o los datos para que termine en una palabra competa.
On Error GoTo AccionesCorrectivas

If Datos = "" Then 'Comprueba que los datos estan en blanco, de estarlo cancela la ejecucion del resto de la funcion para evitar un error.
    MsgBox "Se acabó el libro", vbInformation
    Exit Function
End If

Dim mDepuracion() As String, vLen As Double, vRespuesta As Variant

mDepuracion = Split(Datos, " ") 'Se insertan los datos en la matriz.

vRespuesta = Join(mDepuracion, " ") 'Se vuelven a poner en una variable. Pero quedan tambien en la matriz.

vLen = Len(vRespuesta) - Len(mDepuracion(UBound(mDepuracion))) 'Se usan los datos de la variable y el ultimo id de la matriz para sacar el Len de texto que se devolverá como respuesta.

If Right(Libro.Text, Len(vRespuesta)) = vRespuesta Then
    fDepurandoTexto = prDepurandoTextoUltimoTexto & vRespuesta 'Se devuelve todo por ser el ultimo texto del libro.
    'prDepurandoTextoUltimoTexto = "" 'La propiedad se borra porque ya no es necesario recordar los ultimos datos.
Else
    fDepurandoTexto = prDepurandoTextoUltimoTexto & Left(vRespuesta, vLen) 'Se regresa todo lo que este a la izquierda menos el valor de vLen.
    prDepurandoTextoUltimoTexto = mDepuracion(UBound(mDepuracion)) 'Se guarda el ultimo id de la matriz para regresarlo como parte de la respuesta de la proxima llamada...
'MsgBox prDepurandoTextoUltimoTexto & Left(vRespuesta, vLen)
End If

Exit Function
AccionesCorrectivas:
'MsgBox Err.Description '"Tengo problemas con DepurandoTexto"
End Function


Public Property Get prDepurandoTextoUltimoTexto() As Variant 'Guarda el ultimo id de la matriz local mDepuracion, en la funcion fDepurandoTexto, para luego concatenarlo como parte de la respuesta de la siguiente llamada de la funcion.

On Error GoTo AccionesCorrectivas

prDepurandoTextoUltimoTexto = vDepurandoTexto_ultimo_texto
Exit Property
AccionesCorrectivas:
MsgBox "Tengo problemas con prDepurandoTextoUltimoTextoprDepurandoTextoUltimoTexto"
End Property
Public Property Let prDepurandoTextoUltimoTexto(vNuevosDatos As Variant)  'Guarda el ultimo id de la matriz local mDepuracion, en la funcion fDepurandoTexto, para luego concatenarlo como parte de la respuesta de la siguiente llamada de la funcion.
On Error GoTo AccionesCorrectivas

vDepurandoTexto_ultimo_texto = vNuevosDatos
Exit Property
AccionesCorrectivas:
MsgBox "Tengo problemas con prDepurandoTextoUltimoTextoprDepurandoTextoUltimoTexto"
End Property


Public Function fGuardarLecturaEnProgreso(ByRef RichTextBox_nombre As Object, ByRef Nombre_del_libro As Variant) 'Guarda la lectura en progreso.
On Error GoTo AccionesCorrectivas
RichTextBox_nombre.SaveFile App.Path & "\" & Nombre_del_libro & ".Lectura"
fGuardarLecturaEnProgreso = App.Path & "\" & Nombre_del_libro & ".Lectura"
Exit Function
AccionesCorrectivas:
MsgBox "Tengo problemas con GuardarLecturaEnProgreso"
End Function


Public Function fGuardarBackUp(ByRef RichTextBox_nombre As Object, ByRef Nombre_del_libro As Variant) 'Guarda la copia de seguridad del libro.
On Error GoTo AccionesCorrectivas
RichTextBox_nombre.SaveFile App.Path & "\" & Nombre_del_libro & ".txtLeido"
fGuardarBackUp = App.Path & "\" & Nombre_del_libro & ".txtLeido"
Exit Function
AccionesCorrectivas:
MsgBox "Tengo problemas con GuardarBackUp"
End Function


Public Property Get prCreandoNuevoRegistro() As Boolean 'True se esta creando un registro y False no se esta creando un registro.

On Error GoTo AccionesCorrectivas

prCreandoNuevoRegistro = vCreandoNuevoRegistro
Exit Property
AccionesCorrectivas:
MsgBox "Tengo problemas con prCreandoNuevoRegistroprCreandoNuevoRegistro"
End Property
Public Property Let prCreandoNuevoRegistro(vNuevosDatos As Boolean)  'True se esta creando un registro y False no se esta creando un registro.
On Error GoTo AccionesCorrectivas

vCreandoNuevoRegistro = vNuevosDatos
Exit Property
AccionesCorrectivas:
MsgBox "Tengo problemas con prCreandoNuevoRegistroprCreandoNuevoRegistro"
End Property


Public Sub sSeltxtColor(ByRef RichTextBox_nombre As Object) 'Hace que el texto seleccionado se vuelva negro.
On Error GoTo AccionesCorrectivas

With RichTextBox_nombre
    .SelStart = 0
    .SelLength = Len(RichTextBox_nombre.TextRTF)
    .SelColor = vbBlack
    .SelFontSize = 12
    .SelLength = 0
End With

Exit Sub
AccionesCorrectivas:
MsgBox "Tengo problemas con SeltxtColor"
End Sub


Public Function fCantidadDePalabras(ByRef Texto As Variant) 'Cuenta y regresa el total de palabras en su texto.
On Error GoTo AccionesCorrectivas
fCantidadDePalabras = CantidadDePalabras(Texto)
Exit Function
AccionesCorrectivas:
MsgBox "Tengo problemas con CantidadDePalabras"
End Function


Public Sub sCopiarSelText(ByRef RichTextBox_nombre As Object) 'Copia texto desde el control hacia el portapapeles.
On Error GoTo AccionesCorrectivas
Clipboard.SetText RichTextBox_nombre.SelText
Exit Sub
AccionesCorrectivas:
MsgBox "Tengo problemas con CopiarSelText"
End Sub


Public Sub sAddDiezCargas(ByRef Datos As Variant) 'Agrega 10 fracciones de texto a una matriz que esta en un modulo standar.
On Error GoTo AccionesCorrectivas

Static vConteo As Integer 'Para el conteo de los indices de la matriz.

     
     vConteo = vConteo + 1
     
     If vConteo > 10 Then
        vConteo = 1
     End If

mDiezCargas(vConteo) = "texto numero " & vConteo & "= " & Datos
Exit Sub
AccionesCorrectivas:
MsgBox "Tengo problemas con AddDiezCargas"
End Sub


Public Sub sMostrarUltimas10(Optional Titulo_del_libro As Variant = "Libro") 'Carga un formulario con las ultimas 10 cargas de texto.
'On Error GoTo AccionesCorrectivas
'
'Dim vDiez As Variant
'
'With frmUltimas10
'    .Caption = .Tag & Titulo_del_libro
'    .List1.Clear
'
'    For Each vDiez In mDiezCargas 'vDiez = LBound(vTyLecturas.UltimosDiezSelText) To vTyLecturas.ContarDiezCargas 'UBound(vTyLecturas.UltimosDiezSelText)
'        DoEvents
'
'        If vDiez = "" Then 'vTyLecturas.UltimosDiezSelText(vDiez) = "" Then
'            'No hay texto.
'        Else
'            .List1.AddItem vDiez 'vTyLecturas.UltimosDiezSelText(vDiez)
'        End If
'
'    Next vDiez
'
'    .Show
'End With
'
'vDiez = ""
''vDiez = 0
'
'Exit Sub
'AccionesCorrectivas:
'MsgBox "Tengo problemas con MostrarUltimas10"
End Sub


Public Property Get prNombreAntiguo() As Variant 'Guarda el nombre antiguo de algun libro. Se usa cuando se cambia el nombre de un archivo mediante la interfaz de usuario.
On Error GoTo AccionesCorrectivas

prNombreAntiguo = vNombreAntiguo
Exit Property
AccionesCorrectivas:
MsgBox "Tengo problemas con prNombreAntiguoprNombreAntiguo"
End Property
Public Property Let prNombreAntiguo(vNuevosDatos As Variant)  'Guarda el nombre antiguo de algun libro. Se usa cuando se cambia el nombre de un archivo mediante la interfaz de usuario.
On Error GoTo AccionesCorrectivas

vNombreAntiguo = vNuevosDatos
Exit Property
AccionesCorrectivas:
MsgBox "Tengo problemas con prNombreAntiguoprNombreAntiguo"
End Property


Public Property Get prBackUpGuardado() As Boolean 'True el backUp ya fue guardado por medio de un procedimiento alterno y false no se ha guardado en forma alternativa. Esta propiedad se usará para activarse cuando el nombre del libro cambia y se guarda la copia de seguridad por medio de codigo alternativo.
On Error GoTo AccionesCorrectivas

prBackUpGuardado = vBackUpGuardado
Exit Property
AccionesCorrectivas:
MsgBox "Tengo problemas con prBackUpGuardado"
End Property
Public Property Let prBackUpGuardado(vNuevosDatos As Boolean)  'True el backUp ya fue guardado por medio de un procedimiento alterno y false no se ha guardado en forma alternativa. Esta propiedad se usará para activarse cuando el nombre del libro cambia y se guarda la copia de seguridad por medio de codigo alternativo.
On Error GoTo AccionesCorrectivas

vBackUpGuardado = vNuevosDatos
Exit Property
AccionesCorrectivas:
MsgBox "Tengo problemas con prBackUpGuardado"
End Property



