VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl Libro 
   ClientHeight    =   5400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7395
   ScaleHeight     =   5400
   ScaleWidth      =   7395
   Begin VB.CommandButton CmdUltimoAntesDeCerrar 
      Caption         =   "Recordar"
      Height          =   375
      Left            =   0
      TabIndex        =   18
      ToolTipText     =   "Ultimo texto antes de cerrar el programa"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtSegundosLeyendo 
      DataField       =   "SegundosLeyendo"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1034
         SubFormatType   =   1
      EndProperty
      DataMember      =   "LecturaEnProgreso"
      DataSource      =   "DataEnvironment1"
      Height          =   405
      Left            =   7680
      TabIndex        =   17
      Text            =   "0"
      Top             =   3000
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.TextBox TxtNombreDelLibro 
      DataField       =   "NombreDelLibro"
      DataMember      =   "LecturaEnProgreso"
      DataSource      =   "DataEnvironment1"
      Height          =   375
      Left            =   7680
      TabIndex        =   16
      Text            =   "0"
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxtDepurandoTextoUltimoTexto 
      DataField       =   "DepurandoTextoUltimoTexto"
      DataMember      =   "LecturaEnProgreso"
      DataSource      =   "DataEnvironment1"
      Height          =   375
      Left            =   7680
      TabIndex        =   15
      Text            =   "0"
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxtHorasLeyendo 
      DataField       =   "HorasLeyendo"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1034
         SubFormatType   =   1
      EndProperty
      DataMember      =   "LecturaEnProgreso"
      DataSource      =   "DataEnvironment1"
      Height          =   375
      Left            =   7680
      TabIndex        =   14
      Text            =   "0"
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxtLenDeLectura 
      DataField       =   "LenDeLectura"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1034
         SubFormatType   =   1
      EndProperty
      DataMember      =   "LecturaEnProgreso"
      DataSource      =   "DataEnvironment1"
      Height          =   375
      Left            =   7680
      TabIndex        =   13
      Text            =   "0"
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxtLenIncrementalDelTexto 
      DataField       =   "LenIncrementalDelTexto"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1034
         SubFormatType   =   1
      EndProperty
      DataMember      =   "LecturaEnProgreso"
      DataSource      =   "DataEnvironment1"
      Height          =   375
      Left            =   7680
      TabIndex        =   12
      Text            =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxtMaximoDelValorReal 
      DataField       =   "MaximoDelValorReal"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1034
         SubFormatType   =   1
      EndProperty
      DataMember      =   "LecturaEnProgreso"
      DataSource      =   "DataEnvironment1"
      Height          =   375
      Left            =   8040
      TabIndex        =   11
      Text            =   "0"
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxtMinutosLeyendo 
      DataField       =   "MinutosLeyendo"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1034
         SubFormatType   =   1
      EndProperty
      DataMember      =   "LecturaEnProgreso"
      DataSource      =   "DataEnvironment1"
      Height          =   375
      Left            =   8040
      TabIndex        =   10
      Text            =   "0"
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxtPorcentaje 
      DataField       =   "Porcentaje"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1034
         SubFormatType   =   1
      EndProperty
      DataMember      =   "LecturaEnProgreso"
      DataSource      =   "DataEnvironment1"
      Height          =   375
      Left            =   8040
      TabIndex        =   9
      Text            =   "0"
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxtSelLength 
      DataField       =   "SelLength"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1034
         SubFormatType   =   1
      EndProperty
      DataMember      =   "LecturaEnProgreso"
      DataSource      =   "DataEnvironment1"
      Height          =   375
      Left            =   8040
      TabIndex        =   8
      Text            =   "0"
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxtUltomosDatosLeidos 
      DataField       =   "UltomosDatosLeidos"
      DataMember      =   "LecturaEnProgreso"
      DataSource      =   "DataEnvironment1"
      Height          =   375
      Left            =   8040
      TabIndex        =   7
      Text            =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer tSegundosDeLectura 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar"
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      ToolTipText     =   "Limpia el campo de texto"
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdPegar 
      Caption         =   "Pegar"
      Height          =   375
      Left            =   5160
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdAnexar 
      Caption         =   "Anexar"
      Height          =   375
      Left            =   4080
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "Anexar datos del porta papeles."
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdLeer 
      Caption         =   "Leer"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      ToolTipText     =   "Leer todo el texto del campo de texto:"
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdLeerUltimosDatos 
      Caption         =   "Ultimos datos"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      ToolTipText     =   "Leer los ultimos datos"
      Top             =   3840
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox RtxtUltimoTextoLeido 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   393217
      TextRTF         =   $"Libro.ctx":0000
   End
   Begin RichTextLib.RichTextBox RtxtLibro 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6588
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Libro.ctx":0082
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuAnexarTexto 
         Caption         =   "Anexar texto"
      End
      Begin VB.Menu mnu_1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPegar 
         Caption         =   "Pegar"
      End
      Begin VB.Menu Mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGuardarPropiedadesDelLibro 
         Caption         =   "Guardar propiedades del Libro"
      End
   End
End
Attribute VB_Name = "Libro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private vBackUpGuardado As Boolean
'Permite obtener un si o un no booleano por medio del parametro del evento en fTextoSeleccionado. Para poder continuar con las acciones.

'Las variables de las siguientes lineas se pueden habilitar en caso de modificar sus respectivas propiedades.
'Private vSegundero As Integer 'Para el temporizador.
'Private vDepurandoTexto_ultimo_texto As Variant 'Variable para uso de la propiedad prDepurandoTextoUltimoTexto
'Private vUltomosDatosLeidos As Variant
'Private vLenDeLectura As Double
'Private vHorasLeyendo As Integer
'Private vLen_incremental_del_texto As Double
'Private vPorcentaje As Double
'Private vMaximoDelValorReal As Double
'Private vMinutosLeyendo As Integer

Public Event ClickLeerultimosDatos()
Public Event ClickContinuar(Habilitar As Boolean)
Public Event EveSePuedeContinuar(Se_puede_continuar As Boolean) '
Public Event EveDeshabilitarControles(Habilitados As Boolean)
Public Event ClickIniciarLectura(Datos As Variant)
Public Event EveStatusDeLectura(Estatus As Variant)
Private vSegundos_de_lectura As Double
Public Event EveSegundosLeyendo(Segundos_leyendo As Double)
Public Event TiempoDeLectura(Horas As Integer, Minutos As Integer, Segundos As Integer)

Public Sub GuardarPropiedades()
    'DataEnvironment1.rsLecturaEnProgreso.Update
End Sub
Private Sub sUltimosDatos()
'Por ahora solo dejar pasar el error.
'On Error Resume Next
RtxtLibro.SetFocus
    Dim NivelSuperior As Double
    Dim NivelAnidado1 As Double
    Dim NivelAnidado2 As Double
    
    For NivelSuperior = 0 To Len(RtxtLibro.Text)
        DoEvents
        RtxtLibro.SelStart = NivelSuperior
        RtxtLibro.SelLength = 1
        If RtxtLibro.SelColor = vbBlack Then
            RtxtLibro.SelStart = NivelSuperior - 2
            'Se le restan 500 porque este comando solo se usa en caso de haberse iniciado la lectura.
            NivelAnidado1 = RtxtLibro.SelStart - 500
            Exit For
            GoTo Salir
            
        End If
    Next
    
Salir:
Dim vParar As Boolean
'MsgBox ("Inicio " & NivelAnidado1)
Do While (vParar <> True)
    DoEvents
    NivelAnidado1 = NivelAnidado1 - 1
    If NivelAnidado1 < 0 Then
        NivelAnidado1 = 0
        RtxtLibro.SelStart = NivelAnidado1
        vParar = True
        GoTo S
    End If
    RtxtLibro.SelStart = NivelAnidado1
    RtxtLibro.SelLength = 1
    'Debug.Print RTC & RTC & RtxtLibro.SelText
    If RtxtLibro.SelText = " " Then
S:
        RtxtLibro.SelLength = 500
        'Debug.Print "Estoy en " & NivelAnidado1 & RTC & "=" & RtxtLibro.Text '(prUltimosDatosLeidos)
        vParar = True
        sDecir RtxtLibro.SelText
        RtxtLibro.SelBold = True
        RtxtLibro.SelFontSize = 14
        RtxtLibro.SelColor = vbGreen
        RtxtLibro.SelStart = RtxtLibro.SelStart + RtxtLibro.SelLength
        fGuardarLibro 'Guarda la lectura en progreso. Originalmente se guardaba al guardar en base.
        'RaiseEvent ClickIniciarLectura(RtxtLibro.SelText)
        'MsgBox (NivelAnidado1 & " voy " & RtxtLibro.SelText)
        'Exit Do
    End If
Loop



'        fGuardarLibro 'Guarda la lectura en progreso. Originalmente se guardaba al guardar en base.
'        RaiseEvent ClickIniciarLectura(RtxtLibro.SelText)
    'sDecir prUltimosDatosLeidos
End Sub

Public Property Get prHabilitarPegar() As Boolean
    prHabilitarAnexar = cmdPegar.Enabled
End Property
Public Property Let prHabilitarPegar(Nuevo As Boolean)
    cmdPegar.Enabled = Nuevo
End Property

Public Property Get prHabilitarAnexar() As Boolean
    prHabilitarAnexar = cmdAnexar.Enabled
End Property
Public Property Let prHabilitarAnexar(Nuevo As Boolean)
    cmdAnexar.Enabled = Nuevo
End Property

Public Property Get Text() As Variant
    Text = RtxtLibro.Text
End Property

Public Property Let Text(Nuevo As Variant)
    RtxtLibro.Text = Nuevo
End Property

Public Property Get SelText() As Variant
    SelText = RtxtLibro.SelText
End Property
Public Property Let SelText(Nuevo As Variant)
    MsgBox ("La propiedad es de solo lectura.")
End Property

Public Function FLen() As Double
'On Error Resume Next
'El desbordamiento de pila se debio a que estaba usando un integer cuando en realidad tenia que usar un double o un long.
    FLen = Len(RtxtLibro.Text)
End Function


Public Property Get SelStart() As Integer
    SelStart = RtxtLibro.SelStart
End Property
Public Property Let SelStart(Nuevo As Integer)
    RtxtLibro.SelStart = Nuevo
End Property

Public Property Get PrSelLength() As Long
'Es de solo lectura
    PrSelLength = RtxtLibro.SelLength
    TxtSelLength.Text = RtxtLibro.SelLength
End Property
Public Property Let LelLength(Nuevo As Long)
    RtxtLibro.SelLength = Nuevo
End Property

Public Function fDepurandoTexto(ByRef Datos As Variant, Libro As Object) 'Depura el texto o los datos para que termine en una palabra competa.
'Este es el algoritmo para concatenar palabras incompletas.
'Ademas hace que el texto termine en una palabra completa.
'El parametro Libro es un RichTextBox.

On Error GoTo AccionesCorrectivas

If Datos = "" Then 'Comprueba que los datos estan en blanco, de estarlo cancela la ejecucion del resto de la funcion para evitar un error.
    MsgBox "Se termino el libro", vbInformation
    sEnabledControles True
    tSegundosDeLectura.Enabled = False
    cmdAnexar.Enabled = True
    Exit Function
End If

Dim mDepuracion() As String, vLen As Double, vRespuesta As Variant

mDepuracion = Split(Datos, " ") 'Se insertan los datos en la matriz.

vRespuesta = Join(mDepuracion, " ") 'Se vuelven a poner en una variable. Pero quedan tambien en la matriz.

vLen = Len(vRespuesta) - Len(mDepuracion(UBound(mDepuracion))) 'Se usan los datos de la variable y el ultimo id de la matriz para sacar el Len de texto que se devolver� como respuesta.

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

Private Property Get prDepurandoTextoUltimoTexto() As Variant 'Guarda el ultimo id de la matriz local mDepuracion, en la funcion fDepurandoTexto, para luego concatenarlo como parte de la respuesta de la siguiente llamada de la funcion.
'Guarda la ultima palabra del texto seleccionado sea o no completa.
On Error GoTo AccionesCorrectivas

prDepurandoTextoUltimoTexto = TxtDepurandoTextoUltimoTexto.Text 'vDepurandoTexto_ultimo_texto
Exit Property
AccionesCorrectivas:
MsgBox "Tengo problemas con prDepurandoTextoUltimoTextoprDepurandoTextoUltimoTexto"
End Property
Private Property Let prDepurandoTextoUltimoTexto(vNuevosDatos As Variant)  'Guarda el ultimo id de la matriz local mDepuracion, en la funcion fDepurandoTexto, para luego concatenarlo como parte de la respuesta de la siguiente llamada de la funcion.
On Error GoTo AccionesCorrectivas

'vDepurandoTexto_ultimo_texto = vNuevosDatos
TxtDepurandoTextoUltimoTexto.Text = vNuevosDatos
Exit Property
AccionesCorrectivas:
MsgBox "Tengo problemas con prDepurandoTextoUltimoTextoprDepurandoTextoUltimoTexto"
End Property

Public Function fTextoSeleccionado() As Variant
    Dim vSePuedeContinuar As Boolean 'Para guardar los datos del paramettro del evento.
    
    RaiseEvent EveSePuedeContinuar(vSePuedeContinuar)
    If vSePuedeContinuar = False Then Exit Function 'Detiene la funcion. De lo contrario se seleccionaria una fraccion inecesaria de texto en clsRichlectura_sDecir no se leeria esa ultima fraccion de texto con lo cual quedaria en el olvido.
    
    RtxtLibro.Enabled = False 'Se deshabilita para evitar que el usuario toque la zona de texto durante la seleccion de datos.
    On Error GoTo AccionesCorrectivas
    
    Dim vMaxLen As Integer
    vMaxLen = 500
    
    fStatusDeLaLectura
    'lbStatus = "prSelStart " & prSelStart & " prLenDeLectura " & prLenDeLectura & " Tiempo leyendolo " & prHorasLeyendo & ":" & prMinutosLeyendo  'Datos en la barra de estado lbStatus.
    RaiseEvent EveStatusDeLectura("SelStart " & SelStart & " prLenDeLectura " & prLenDeLectura & " Tiempo leyendolo " & prHorasLeyendo & ":" & prMinutosLeyendo)
    
    If SelStart < FLen Then
    
        If vMaxLen > FLen Then vMaxLen = FLen
        
            With RtxtLibro
                   .SelStart = SelStart 'Inicialmente la propiedad vale cero.
                   .SelLength = vMaxLen
                   .SelBold = True
                   .SelColor = vbGreen 'El texto leido se vuelve color verde.
                   .SelFontSize = 14
                   'Los ultimos datos y se leeran en una nueva secion.
               
                   fTextoSeleccionado = .SelText 'vControles.fDepurandoTexto(RtxtLibro.SelText)
               
                   SelStart = SelStart + 500 'Incrementa el punto del inicio de la seleccion
                   .SelLength = 0
                   .SelColor = vbBlack 'Quita la seleccion sin borrar los datos. Y ademas regresa la propiedad a un color que indica que el texto no ha sido leido.
                   .SelStart = SelStart 'Salta al punto exacto donde va la lectura.
               End With 'RtxtLibro
        
        'clsRichlectura_sProgressBarStatus ProgressBar1, prSelStart 'Muestra el avance del texto que ya se ha leido.
        
    End If
    RtxtLibro.Enabled = True 'Despues de finalizadas las tareas de seleccion, se habilita la zona de texto.
    
    Exit Function
AccionesCorrectivas:     'A ver si con esto corrijo o descubro que hacer cuando se presente este error.
    
    If vMaxLen > FLen Then vMaxLen = FLen

    With RtxtLibro
        .SelStart = SelStart 'Inicialmente la propiedad vale cero.
        .SelLength = vMaxLen
        .SelBold = True
        .SelColor = vbRed
        .SelFontSize = 10
    
        SelStart = SelStart + 500 'Incrementa el punto del inicio de la seleccion.
    
        .SelLength = 0
        .SelColor = vbBlack 'Quita la seleccion sin borrar los datos.
        .SelStart = SelStart 'Salta al punto exacto donde va la lectura.
        'clsRichlectura_sProgressBarStatus ProgressBar1, prSelStart 'Muestra el avance del texto que ya se ha leido.
    End With
    RtxtLibro.Enabled = True
End Function

Public Sub sEnabledControles(Habilitados As Boolean)
    RtxtLibro.Enabled = Habilitados
    cmdLeerUltimosDatos.Enabled = Habilitados
    cmdLeer.Enabled = Habilitados
    cmdAnexar.Enabled = Habilitados
    cmdPegar.Enabled = Habilitados
    cmdLimpiar.Enabled = Habilitados
    
End Sub

Public Function fContinuarLeyendo() As Variant
'permite continuar leyendo una vez inicializada la lectura.
    sDecir fTextoSeleccionado
    fContinuarLeyendo = prUltimosDatosLeidos
End Function

Private Sub sDecir(Datos As Variant)
'Funciona as�: sDecir clsRichlectura_fTextoSeleccionado
    On Error GoTo AccionesCorrectivas
    Dim vSePuedeContinuar As Boolean 'Para guardar los datos del paramettro del evento.

    RaiseEvent EveSePuedeContinuar(vSePuedeContinuar)
    If vSePuedeContinuar = True Then
        sEnabledControles (False)
        RaiseEvent EveDeshabilitarControles(True)
        
        
        
        'DirectSS1.AudioReset
        prUltimosDatosLeidos = fDepurandoTexto(Datos, RtxtLibro)  'Copia del ultimo texto leido.
        'DirectSS1.Speak "Hola Rafa"
        'DirectSS1.Speak prUltimosDatosLeidos   'vControles.fDepurandoTexto(Datos, RtxtLibro)  'Los datos se depuran en el Sub Decir; porque Decir se llama para leer los ultimos datos en una nueva secion. Los ultimos datos leidos estan sin depurar para evitar perdidas de informacion.
        fGuardarLibro 'Guarda la lectura en progreso. Originalmente se guardaba al guardar en base.
        RaiseEvent ClickIniciarLectura(prUltimosDatosLeidos)
        
        'Los ultimos datos se guardan automaticamente. Para evitar perdidas.
        'Reactivar la siguiente linea al finalizar.
        'CBase.LetF prUltimosDatosLeidos, "UltimosDatos"
        
        '___________________________________________________________________
        
        'vControles.sAddDiezCargas prUltimosDatosLeidos 'Agrega 10 cargas de texto y reinicia la lista.
        
        'Habilitar la siguiente linea
        'CBase.LetF prDepurandoTextoUltimoTexto, "DepurandoTexto_ultimo_texto"
    
    Else
       ' MsgBox "No puedo decir nada porque la opcion esta en pausa o en parar.", vbInformation
    End If
    
    Exit Sub
AccionesCorrectivas:
    'Marcar con rojo el texto que causa el error.
    sEnabledControles True
End Sub

Public Property Get prLenDeLectura() As Double 'La cantidad de texto a leer.
    On Error GoTo AccionesCorrectivas
    prLenDeLectura = TxtLenDeLectura.Text 'vLenDeLectura
    Exit Property
AccionesCorrectivas:
    MsgBox "Tengo problemas con prLenDeLecturaprLenDeLectura"
End Property
Public Property Let prLenDeLectura(vNuevosDatos As Double)  'La cantidad de texto a leer.
    On Error GoTo AccionesCorrectivas
    TxtLenDeLectura.Text = vNuevosDatos
    'vLenDeLectura = vNuevosDatos
    'Habilitarlo luego
    'fProgressBarMax ProgressBar1, prLenDeLectura  'Se carga el Max del ProgressBar mediante un procedimiento que verifica una serie de valores para evitar errores o que esta propiedad se detenga.
  
    PropertyChanged "prLenDeLectura"
    Exit Property
AccionesCorrectivas:
    MsgBox "Tengo problemas con prLenDeLecturaprLenDeLectura"
End Property

Private Function fProgressBarMax(ProgressBar_nombre As Object, Optional ByVal Max As Double = 1#) As Variant
    On Error GoTo N
    
    Dim vMax As Double
    vMax = Max
    
    If vMax = 0 Then vMax = 1 'El Max jamas puede tener un valor cero, asi que se le da el valor 1.
    
    ProgressBar_nombre.Max = vMax
    
    Exit Function
N:
    MsgBox Err.Description & RTC & "clsRichlectura_fProgressBarMax"

End Function

Private Sub cmdAnexar_Click()
    'La propiedad se carga con el Len de RtxtLibro.
     prLenDeLectura = fGetTextoDelPortapapeles(RtxtLibro)
    'Luego se llama a la propiedad para mostrar informacion en lbStatus.
    'lbStatus = "Len del texto copiado " & prLenDeLectura 'vControles.fGetTextoDelPortapapeles(RtxtLibro)
    
    prBackUpGuardado = False
End Sub


Public Property Get prHorasLeyendo() As Integer 'Cantidad de horas leyendo el libro.
    On Error GoTo AccionesCorrectivas
    prHorasLeyendo = TxtHorasLeyendo.Text 'vHorasLeyendo
    Exit Property
AccionesCorrectivas:
    MsgBox "Tengo problemas con prHorasLeyendoprHorasLeyendo"
End Property
Public Property Let prHorasLeyendo(vNuevosDatos As Integer)  'Cantidad de horas leyendo el libro.
    On Error GoTo AccionesCorrectivas
    TxtHorasLeyendo.Text = vNuevosDatos
    'vHorasLeyendo = vNuevosDatos
    PropertyChanged "prHorasLeyendo"
    Exit Property
AccionesCorrectivas:
    MsgBox "Tengo problemas con prHorasLeyendoprHorasLeyendo"
End Property


Private Sub cmdAnexar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPopUp
    End If
End Sub

Private Sub cmdLeer_Click()
    cmdAnexar.Enabled = False
    tSegundosDeLectura.Enabled = True 'no va aqu�. Va en el evento DirectSS1_AudioStart.
    prDepurandoTextoUltimoTexto = "" 'Se borra la informacion de la propiedad porque ya ha sido concatenada al final de la propiedad prUltimosDatosLeidos.
    
    'OptContinuar_Click 'Activo el Option de la lectura
    RaiseEvent ClickContinuar(True)
    sSeltxtColor RtxtLibro  'Se pone el color negro al texto en caso que se haya iniciado una lectura anterior.
    
    With Me 'Tareas en el inicio de una nueva lectura.
        .SelStart = 0 'La propiedad se pone a cero para iniciar una nueva lectura.
        .prHorasLeyendo = 0
        .prMinutosLeyendo = 0
        'Darle el Max al ProgressBar y guardarlo en la propiedad prLenDeLectura.
        .prLenDeLectura = Len(RtxtLibro.Text)
    End With
    
    fGuardarBackUp RtxtLibro, "Libro.LibroBkUp" 'Se guarda una copia de seguridad antes de iniciar la lectura.
    sDecir fTextoSeleccionado
End Sub

Private Sub cmdLeer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPopUp
    End If
End Sub


Private Sub cmdLeerUltimosDatos_Click()
    tSegundosDeLectura.Enabled = True 'no va aqu�. Va en el evento DirectSS1_AudioStart.
    
    prDepurandoTextoUltimoTexto = "" 'Se borra la informacion de la propiedad porque ya ha sido concatenada al final de la propiedad prUltimosDatosLeidos.
    'ProgressBar1.Max = Len(RtxtLibro.Text)'No es necesario porque el Max se carga al cargar los datos de la base de datos.
    'OptParar_Click
    'OptContinuar_Click
    
    RaiseEvent ClickLeerultimosDatos
    sDecir prUltimosDatosLeidos
End Sub


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

Public Function fGuardarBackUp(ByRef RichTextBox_nombre As Object, ByRef Nombre_del_libro As Variant) 'Guarda la copia de seguridad del libro.
    'Sirve para guardar libros.
    On Error GoTo AccionesCorrectivas
    RichTextBox_nombre.SaveFile App.Path & "\" & Nombre_del_libro & ".txtLeido"
    fGuardarBackUp = App.Path & "\" & Nombre_del_libro & ".txtLeido"
    Exit Function
AccionesCorrectivas:
    MsgBox "Tengo problemas con GuardarBackUp"
End Function


Public Property Get prBackUpGuardado() As Boolean 'True el backUp ya fue guardado por medio de un procedimiento alterno y false no se ha guardado en forma alternativa. Esta propiedad se usar� para activarse cuando el nombre del libro cambia y se guarda la copia de seguridad por medio de codigo alternativo.
    On Error GoTo AccionesCorrectivas
    
    prBackUpGuardado = vBackUpGuardado
    Exit Property
AccionesCorrectivas:
    MsgBox "Tengo problemas con prBackUpGuardado"
End Property
Public Property Let prBackUpGuardado(vNuevosDatos As Boolean)  'True el backUp ya fue guardado por medio de un procedimiento alterno y false no se ha guardado en forma alternativa. Esta propiedad se usar� para activarse cuando el nombre del libro cambia y se guarda la copia de seguridad por medio de codigo alternativo.
    On Error GoTo AccionesCorrectivas
    
    vBackUpGuardado = vNuevosDatos
    Exit Property
AccionesCorrectivas:
    MsgBox "Tengo problemas con prBackUpGuardado"
End Property

Public Function fGetTextoDelPortapapeles(ByRef Control_con_una_propiedad_Text As Object) As Double 'Carga al RichTextBox con nuevos datos sin borrar los anteriores. Ademas devuelve el Len de todo el texto.
    On Error GoTo AccionesCorrectivas
    
    Dim vLen As Double 'El uso de variables locales evita la sobrecarga. Y as� Len(RichTextBox1.Text) solo se carga una vez.
    vLen = Len(Control_con_una_propiedad_Text.Text) 'Se mide al principio y al final despues de cargar los datos.
    
    On Error GoTo N
    
    If Control_con_una_propiedad_Text.Text = "" Or Control_con_una_propiedad_Text.Text = "0" Then
        
        Control_con_una_propiedad_Text.Text = Clipboard.GetText
    Else
    
N:
        On Error Resume Next 'No es muy correcto usar On Error Resume Next, es mejor usar On Error GoTo N; pero por el momento usar� esta linea aqu�.
        
            
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

Public Property Get prLenIncrementalDelTexto() As Double 'Se carga con el Len del texto copiado del portapapeles con el Function fGetTextoDelPortapapeles.
    On Error GoTo AccionesCorrectivas
    prLenIncrementalDelTexto = TxtLenIncrementalDelTexto.Text 'vLen_incremental_del_texto
    Exit Property
AccionesCorrectivas:
    MsgBox "Tengo problemas con prLenIncrementalDelTextoprLenIncrementalDelTexto"
End Property
Public Property Let prLenIncrementalDelTexto(vNuevosDatos As Double)  'Se carga con el Len del texto copiado del portapapeles con el Function fGetTextoDelPortapapeles.
    On Error GoTo AccionesCorrectivas
    TxtLenIncrementalDelTexto.Text = vNuevosDatos
    'vLen_incremental_del_texto = vNuevosDatos
    Exit Property
AccionesCorrectivas:
    MsgBox "Tengo problemas con prLenIncrementalDelTextoprLenIncrementalDelTexto"
End Property


Private Function fStatusDeLaLectura() As Variant
    fStatusDeLaLectura = prPorcentaje & "% de prSelStart=" & SelStart & " prLenDeLectura=" & prLenDeLectura & " Tiempo leyendolo " & prHorasLeyendo & ":" & prMinutosLeyendo & ":" & prSegundosDeLectura   'Datos en la barra de estado lbStatus.
End Function

Private Function fValeCero(ByRef Nombre_del_procedimiento As Variant, ByVal Objeto_numerico_a_comprobar As Double, ByRef Nombre_del_objeto_numerico As Variant) As Boolean 'Comprueba la propiedad prMaximoDelValorReal a ver si vale cero o m�s...
    On Error GoTo AccionesCorrectivas
    
    
    If Objeto_numerico_a_comprobar = 0 Then
        'MsgBox "La propiedad vale cero, no se puede seguir con la operacion. " & Nombre_del_objeto_numerico & "=0", vbInformation, Nombre_del_procedimiento
        fValeCero = True
    Else
        fValeCero = False
    End If
    
    Exit Function
AccionesCorrectivas:
    MsgBox "Tengo problemas con fValeCero"
End Function

Public Property Get prMaximoDelValorReal() As Double 'Valor maximo del numero que desea trabajar con porcentajes. Por ejemplo el Maximo de 38 es 38...
    On Error GoTo AccionesCorrectivas
    prMaximoDelValorReal = TxtMaximoDelValorReal.Text 'vMaximoDelValorReal
    Exit Property
AccionesCorrectivas:
    MsgBox "Tengo problemas con prMaximoDelValorReal"
End Property
Public Property Let prMaximoDelValorReal(vNuevosDatos As Double)  'Valor maximo del numero que desea trabajar con porcentajes. Por ejemplo el Maximo de 38 es 38...
    On Error GoTo AccionesCorrectivas
    TxtMaximoDelValorReal.Text = vNuevosDatos
    'vMaximoDelValorReal = vNuevosDatos
    Exit Property
AccionesCorrectivas:
    MsgBox "Tengo problemas con prMaximoDelValorReal"
End Property

Private Function fPorcentaje(ByVal Valor_real As Double, Optional Solo_numeros_enteros As Boolean = True) As Double '(ByVal Max_del_valor_real As double, ByVal Valor_real As double) As double 'Da el porcentaje de la cantidad que se le esta pasando con respecto al Max
    
    If fValeCero("fValorRealDeUnPorcentaje", Valor_real, "Valor_real") = True Then
        Exit Function
    End If
    
    If fValeCero("fValorRealDeUnPorcentaje", prMaximoDelValorReal, "prMaximoDelValorReal") = True Then
        Exit Function
    End If
    
    If Valor_real = 0 Then Exit Function
    
    If Solo_numeros_enteros Then
        prPorcentaje = CInt(Valor_real * 100 / prMaximoDelValorReal)
        'fPorcentaje = CInt(Valor_real * 100 / prMaximoDelValorReal) 'Max_del_valor_real
        
    Else
        prPorcentaje = Valor_real * 100 / prMaximoDelValorReal
        'fPorcentaje = Valor_real * 100 / prMaximoDelValorReal 'Max_del_valor_real
    End If
    
    fPorcentaje = prPorcentaje
    'Saca el porcentaje del valor con respecto al maximo o 100% de un numero por ejemplo 34.
End Function

Public Property Get prMinutosLeyendo() As Integer 'Cantidad de minutos leyendo.
    On Error GoTo AccionesCorrectivas
    prMinutosLeyendo = TxtMinutosLeyendo.Text 'vMinutosLeyendo
    Exit Property
AccionesCorrectivas:
    MsgBox "Tengo problemas con prMinutosLeyendoprMinutosLeyendo"
End Property
Public Property Let prMinutosLeyendo(vNuevosDatos As Integer)  'Cantidad de minutos leyendo.
    On Error GoTo AccionesCorrectivas
    TxtMinutosLeyendo.Text = Val(Left(CStr(vNuevosDatos), 2)) 'Hace que solo se muestren dos digitos.
    'vMinutosLeyendo = Val(Left(CStr(vNuevosDatos), 2)) 'Hace que solo se muestren dos digitos.
    PropertyChanged "prMinutosLeyendo"
    Exit Property
AccionesCorrectivas:
    MsgBox "Tengo problemas con prMinutosLeyendoprMinutosLeyendo"
End Property

Public Property Get prUltimosDatosLeidos() As Variant 'Guarda los ultimos datos que se leyeron.
    On Error GoTo N
    prUltimosDatosLeidos = TxtUltomosDatosLeidos.Text 'vUltomosDatosLeidos
    Exit Property
N:
    MsgBox "Tengo problemas con prUltimosDatosLeidos"
End Property
Public Property Let prUltimosDatosLeidos(vNuevosDatos As Variant)  'Guarda los ultimos datos que se leyeron.
    On Error GoTo N
    TxtUltomosDatosLeidos.Text = vNuevosDatos
    'vUltomosDatosLeidos = vNuevosDatos
    PropertyChanged "prUltimosDatosLeidos"
    Exit Property
N:
    MsgBox "Tengo problemas con prUltimosDatosLeidos"
End Property

Public Property Get prPorcentaje() As Double 'Se guarda el porcentaje calculado por si necesita llamarse de nuevo en otra parte del codigo.
    On Error GoTo AccionesCorrectivas
    prPorcentaje = TxtPorcentaje.Text 'vPorcentaje
    Exit Property
AccionesCorrectivas:
    MsgBox "Tengo problemas con prPorcentaje"
End Property
Public Property Let prPorcentaje(vNuevosDatos As Double)  'Se guarda el porcentaje calculado por si necesita llamarse de nuevo en otra parte del codigo.
    On Error GoTo AccionesCorrectivas
    TxtPorcentaje.Text = vNuevosDatos
    'vPorcentaje = vNuevosDatos
    Exit Property
AccionesCorrectivas:
    MsgBox "Tengo problemas con prPorcentaje"
End Property

Public Sub LoadLecturaEnProgreso()
'Originalmente permitia guardar varios libros a la vez.
'Lo que permitia leer varios libros en forma alternada.
     sAbrirLecturaEnProgreso RtxtLibro, "LecturaEnProgreso" 'CmbTitulo.Text
End Sub

Private Sub sAbrirLecturaEnProgreso(ByRef RichTextBox_nombre As Object, Nombre_del_libro As Variant) 'Abre el libro que se esta leyendo.
    On Error GoTo AccionesCorrectivas
        RtxtUltimoTextoLeido.LoadFile App.Path & "\" & Nombre_del_libro & "UltimosDatos.Lectura"   'Abre el archivo
        RichTextBox_nombre.LoadFile App.Path & "\" & Nombre_del_libro & ".Lectura"   'Abre el archivo
    Exit Sub
AccionesCorrectivas:
        MsgBox "Tengo problemas con AbrirLecturaEnProgreso"
End Sub

Public Sub sAbrirLectura()
    Const NOMBRE = "LecturaEnProgreso"
    sAbrirLecturaEnProgreso RtxtLibro, NOMBRE
End Sub

Public Sub sAbrirBackUp(ByRef RichTextBox_nombre As Object, Nombre_del_libro As Variant) 'Abre el BackUp del libro que se esta leyendo.
    On Error GoTo AccionesCorrectivas
        RichTextBox_nombre.Text = ""
        RichTextBox_nombre.LoadFile App.Path & "\" & Nombre_del_libro & ".txtLeido"
    Exit Sub
AccionesCorrectivas:
        MsgBox "Tengo problemas con AbrirBackUp"
End Sub


Private Function fGuardarLecturaEnProgreso(ByRef RichTextBox_nombre As Object, ByRef Nombre_del_libro As Variant) 'Guarda la lectura en progreso.
    On Error GoTo AccionesCorrectivas
    RtxtUltimoTextoLeido.SaveFile App.Path & "\" & Nombre_del_libro & "UltimosDatos.Lectura"
    TxtNombreDelLibro.Text = Nombre_del_libro
    RichTextBox_nombre.SaveFile App.Path & "\" & Nombre_del_libro & ".Lectura"
    fGuardarLecturaEnProgreso = App.Path & "\" & Nombre_del_libro & ".Lectura"
    'No se agregan registros porque solo existe un registro para este programa por el momento.
    'DataEnvironment1.rsLecturaEnProgreso.Update
    Exit Function
AccionesCorrectivas:
    MsgBox "Tengo problemas con GuardarLecturaEnProgreso" & RTC & Err.Description
End Function



Public Function fGuardarLibro()
'Se llama cuando se guarda en la base.
    fGuardarLibro = fGuardarLecturaEnProgreso(RtxtLibro, "LecturaEnProgreso") 'CmbTitulo.Text)
End Function

Public Sub sLoadBackUp()
    'El texto inicial o BackUp se carga en RichtxtLecturaEnProgreso.
    sAbrirBackUp RtxtLibro, "Libro.LibroBkUp" 'CmbTitulo.Text
    'RichtxtLecturaEnProgreso = RichtxtBackUp
End Sub

Public Property Get prSegundosDeLectura() As Double 'Guarda la cantidad de tiempo de cada fraccion de lectura.
    On Error GoTo N
    prSegundosDeLectura = txtSegundosLeyendo.Text 'vSegundos_de_lectura
    Exit Property
N:
    MsgBox "Tengo problemas con prSegundosDeLectura"
End Property
Public Property Let prSegundosDeLectura(vNuevosDatos As Double)  'Guarda la cantidad de tiempo de cada fraccion de lectura.
    On Error GoTo N
    txtSegundosLeyendo.Text = vNuevosDatos
    'vSegundos_de_lectura = vNuevosDatos
    PropertyChanged "prSegundosDeLectura"
    Exit Property
N:
    MsgBox "Tengo problemas con prSegundosDeLectura"
End Property

Private Sub cmdLeerUltimosDatos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPopUp
    End If
End Sub

Private Sub cmdLimpiar_Click()
    RtxtLibro.Text = ""
    cmdPegar.Enabled = True
End Sub

Private Sub cmdLimpiar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPopUp
    End If
End Sub


Private Sub cmdPegar_Click()
'    With Me
'        .prHorasLeyendo = 0
'        .prMinutosLeyendo = 0
'        .prSegundosDeLectura = 0
'    End With
'    RtxtLibro.Text = Clipboard.GetText
'
'    prLenDeLectura = Len(RtxtLibro.Text)
    cmdPegar.Enabled = False
    RtxtLibro.Text = ""
    fGetTextoDelPortapapeles RtxtLibro
End Sub

Private Sub cmdPegar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPopUp
    End If
End Sub


Private Sub CmdUltimoAntesDeCerrar_Click()
    sUltimosDatos
End Sub

Private Sub mnuAnexarTexto_Click()
    cmdAnexar_Click
End Sub

Private Sub MnuGuardarPropiedadesDelLibro_Click()
    GuardarPropiedades
End Sub

Private Sub MnuPegar_Click()
    cmdPegar_Click
End Sub

'Private Sub RtxtLibro_KeyPress(KeyAscii As Integer)
'    Debug.Print KeyAscii
'End Sub

Private Sub RtxtLibro_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPopUp
    End If
End Sub


Private Sub tSegundosDeLectura_Timer()
    prSegundosDeLectura = prSegundosDeLectura + 1
    'lbTiempoTranscurido = prSegundosDeLectura
    RaiseEvent EveSegundosLeyendo(prSegundosDeLectura)
    
    'vSegundero = vSegundero + 1
    prSegundosDeLectura = prSegundosDeLectura + 1
    If prSegundosDeLectura = 60 Then
    
        prSegundosDeLectura = 0
        prMinutosLeyendo = prMinutosLeyendo + 1
        
        If prMinutosLeyendo >= 60 Then
            prHorasLeyendo = prHorasLeyendo + 1 'Se suma una hora m�s al cabo de 60 minutos.
            prMinutosLeyendo = 0 'Despues de 60 minutos, los minutos vuelven a cero.
        End If
    End If
RaiseEvent TiempoDeLectura(prHorasLeyendo, prMinutosLeyendo, prSegundosDeLectura)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPopUp
    End If
End Sub


Private Sub sEliminarLibro() 'Elimina la lectura en progreso y el BackUp.
    On Error GoTo AccionesCorrectivas
        
        'Luego habilitar la siguiente linea para modificarla.
        'Kill txtDireccion 'App.Path & "\" & CmbTitulo.Text & ".txtLeido"
        
        'La siguiente linea elimina el libro.
        'Se debe habilitar y ,modificar.
        'Kill App.Path & "\" & CmbTitulo.Text & ".txtLeido"
        
        
        prMinutosLeyendo = 0
        prHorasLeyendo = 0
    Exit Sub
AccionesCorrectivas:
        MsgBox Err.Description '"Tengo problemas con clsRichlectura_sEliminarLibro" & RTC & "No se pudo eliminar el libro.", vbCritical
End Sub

Public Property Get PrHabilitarLeer() As Boolean
    PrHabilitarLeer = cmdLeer.Enabled
End Property

Public Property Let PrHabilitarLeer(Nuevo As Boolean)
    cmdLeer.Enabled = Nuevo
End Property

Public Property Get PrHabilitarUltimosDatos() As Boolean
    PrHabilitarUltimosDatos = cmdLeerUltimosDatos.Enabled
End Property
Public Property Let PrHabilitarUltimosDatos(Nuevo As Boolean)
    cmdLeerUltimosDatos.Enabled = Nuevo
    RtxtLibro.Enabled = Nuevo
End Property
