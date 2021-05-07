Attribute VB_Name = "MdlGlobales"
Option Explicit


Public Enum Iconos
    Bola_roja = 101
    Bola_verde = 102
End Enum

Public Const DebeSeleccionarElTxt = "Debe seleccionar el texto que desea unir."
Public Const RecomendacionOnAdd = "Este comando es como cuando las palabras estan separadas en letras, esto pasa muchas veces cuando se pega el texto entonces puede unir ese texto para luego leerlo."
Public Const NoLeerUltimos = "No puedo leer los ultimos datos porque tiene activada la opcion "
Public Const ActivarParar = " Debe activar la opcion parrar lectura."

'Los datos han cambiado o no.
Public vLosDatosHanCambiado As Boolean 'Para comprobar si ya se guardaron los datos al salir.

'Crea un salto de linea.
Public Function RTC()
    RTC = Chr(13) + Chr(10)
End Function

Public Function VarTipo(Variable_nombre As Variant)
    'Comprueva la variable.
    VarTipo = TypeName(Variable_nombre)
End Function

Public Sub Ejemplo()
    'este es muy bueno para comprobar datos.
    Dim cualquierNúmero
    Do
       cualquierNúmero = InputBox("Escriba algo")
       DoEvents
    Loop Until cualquierNúmero = "angel" 'El bucle termina cuando se escribe la palabra angel.
    MsgBox cualquierNúmero

End Sub

Public Sub Main()
    'MsgBox Command
    
    'frmLecturaDelAyer.Show
    'PlaySound I_Waiting_orders
End Sub


Public Function ReemplazarTexto(Texto As Variant, Buscar_letra_o_texto As Variant, Reemplazar_con As Variant)
    On Error GoTo N
    ReemplazarTexto = Replace(Texto, Buscar_letra_o_texto, Reemplazar_con)
    Exit Function
N:
    MsgBox Err.Description, vbExclamation
End Function

'Public Function LoadIcon(Id As Iconos) As StdPicture
'Set LoadIcon = LoadResPicture(Id, vbResIcon)
'End Function

Public Function Unir(Datos As Variant)
    'Une el texto ceparado por espacios.
    Unir = ReemplazarTexto(Datos, Chr(32), "")
End Function



Public Sub UnirTxtOnTextBox(vTextBox As TextBox)
    'Une el texto dentro de un TextBox.
    Dim X As Variant
    X = vTextBox.SelText
    If X = "" Then
        MsgBox DebeSeleccionarElTxt & RTC & RecomendacionOnAdd, vbInformation
    Else
        vTextBox.SelText = Unir(vTextBox.SelText)
    End If
End Sub
 
Public Function CantidadDePalabras(Texto As Variant) 'Si el texto es muy grande, proboca desbordamiento de pila.
    'Permite contar la cantidad de palabras que hay en el texto que le pasen.
    'On Error GoTo N
    Dim ConteoDePalabras As Double, PalabraPorPalabra() As String
    PalabraPorPalabra() = Split(Texto, " ")
    'Dim fPrg As New FrmProgress
    'While UBound(PalabraPorPalabra) > ConteoDePalabras
    'Do
    'fPrg.Show
    'fPrg.Progres1.Max = UBound(PalabraPorPalabra)
    For ConteoDePalabras = 1 To UBound(PalabraPorPalabra) 'El más eficiente es el For Next.
        'On Error GoTo N
        ConteoDePalabras = ConteoDePalabras + Val(1)
        DoEvents
       'fPrg.Progres1.AvanzandoDouble ConteoDePalabras
    Next ConteoDePalabras
    'Loop Until UBound(PalabraPorPalabra) <> ConteoDePalabras
    'Wend
    'Unload fPrg
    'N:
    
    CantidadDePalabras = ConteoDePalabras
    ConteoDePalabras = 0
    'On Error Resume Next 'Por si al descargar el form este esta en Nothing y causa un error.
    'Unload fPrg
End Function

Public Sub sAgarrarDatos(Data As DataObject, ByRef Control_con_propiedad_text As Object) 'Obtiene el texto de otra ventana mediante el arrastre.
    On Error GoTo AccionesCorrectivas
    Control_con_propiedad_text.Text = Data.GetData(vbCFText)
    Exit Sub
AccionesCorrectivas:
    MsgBox "Tengo problemas con sAgarrarDatos"
End Sub
