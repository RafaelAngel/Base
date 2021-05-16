VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCampos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private vFieldDeBusqueda As Variant
Private mvarField1 As Variant 'copia local
Private mvarField2 As Variant 'copia local
Private mvarField4 As Variant 'copia local
Private mvarField5 As Variant 'copia local
Private mvarField6 As Variant 'copia local
Private mvarField7 As Variant 'copia local
Private mvarField8 As Variant 'copia local
Private mvarField9 As Variant 'copia local
Private mvarField10 As Variant 'copia local
Private mvarField11 As Variant 'copia local
Private mvarField12 As Variant 'copia local
Private mvarField13 As Variant 'copia local
Private mvarField14 As Variant 'copia local
Private mvarField15 As Variant 'copia local
Private vBase As Variant
Private vQCampo As Single 'Para indicar en cual campo s ebuscará.
'Para activar este evento, use RaiseEvent con la siguiente sintaxis:
'RaiseEvent Fieldsevent[(arg1, arg2, ... , argn)]
Public Event Fieldsevent(Datos_del_field As Variant)
'variables locales para almacenar los valores de las propiedades
Private mvarField3 As Variant 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarTabla As Variant 'copia local

Public Property Let Tabla(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Tabla = 5
    mvarTabla = vData
End Property

Public Property Set Tabla(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Tabla = Form1
    Set mvarTabla = vData
End Property


Public Property Get Tabla() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Tabla
    If IsObject(mvarTabla) Then
        Set Tabla = mvarTabla
    Else
        Tabla = mvarTabla
    End If
End Property



Public Property Let Field3(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Field3 = 5
    mvarField3 = vData
End Property


Public Property Set Field3(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Field3 = Form1
    Set mvarField3 = vData
End Property


Public Property Get Field3() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Field3
    If IsObject(mvarField3) Then
        Set Field3 = mvarField3
    Else
        Field3 = mvarField3
    End If
End Property





Public Property Let Field15(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Field15 = 5
    mvarField15 = vData
End Property


Public Property Set Field15(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Field15 = Form1
    Set mvarField15 = vData
End Property


Public Property Get Field15() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Field15
    If IsObject(mvarField15) Then
        Set Field15 = mvarField15
    Else
        Field15 = mvarField15
    End If
End Property

Public Property Let Field14(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Field14 = 5
    mvarField14 = vData
End Property


Public Property Set Field14(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Field14 = Form1
    Set mvarField14 = vData
End Property


Public Property Get Field14() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Field14
    If IsObject(mvarField14) Then
        Set Field14 = mvarField14
    Else
        Field14 = mvarField14
    End If
End Property

Public Property Let Field13(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Field13 = 5
    mvarField13 = vData
End Property


Public Property Set Field13(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Field13 = Form1
    Set mvarField13 = vData
End Property


Public Property Get Field13() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Field13
    If IsObject(mvarField13) Then
        Set Field13 = mvarField13
    Else
        Field13 = mvarField13
    End If
End Property

Public Property Let Field12(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Field12 = 5
    mvarField12 = vData
End Property

Public Property Set Field12(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Field12 = Form1
    Set mvarField12 = vData
End Property

Public Property Get Field12() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Field12
    If IsObject(mvarField12) Then
        Set Field12 = mvarField12
    Else
        Field12 = mvarField12
    End If
End Property

Public Property Let Field11(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Field11 = 5
    mvarField11 = vData
End Property


Public Property Set Field11(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Field11 = Form1
    Set mvarField11 = vData
End Property

Public Property Get Field11() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Field11
    If IsObject(mvarField11) Then
        Set Field11 = mvarField11
    Else
        Field11 = mvarField11
    End If
End Property

Public Property Let Field10(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Field10 = 5
    mvarField10 = vData
End Property


Public Property Set Field10(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Field10 = Form1
    Set mvarField10 = vData
End Property


Public Property Get Field10() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Field10
    If IsObject(mvarField10) Then
        Set Field10 = mvarField10
    Else
        Field10 = mvarField10
    End If
End Property

Public Property Let Field9(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Field9 = 5
    mvarField9 = vData
End Property


Public Property Set Field9(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Field9 = Form1
    Set mvarField9 = vData
End Property


Public Property Get Field9() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Field9
    If IsObject(mvarField9) Then
        Set Field9 = mvarField9
    Else
        Field9 = mvarField9
    End If
End Property



Public Property Let Field8(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Field8 = 5
    mvarField8 = vData
End Property


Public Property Set Field8(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Field8 = Form1
    Set mvarField8 = vData
End Property


Public Property Get Field8() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Field8
    If IsObject(mvarField8) Then
        Set Field8 = mvarField8
    Else
        Field8 = mvarField8
    End If
End Property



Public Property Let Field7(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Field7 = 5
    mvarField7 = vData
End Property


Public Property Set Field7(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Field7 = Form1
    Set mvarField7 = vData
End Property


Public Property Get Field7() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Field7
    If IsObject(mvarField7) Then
        Set Field7 = mvarField7
    Else
        Field7 = mvarField7
    End If
End Property



Public Property Let Field6(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Field6 = 5
    mvarField6 = vData
End Property


Public Property Set Field6(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Field6 = Form1
    Set mvarField6 = vData
End Property


Public Property Get Field6() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Field6
    If IsObject(mvarField6) Then
        Set Field6 = mvarField6
    Else
        Field6 = mvarField6
    End If
End Property



Public Property Let Field5(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Field5 = 5
    mvarField5 = vData
End Property


Public Property Set Field5(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Field5 = Form1
    Set mvarField5 = vData
End Property


Public Property Get Field5() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Field5
    If IsObject(mvarField5) Then
        Set Field5 = mvarField5
    Else
        Field5 = mvarField5
    End If
End Property



Public Property Let Field4(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Field4 = 5
    mvarField4 = vData
End Property


Public Property Set Field4(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Field4 = Form1
    Set mvarField4 = vData
End Property


Public Property Get Field4() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Field4
    If IsObject(mvarField4) Then
        Set Field4 = mvarField4
    Else
        Field4 = mvarField4
    End If
End Property



Public Property Let Field2(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Field2 = 5
    mvarField2 = vData
End Property


Public Property Set Field2(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Field2 = Form1
    Set mvarField2 = vData
End Property


Public Property Get Field2() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Field2
    If IsObject(mvarField2) Then
        Set Field2 = mvarField2
    Else
        Field2 = mvarField2
    End If
End Property



Public Property Let Field1(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Field1 = 5
    mvarField1 = vData
End Property


Public Property Set Field1(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Field1 = Form1
    Set mvarField1 = vData
End Property


Public Property Get Field1() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Field1
    If IsObject(mvarField1) Then
        Set Field1 = mvarField1
    Else
        Field1 = mvarField1
    End If
End Property




Public Property Get Base() As Variant
Base = vBase
End Property

Public Property Let Base(ByVal vNewValue As Variant)
vBase = vNewValue
End Property



Public Property Get QCampo() As Single
QCampo = vQCampo
End Property
Public Property Let QCampo(ByVal vNewValue As Single)
vQCampo = vNewValue
End Property

Public Property Get FieldDeBusqueda() As Variant
FieldDeBusqueda = vFieldDeBusqueda
End Property

Public Property Let FieldDeBusqueda(ByVal vNewValue As Variant)
vFieldDeBusqueda = vNewValue
End Property



