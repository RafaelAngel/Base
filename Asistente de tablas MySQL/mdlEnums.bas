Attribute VB_Name = "mdlEnums"
Option Explicit

Public Enum TipoDeConsulta
    ConsultaSelect = 1
    ConsultaInsert = 2
    ConsultaUpdate = 3
    ConsultaDelete = 4
End Enum
'
'Public Function fOptenerConsulta(Tipo_de_consulta_de_retorno As TipoDeConsulta)
''Solo si ya se han insertado campos.
'    Select Case Tipo_de_consulta_de_retorno
'    Case TipoDeConsulta.ConsultaSelect
'        fOptenerConsulta = "Select " & Me.prListaDeCampos & " from " & Me.prTabla & fWhere.prCondicionesWhere & ";"
'    Case TipoDeConsulta.ConsultaInsert
'    'Se debe crear un form que se lanzara para editar los datos.
'    Case TipoDeConsulta.ConsultaUpdate
'    'Se debe crear un form que se lanzara para editar los datos.
'
'    Case TipoDeConsulta.ConsultaDelete
'        fOptenerConsulta = "Delete from " & Me.prTabla & fWhere.prCondicionesWhere & ";"
'    End Select
'End Function
