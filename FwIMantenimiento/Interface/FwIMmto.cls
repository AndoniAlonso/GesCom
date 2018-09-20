VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FwIMmto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum EntityRestriction
    DenyRead = 1
    DenyInsert = 2
    DenyUpdate = 4
    DenyDelete = 8
    DenyNavegate = 16
    DenyManagesCounters = 32
End Enum

Public Enum RecordNavigator
    fwGoFirst = 1
    fwGoPrevious = 2
    fwGoNext = 3
    fwGoLast = 4
    fwGoRecord = 5
End Enum

Public Enum UserMessageCodes
    UserMessageBaseCode = vbObjectError + 512
    UserMessageLimitCode = vbObjectError + 65536
End Enum

Public Enum fwmActionResult
    fwmActionError = 0
    fwmActionSucceeded = -1
    fwmActionCanceled = -2
End Enum

Public Enum fwmUpdateAction
    fwmInsert = 0
    fwmDelete = 1
    fwmUpdate = 2
End Enum

Public Enum fwSelectedLinesCounterMessages
    fwSelectedMessage = 55121
    fwCheckedMessage = 55120
End Enum

Public Function CounterRs() As Recordset
Attribute CounterRs.VB_Description = "Retorna un recordset desconectado con el código, descripción y valor provisional de todos los contadores asociados a la entidad correspondiente a la clase que lo implementa."
End Function
Public Function CounterDefault() As Recordset
Attribute CounterDefault.VB_Description = "Retorna un recordset desconectado con el código, descripción y valor provisional del contador predeterminado de la entidad correspondiente a la clase que lo implementa. Si la entidad correspondiente no tiene una contador predeteminado devuelve un records"
End Function
Public Function CounterValue(ByVal strIDContador As String) As String
Attribute CounterValue.VB_Description = "Retorna un recordset desconectado con el código, descripción y valor provisional para una contador dado."
End Function
Public Function SelOnPkRs(ByVal rs As Recordset) As Recordset
Attribute SelOnPkRs.VB_Description = "Realiza búsquedas por la clave principal en la tabla de la entidad correspondiente a la clase que lo implementa. Para ello toma el valor de la clave principal del recordset que se la pasa."
End Function
Public Function PkRs() As Recordset
Attribute PkRs.VB_Description = "Retorna un recordset desconectado y sin registros con la estructura de los campos correspondientes a la clave principal de la tabla de la entidad correspondiente a la clase que lo implementa"
End Function
Public Function InsertRs(ByVal rs As Recordset) As Recordset
Attribute InsertRs.VB_Description = "Toma un recordset desconectado, lo conecta y actualiza el contenido del mismo sobre la tabla de la base de datos correspondiente a la entidad de la clase que lo implementa. Al actualizarlo sobre la base de datos habrá campos del recordset que sufran modi"
End Function
Public Function UpdateRs(ByVal rs As Recordset) As Recordset
Attribute UpdateRs.VB_Description = "Toma un recordset desconectado, lo conecta y actualiza el contenido del mismo sobre la tabla de la base de datos correspondiente a la entidad de la clase que lo implementa. Al actualizarlo sobre la base de datos habrá campos del recordset que sufran modi"
End Function
Public Function DeleteRs(ByVal rs As Recordset) As Long
Attribute DeleteRs.VB_Description = "Elimina un registro de una entidad. Para ello se le pasa un recordset desconectado con el registro que se desea eliminar. Si el recordset contiene varios registros eliminará el registro activo."
End Function
Public Function Filter(Optional ByVal strSelect As String, _
                       Optional ByVal strWhere As String, _
                       Optional ByVal strOrderBy As String) As Recordset
Attribute Filter.VB_Description = "Devuelve un recordset desconectado con los registros  resultantes de la consulta aplicada sobre la entidad correspondiente a la clase que lo implementa."
End Function
Public Function NewRs() As Recordset
Attribute NewRs.VB_Description = "Retorna una recordset desconectado y sin registros. Unícamente tiene la estructa de campos de la tabla de la entidad correspondiente a la clase que lo implementa"
End Function
Public Function NewRsForm() As Recordset
Attribute NewRsForm.VB_Description = "Retorna un recordset desconectado con un registro  de la tabla de la entidad correspondiente a la clase que lo implementa. Todos los campos del resgistro del recordset retornado estarán vacios, excepto los campos que sean de tipo Autonumérico y los campo"
End Function
Public Function GoToRecord(ByVal enRecordNavigator As RecordNavigator, _
                           Optional ByVal rs As Recordset, _
                           Optional ByVal strFilter As String) As Recordset
End Function

