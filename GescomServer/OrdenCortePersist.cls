VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "OrdenCortePersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext

Private mudtProps As OrdenCorteProps

Private Function GetState() As String
    Dim udtData As OrdenCorteData
    
    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As OrdenCorteData
    
    udtData.Buffer = Buffer
    LSet mudtProps = udtData
    
End Sub

Public Function Save(ByVal Buffer As String, ByRef ItemsBuffer As String) As String
    Dim rsDocumento As Recordset
    Dim strSQL As String
    Dim objPersist As OrdenCorteItemsPersist
    
    On Error GoTo errh
    
    SetState Buffer
    
    ' Si vamos a insertar un registro nuevo, comprobamos que el documento a insertar
    ' no existe en la base de datos (lo que nos daría un error de clave duplicada).
    ' Para ello obtenemos el mayor numero de la tabla de Ordenes de corte y de los contadores
    ' de la empresa.
'    If mudtProps.IsNew Then mudtProps.Numero = GetMaxOrdenCorte()
    
    Set rsDocumento = New Recordset
    strSQL = "SELECT * FROM OrdenesCorte WHERE OrdenCorteID=" & _
        mudtProps.OrdenCorteID
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsDocumento.AddNew
    
    With rsDocumento
        .Fields("Fecha") = mudtProps.Fecha
        .Fields("FechaCorte") = mudtProps.FechaCorte
        .Fields("Numero") = mudtProps.Numero
        .Fields("ArticuloID") = mudtProps.ArticuloID
        .Fields("Observaciones") = mudtProps.Observaciones & vbNullString
        .Fields("TemporadaID") = mudtProps.TemporadaID
        .Fields("EmpresaID") = mudtProps.EmpresaID
    
        .Update
    
        If mudtProps.IsNew Then
           .Bookmark = .Bookmark
           mudtProps.OrdenCorteID = .Fields("OrdenCorteID")
        End If
        .Close
    End With
    
    Set rsDocumento = Nothing
    Save = GetState
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.OrdenCorteItemsPersist")
    ItemsBuffer = objPersist.Save(mudtProps.OrdenCorteID, ItemsBuffer)
    Set objPersist = Nothing
    
    mobjContext.SetComplete
    
    Exit Function
    
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

Public Sub DeleteObject(ByVal OrdenCorteID As Long, ByVal ItemsBuffer As String)
    Dim cnGescom As Connection
    Dim objPersist As OrdenCorteItemsPersist
    Dim strSave As String

    On Error GoTo errh
    
    ' Provocamos el borrado de los items de orden de corte
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.OrdenCorteItemsPersist")
    strSave = objPersist.Save(OrdenCorteID, ItemsBuffer)
    Set objPersist = Nothing
    
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM OrdenesCorte WHERE OrdenCorteID=" & _
        OrdenCorteID

    cnGescom.Close
    Set cnGescom = Nothing
    
    mobjContext.SetComplete
    
    Exit Sub
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Sub

Public Function Fetch(ByVal OrdenCorteID As Long) As String
    Dim rsDocumento As Recordset
    Dim strSQL As String
    
    On Error GoTo errh
    strSQL = "SELECT * FROM vOrdenesCorte WHERE OrdenCorteID=" & _
        OrdenCorteID
    Set rsDocumento = New Recordset
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    With rsDocumento
        mudtProps.OrdenCorteID = .Fields("OrdenCorteID")
        mudtProps.Fecha = .Fields("Fecha")
        mudtProps.FechaCorte = .Fields("FechaCorte")
        mudtProps.Numero = .Fields("Numero")
        mudtProps.ArticuloID = .Fields("ArticuloID")
        mudtProps.Observaciones = .Fields("Observaciones") & vbNullString
        mudtProps.Nombre = .Fields("Nombre") & vbNullString
        mudtProps.TemporadaID = .Fields("TemporadaID")
        mudtProps.EmpresaID = .Fields("EmpresaID")

        .Close
    End With

    Set rsDocumento = Nothing
    
    Fetch = GetState
    mobjContext.SetComplete
    
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function


Public Function GetMaxOrdenCorte() As Long
    Dim rsDocumento As Recordset
    Dim strSQL As String
    Dim MaxNumero As Long
    
    On Error GoTo errh
    Set rsDocumento = New Recordset
    strSQL = "SELECT max(numero) as numero FROM OrdenesCorte WHERE EmpresaID=" & _
      mudtProps.EmpresaID
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    MaxNumero = IIf(rsDocumento.Fields("Numero") >= mudtProps.Numero, _
                           rsDocumento.Fields("Numero") + 1, _
                           mudtProps.Numero)
    
    rsDocumento.Close
    
    Set rsDocumento = Nothing
    
    GetMaxOrdenCorte = MaxNumero
    mobjContext.SetComplete
    
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

Private Function ObjectControl_CanBePooled() As Boolean

    ObjectControl_CanBePooled = False
  
End Function

Private Sub ObjectControl_Activate()

    Set mobjContext = GetObjectContext

End Sub

Private Sub ObjectControl_Deactivate()

    Set mobjContext = Nothing

End Sub

