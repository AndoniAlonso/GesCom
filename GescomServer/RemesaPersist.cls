VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "RemesaPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext

Private mudtProps As RemesaProps

Private Function GetState() As String
    Dim udtData As RemesaData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

'Private Sub SetState(Buffer As String)
'    Dim udtData As RemesaData
'
'    udtData.Buffer = Buffer
'    LSet mudtProps = udtData
'
'End Sub
'
Public Sub Save(ByRef ItemsBuffer As String)
    Dim objPersist As CobrosPagosPersist
    
    On Error GoTo errh
        
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.CobrosPagosPersist")
    ItemsBuffer = objPersist.Save(0, ItemsBuffer)
    Set objPersist = Nothing
    
    mobjContext.SetComplete
    
    Exit Sub
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Sub

'Public Sub DeleteObject(ByVal RemesaID As Long, ByVal DatoComercialID As Long, ByVal ItemsBuffer As String, ByRef CobrosPagosBuffer As String)
'    Dim objPersist As CobrosPagosPersist
'    Dim strSave As String
'
'    On Error GoTo errh
'
'    ' Borramos los items de factura
'    Set objPersist = mobjContext.CreateInstance("GCServerMTS.CobrosPagosPersist")
'    strSave = objPersist.Save(RemesaID, ItemsBuffer)
'    Set objPersist = Nothing
'
'    mobjContext.SetComplete
'
'    Exit Sub
'
'errh:
'    mobjContext.SetAbort
'    With Err
'        .Raise .Number, .Source, .Description
'    End With
'
'End Sub
'
Public Function Fetch(ByVal BancoID As Long, ByVal FechaDomiciliacion As Date) As String
    Dim rsDocumento As Recordset
    Dim strSQL As String
    
    On Error GoTo errh
    strSQL = "SELECT * FROM vRemesas WHERE BancoID=" & _
        BancoID & " AND FechaDomiciliacion = '" & FechaDomiciliacion & "' "
    Set rsDocumento = New Recordset
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    With rsDocumento
        mudtProps.FechaDomiciliacion = .Fields("FechaDomiciliacion")
        mudtProps.BancoID = .Fields("BancoID")
        mudtProps.NombreEntidad = .Fields("NombreEntidad") & vbNullString
        mudtProps.SituacionComercial = .Fields("SituacionComercial") & vbNullString
        mudtProps.NumeroEfectos = .Fields("NumeroEfectos")
        mudtProps.ImporteEUR = .Fields("ImporteEUR")
        mudtProps.ImportePTA = .Fields("ImportePTA")
        
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

Public Sub Contabilizar(ByVal BancoID As Long, ByVal FechaDomiciliacion As Date, _
                        ByVal SituacionContable As String, _
                        Buffer() As Byte)
    Dim objAsientoPersist As AsientoPersist
    Dim cnGescom As Connection
    
    On Error GoTo errh
    Set objAsientoPersist = mobjContext.CreateInstance("GCServerMTS.AsientoPersist")
    objAsientoPersist.Save Buffer
    Set objAsientoPersist = Nothing
        
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"
    cnGescom.Execute "UPDATE CobrosPagos SET SituacionContable ='" & SituacionContable & _
                     "' WHERE BancoID=" & BancoID & " AND " & _
                     "FechaDomiciliacion = '" & FechaDomiciliacion & "'"
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

Private Function ObjectControl_CanBePooled() As Boolean

    ObjectControl_CanBePooled = False
  
End Function

Private Sub ObjectControl_Activate()

    Set mobjContext = GetObjectContext

End Sub

Private Sub ObjectControl_Deactivate()

    Set mobjContext = Nothing

End Sub

