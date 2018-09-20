VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "CobrosPagosPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ObjectControl

Private mobjContext As ObjectContext

Public Function Fetch(ByVal FacturaID As Long, Tipo As String) As String
    Dim strSQL As String
    Dim rsCobroPago As Recordset
    Dim objBuffer As Buffer
    Dim objPersist As CobroPagoPersist
    Dim udtProps As CobroPagoProps
    
    On Error GoTo errh
    Set objBuffer = New Buffer
    strSQL = "SELECT CobroPagoID FROM CobrosPagos WHERE FacturaID=" & FacturaID & " and tipo = '" & Tipo & "'"
    Set rsCobroPago = New Recordset
    rsCobroPago.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.CobroPagoPersist")
    
    With objPersist
        objBuffer.Initialize Len(udtProps), 20
        
        Do While Not rsCobroPago.EOF
            objBuffer.Add .Fetch(rsCobroPago("CobroPagoID"))
            rsCobroPago.MoveNext
        Loop
        
    End With
    
    Set objPersist = Nothing
    rsCobroPago.Close
    Set rsCobroPago = Nothing
    Fetch = objBuffer.GetState
    mobjContext.SetComplete
    
    Exit Function
    
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

Public Function FetchRemesa(ByVal BancoID As Long, ByVal FechaDomiciliacion As Date) As String
    Dim strSQL As String
    Dim rsCobroPago As Recordset
    Dim objBuffer As Buffer
    Dim objPersist As CobroPagoPersist
    Dim udtProps As CobroPagoProps
    
    On Error GoTo errh
    Set objBuffer = New Buffer
    strSQL = "SELECT CobroPagoID FROM CobrosPagos WHERE BancoID = " & BancoID & " AND FechaDomiciliacion = '" & FechaDomiciliacion & "' "
    Set rsCobroPago = New Recordset
    rsCobroPago.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.CobroPagoPersist")
    
    With objPersist
        objBuffer.Initialize Len(udtProps), 20
        
        Do While Not rsCobroPago.EOF
            objBuffer.Add .Fetch(rsCobroPago("CobroPagoID"))
            rsCobroPago.MoveNext
        Loop
        
    End With
    
    Set objPersist = Nothing
    rsCobroPago.Close
    Set rsCobroPago = Nothing
    FetchRemesa = objBuffer.GetState
    mobjContext.SetComplete
    
    Exit Function
    
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

Public Function Save(ByVal FacturaID As Long, ByRef Buffer As String) As String
    Dim objPersist As CobroPagoPersist
    Dim objBuffer As Buffer
    Dim objUpdatedCobrosPagos As Buffer
    Dim lngIndex As Long
    Dim strBuffer As String
    
    On Error GoTo errh
    Set objBuffer = New Buffer
    objBuffer.SetState Buffer
    Set objUpdatedCobrosPagos = New Buffer
    objUpdatedCobrosPagos.Initialize objBuffer.Length, objBuffer.Count
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.CobroPagoPersist")
    
    For lngIndex = 1 To objBuffer.Count
        With objPersist
            strBuffer = .SaveChild(FacturaID, objBuffer.Item(lngIndex))
            If Len(strBuffer) > 0 Then _
                objUpdatedCobrosPagos.Add strBuffer
        End With
        
    Next
    
    Set objPersist = Nothing
    Save = objUpdatedCobrosPagos.GetState
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


