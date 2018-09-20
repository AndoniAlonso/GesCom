VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CobroPagoPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext

Private mudtProps As CobroPagoProps

Private Function GetState() As String
    Dim udtData As CobroPagoData
    
    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As CobroPagoData
    
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String
    Dim rsCobroPago As Recordset
    Dim strSQL As String
    
    On Error GoTo errh
    
    SetState Buffer
    
    Set rsCobroPago = New Recordset
    strSQL = "SELECT * FROM CobrosPagos WHERE CobroPagoID=" & _
      mudtProps.CobroPagoID
    rsCobroPago.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsCobroPago.AddNew
    
    With rsCobroPago
      .Fields("Tipo") = mudtProps.Tipo
      .Fields("Vencimiento") = mudtProps.Vencimiento
      .Fields("PersonaID") = mudtProps.PersonaID
      .Fields("FormaPagoID") = mudtProps.FormaPagoID
      .Fields("FacturaID") = IIf(mudtProps.FacturaID, mudtProps.FacturaID, Null)
      .Fields("NumeroGiro") = mudtProps.NumeroGiro
      .Fields("SituacionComercial") = mudtProps.SituacionComercial
      .Fields("SituacionContable") = mudtProps.SituacionContable
      .Fields("ImportePTA") = mudtProps.ImportePTA
      .Fields("ImporteEUR") = mudtProps.ImporteEUR
      .Fields("FechaEmision") = mudtProps.FechaEmision
      .Fields("FechaDomiciliacion") = mudtProps.FechaDomiciliacion
      .Fields("FechaContable") = mudtProps.FechaContable
      .Fields("BancoID") = mudtProps.BancoID
      .Fields("EmpresaID") = mudtProps.EmpresaID
      .Fields("TemporadaID") = mudtProps.TemporadaID
      .Fields("MedioPagoID") = IIf(mudtProps.MedioPagoID, mudtProps.MedioPagoID, Null)

      .Update
    
      If mudtProps.IsNew Then
         .Bookmark = .Bookmark
         mudtProps.CobroPagoID = .Fields("CobroPagoID")
      End If
      .Close
    End With
    
    Set rsCobroPago = Nothing
    Save = GetState
    mobjContext.SetComplete
    
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

Public Sub DeleteObject(ByVal CobroPagoID As Long)
    Dim cnGescom As Connection
     
    On Error GoTo errh
    
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    cnGescom.Execute "DELETE FROM CobrosPagos WHERE CobroPagoID=" & _
      CobroPagoID
    
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

Public Function Fetch(ByVal CobroPagoID As Long) As String
    Dim rsCobroPago As Recordset
    Dim strSQL As String
    
    On Error GoTo errh
    
    strSQL = "SELECT * FROM CobrosPagos WHERE CobroPagoID=" & _
      CobroPagoID
    Set rsCobroPago = New Recordset
    rsCobroPago.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    With rsCobroPago
      mudtProps.CobroPagoID = .Fields("CobroPagoID")
      mudtProps.Tipo = .Fields("Tipo") & vbNullString
      mudtProps.Vencimiento = .Fields("Vencimiento")
      mudtProps.PersonaID = .Fields("PersonaID")
      mudtProps.FormaPagoID = .Fields("FormaPagoID")
      mudtProps.FacturaID = IIf(IsNull(.Fields("FacturaID")), 0, .Fields("FacturaID"))
      mudtProps.NumeroGiro = .Fields("NumeroGiro")
      mudtProps.SituacionComercial = .Fields("SituacionComercial") & vbNullString
      mudtProps.SituacionContable = .Fields("SituacionContable") & vbNullString
      mudtProps.ImportePTA = .Fields("ImportePTA")
      mudtProps.ImporteEUR = .Fields("ImporteEUR")
      mudtProps.FechaEmision = .Fields("FechaEmision")
      mudtProps.FechaDomiciliacion = .Fields("FechaDomiciliacion")
      mudtProps.FechaContable = .Fields("FechaContable")
      mudtProps.BancoID = .Fields("BancoID")
      mudtProps.EmpresaID = .Fields("EmpresaID")
      mudtProps.TemporadaID = .Fields("TemporadaID")
      mudtProps.MedioPagoID = IIf(IsNull(.Fields("MedioPagoID")), 0, .Fields("MedioPagoID"))
      
      .Close
    End With
    
    Set rsCobroPago = Nothing
    Fetch = GetState
    mobjContext.SetComplete
    
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function


Public Function SaveChild(FacturaID As Long, Buffer As String) As String

    On Error GoTo errh
    
    SetState Buffer
    
    With mudtProps
      If .IsDeleted Then
        DeleteObject .CobroPagoID
        SaveChild = vbNullString
    
      ElseIf .IsDirty Or .IsNew Then
        ' Cuando viene de una remesa no se asigna el numero de factura.
        If FacturaID <> 0 Then .FacturaID = FacturaID
        SaveChild = Save(GetState)
    
      Else
        SaveChild = Buffer
        mobjContext.SetComplete
     
      End If
    
    End With
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

