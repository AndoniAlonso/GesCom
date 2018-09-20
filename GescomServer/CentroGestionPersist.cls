VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CentroGestionPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
 
Private mudtProps As CentroGestionProps

Private Function GetState() As String
    Dim udtData As CentroGestionData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As CentroGestionData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String
    Dim rsCentroGestion As Recordset
    Dim strSQL As String
    Dim objDireccionPersist As DireccionPersist
    Dim udtDireccionProps As DireccionProps
  
    On Error GoTo errh
    SetState Buffer
  
    Set objDireccionPersist = mobjContext.CreateInstance("GCServerMTS.DireccionPersist")
    mudtProps.Direccion.Buffer = objDireccionPersist.Save(mudtProps.Direccion.Buffer)
    LSet udtDireccionProps = mudtProps.Direccion
    mudtProps.DireccionID = udtDireccionProps.DireccionID
  
    Set rsCentroGestion = New Recordset
    strSQL = "SELECT * FROM CentrosGestion WHERE CentroGestionID=" & _
        mudtProps.CentroGestionID
    rsCentroGestion.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsCentroGestion.AddNew
  
    With rsCentroGestion
        .Fields("Nombre") = mudtProps.Nombre
        .Fields("DireccionID") = mudtProps.DireccionID
        .Fields("ContadorTicketID") = mudtProps.DireccionID
        .Fields("ContadorPedidoVentaID") = IIf(mudtProps.ContadorPedidoVentaID = 0, Null, mudtProps.ContadorPedidoVentaID)
        .Fields("ContadorAlbaranVentaID") = IIf(mudtProps.ContadorAlbaranVentaID = 0, Null, mudtProps.ContadorAlbaranVentaID)
        .Fields("ContadorFacturaVentaID") = IIf(mudtProps.ContadorFacturaVentaID = 0, Null, mudtProps.ContadorFacturaVentaID)
        .Fields("ContadorPedidoCompraID") = IIf(mudtProps.ContadorPedidoCompraID = 0, Null, mudtProps.ContadorPedidoCompraID)
        .Fields("ContadorAlbaranCompraID") = IIf(mudtProps.ContadorAlbaranCompraID = 0, Null, mudtProps.ContadorAlbaranCompraID)
        .Fields("ContadorFacturaCompraID") = IIf(mudtProps.ContadorFacturaCompraID = 0, Null, mudtProps.ContadorFacturaCompraID)
        .Fields("SedeCentral") = mudtProps.SedeCentral
        .Fields("EmpresaID") = mudtProps.EmpresaID
        
        .Update

        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.CentroGestionID = .Fields("CentroGestionID")
        End If
        .Close
    End With

    Set rsCentroGestion = Nothing
    Save = GetState
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal CentroGestionID As Long, DireccionID As Long)
    Dim cnGescom As Connection
    Dim objDireccionPersist As DireccionPersist

    On Error GoTo errh
    
    Set objDireccionPersist = mobjContext.CreateInstance("GCServerMTS.DireccionPersist")
    objDireccionPersist.DeleteObject (DireccionID)
    Set objDireccionPersist = Nothing

    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM CentrosGestion WHERE CentroGestionID=" & _
        CentroGestionID

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

Public Function Fetch(ByVal CentroGestionID As Long) As String
    Dim rsCentroGestion As Recordset
    Dim strSQL As String
    Dim objDireccionPersist As DireccionPersist
  
    On Error GoTo errh
    
     strSQL = "SELECT * FROM CentrosGestion WHERE CentroGestionID=" & _
         CentroGestionID

    Set rsCentroGestion = New Recordset
    rsCentroGestion.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsCentroGestion
        mudtProps.CentroGestionID = .Fields("CentroGestionID")
        mudtProps.Nombre = .Fields("Nombre")
        mudtProps.DireccionID = .Fields("DireccionID")
        mudtProps.ContadorTicketID = IIf(IsNull(.Fields("ContadorTicketID")), 0, .Fields("ContadorTicketID"))
        mudtProps.ContadorPedidoVentaID = IIf(IsNull(.Fields("ContadorPedidoVentaID")), 0, .Fields("ContadorPedidoVentaID"))
        mudtProps.ContadorAlbaranVentaID = IIf(IsNull(.Fields("ContadorAlbaranVentaID")), 0, .Fields("ContadorAlbaranVentaID"))
        mudtProps.ContadorFacturaVentaID = IIf(IsNull(.Fields("ContadorFacturaVentaID")), 0, .Fields("ContadorFacturaVentaID"))
        mudtProps.ContadorPedidoCompraID = IIf(IsNull(.Fields("ContadorPedidoCompraID")), 0, .Fields("ContadorPedidoCompraID"))
        mudtProps.ContadorAlbaranCompraID = IIf(IsNull(.Fields("ContadorAlbaranCompraID")), 0, .Fields("ContadorAlbaranCompraID"))
        mudtProps.ContadorFacturaCompraID = IIf(IsNull(.Fields("ContadorFacturaCompraID")), 0, .Fields("ContadorFacturaCompraID"))
        mudtProps.SedeCentral = .Fields("SedeCentral")
        mudtProps.EmpresaID = .Fields("EmpresaID")
        
        .Close
    End With

    Set rsCentroGestion = Nothing
    
    
    Set objDireccionPersist = mobjContext.CreateInstance("GCServerMTS.DireccionPersist")
    mudtProps.Direccion.Buffer = objDireccionPersist.Fetch(mudtProps.DireccionID)
    Set objDireccionPersist = Nothing
    
    Fetch = GetState
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
