VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "EmpresaPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
  
Private mudtProps As EmpresaProps

Private Function GetState() As String

    Dim udtData As EmpresaData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

    Dim udtData As EmpresaData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String
    Dim rsEmpresa As Recordset
    Dim strSQL As String
    Dim objDireccionPersist As DireccionPersist
    Dim udtDireccionProps As DireccionProps
  

    On Error GoTo errh
    SetState Buffer
  
    Set objDireccionPersist = mobjContext.CreateInstance("GCServerMTS.DireccionPersist")
    mudtProps.Direccion.Buffer = objDireccionPersist.Save(mudtProps.Direccion.Buffer)
    LSet udtDireccionProps = mudtProps.Direccion
    mudtProps.DireccionID = udtDireccionProps.DireccionID
    Set objDireccionPersist = Nothing
  
    Set rsEmpresa = New Recordset
    strSQL = "SELECT * FROM Empresas WHERE EmpresaID=" & _
        mudtProps.EmpresaID
    rsEmpresa.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsEmpresa.AddNew
  
    With rsEmpresa
        .Fields("Nombre") = mudtProps.Nombre
        .Fields("codigo") = mudtProps.Codigo
        .Fields("Titular") = mudtProps.Titular
        .Fields("DNINIF") = mudtProps.DNINIF
        .Fields("Actividad") = mudtProps.Actividad
        .Fields("DireccionID") = mudtProps.DireccionID
        .Fields("PedidoCompras") = mudtProps.PedidoCompras
        .Fields("PedidoVentas") = mudtProps.PedidoVentas
        .Fields("AlbaranCompras") = mudtProps.AlbaranCompras
        .Fields("AlbaranVentas") = mudtProps.AlbaranVentas
        .Fields("FacturaCompras") = mudtProps.FacturaCompras
        .Fields("FacturaVentas") = mudtProps.FacturaVentas
        .Fields("OrdenCorte") = mudtProps.OrdenCorte
        .Fields("TratamientoIVA") = mudtProps.TratamientoIVA
        .Fields("EmpresaContawin") = mudtProps.EmpresaContawin
        .Fields("CodigoContawin") = mudtProps.CodigoContawin
        .Update

        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.EmpresaID = .Fields("EmpresaID")
        End If
        .Close
    End With

    Set rsEmpresa = Nothing
    Save = GetState
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal EmpresaID As Long, ByVal DireccionID As Long)
    Dim objDireccionPersist As DireccionPersist
    Dim cnGescom As Connection


    On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM Empresas WHERE EmpresaID=" & _
        EmpresaID
  
    Set objDireccionPersist = mobjContext.CreateInstance("GCServerMTS.DireccionPersist")
    objDireccionPersist.DeleteObject (DireccionID)
    Set objDireccionPersist = Nothing

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

Public Function Fetch(ByVal EmpresaID As Long) As String

    Dim rsEmpresa As Recordset
    Dim strSQL As String
    Dim objDireccionPersist As DireccionPersist
  

  On Error GoTo errh
    strSQL = "SELECT * FROM Empresas WHERE EmpresaID=" & _
        EmpresaID
    Set rsEmpresa = New Recordset
    rsEmpresa.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsEmpresa
        mudtProps.EmpresaID = .Fields("EmpresaID")
        mudtProps.Codigo = .Fields("codigo") & vbNullString
        mudtProps.Nombre = .Fields("Nombre") & vbNullString
        mudtProps.Titular = .Fields("Titular") & vbNullString
        mudtProps.DNINIF = .Fields("DNINIF") & vbNullString
        mudtProps.Actividad = .Fields("Actividad") & vbNullString
        mudtProps.DireccionID = .Fields("DireccionID")
        mudtProps.PedidoCompras = .Fields("PedidoCompras")
        mudtProps.PedidoVentas = .Fields("PedidoVentas")
        mudtProps.AlbaranCompras = .Fields("AlbaranCompras")
        mudtProps.AlbaranVentas = .Fields("AlbaranVentas")
        mudtProps.FacturaCompras = .Fields("FacturaCompras")
        mudtProps.FacturaVentas = .Fields("FacturaVentas")
        mudtProps.OrdenCorte = .Fields("OrdenCorte")
        mudtProps.TratamientoIVA = .Fields("TratamientoIVA")
        mudtProps.EmpresaContawin = .Fields("EmpresaContawin")
        mudtProps.CodigoContawin = .Fields("CodigoContawin")
        .Close
    End With

    Set rsEmpresa = Nothing
        
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


           
