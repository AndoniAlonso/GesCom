VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ClientePersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
  
Private mudtProps As ClienteProps

Private Function GetState() As String
    Dim udtData As ClienteData
   
    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As ClienteData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String
    Dim rsCliente As Recordset
    Dim strSQL As String
    Dim objDireccionPersist As DireccionPersist
    Dim udtDireccionProps As DireccionProps
    Dim objCuentaBancariaPersist As CuentaBancariaPersist
    Dim udtCuentaBancariaProps As CuentaBancariaProps
    Dim objDatoComercialPersist As DatoComercialPersist
    Dim udtDatoComercialProps As DatoComercialProps
    Dim objDatoComercialBPersist As DatoComercialPersist
    Dim udtDatoComercialBProps As DatoComercialProps
  

    On Error GoTo errh
    SetState Buffer
  
    Set objDireccionPersist = mobjContext.CreateInstance("GCServerMTS.DireccionPersist")
    mudtProps.DireccionFiscal.Buffer = objDireccionPersist.Save(mudtProps.DireccionFiscal.Buffer)
    LSet udtDireccionProps = mudtProps.DireccionFiscal
    mudtProps.DireccionFiscalID = udtDireccionProps.DireccionID
  
    mudtProps.DireccionEntrega.Buffer = objDireccionPersist.Save(mudtProps.DireccionEntrega.Buffer)
    LSet udtDireccionProps = mudtProps.DireccionEntrega
    mudtProps.DireccionEntregaID = udtDireccionProps.DireccionID
    Set objDireccionPersist = Nothing
  
    Set objCuentaBancariaPersist = mobjContext.CreateInstance("GCServerMTS.CuentaBancariaPersist")
    mudtProps.CuentaBancaria.Buffer = objCuentaBancariaPersist.Save(mudtProps.CuentaBancaria.Buffer)
    LSet udtCuentaBancariaProps = mudtProps.CuentaBancaria
    mudtProps.CuentaBancariaID = udtCuentaBancariaProps.CuentaBancariaID
    Set objCuentaBancariaPersist = Nothing
  
    Set objDatoComercialPersist = mobjContext.CreateInstance("GCServerMTS.DatoComercialPersist")
    mudtProps.DatoComercial.Buffer = objDatoComercialPersist.Save(mudtProps.DatoComercial.Buffer)
    LSet udtDatoComercialProps = mudtProps.DatoComercial
    mudtProps.DatoComercialID = udtDatoComercialProps.DatoComercialID
    Set objDatoComercialPersist = Nothing
    
    Set objDatoComercialBPersist = mobjContext.CreateInstance("GCServerMTS.DatoComercialPersist")
    mudtProps.DatoComercialB.Buffer = objDatoComercialBPersist.Save(mudtProps.DatoComercialB.Buffer)
    LSet udtDatoComercialBProps = mudtProps.DatoComercialB
    mudtProps.DatoComercialBID = udtDatoComercialBProps.DatoComercialID
    Set objDatoComercialBPersist = Nothing
    
    Set rsCliente = New Recordset
    strSQL = "SELECT * FROM Clientes WHERE ClienteID=" & _
        mudtProps.ClienteID
    rsCliente.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsCliente.AddNew
  
    With rsCliente
        .Fields("Nombre") = mudtProps.Nombre
        .Fields("Titular") = mudtProps.Titular
        .Fields("Contacto") = mudtProps.Contacto
        .Fields("DNINIF") = mudtProps.DNINIF
        .Fields("DireccionFiscalID") = mudtProps.DireccionFiscalID
        .Fields("DireccionEntregaID") = mudtProps.DireccionEntregaID
        .Fields("TransportistaID") = mudtProps.TransportistaID
        .Fields("RepresentanteID") = mudtProps.RepresentanteID
        .Fields("CuentaBancariaID") = mudtProps.CuentaBancariaID
        .Fields("FormaPagoID") = mudtProps.FormaPagoID
        .Fields("CuentaContable") = mudtProps.CuentaContable
        .Fields("DatoComercialID") = mudtProps.DatoComercialID
        .Fields("DatoComercialBID") = mudtProps.DatoComercialBID
        .Fields("DiaPago1") = mudtProps.DiaPago1
        .Fields("DiaPago2") = mudtProps.DiaPago2
        .Fields("DiaPago3") = mudtProps.DiaPago3
        .Fields("PorcFacturacionAB") = mudtProps.PorcFacturacionAB
        .Update

        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.ClienteID = .Fields("ClienteID")
        End If
        .Close
  End With
  
  Set rsCliente = Nothing
  
  Save = GetState
    mobjContext.SetComplete
  
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal ClienteID As Long, ByVal DireccionFiscalID As Long, _
                        ByVal DireccionEntregaID As Long, ByVal CuentaBancariaID As Long, _
                        ByVal DatoComercialID As Long, _
                        ByVal DatoComercialBID As Long)
    Dim objDireccionPersist As DireccionPersist
    Dim objCuentaBancariaPersist As CuentaBancariaPersist
    Dim objDatoComercialPersist As DatoComercialPersist
    Dim objDatoComercialBPersist As DatoComercialPersist
    Dim cnGescom As Connection


    On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"
  
    cnGescom.Execute "DELETE FROM Clientes WHERE ClienteID=" & _
        ClienteID
    
    Set objDireccionPersist = mobjContext.CreateInstance("GCServerMTS.DireccionPersist")
    objDireccionPersist.DeleteObject (DireccionFiscalID)
    objDireccionPersist.DeleteObject (DireccionEntregaID)
    Set objDireccionPersist = Nothing

    Set objCuentaBancariaPersist = mobjContext.CreateInstance("GCServerMTS.CuentaBancariaPersist")
    objCuentaBancariaPersist.DeleteObject (CuentaBancariaID)
    Set objCuentaBancariaPersist = Nothing

    Set objDatoComercialPersist = mobjContext.CreateInstance("GCServerMTS.DatoComercialPersist")
    objDatoComercialPersist.DeleteObject (DatoComercialID)
    Set objDatoComercialPersist = Nothing

    Set objDatoComercialBPersist = mobjContext.CreateInstance("GCServerMTS.DatoComercialPersist")
    objDatoComercialBPersist.DeleteObject (DatoComercialBID)
    Set objDatoComercialBPersist = Nothing

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

Public Function Fetch(ByVal ClienteID As Long) As String
    Dim rsCliente As Recordset
    Dim strSQL As String
    Dim objDireccionPersist As DireccionPersist
    Dim objCuentaBancariaPersist As CuentaBancariaPersist
    Dim objDatoComercialPersist As DatoComercialPersist
    Dim objDatoComercialBPersist As DatoComercialPersist
  

    On Error GoTo errh
    strSQL = "SELECT * FROM Clientes WHERE ClienteID=" & _
        ClienteID
    Set rsCliente = New Recordset
    rsCliente.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsCliente
        mudtProps.ClienteID = .Fields("ClienteID")
        mudtProps.Nombre = .Fields("Nombre") & vbNullString
        mudtProps.Titular = .Fields("Titular") & vbNullString
        mudtProps.Contacto = .Fields("Contacto") & vbNullString
        mudtProps.DNINIF = .Fields("DNINIF") & vbNullString
        mudtProps.DireccionFiscalID = .Fields("DireccionFiscalID")
        mudtProps.DireccionEntregaID = .Fields("DireccionEntregaID")
        mudtProps.TransportistaID = .Fields("TransportistaID")
        mudtProps.RepresentanteID = .Fields("RepresentanteID")
        mudtProps.CuentaBancariaID = .Fields("CuentaBancariaID")
        mudtProps.FormaPagoID = .Fields("FormaPagoID")
        mudtProps.CuentaContable = .Fields("CuentaContable") & vbNullString
        mudtProps.DatoComercialID = .Fields("DatoComercialID")
        mudtProps.DatoComercialBID = .Fields("DatoComercialBID")
        mudtProps.DiaPago1 = .Fields("DiaPago1")
        mudtProps.DiaPago2 = .Fields("DiaPago2")
        mudtProps.DiaPago3 = .Fields("DiaPago3")
        mudtProps.PorcFacturacionAB = .Fields("PorcFacturacionAB")
        .Close
    End With

    Set rsCliente = Nothing
  
    Set objDireccionPersist = mobjContext.CreateInstance("GCServerMTS.DireccionPersist")
    mudtProps.DireccionFiscal.Buffer = objDireccionPersist.Fetch(mudtProps.DireccionFiscalID)
    mudtProps.DireccionEntrega.Buffer = objDireccionPersist.Fetch(mudtProps.DireccionEntregaID)
    Set objDireccionPersist = Nothing
  
    Set objCuentaBancariaPersist = mobjContext.CreateInstance("GCServerMTS.CuentaBancariaPersist")
    mudtProps.CuentaBancaria.Buffer = objCuentaBancariaPersist.Fetch(mudtProps.CuentaBancariaID)
    Set objCuentaBancariaPersist = Nothing
  
    Set objDatoComercialPersist = mobjContext.CreateInstance("GCServerMTS.DatoComercialPersist")
    mudtProps.DatoComercial.Buffer = objDatoComercialPersist.Fetch(mudtProps.DatoComercialID)
    Set objDatoComercialPersist = Nothing
  
    Set objDatoComercialBPersist = mobjContext.CreateInstance("GCServerMTS.DatoComercialPersist")
    mudtProps.DatoComercialB.Buffer = objDatoComercialBPersist.Fetch(mudtProps.DatoComercialBID)
    Set objDatoComercialBPersist = Nothing
  
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

