VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "ProveedorPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
  
Private mudtProps As ProveedorProps

Private Function GetState() As String
    Dim udtData As ProveedorData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As ProveedorData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String
    Dim rsProveedor As Recordset
    Dim strSQL As String
    Dim objDireccionPersist As DireccionPersist
    Dim udtDireccionProps As DireccionProps
    Dim objCuentaBancariaPersist As CuentaBancariaPersist
    Dim udtCuentaBancariaProps As CuentaBancariaProps
    Dim objDatoComercialPersist As DatoComercialPersist
    Dim udtDatoComercialProps As DatoComercialProps
  

    On Error GoTo errh
    SetState Buffer
  
    Set objDireccionPersist = mobjContext.CreateInstance("GCServerMTS.DireccionPersist")
    mudtProps.Direccion.Buffer = objDireccionPersist.Save(mudtProps.Direccion.Buffer)
    LSet udtDireccionProps = mudtProps.Direccion
    mudtProps.DireccionID = udtDireccionProps.DireccionID
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
  
    Set rsProveedor = New Recordset
    strSQL = "SELECT * FROM Proveedores WHERE ProveedorID=" & _
        mudtProps.ProveedorID
    rsProveedor.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsProveedor.AddNew
    
    With rsProveedor
        .Fields("Nombre") = mudtProps.Nombre
        .Fields("Titular") = mudtProps.Titular
        .Fields("Contacto") = mudtProps.Contacto
        .Fields("DNINIF") = mudtProps.DNINIF
        .Fields("DireccionID") = mudtProps.DireccionID
        .Fields("BancoID") = mudtProps.BancoID
        .Fields("CuentaBancariaID") = mudtProps.CuentaBancariaID
        .Fields("TransportistaID") = mudtProps.TransportistaID
        .Fields("FormaPagoID") = mudtProps.FormaPagoID
        .Fields("MedioPagoID") = IIf(mudtProps.MedioPagoID = 0, Null, mudtProps.MedioPagoID)
        .Fields("CuentaContable") = mudtProps.CuentaContable
        .Fields("CuentaContrapartida") = mudtProps.CuentaContrapartida
        .Fields("DatoComercialID") = mudtProps.DatoComercialID
        .Fields("Codigo") = mudtProps.Codigo
        .Fields("TipoProveedor") = mudtProps.TipoProveedor
        .Update

        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.ProveedorID = .Fields("ProveedorID")
        End If
        .Close
    End With

    Set rsProveedor = Nothing
    Save = GetState
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal ProveedorID As Long, ByVal DireccionID As Long, _
                        ByVal CuentaBancariaID As Long, ByVal DatoComercialID As Long)
    Dim objDireccionPersist As DireccionPersist
    Dim objCuentaBancariaPersist As CuentaBancariaPersist
    Dim objDatoComercialPersist As DatoComercialPersist
    Dim cnGescom As Connection


    On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM Proveedores WHERE ProveedorID=" & _
        ProveedorID
  
    Set objDireccionPersist = mobjContext.CreateInstance("GCServerMTS.DireccionPersist")
    objDireccionPersist.DeleteObject (DireccionID)
    Set objDireccionPersist = Nothing

    Set objCuentaBancariaPersist = mobjContext.CreateInstance("GCServerMTS.CuentaBancariaPersist")
    objCuentaBancariaPersist.DeleteObject (CuentaBancariaID)
    Set objCuentaBancariaPersist = Nothing

    Set objDatoComercialPersist = mobjContext.CreateInstance("GCServerMTS.DatoComercialPersist")
    objDatoComercialPersist.DeleteObject (DatoComercialID)
    Set objDatoComercialPersist = Nothing

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

Public Function Fetch(ByVal ProveedorID As Long) As String
    Dim rsProveedor As Recordset
    Dim strSQL As String
    Dim objDireccionPersist As DireccionPersist
    Dim objCuentaBancariaPersist As CuentaBancariaPersist
    Dim objDatoComercialPersist As DatoComercialPersist
    

    On Error GoTo errh
    strSQL = "SELECT * FROM Proveedores WHERE ProveedorID=" & _
        ProveedorID
    Set rsProveedor = New Recordset
    rsProveedor.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsProveedor
        mudtProps.ProveedorID = .Fields("ProveedorID")
        mudtProps.Nombre = .Fields("Nombre") & vbNullString
        mudtProps.Titular = .Fields("Titular") & vbNullString
        mudtProps.Contacto = .Fields("Contacto") & vbNullString
        mudtProps.DNINIF = .Fields("DNINIF") & vbNullString
        mudtProps.DireccionID = .Fields("DireccionID")
        mudtProps.BancoID = .Fields("BancoID")
        mudtProps.TransportistaID = .Fields("TransportistaID")
        mudtProps.CuentaBancariaID = .Fields("CuentaBancariaID")
        mudtProps.FormaPagoID = .Fields("FormaPagoID")
        mudtProps.MedioPagoID = IIf(IsNull(.Fields("MedioPagoID")), 0, .Fields("MedioPagoID"))
        mudtProps.CuentaContable = .Fields("CuentaContable") & vbNullString
        mudtProps.CuentaContrapartida = .Fields("CuentaContrapartida") & vbNullString
        mudtProps.DatoComercialID = .Fields("DatoComercialID")
        mudtProps.Codigo = .Fields("Codigo")
        mudtProps.TipoProveedor = .Fields("TipoProveedor")
        .Close
    End With

    Set rsProveedor = Nothing
  
    Set objDireccionPersist = mobjContext.CreateInstance("GCServerMTS.DireccionPersist")
    mudtProps.Direccion.Buffer = objDireccionPersist.Fetch(mudtProps.DireccionID)
    Set objDireccionPersist = Nothing
    
    Set objCuentaBancariaPersist = mobjContext.CreateInstance("GCServerMTS.CuentaBancariaPersist")
    mudtProps.CuentaBancaria.Buffer = objCuentaBancariaPersist.Fetch(mudtProps.CuentaBancariaID)
    Set objCuentaBancariaPersist = Nothing
    
    Set objDatoComercialPersist = mobjContext.CreateInstance("GCServerMTS.DatoComercialPersist")
    mudtProps.DatoComercial.Buffer = objDatoComercialPersist.Fetch(mudtProps.DatoComercialID)
    Set objDatoComercialPersist = Nothing
    
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

