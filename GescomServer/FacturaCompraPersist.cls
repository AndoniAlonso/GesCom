VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "FacturaCompraPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ObjectControl

Private mobjContext As ObjectContext

Private mudtProps As FacturaCompraProps

Private Function GetState() As String
    Dim udtData As FacturaCompraData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As FacturaCompraData

    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String, ByRef ItemsBuffer As String, ByRef CobrosPagosBuffer As String) As String
    Dim rsDocumento As Recordset
    Dim objDatoComercialPersist As DatoComercialPersist
    Dim udtDatoComercialProps As DatoComercialProps
    Dim strSQL As String
    Dim objPersist As FacturaCompraItemsPersist
    Dim objCobrosPagosPersist As CobrosPagosPersist

    On Error GoTo errh
  
    SetState Buffer
    Set objDatoComercialPersist = mobjContext.CreateInstance("GCServerMTS.DatoComercialPersist")
    mudtProps.DatoComercial.Buffer = objDatoComercialPersist.Save(mudtProps.DatoComercial.Buffer)
    LSet udtDatoComercialProps = mudtProps.DatoComercial
    mudtProps.DatoComercialID = udtDatoComercialProps.DatoComercialID
    Set objDatoComercialPersist = Nothing
  
    ' Si vamos a insertar un registro nuevo, comprobamos que el documento a insertar
    ' no existe en la base de datos (lo que nos daría un error de clave duplicada).
    ' Para ello obtenemos el mayor numero de la tabla de Facturas y de los contadores
    ' de la empresa.
'    If mudtProps.IsNew Then _
'        mudtProps.Numero = GetMaxFacturaCompra()
    
    Set rsDocumento = New Recordset
    strSQL = "SELECT * FROM FacturasCompra WHERE FacturaCompraID=" & _
        mudtProps.FacturaCompraID
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsDocumento.AddNew

    With rsDocumento
        .Fields("ProveedorID") = mudtProps.ProveedorID
        .Fields("Fecha") = mudtProps.Fecha
        .Fields("FechaContable") = mudtProps.FechaContable
        .Fields("Numero") = mudtProps.Numero
        .Fields("Sufijo") = mudtProps.Sufijo
        .Fields("NuestraReferencia") = mudtProps.NuestraReferencia
        .Fields("SuReferencia") = mudtProps.SuReferencia
        .Fields("Observaciones") = mudtProps.Observaciones
        .Fields("SituacionContable") = mudtProps.SituacionContable
        .Fields("BrutoPTA") = mudtProps.BrutoPTA
        .Fields("BrutoEUR") = mudtProps.BrutoEUR
        .Fields("DescuentoPTA") = mudtProps.DescuentoPTA
        .Fields("DescuentoEUR") = mudtProps.DescuentoEUR
        .Fields("PortesPTA") = mudtProps.PortesPTA
        .Fields("PortesEUR") = mudtProps.PortesEUR
        .Fields("EmbalajesPTA") = mudtProps.EmbalajesPTA
        .Fields("EmbalajesEUR") = mudtProps.EmbalajesEUR
        .Fields("BaseImponiblePTA") = mudtProps.BaseImponiblePTA
        .Fields("BaseImponibleEUR") = mudtProps.BaseImponibleEUR
        .Fields("IVAPTA") = mudtProps.IVAPTA
        .Fields("IVAEUR") = mudtProps.IVAEUR
        .Fields("RecargoPTA") = mudtProps.RecargoPTA
        .Fields("RecargoEUR") = mudtProps.RecargoEUR
        .Fields("NetoPTA") = mudtProps.NetoPTA
        .Fields("NetoEUR") = mudtProps.NetoEUR
        .Fields("BancoID") = IIf(mudtProps.BancoID, mudtProps.BancoID, Null)
        .Fields("TransportistaID") = mudtProps.TransportistaID
        .Fields("FormaPagoID") = mudtProps.FormaPagoID
        .Fields("DatoComercialID") = mudtProps.DatoComercialID
        .Fields("TemporadaID") = mudtProps.TemporadaID
        .Fields("EmpresaID") = mudtProps.EmpresaID
        .Fields("MedioPagoID") = mudtProps.MedioPagoID
    
        .Update

        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.FacturaCompraID = .Fields("FacturaCompraID")
        End If
        .Close
    End With

    Set rsDocumento = Nothing
    Save = GetState
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.FacturaCompraItemsPersist")
    ItemsBuffer = objPersist.Save(mudtProps.FacturaCompraID, ItemsBuffer)
    Set objPersist = Nothing
    
    Set objCobrosPagosPersist = mobjContext.CreateInstance("GCServerMTS.CobrosPagosPersist")
    CobrosPagosBuffer = objCobrosPagosPersist.Save(mudtProps.FacturaCompraID, CobrosPagosBuffer)
    Set objCobrosPagosPersist = Nothing
    
    mobjContext.SetComplete
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal FacturaCompraID As Long, ByVal DatoComercialID As Long, ByVal ItemsBuffer As String, ByRef CobrosPagosBuffer As String)
    Dim objDatoComercialPersist As DatoComercialPersist
    Dim cnGescom As Connection
    Dim objPersist As FacturaCompraItemsPersist
    Dim objCobrosPagosPersist As CobrosPagosPersist
    Dim strSave As String

    On Error GoTo errh
    
    ' Borramos los items de factura
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.FacturaCompraItemsPersist")
    strSave = objPersist.Save(FacturaCompraID, ItemsBuffer)
    Set objPersist = Nothing
    
    ' Borramos los items de cobros y pagos
    Set objCobrosPagosPersist = mobjContext.CreateInstance("GCServerMTS.CobrosPagosPersist")
    strSave = objCobrosPagosPersist.Save(FacturaCompraID, CobrosPagosBuffer)
    Set objCobrosPagosPersist = Nothing
    
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM FacturasCompra WHERE FacturaCompraID=" & _
    FacturaCompraID

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

Public Function Fetch(ByVal FacturaCompraID As Long) As String
    Dim rsDocumento As Recordset
    Dim strSQL As String
    Dim objDatoComercialPersist As DatoComercialPersist
    
    On Error GoTo errh
    strSQL = "SELECT * FROM FacturasCompra WHERE FacturaCompraID=" & _
        FacturaCompraID
    Set rsDocumento = New Recordset
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    With rsDocumento
        mudtProps.FacturaCompraID = .Fields("FacturaCompraID")
        mudtProps.ProveedorID = .Fields("ProveedorID")
        mudtProps.Fecha = .Fields("Fecha")
        mudtProps.FechaContable = .Fields("FechaContable")
        mudtProps.Numero = .Fields("Numero")
        mudtProps.Sufijo = .Fields("Sufijo")
        mudtProps.NuestraReferencia = .Fields("NuestraReferencia") & vbNullString
        mudtProps.SuReferencia = .Fields("SuReferencia") & vbNullString
        mudtProps.Observaciones = .Fields("Observaciones") & vbNullString
        mudtProps.SituacionContable = .Fields("SituacionContable") & vbNullString
        mudtProps.BrutoPTA = .Fields("BrutoPTA")
        mudtProps.BrutoEUR = .Fields("BrutoEUR")
        mudtProps.DescuentoPTA = .Fields("DescuentoPTA")
        mudtProps.DescuentoEUR = .Fields("DescuentoEUR")
        mudtProps.PortesPTA = .Fields("PortesPTA")
        mudtProps.PortesEUR = .Fields("PortesEUR")
        mudtProps.EmbalajesPTA = .Fields("EmbalajesPTA")
        mudtProps.EmbalajesEUR = .Fields("EmbalajesEUR")
        mudtProps.BaseImponiblePTA = .Fields("BaseImponiblePTA")
        mudtProps.BaseImponibleEUR = .Fields("BaseImponibleEUR")
        mudtProps.IVAPTA = .Fields("IVAPTA")
        mudtProps.IVAEUR = .Fields("IVAEUR")
        mudtProps.RecargoPTA = .Fields("RecargoPTA")
        mudtProps.RecargoEUR = .Fields("RecargoEUR")
        mudtProps.NetoPTA = .Fields("NetoPTA")
        mudtProps.NetoEUR = .Fields("NetoEUR")
        mudtProps.BancoID = IIf(IsNull(.Fields("BancoID")), 0, .Fields("BancoID"))
        mudtProps.TransportistaID = .Fields("TransportistaID")
        mudtProps.FormaPagoID = .Fields("FormaPagoID")
        mudtProps.DatoComercialID = .Fields("DatoComercialID")
        mudtProps.TemporadaID = .Fields("TemporadaID")
        mudtProps.EmpresaID = .Fields("EmpresaID")
        mudtProps.MedioPagoID = .Fields("MedioPagoID")
    
      .Close
    End With

    Set rsDocumento = Nothing
    
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

Public Sub Contabilizar(ByVal FacturaCompraID As Long, _
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
    cnGescom.Execute "UPDATE FacturasCompra SET SituacionContable ='" & SituacionContable & _
                     "' WHERE FacturaCompraID=" & FacturaCompraID
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

Public Function GetMaxFacturaCompra() As Long
    Dim rsDocumento As Recordset
    Dim strSQL As String
    Dim MaxNumero As Long

    ' URGENTE: realizar un tratamiento de errores coherente!!!!

    Set rsDocumento = New Recordset
    strSQL = "SELECT MAX(Numero) as Numero FROM FacturasCompra WHERE EmpresaID=" & _
        mudtProps.EmpresaID & " AND TemporadaID=" & mudtProps.TemporadaID
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    MaxNumero = IIf(rsDocumento.Fields("Numero") >= mudtProps.Numero, _
        rsDocumento.Fields("Numero") + 1, _
        mudtProps.Numero)

    rsDocumento.Close

    Set rsDocumento = Nothing

    GetMaxFacturaCompra = MaxNumero
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
