VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "FacturaVentaPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext

Private mudtProps As FacturaVentaProps

Private Function GetState() As String
    Dim udtData As FacturaVentaData
    
    LSet udtData = mudtProps
    GetState = udtData.Buffer
    
End Function

Private Sub SetState(Buffer As String)
    Dim udtData As FacturaVentaData
    
    udtData.Buffer = Buffer
    LSet mudtProps = udtData
    
End Sub

Public Function Save(ByVal Buffer As String, ByRef ItemsBuffer As String, ByRef CobrosPagosBuffer As String) As String
    Dim rsDocumento As Recordset
    Dim objDatoComercialPersist As DatoComercialPersist
    Dim udtDatoComercialProps As DatoComercialProps
    Dim strSQL As String
    Dim objPersist As FacturaVentaItemsPersist
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
    If mudtProps.IsNew Then _
        mudtProps.Numero = GetMaxFacturaVenta(mudtProps.TemporadaID, mudtProps.EmpresaID, mudtProps.Numero)
      
    Set rsDocumento = New Recordset
    strSQL = "SELECT * FROM FacturasVenta WHERE FacturaVentaID=" & _
        mudtProps.FacturaVentaID
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsDocumento.AddNew
    
    With rsDocumento
        .Fields("ClienteID") = mudtProps.ClienteID
        .Fields("Fecha") = mudtProps.Fecha
        .Fields("Numero") = mudtProps.Numero
        .Fields("NuestraReferencia") = mudtProps.NuestraReferencia
        .Fields("SuReferencia") = mudtProps.SuReferencia
        .Fields("Observaciones") = mudtProps.Observaciones
        .Fields("SituacionContable") = mudtProps.SituacionContable
        .Fields("Bultos") = mudtProps.Bultos
        .Fields("PesoNeto") = mudtProps.PesoNeto
        .Fields("PesoBruto") = mudtProps.PesoBruto
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
        .Fields("RepresentanteID") = mudtProps.RepresentanteID
        .Fields("TransportistaID") = mudtProps.TransportistaID
        .Fields("FormaPagoID") = mudtProps.FormaPagoID
        .Fields("DatoComercialID") = mudtProps.DatoComercialID
        .Fields("TemporadaID") = mudtProps.TemporadaID
        .Fields("EmpresaID") = mudtProps.EmpresaID
        .Fields("FacturaVentaIDAB") = IIf(mudtProps.FacturaVentaIDAB, mudtProps.FacturaVentaIDAB, Null)
        .Fields("AlmacenID") = IIf(mudtProps.AlmacenID, mudtProps.AlmacenID, Null)
        .Fields("CentroGestionID") = IIf(mudtProps.CentroGestionID, mudtProps.CentroGestionID, Null)
        .Fields("TerminalID") = IIf(mudtProps.TerminalID, mudtProps.TerminalID, Null)
        .Fields("MedioPagoID") = mudtProps.MedioPagoID
        
        .Update
        
        If mudtProps.IsNew Then
           .Bookmark = .Bookmark
           mudtProps.FacturaVentaID = .Fields("FacturaVentaID")
        End If
        .Close
    End With
    
    Set rsDocumento = Nothing
    Save = GetState
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.FacturaVentaItemsPersist")
    ItemsBuffer = objPersist.Save(mudtProps.FacturaVentaID, ItemsBuffer)
    Set objPersist = Nothing
    
    Set objCobrosPagosPersist = mobjContext.CreateInstance("GCServerMTS.CobrosPagosPersist")
    CobrosPagosBuffer = objCobrosPagosPersist.Save(mudtProps.FacturaVentaID, CobrosPagosBuffer)
    Set objCobrosPagosPersist = Nothing
    
    mobjContext.SetComplete
    
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

Public Sub DeleteObject(ByVal FacturaVentaID As Long, ByVal DatoComercialID As Long, ByVal ItemsBuffer As String, ByRef CobrosPagosBuffer As String)
    Dim objDatoComercialPersist As DatoComercialPersist
    Dim cnGescom As Connection
    Dim objPersist As FacturaVentaItemsPersist
    Dim objCobrosPagosPersist As CobrosPagosPersist
    Dim strSave As String

    On Error GoTo errh
    
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    ' Eliminamos las restricciones que pueda haber con los albaranes
    cnGescom.Execute "UPDATE AlbaranVentaItems SET FacturaVentaItemIDA=NULL, FacturadoAB=0 FROM FacturaVentaItems WHERE FacturaVentaItemIDA=FacturaVentaItems.FacturaVentaID AND FacturaVentaItems.FacturaVentaID=" & _
        FacturaVentaID

    ' Eliminamos las restricciones que pueda haber con los albaranes
    cnGescom.Execute "UPDATE AlbaranVentaItems SET FacturaVentaItemIDB=NULL, FacturadoAB=0 FROM FacturaVentaItems WHERE FacturaVentaItemIDB=FacturaVentaItems.FacturaVentaID AND FacturaVentaItems.FacturaVentaID=" & _
        FacturaVentaID

    ' Eliminamos las restricciones que pueda haber con los albaranes
    cnGescom.Execute "UPDATE AlbaranesVenta SET FacturaVentaIDA=NULL, FacturadoAB=0 FROM FacturasVenta WHERE AlbaranesVenta.FacturaVentaIDA=" & _
        FacturaVentaID

    ' Eliminamos las restricciones que pueda haber con los albaranes
    cnGescom.Execute "UPDATE AlbaranesVenta SET FacturaVentaIDB=NULL, FacturadoAB=0 FROM FacturasVenta WHERE AlbaranesVenta.FacturaVentaIDB=" & _
        FacturaVentaID

    ' Borramos los items de factura
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.FacturaVentaItemsPersist")
    strSave = objPersist.Save(FacturaVentaID, ItemsBuffer)
    Set objPersist = Nothing
    
    ' Borramos los items de cobros y pagos
    Set objCobrosPagosPersist = mobjContext.CreateInstance("GCServerMTS.CobrosPagosPersist")
    strSave = objCobrosPagosPersist.Save(FacturaVentaID, CobrosPagosBuffer)
    Set objCobrosPagosPersist = Nothing
    
    cnGescom.Execute "DELETE FROM FacturasVenta WHERE FacturaVentaID=" & _
        FacturaVentaID

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

Public Sub DeleteFacturaAB(ByVal FacturaVentaIDA As Long, ByVal DatoComercialIDA As Long, ByVal ItemsBufferA As String, ByRef CobrosPagosBufferA As String, _
                           ByVal FacturaVentaIDB As Long, ByVal DatoComercialIDB As Long, ByVal ItemsBufferB As String, ByRef CobrosPagosBufferB As String)
    
    Me.DeleteObject FacturaVentaIDA, ByVal DatoComercialIDA, ByVal ItemsBufferA, CobrosPagosBufferA
    Me.DeleteObject FacturaVentaIDB, ByVal DatoComercialIDB, ByVal ItemsBufferB, CobrosPagosBufferB

End Sub

Public Function Fetch(ByVal FacturaVentaID As Long) As String
    Dim rsDocumento As Recordset
    Dim strSQL As String
    Dim objDatoComercialPersist As DatoComercialPersist
    
    On Error GoTo errh
    strSQL = "SELECT * FROM FacturasVenta WHERE FacturaVentaID=" & _
        FacturaVentaID
    Set rsDocumento = New Recordset
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    With rsDocumento
        mudtProps.FacturaVentaID = .Fields("FacturaVentaID")
        mudtProps.ClienteID = .Fields("ClienteID")
        mudtProps.Fecha = .Fields("Fecha")
        mudtProps.Numero = .Fields("Numero")
        mudtProps.NuestraReferencia = .Fields("NuestraReferencia") & vbNullString
        mudtProps.SuReferencia = .Fields("SuReferencia") & vbNullString
        mudtProps.Observaciones = .Fields("Observaciones") & vbNullString
        mudtProps.SituacionContable = .Fields("SituacionContable") & vbNullString
        mudtProps.Bultos = .Fields("Bultos")
        mudtProps.PesoNeto = .Fields("PesoNeto")
        mudtProps.PesoBruto = .Fields("PesoBruto")
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
        mudtProps.RepresentanteID = .Fields("RepresentanteID")
        mudtProps.TransportistaID = .Fields("TransportistaID")
        mudtProps.FormaPagoID = .Fields("FormaPagoID")
        mudtProps.DatoComercialID = .Fields("DatoComercialID")
        mudtProps.TemporadaID = .Fields("TemporadaID")
        mudtProps.EmpresaID = .Fields("EmpresaID")
        mudtProps.FacturaVentaIDAB = IIf(IsNull(.Fields("FacturaVentaIDAB")), 0, .Fields("FacturaVentaIDAB"))
        mudtProps.AlmacenID = IIf(IsNull(.Fields("AlmacenID")), 0, .Fields("AlmacenID"))
        mudtProps.CentroGestionID = IIf(IsNull(.Fields("CentroGestionID")), 0, .Fields("CentroGestionID"))
        mudtProps.TerminalID = IIf(IsNull(.Fields("TerminalID")), 0, .Fields("TerminalID"))
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

Public Sub Contabilizar(ByVal FacturaVentaID As Long, _
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
    cnGescom.Execute "UPDATE FacturasVenta SET SituacionContable ='" & SituacionContable & _
                     "' WHERE FacturaVentaID=" & FacturaVentaID
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

Public Function GetMaxFacturaVenta(TemporadaID As Long, EmpresaID As Long, Numero As Long) As Long
    Dim cnGescom As ADODB.Connection
    Dim cmdGescom As ADODB.Command
    Dim prmGescom As ADODB.Parameter
    
    On Error GoTo errh
    
    ' Abre una conexión.
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    cnGescom.CursorLocation = adUseClient
    
    ' Abre el objeto Command con un parámetro.
    Set cmdGescom = New ADODB.Command
    cmdGescom.CommandText = "spGetMaxFacturaVenta"
    cmdGescom.CommandType = adCmdStoredProc
    
    ' Obtiene el valor del parámetro y adjunta el parámetro.
    Set prmGescom = cmdGescom.CreateParameter("TemporadaID", adInteger, adParamInput)
    cmdGescom.Parameters.Append prmGescom
    prmGescom.Value = TemporadaID
   
    Set prmGescom = cmdGescom.CreateParameter("EmpresaID", adInteger, adParamInput)
    cmdGescom.Parameters.Append prmGescom
    prmGescom.Value = EmpresaID
   
    Set prmGescom = cmdGescom.CreateParameter("Numero", adInteger, adParamInputOutput)
    cmdGescom.Parameters.Append prmGescom
    prmGescom.Value = Numero
   
    ' Crea un objeto Recordset al ejecutar el comando.
    Set cmdGescom.ActiveConnection = cnGescom
    
    cmdGescom.Execute
   
    GetMaxFacturaVenta = prmGescom.Value
    
    cnGescom.Close
    
    Set prmGescom = Nothing
    Set cmdGescom = Nothing
    Set cnGescom = Nothing
    
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

