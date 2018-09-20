VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "AlbaranVentaPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext

Private mudtProps As AlbaranVentaProps

Private Function GetState() As String
    Dim udtData As AlbaranVentaData
    
    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As AlbaranVentaData
    
    udtData.Buffer = Buffer
    LSet mudtProps = udtData
    
End Sub

Public Function Save(ByVal Buffer As String, ByRef ItemsBuffer As String) As String
    Dim rsDocumento As Recordset
    Dim objDatoComercialPersist As DatoComercialPersist
    Dim udtDatoComercialProps As DatoComercialProps
    Dim strSQL As String
    Dim objPersist As AlbaranVentaItemsPersist
    
    On Error GoTo errh
    
    SetState Buffer
    Set objDatoComercialPersist = mobjContext.CreateInstance("GCServerMTS.DatoComercialPersist")
    mudtProps.DatoComercial.Buffer = objDatoComercialPersist.Save(mudtProps.DatoComercial.Buffer)
    LSet udtDatoComercialProps = mudtProps.DatoComercial
    mudtProps.DatoComercialID = udtDatoComercialProps.DatoComercialID
    Set objDatoComercialPersist = Nothing
    
    ' Si vamos a insertar un registro nuevo, comprobamos que el documento a insertar
    ' no existe en la base de datos (lo que nos daría un error de clave duplicada).
    ' Para ello obtenemos el mayor numero de la tabla de Albaranes y de los contadores
    ' de la empresa.
'    If mudtProps.IsNew Then _
'        mudtProps.Numero = GetMaxAlbaranVenta()
    
    Set rsDocumento = New Recordset
    strSQL = "SELECT * FROM AlbaranesVenta WHERE AlbaranVentaID=" & _
        mudtProps.AlbaranVentaID
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsDocumento.AddNew
    
    With rsDocumento
        .Fields("ClienteID") = mudtProps.ClienteID
        .Fields("Fecha") = mudtProps.Fecha
        .Fields("Numero") = mudtProps.Numero
        .Fields("NuestraReferencia") = mudtProps.NuestraReferencia
        .Fields("SuReferencia") = mudtProps.SuReferencia
        .Fields("Observaciones") = mudtProps.Observaciones
        .Fields("Bultos") = mudtProps.Bultos
        .Fields("PesoNeto") = mudtProps.PesoNeto
        .Fields("PesoBruto") = mudtProps.PesoBruto
        .Fields("PortesPTA") = mudtProps.PortesPTA
        .Fields("PortesEUR") = mudtProps.PortesEUR
        .Fields("EmbalajesPTA") = mudtProps.EmbalajesPTA
        .Fields("EmbalajesEUR") = mudtProps.EmbalajesEUR
        .Fields("TotalBrutoPTA") = mudtProps.TotalBrutoPTA
        .Fields("TotalBrutoEUR") = mudtProps.TotalBrutoEUR
        .Fields("RepresentanteID") = mudtProps.RepresentanteID
        .Fields("TransportistaID") = mudtProps.TransportistaID
        .Fields("FormaPagoID") = mudtProps.FormaPagoID
        .Fields("DatoComercialID") = mudtProps.DatoComercialID
        .Fields("TemporadaID") = mudtProps.TemporadaID
        .Fields("EmpresaID") = mudtProps.EmpresaID
        .Fields("FacturadoAB") = mudtProps.FacturadoAB
        .Fields("FacturaVentaIDA") = IIf(mudtProps.FacturaVentaIDA, mudtProps.FacturaVentaIDA, Null)
        .Fields("FacturaVentaIDB") = IIf(mudtProps.FacturaVentaIDB, mudtProps.FacturaVentaIDB, Null)
        .Fields("AlmacenID") = IIf(mudtProps.AlmacenID, mudtProps.AlmacenID, Null)
        .Fields("CentroGestionID") = IIf(mudtProps.CentroGestionID, mudtProps.CentroGestionID, Null)
        .Fields("TerminalID") = IIf(mudtProps.TerminalID, mudtProps.TerminalID, Null)
        .Fields("MedioPagoID") = mudtProps.MedioPagoID
    
        .Update
    
        If mudtProps.IsNew Then
           .Bookmark = .Bookmark
           mudtProps.AlbaranVentaID = .Fields("AlbaranVentaID")
        End If
        .Close
    End With
    
    Set rsDocumento = Nothing
    Save = GetState
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.AlbaranVentaItemsPersist")
    ItemsBuffer = objPersist.Save(mudtProps.AlbaranVentaID, ItemsBuffer)
    Set objPersist = Nothing
    
    mobjContext.SetComplete
    
    Exit Function
    
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

Public Sub DeleteObject(ByVal AlbaranVentaID As Long, ByVal DatoComercialID As Long, ByVal ItemsBuffer As String)
    Dim objDatoComercialPersist As DatoComercialPersist
    Dim cnGescom As Connection
    Dim objPersist As AlbaranVentaItemsPersist
    Dim strSave As String

    On Error GoTo errh
    
    ' Provocamos el borrado de los items de albarán
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.AlbaranVentaItemsPersist")
    strSave = objPersist.Save(AlbaranVentaID, ItemsBuffer)
    Set objPersist = Nothing
    
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM AlbaranesVenta WHERE AlbaranVentaID=" & _
        AlbaranVentaID

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

Public Function Fetch(ByVal AlbaranVentaID As Long) As String
    
    Dim rsDocumento As Recordset
    Dim strSQL As String
    Dim objDatoComercialPersist As DatoComercialPersist
    
    On Error GoTo errh
    strSQL = "SELECT * FROM AlbaranesVenta WHERE AlbaranVentaID=" & _
        AlbaranVentaID
    Set rsDocumento = New Recordset
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    With rsDocumento
        mudtProps.AlbaranVentaID = .Fields("AlbaranVentaID")
        mudtProps.ClienteID = .Fields("ClienteID")
        mudtProps.Fecha = .Fields("Fecha")
        mudtProps.Numero = .Fields("Numero")
        mudtProps.NuestraReferencia = .Fields("NuestraReferencia") & vbNullString
        mudtProps.SuReferencia = .Fields("SuReferencia") & vbNullString
        mudtProps.Observaciones = .Fields("Observaciones") & vbNullString
        mudtProps.Bultos = .Fields("Bultos")
        mudtProps.PesoNeto = .Fields("PesoNeto")
        mudtProps.PesoBruto = .Fields("PesoBruto")
        mudtProps.PortesPTA = .Fields("PortesPTA")
        mudtProps.PortesEUR = .Fields("PortesEUR")
        mudtProps.EmbalajesPTA = .Fields("EmbalajesPTA")
        mudtProps.EmbalajesEUR = .Fields("EmbalajesEUR")
        mudtProps.TotalBrutoPTA = .Fields("TotalBrutoPTA")
        mudtProps.TotalBrutoEUR = .Fields("TotalBrutoEUR")
        mudtProps.RepresentanteID = .Fields("RepresentanteID")
        mudtProps.TransportistaID = .Fields("TransportistaID")
        mudtProps.FormaPagoID = .Fields("FormaPagoID")
        mudtProps.DatoComercialID = .Fields("DatoComercialID")
        mudtProps.TemporadaID = .Fields("TemporadaID")
        mudtProps.EmpresaID = .Fields("EmpresaID")
        mudtProps.FacturadoAB = .Fields("FacturadoAB")
        mudtProps.FacturaVentaIDA = IIf(IsNull(.Fields("FacturaVentaIDA")), 0, .Fields("FacturaVentaIDA"))
        mudtProps.FacturaVentaIDB = IIf(IsNull(.Fields("FacturaVentaIDB")), 0, .Fields("FacturaVentaIDB"))
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


Public Function GetMaxAlbaranVenta() As Long

    Dim rsDocumento As Recordset
    Dim strSQL As String
    Dim MaxNumero As Long
    
    On Error GoTo errh
    Set rsDocumento = New Recordset
    strSQL = "SELECT max(numero) as numero FROM AlbaranesVenta WHERE EmpresaID=" & _
      mudtProps.EmpresaID & " AND TemporadaID=" & mudtProps.TemporadaID
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    MaxNumero = IIf(rsDocumento.Fields("Numero") >= mudtProps.Numero, _
                           rsDocumento.Fields("Numero") + 1, _
                           mudtProps.Numero)
    
    rsDocumento.Close
    
    Set rsDocumento = Nothing
    
    GetMaxAlbaranVenta = MaxNumero
    mobjContext.SetComplete
    
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

Public Sub SaveAlbaranAB(ByVal BufferFacturaVentaA As String, ByVal ItemsBufferFacturaVentaA As String, ByVal CobrosPagosBufferFacturaVentaA As String, bolGenerarFacturaA As Boolean, _
                         ByVal BufferFacturaVentaB As String, ByVal ItemsBufferFacturaVentaB As String, ByVal CobrosPagosBufferFacturaVentaB As String, bolGenerarFacturaB As Boolean, _
                         ByVal BufferAlbaranVenta As String, ByVal ItemsBufferAlbaranVenta As String)
'Dim objFacturaVentaAPersist As FacturaVentaPersist
'Dim objFacturaVentaBPersist As FacturaVentaPersist
'Dim strBufferFacturaVentaA As String
'Dim strBufferFacturaVentaB As String
'Dim mudtPropsA As FacturaVentaProps
'Dim mudtPropsB As FacturaVentaProps
'Dim udtData As FacturaVentaData
'
'
'
'    On Error GoTo errh
'    Set objFacturaVentaAPersist = mobjContext.CreateInstance("GCServerMTS.FacturaVentaPersist")
'    Set objFacturaVentaBPersist = mobjContext.CreateInstance("GCServerMTS.FacturaVentaPersist")
'
'    If bolGenerarFacturaA Then strBufferFacturaVentaA = objFacturaVentaAPersist.Save(BufferFacturaVentaA, ItemsBufferFacturaVentaA, CobrosPagosBufferFacturaVentaA)
'    If bolGenerarFacturaB Then strBufferFacturaVentaB = objFacturaVentaBPersist.Save(BufferFacturaVentaB, ItemsBufferFacturaVentaB, CobrosPagosBufferFacturaVentaB)
'
'    udtData.Buffer = strBufferFacturaVentaA
'    LSet mudtPropsA = udtData
'
'    udtData.Buffer = strBufferFacturaVentaB
'    LSet mudtPropsB = udtData
'
'
'    SetState BufferAlbaranVenta
'    mudtProps.IsDirty = True
'    mudtProps.FacturadoAB = True
'    If bolGenerarFacturaA Then mudtProps.FacturaVentaIDA = mudtPropsA.FacturaVentaID
'    If bolGenerarFacturaB Then mudtProps.FacturaVentaIDB = mudtPropsB.FacturaVentaID
'    BufferAlbaranVenta = GetState
'
'    Me.Save BufferAlbaranVenta, ItemsBufferAlbaranVenta
'    mobjContext.SetComplete
'    Exit Sub
'errh:
'    mobjContext.SetAbort
'    With Err
'        .Raise .Number, .Source, .Description
'    End With
End Sub

Public Function SaveAlbaranesAB(ByVal BufferFacturaVentaA As String, ByVal ItemsBufferFacturaVentaA As String, ByVal CobrosPagosBufferFacturaVentaA As String, bolGenerarFacturaA As Boolean, _
                         ByVal BufferFacturaVentaB As String, ByVal ItemsBufferFacturaVentaB As String, ByVal CobrosPagosBufferFacturaVentaB As String, bolGenerarFacturaB As Boolean, _
                         ByVal colAlbaranesPersist As Collection) As String
Dim objFacturaVentaAPersist As FacturaVentaPersist
Dim objFacturaVentaBPersist As FacturaVentaPersist
Dim strBufferFacturaVentaA As String
Dim strBufferFacturaVentaB As String
Dim mudtPropsA As FacturaVentaProps
Dim mudtPropsB As FacturaVentaProps
Dim udtData As FacturaVentaData
Dim i As Integer
    

    On Error GoTo errh
    Set objFacturaVentaAPersist = mobjContext.CreateInstance("GCServerMTS.FacturaVentaPersist")
    Set objFacturaVentaBPersist = mobjContext.CreateInstance("GCServerMTS.FacturaVentaPersist")
    
    If bolGenerarFacturaA Then
        strBufferFacturaVentaA = objFacturaVentaAPersist.Save(BufferFacturaVentaA, ItemsBufferFacturaVentaA, CobrosPagosBufferFacturaVentaA)
        udtData.Buffer = strBufferFacturaVentaA
        LSet mudtPropsA = udtData
    End If
    
    If bolGenerarFacturaB Then
        strBufferFacturaVentaB = objFacturaVentaBPersist.Save(BufferFacturaVentaB, ItemsBufferFacturaVentaB, CobrosPagosBufferFacturaVentaB)
        udtData.Buffer = strBufferFacturaVentaB
        LSet mudtPropsB = udtData
    End If
    
    EnlacesAlbaranes mudtPropsA.FacturaVentaID, mudtPropsB.FacturaVentaID
    SaveAlbaranesAB = CStr(mudtPropsA.FacturaVentaID) & "|" & CStr(mudtPropsB.FacturaVentaID)
    mobjContext.SetComplete
    Exit Function
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

' Al crear una factura de venta, se ocupa de registrar las relaciones entre los albaranes facturados y las facturas creadas.
Private Sub EnlacesAlbaranes(ByVal FacturaVentaIDA As Long, ByVal FacturaVentaIDB As Long)
    Dim cnGescom As Connection

    On Error GoTo errh
    
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    ' Primero actualizamos las lineas de albaranes de la factura A
    cnGescom.Execute "UPDATE AlbaranVentaItems SET Situacion = 'F', " & _
                     "       FacturaVentaItemIDA = FacturaVentaItemID, " & _
                     "       FacturadoAB = 1 " & _
                     "FROM FacturaVentaItems " & _
                     "WHERE FacturaVentaItems.FacturaVentaID = " & FacturaVentaIDA & " " & _
                     "AND  AlbaranVentaItems.AlbaranVentaItemID = FacturaVentaItems.AlbaranVentaItemID"

    ' despues los de la factura B
    cnGescom.Execute "UPDATE AlbaranVentaItems SET Situacion = 'F', " & _
                     "       FacturaVentaItemIDB = FacturaVentaItemID, " & _
                     "       FacturadoAB = 1 " & _
                     "FROM FacturaVentaItems " & _
                     "WHERE FacturaVentaItems.FacturaVentaID = " & FacturaVentaIDB & " " & _
                     "AND  AlbaranVentaItems.AlbaranVentaItemID = FacturaVentaItems.AlbaranVentaItemID"
                     
    ' despues las cabeceras de los albaranes A y B
    cnGescom.Execute "UPDATE AlbaranesVenta " & _
                    "SET FacturadoAB = 1," & _
                    "    FacturaVentaIDA = " & IIf(FacturaVentaIDA, FacturaVentaIDA, "NULL") & "," & _
                    "    FacturaVentaIDB = " & IIf(FacturaVentaIDB, FacturaVentaIDB, "NULL") & " " & _
                    "FROM albaranventaitems," & _
                    "     facturaventaitems " & _
                    "WHERE FacturaVentaItems.FacturaVentaID in (" & FacturaVentaIDA & ", " & FacturaVentaIDB & ") " & _
                    "AND AlbaranVentaItems.AlbaranVentaItemID = FacturaVentaItems.AlbaranVentaItemID " & _
                    "AND AlbaranVentaItems.AlbaranVentaID = AlbaranesVenta.AlbaranVentaID"
                   
    ' despues el ID de factura "gemela" de la factura A
    cnGescom.Execute "UPDATE FacturasVenta " & _
                    "SET FacturaVentaIDAB = " & FacturaVentaIDB & _
                    "WHERE FacturaVentaID = " & FacturaVentaIDA
                    
    ' despues el ID de factura "gemela" de la factura B
    cnGescom.Execute "UPDATE FacturasVenta " & _
                    "SET FacturaVentaIDAB = " & FacturaVentaIDA & _
                    "WHERE FacturaVentaID = " & FacturaVentaIDB
                    
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

