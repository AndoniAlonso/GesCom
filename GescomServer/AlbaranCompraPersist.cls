VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "AlbaranCompraPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ObjectControl

Private mobjContext As ObjectContext

Private mudtProps As AlbaranCompraProps

Private Function GetState() As String
    Dim udtData As AlbaranCompraData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As AlbaranCompraData

    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String, ByRef ItemsBuffer As String, ByRef ArticulosBuffer As String) As String
    Dim rsDocumento As Recordset
    Dim objDatoComercialPersist As DatoComercialPersist
    Dim udtDatoComercialProps As DatoComercialProps
    Dim strSQL As String
    Dim objPersist As AlbaranCompraItemsPersist
    
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
'        mudtProps.Numero = GetMaxAlbaranCompra()
        
    Set rsDocumento = New Recordset
    strSQL = "SELECT * FROM AlbaranesCompra WHERE AlbaranCompraID=" & _
        mudtProps.AlbaranCompraID
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsDocumento.AddNew
    
    With rsDocumento
        .Fields("ProveedorID") = mudtProps.ProveedorID
        .Fields("Fecha") = mudtProps.Fecha
        .Fields("Numero") = mudtProps.Numero
        .Fields("NuestraReferencia") = mudtProps.NuestraReferencia
        .Fields("SuReferencia") = mudtProps.SuReferencia
        .Fields("Observaciones") = mudtProps.Observaciones
        .Fields("PortesPTA") = mudtProps.PortesPTA
        .Fields("PortesEUR") = mudtProps.PortesEUR
        .Fields("EmbalajesPTA") = mudtProps.EmbalajesPTA
        .Fields("EmbalajesEUR") = mudtProps.EmbalajesEUR
        .Fields("TotalBrutoPTA") = mudtProps.TotalBrutoPTA
        .Fields("TotalBrutoEUR") = mudtProps.TotalBrutoEUR
        .Fields("TransportistaID") = mudtProps.TransportistaID
        .Fields("DatoComercialID") = mudtProps.DatoComercialID
        .Fields("TemporadaID") = mudtProps.TemporadaID
        .Fields("EmpresaID") = mudtProps.EmpresaID
    
        .Update
    
        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.AlbaranCompraID = .Fields("AlbaranCompraID")
        End If
        .Close
    End With

    Set rsDocumento = Nothing
    Save = GetState
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.AlbaranCompraItemsPersist")
    ItemsBuffer = objPersist.Save(mudtProps.AlbaranCompraID, ItemsBuffer)
    ArticulosBuffer = objPersist.SaveArticulos(mudtProps.AlbaranCompraID, ArticulosBuffer)
    Set objPersist = Nothing
    
    mobjContext.SetComplete
  
  Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

Public Sub DeleteObject(ByVal AlbaranCompraID As Long, ByVal DatoComercialID As Long, ByVal ItemsBuffer As String, ByVal ArticulosBuffer As String)
    Dim objDatoComercialPersist As DatoComercialPersist
    Dim cnGescom As Connection
    Dim objPersist As AlbaranCompraItemsPersist
    Dim strSave As String
    Dim strArticulos As String

    On Error GoTo errh
    
    ' Provocamos el borrado de los items de albarán
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.AlbaranCompraItemsPersist")
    strSave = objPersist.Save(AlbaranCompraID, ItemsBuffer)
    strArticulos = objPersist.SaveArticulos(AlbaranCompraID, ArticulosBuffer)
    Set objPersist = Nothing
    
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM AlbaranesCompra WHERE AlbaranCompraID=" & _
        AlbaranCompraID

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

Public Function Fetch(ByVal AlbaranCompraID As Long) As String

    Dim rsDocumento As Recordset
    Dim strSQL As String
    Dim objDatoComercialPersist As DatoComercialPersist
    
    On Error GoTo errh
    strSQL = "SELECT * FROM AlbaranesCompra WHERE AlbaranCompraID=" & _
        AlbaranCompraID
    Set rsDocumento = New Recordset
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    With rsDocumento
        mudtProps.AlbaranCompraID = .Fields("AlbaranCompraID")
        mudtProps.ProveedorID = .Fields("ProveedorID")
        mudtProps.Fecha = .Fields("Fecha")
        mudtProps.Numero = .Fields("Numero")
        mudtProps.NuestraReferencia = .Fields("NuestraReferencia") & vbNullString
        mudtProps.SuReferencia = .Fields("SuReferencia") & vbNullString
        mudtProps.Observaciones = .Fields("Observaciones") & vbNullString
        mudtProps.PortesPTA = .Fields("PortesPTA")
        mudtProps.PortesEUR = .Fields("PortesEUR")
        mudtProps.EmbalajesPTA = .Fields("EmbalajesPTA")
        mudtProps.EmbalajesEUR = .Fields("EmbalajesEUR")
        mudtProps.TotalBrutoPTA = .Fields("TotalBrutoPTA")
        mudtProps.TotalBrutoEUR = .Fields("TotalBrutoEUR")
        mudtProps.TransportistaID = .Fields("TransportistaID")
        mudtProps.DatoComercialID = .Fields("DatoComercialID")
        mudtProps.TemporadaID = .Fields("TemporadaID")
        mudtProps.EmpresaID = .Fields("EmpresaID")
    
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

Public Function GetMaxAlbaranCompra() As Long

    Dim rsDocumento As Recordset
    Dim strSQL As String
    Dim MaxNumero As Long
    
    On Error GoTo errh
    Set rsDocumento = New Recordset
    strSQL = "SELECT max(numero) as numero FROM AlbaranesCompra WHERE EmpresaID=" & _
        mudtProps.EmpresaID & " AND TemporadaID=" & mudtProps.TemporadaID
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
  
    MaxNumero = IIf(rsDocumento.Fields("Numero") >= mudtProps.Numero, _
        rsDocumento.Fields("Numero") + 1, _
        mudtProps.Numero)

    rsDocumento.Close
  
    Set rsDocumento = Nothing
  
    GetMaxAlbaranCompra = MaxNumero
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

