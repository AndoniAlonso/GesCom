VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "PedidoVentaPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext

Private mudtProps As PedidoVentaProps

Private Function GetState() As String
    Dim udtData As PedidoVentaData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As PedidoVentaData

    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String, ByRef ItemsBuffer As String) As String
    Dim rsDocumento As Recordset
    Dim objDatoComercialPersist As DatoComercialPersist
    Dim udtDatoComercialProps As DatoComercialProps
    Dim strSQL As String
    Dim objPersist As PedidoVentaItemsPersist
    
    On Error GoTo errh
    
    SetState Buffer
    Set objDatoComercialPersist = mobjContext.CreateInstance("GCServerMTS.DatoComercialPersist")
    mudtProps.DatoComercial.Buffer = objDatoComercialPersist.Save(mudtProps.DatoComercial.Buffer)
    LSet udtDatoComercialProps = mudtProps.DatoComercial
    mudtProps.DatoComercialID = udtDatoComercialProps.DatoComercialID
    Set objDatoComercialPersist = Nothing
  
    ' Si vamos a insertar un registro nuevo, comprobamos que el documento a insertar
    ' no existe en la base de datos (lo que nos daría un error de clave duplicada).
    ' Para ello obtenemos el mayor numero de la tabla de pedidos y de los contadores
    ' de la empresa.
'    If mudtProps.IsNew Then _
'        mudtProps.Numero = GetMaxPedidoVenta()
'
    Set rsDocumento = New Recordset
    strSQL = "SELECT * FROM PedidosVenta WHERE PedidoVentaID=" & _
        mudtProps.PedidoVentaID
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsDocumento.AddNew
    
    With rsDocumento
        .Fields("ClienteID") = mudtProps.ClienteID
        .Fields("Fecha") = mudtProps.Fecha
        .Fields("FechaEntrega") = mudtProps.FechaEntrega
        .Fields("FechaTopeServicio") = mudtProps.FechaTopeServicio
        .Fields("Numero") = mudtProps.Numero
        .Fields("Observaciones") = mudtProps.Observaciones
        .Fields("TotalBrutoPTA") = mudtProps.TotalBrutoPTA
        .Fields("TotalBrutoEUR") = mudtProps.TotalBrutoEUR
        .Fields("RepresentanteID") = mudtProps.RepresentanteID
        .Fields("TransportistaID") = mudtProps.TransportistaID
        .Fields("FormaPagoID") = mudtProps.FormaPagoID
        .Fields("DatoComercialID") = mudtProps.DatoComercialID
        .Fields("TemporadaID") = mudtProps.TemporadaID
        .Fields("EmpresaID") = mudtProps.EmpresaID
        
        .Update

        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.PedidoVentaID = .Fields("PedidoVentaID")
        End If
        .Close
    End With

    Set rsDocumento = Nothing
    Save = GetState
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.PedidoVentaItemsPersist")
    ItemsBuffer = objPersist.Save(mudtProps.PedidoVentaID, ItemsBuffer)
    Set objPersist = Nothing
    
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

Public Sub DeleteObject(ByVal PedidoVentaID As Long, ByVal DatoComercialID As Long, ByVal ItemsBuffer As String)
    Dim objDatoComercialPersist As DatoComercialPersist
    Dim cnGescom As Connection
    Dim objPersist As PedidoVentaItemsPersist
    Dim strSave As String

    On Error GoTo errh
    
    ' Provocamos el borrado de los items de pedido
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.PedidoVentaItemsPersist")
    strSave = objPersist.Save(PedidoVentaID, ItemsBuffer)
    Set objPersist = Nothing
    
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"
  
    cnGescom.Execute "DELETE FROM PedidosVenta WHERE PedidoVentaID=" & _
        PedidoVentaID

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

Public Function Fetch(ByVal PedidoVentaID As Long, _
        Optional ByVal Numero As Long, Optional ByVal EmpresaID As Long) As String
    Dim rsDocumento As Recordset
    Dim strSQL As String
    Dim objDatoComercialPersist As DatoComercialPersist

    On Error GoTo errh
    If PedidoVentaID = 0 Then
        strSQL = "SELECT * FROM PedidosVenta WHERE Numero=" & _
            Numero & " AND EmpresaID=" & EmpresaID
    Else
        strSQL = "SELECT * FROM PedidosVenta WHERE PedidoVentaID=" & _
            PedidoVentaID
    End If
    Set rsDocumento = New Recordset
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsDocumento
        mudtProps.PedidoVentaID = .Fields("PedidoVentaID")
        mudtProps.ClienteID = .Fields("ClienteID")
        mudtProps.Fecha = .Fields("Fecha")
        mudtProps.FechaEntrega = .Fields("FechaEntrega")
        mudtProps.FechaTopeServicio = IIf(IsNull(.Fields("FechaTopeServicio")), CDate("00:00:00"), .Fields("FechaTopeServicio"))
        mudtProps.Numero = .Fields("Numero")
        mudtProps.Observaciones = .Fields("Observaciones") & vbNullString
        mudtProps.TotalBrutoEUR = .Fields("TotalBrutoEUR")
        mudtProps.TotalBrutoPTA = .Fields("TotalBrutoPTA")
        mudtProps.RepresentanteID = .Fields("RepresentanteID")
        mudtProps.TransportistaID = .Fields("TransportistaID")
        mudtProps.FormaPagoID = .Fields("FormaPagoID")
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

Public Function GetMaxPedidoVenta() As Long
    Dim rsDocumento As Recordset
    Dim strSQL As String
    Dim MaxNumero As Long

    On Error GoTo errh
    Set rsDocumento = New Recordset
    strSQL = "SELECT MAX(Numero) as Numero FROM PedidosVenta WHERE EmpresaID=" & _
        mudtProps.EmpresaID & " AND TemporadaID=" & mudtProps.TemporadaID
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
  
    MaxNumero = IIf(rsDocumento.Fields("Numero") >= mudtProps.Numero, _
        rsDocumento.Fields("Numero") + 1, _
        mudtProps.Numero)

    rsDocumento.Close
  
    Set rsDocumento = Nothing
  
    GetMaxPedidoVenta = MaxNumero
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

