VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "MoviArticuloPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
  
Private mudtProps As MoviArticuloProps

Private Function GetState() As String
    Dim udtData As MoviArticuloData
    
    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As MoviArticuloData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByRef Buffer As String) As String
    Dim rsMovimiento As Recordset
    Dim strSQL As String
    Dim objArticuloColorPersist As ArticuloColorPersist
    Dim udtArticuloColorData As ArticuloColorData
    Dim udtArticuloColor As ArticuloColorProps
    Dim objArticuloPersist As ArticuloPersist
    Dim udtArticuloData As ArticuloData
    Dim udtArticulo As ArticuloProps

    On Error GoTo errh
    SetState Buffer
  
    Set rsMovimiento = New Recordset
    strSQL = "SELECT * FROM MoviArticulos WHERE MoviArticuloID=" & _
        mudtProps.MoviArticuloID
    rsMovimiento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsMovimiento.AddNew
    
    If mudtProps.AlmacenID <> 0 Then
        ActualizarArticuloColorAlmacen
    End If
   ' Leemos el articulo-color del movimiento
    Set objArticuloColorPersist = mobjContext.CreateInstance("GCServerMTS.ArticuloColorPersist")
    udtArticuloColorData.Buffer = objArticuloColorPersist.Fetch(mudtProps.ArticuloColorID)
    LSet udtArticuloColor = udtArticuloColorData
    
    ' Calculamos el nuevo stock
    udtArticuloColor.IsDeleted = False
    udtArticuloColor.IsDirty = True
    udtArticuloColor.IsNew = False
    
    Select Case mudtProps.Tipo
        Case TMMReserva
            udtArticuloColor.StockPendienteT36 = udtArticuloColor.StockPendienteT36 + mudtProps.CantidadT36
            udtArticuloColor.StockPendienteT38 = udtArticuloColor.StockPendienteT38 + mudtProps.CantidadT38
            udtArticuloColor.StockPendienteT40 = udtArticuloColor.StockPendienteT40 + mudtProps.CantidadT40
            udtArticuloColor.StockPendienteT42 = udtArticuloColor.StockPendienteT42 + mudtProps.CantidadT42
            udtArticuloColor.StockPendienteT44 = udtArticuloColor.StockPendienteT44 + mudtProps.CantidadT44
            udtArticuloColor.StockPendienteT46 = udtArticuloColor.StockPendienteT46 + mudtProps.CantidadT46
            udtArticuloColor.StockPendienteT48 = udtArticuloColor.StockPendienteT48 + mudtProps.CantidadT48
            udtArticuloColor.StockPendienteT50 = udtArticuloColor.StockPendienteT50 + mudtProps.CantidadT50
            udtArticuloColor.StockPendienteT52 = udtArticuloColor.StockPendienteT52 + mudtProps.CantidadT52
            udtArticuloColor.StockPendienteT54 = udtArticuloColor.StockPendienteT54 + mudtProps.CantidadT54
            udtArticuloColor.StockPendienteT56 = udtArticuloColor.StockPendienteT56 + mudtProps.CantidadT56
        
        Case TMMEntrega
            udtArticuloColor.StockPendienteT36 = udtArticuloColor.StockPendienteT36 - mudtProps.CantidadT36
            udtArticuloColor.StockPendienteT38 = udtArticuloColor.StockPendienteT38 - mudtProps.CantidadT38
            udtArticuloColor.StockPendienteT40 = udtArticuloColor.StockPendienteT40 - mudtProps.CantidadT40
            udtArticuloColor.StockPendienteT42 = udtArticuloColor.StockPendienteT42 - mudtProps.CantidadT42
            udtArticuloColor.StockPendienteT44 = udtArticuloColor.StockPendienteT44 - mudtProps.CantidadT44
            udtArticuloColor.StockPendienteT46 = udtArticuloColor.StockPendienteT46 - mudtProps.CantidadT46
            udtArticuloColor.StockPendienteT48 = udtArticuloColor.StockPendienteT48 - mudtProps.CantidadT48
            udtArticuloColor.StockPendienteT50 = udtArticuloColor.StockPendienteT50 - mudtProps.CantidadT50
            udtArticuloColor.StockPendienteT52 = udtArticuloColor.StockPendienteT52 - mudtProps.CantidadT52
            udtArticuloColor.StockPendienteT54 = udtArticuloColor.StockPendienteT54 - mudtProps.CantidadT54
            udtArticuloColor.StockPendienteT56 = udtArticuloColor.StockPendienteT56 - mudtProps.CantidadT56
        
        Case TMMEntrada
            udtArticuloColor.StockActualT36 = udtArticuloColor.StockActualT36 + mudtProps.CantidadT36
            udtArticuloColor.StockActualT38 = udtArticuloColor.StockActualT38 + mudtProps.CantidadT38
            udtArticuloColor.StockActualT40 = udtArticuloColor.StockActualT40 + mudtProps.CantidadT40
            udtArticuloColor.StockActualT42 = udtArticuloColor.StockActualT42 + mudtProps.CantidadT42
            udtArticuloColor.StockActualT44 = udtArticuloColor.StockActualT44 + mudtProps.CantidadT44
            udtArticuloColor.StockActualT46 = udtArticuloColor.StockActualT46 + mudtProps.CantidadT46
            udtArticuloColor.StockActualT48 = udtArticuloColor.StockActualT48 + mudtProps.CantidadT48
            udtArticuloColor.StockActualT50 = udtArticuloColor.StockActualT50 + mudtProps.CantidadT50
            udtArticuloColor.StockActualT52 = udtArticuloColor.StockActualT52 + mudtProps.CantidadT52
            udtArticuloColor.StockActualT54 = udtArticuloColor.StockActualT54 + mudtProps.CantidadT54
            udtArticuloColor.StockActualT56 = udtArticuloColor.StockActualT56 + mudtProps.CantidadT56
        
        Case TMMSalida
            udtArticuloColor.StockActualT36 = udtArticuloColor.StockActualT36 - mudtProps.CantidadT36
            udtArticuloColor.StockActualT38 = udtArticuloColor.StockActualT38 - mudtProps.CantidadT38
            udtArticuloColor.StockActualT40 = udtArticuloColor.StockActualT40 - mudtProps.CantidadT40
            udtArticuloColor.StockActualT42 = udtArticuloColor.StockActualT42 - mudtProps.CantidadT42
            udtArticuloColor.StockActualT44 = udtArticuloColor.StockActualT44 - mudtProps.CantidadT44
            udtArticuloColor.StockActualT46 = udtArticuloColor.StockActualT46 - mudtProps.CantidadT46
            udtArticuloColor.StockActualT48 = udtArticuloColor.StockActualT48 - mudtProps.CantidadT48
            udtArticuloColor.StockActualT50 = udtArticuloColor.StockActualT50 - mudtProps.CantidadT50
            udtArticuloColor.StockActualT52 = udtArticuloColor.StockActualT52 - mudtProps.CantidadT52
            udtArticuloColor.StockActualT54 = udtArticuloColor.StockActualT54 - mudtProps.CantidadT54
            udtArticuloColor.StockActualT56 = udtArticuloColor.StockActualT56 - mudtProps.CantidadT56
        
    End Select
   
    ' grabamos el registro resultante
    LSet udtArticuloColorData = udtArticuloColor
    udtArticuloColorData.Buffer = objArticuloColorPersist.Save(udtArticuloColorData.Buffer)
    LSet udtArticuloColor = udtArticuloColorData
    
    Set objArticuloColorPersist = Nothing
    
    ' Leemos el articulo del movimiento
    Set objArticuloPersist = mobjContext.CreateInstance("GCServerMTS.ArticuloPersist")
    udtArticuloData.Buffer = objArticuloPersist.Fetch(udtArticuloColor.ArticuloID)
    LSet udtArticulo = udtArticuloData
    
    ' Calculamos el nuevo stock
    udtArticulo.IsDeleted = False
    udtArticulo.IsDirty = True
    udtArticulo.IsNew = False
    
    Select Case mudtProps.Tipo
        Case TMMReserva
            udtArticulo.StockPendiente = udtArticulo.StockPendiente + mudtProps.CantidadT36 _
            + mudtProps.CantidadT38 + mudtProps.CantidadT40 + mudtProps.CantidadT42 _
            + mudtProps.CantidadT44 + mudtProps.CantidadT46 + mudtProps.CantidadT48 _
            + mudtProps.CantidadT50 + mudtProps.CantidadT52 + mudtProps.CantidadT54 _
            + mudtProps.CantidadT56
        
        Case TMMEntrega
            udtArticulo.StockPendiente = udtArticulo.StockPendiente - mudtProps.CantidadT36 _
            - mudtProps.CantidadT38 - mudtProps.CantidadT40 - mudtProps.CantidadT42 _
            - mudtProps.CantidadT44 - mudtProps.CantidadT46 - mudtProps.CantidadT48 _
            - mudtProps.CantidadT50 - mudtProps.CantidadT52 - mudtProps.CantidadT54 _
            - mudtProps.CantidadT56
        
        Case TMMEntrada
            udtArticulo.StockActual = udtArticulo.StockActual + mudtProps.CantidadT36 _
            + mudtProps.CantidadT38 + mudtProps.CantidadT40 + mudtProps.CantidadT42 _
            + mudtProps.CantidadT44 + mudtProps.CantidadT46 + mudtProps.CantidadT48 _
            + mudtProps.CantidadT50 + mudtProps.CantidadT52 + mudtProps.CantidadT54 _
            + mudtProps.CantidadT56
        
        Case TMMSalida
            udtArticulo.StockActual = udtArticulo.StockActual - mudtProps.CantidadT36 _
            - mudtProps.CantidadT38 - mudtProps.CantidadT40 - mudtProps.CantidadT42 _
            - mudtProps.CantidadT44 - mudtProps.CantidadT46 - mudtProps.CantidadT48 _
            - mudtProps.CantidadT50 - mudtProps.CantidadT52 - mudtProps.CantidadT54 _
            - mudtProps.CantidadT56
        
    End Select
   
    ' grabamos el registro resultante
    LSet udtArticuloData = udtArticulo
    udtArticuloData.Buffer = objArticuloPersist.Save(udtArticuloData.Buffer)
    LSet udtArticulo = udtArticuloData
    
    Set objArticuloPersist = Nothing
    
     
    With rsMovimiento
        'Creo que lo mejor es que obtenga la fecha real aqui.
        mudtProps.Fecha = Now
        .Fields("Fecha") = mudtProps.Fecha
        .Fields("ArticuloColorID") = mudtProps.ArticuloColorID
        .Fields("Tipo") = mudtProps.Tipo
        .Fields("Concepto") = mudtProps.Concepto
        .Fields("CantidadT36") = mudtProps.CantidadT36
        .Fields("CantidadT38") = mudtProps.CantidadT38
        .Fields("CantidadT40") = mudtProps.CantidadT40
        .Fields("CantidadT42") = mudtProps.CantidadT42
        .Fields("CantidadT44") = mudtProps.CantidadT44
        .Fields("CantidadT46") = mudtProps.CantidadT46
        .Fields("CantidadT48") = mudtProps.CantidadT48
        .Fields("CantidadT50") = mudtProps.CantidadT50
        .Fields("CantidadT52") = mudtProps.CantidadT52
        .Fields("CantidadT54") = mudtProps.CantidadT54
        .Fields("CantidadT56") = mudtProps.CantidadT56
        Select Case mudtProps.Tipo
        Case TMMReserva, TMMEntrega
            .Fields("StockFinal") = udtArticulo.StockPendiente
        
        Case TMMEntrada, TMMSalida
            .Fields("StockFinal") = udtArticulo.StockActual
        
        End Select
        .Fields("PrecioEUR") = mudtProps.PrecioEUR
        .Fields("PrecioVentaEUR") = udtArticulo.PrecioVentaEUR
        .Fields("PrecioCosteEUR") = udtArticulo.PrecioCosteEUR
        .Fields("AlmacenID") = IIf(mudtProps.AlmacenID <> 0, mudtProps.AlmacenID, Null)

        .Update
  
        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.MoviArticuloID = .Fields("MoviArticuloID")
        End If
        .Close
    End With
    
    Set rsMovimiento = Nothing
    Save = GetState
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Private Sub ActualizarArticuloColorAlmacen()
    Dim objArticuloColorAlmacenPersist As ArticuloColorAlmacenPersist
    Dim udtArticuloColorAlmacenData As ArticuloColorAlmacenData
    Dim udtArticuloColorAlmacen As ArticuloColorAlmacenProps
  
    ' Leemos al articulo-color-almacen del movimiento
    Set objArticuloColorAlmacenPersist = mobjContext.CreateInstance("GCServerMTS.ArticuloColorAlmacenPersist")
    udtArticuloColorAlmacenData.Buffer = objArticuloColorAlmacenPersist.FetchCreateIfNotExists(mudtProps.ArticuloColorID, mudtProps.AlmacenID)
    LSet udtArticuloColorAlmacen = udtArticuloColorAlmacenData
    
    ' Calculamos el nuevo stock
    udtArticuloColorAlmacen.IsDeleted = False
    udtArticuloColorAlmacen.IsDirty = True
    udtArticuloColorAlmacen.IsNew = False
    
    Select Case mudtProps.Tipo
        Case TMMReserva
            udtArticuloColorAlmacen.StockPendienteT36 = udtArticuloColorAlmacen.StockPendienteT36 + mudtProps.CantidadT36
            udtArticuloColorAlmacen.StockPendienteT38 = udtArticuloColorAlmacen.StockPendienteT38 + mudtProps.CantidadT38
            udtArticuloColorAlmacen.StockPendienteT40 = udtArticuloColorAlmacen.StockPendienteT40 + mudtProps.CantidadT40
            udtArticuloColorAlmacen.StockPendienteT42 = udtArticuloColorAlmacen.StockPendienteT42 + mudtProps.CantidadT42
            udtArticuloColorAlmacen.StockPendienteT44 = udtArticuloColorAlmacen.StockPendienteT44 + mudtProps.CantidadT44
            udtArticuloColorAlmacen.StockPendienteT46 = udtArticuloColorAlmacen.StockPendienteT46 + mudtProps.CantidadT46
            udtArticuloColorAlmacen.StockPendienteT48 = udtArticuloColorAlmacen.StockPendienteT48 + mudtProps.CantidadT48
            udtArticuloColorAlmacen.StockPendienteT50 = udtArticuloColorAlmacen.StockPendienteT50 + mudtProps.CantidadT50
            udtArticuloColorAlmacen.StockPendienteT52 = udtArticuloColorAlmacen.StockPendienteT52 + mudtProps.CantidadT52
            udtArticuloColorAlmacen.StockPendienteT54 = udtArticuloColorAlmacen.StockPendienteT54 + mudtProps.CantidadT54
            udtArticuloColorAlmacen.StockPendienteT56 = udtArticuloColorAlmacen.StockPendienteT56 + mudtProps.CantidadT56
        
        Case TMMEntrega
            udtArticuloColorAlmacen.StockPendienteT36 = udtArticuloColorAlmacen.StockPendienteT36 - mudtProps.CantidadT36
            udtArticuloColorAlmacen.StockPendienteT38 = udtArticuloColorAlmacen.StockPendienteT38 - mudtProps.CantidadT38
            udtArticuloColorAlmacen.StockPendienteT40 = udtArticuloColorAlmacen.StockPendienteT40 - mudtProps.CantidadT40
            udtArticuloColorAlmacen.StockPendienteT42 = udtArticuloColorAlmacen.StockPendienteT42 - mudtProps.CantidadT42
            udtArticuloColorAlmacen.StockPendienteT44 = udtArticuloColorAlmacen.StockPendienteT44 - mudtProps.CantidadT44
            udtArticuloColorAlmacen.StockPendienteT46 = udtArticuloColorAlmacen.StockPendienteT46 - mudtProps.CantidadT46
            udtArticuloColorAlmacen.StockPendienteT48 = udtArticuloColorAlmacen.StockPendienteT48 - mudtProps.CantidadT48
            udtArticuloColorAlmacen.StockPendienteT50 = udtArticuloColorAlmacen.StockPendienteT50 - mudtProps.CantidadT50
            udtArticuloColorAlmacen.StockPendienteT52 = udtArticuloColorAlmacen.StockPendienteT52 - mudtProps.CantidadT52
            udtArticuloColorAlmacen.StockPendienteT54 = udtArticuloColorAlmacen.StockPendienteT54 - mudtProps.CantidadT54
            udtArticuloColorAlmacen.StockPendienteT56 = udtArticuloColorAlmacen.StockPendienteT56 - mudtProps.CantidadT56
        
        Case TMMEntrada
            udtArticuloColorAlmacen.StockActualT36 = udtArticuloColorAlmacen.StockActualT36 + mudtProps.CantidadT36
            udtArticuloColorAlmacen.StockActualT38 = udtArticuloColorAlmacen.StockActualT38 + mudtProps.CantidadT38
            udtArticuloColorAlmacen.StockActualT40 = udtArticuloColorAlmacen.StockActualT40 + mudtProps.CantidadT40
            udtArticuloColorAlmacen.StockActualT42 = udtArticuloColorAlmacen.StockActualT42 + mudtProps.CantidadT42
            udtArticuloColorAlmacen.StockActualT44 = udtArticuloColorAlmacen.StockActualT44 + mudtProps.CantidadT44
            udtArticuloColorAlmacen.StockActualT46 = udtArticuloColorAlmacen.StockActualT46 + mudtProps.CantidadT46
            udtArticuloColorAlmacen.StockActualT48 = udtArticuloColorAlmacen.StockActualT48 + mudtProps.CantidadT48
            udtArticuloColorAlmacen.StockActualT50 = udtArticuloColorAlmacen.StockActualT50 + mudtProps.CantidadT50
            udtArticuloColorAlmacen.StockActualT52 = udtArticuloColorAlmacen.StockActualT52 + mudtProps.CantidadT52
            udtArticuloColorAlmacen.StockActualT54 = udtArticuloColorAlmacen.StockActualT54 + mudtProps.CantidadT54
            udtArticuloColorAlmacen.StockActualT56 = udtArticuloColorAlmacen.StockActualT56 + mudtProps.CantidadT56
        
        Case TMMSalida
            udtArticuloColorAlmacen.StockActualT36 = udtArticuloColorAlmacen.StockActualT36 - mudtProps.CantidadT36
            udtArticuloColorAlmacen.StockActualT38 = udtArticuloColorAlmacen.StockActualT38 - mudtProps.CantidadT38
            udtArticuloColorAlmacen.StockActualT40 = udtArticuloColorAlmacen.StockActualT40 - mudtProps.CantidadT40
            udtArticuloColorAlmacen.StockActualT42 = udtArticuloColorAlmacen.StockActualT42 - mudtProps.CantidadT42
            udtArticuloColorAlmacen.StockActualT44 = udtArticuloColorAlmacen.StockActualT44 - mudtProps.CantidadT44
            udtArticuloColorAlmacen.StockActualT46 = udtArticuloColorAlmacen.StockActualT46 - mudtProps.CantidadT46
            udtArticuloColorAlmacen.StockActualT48 = udtArticuloColorAlmacen.StockActualT48 - mudtProps.CantidadT48
            udtArticuloColorAlmacen.StockActualT50 = udtArticuloColorAlmacen.StockActualT50 - mudtProps.CantidadT50
            udtArticuloColorAlmacen.StockActualT52 = udtArticuloColorAlmacen.StockActualT52 - mudtProps.CantidadT52
            udtArticuloColorAlmacen.StockActualT54 = udtArticuloColorAlmacen.StockActualT54 - mudtProps.CantidadT54
            udtArticuloColorAlmacen.StockActualT56 = udtArticuloColorAlmacen.StockActualT56 - mudtProps.CantidadT56
        
    End Select
       
    ' grabamos el registro resultante
    LSet udtArticuloColorAlmacenData = udtArticuloColorAlmacen
    udtArticuloColorAlmacenData.Buffer = objArticuloColorAlmacenPersist.Save(udtArticuloColorAlmacenData.Buffer)
    LSet udtArticuloColorAlmacen = udtArticuloColorAlmacenData
    
    Set objArticuloColorAlmacenPersist = Nothing
End Sub



Public Sub DeleteObject(ByVal MoviArticuloID As Long)
    Dim cnGescom As Connection


    On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM MoviArticulos WHERE MoviArticuloID=" & _
        MoviArticuloID

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

Public Function Fetch(ByVal MoviArticuloID As Long) As String
    Dim rsMovimiento As Recordset
    Dim strSQL As String
  

    On Error GoTo errh
    strSQL = "SELECT * FROM MoviArticulos WHERE MoviArticuloID=" & _
        MoviArticuloID
    Set rsMovimiento = New Recordset
    rsMovimiento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsMovimiento
        mudtProps.ArticuloColorID = .Fields("ArticuloColorID")
        mudtProps.Tipo = .Fields("Tipo")
        mudtProps.Concepto = .Fields("Concepto")
        mudtProps.CantidadT36 = .Fields("CantidadT36")
        mudtProps.CantidadT38 = .Fields("CantidadT38")
        mudtProps.CantidadT40 = .Fields("CantidadT40")
        mudtProps.CantidadT42 = .Fields("CantidadT42")
        mudtProps.CantidadT44 = .Fields("CantidadT44")
        mudtProps.CantidadT46 = .Fields("CantidadT46")
        mudtProps.CantidadT48 = .Fields("CantidadT48")
        mudtProps.CantidadT50 = .Fields("CantidadT50")
        mudtProps.CantidadT52 = .Fields("CantidadT52")
        mudtProps.CantidadT54 = .Fields("CantidadT54")
        mudtProps.CantidadT56 = .Fields("CantidadT56")
        mudtProps.StockFinal = .Fields("StockFinal")
        mudtProps.PrecioEUR = .Fields("PrecioEUR")
        mudtProps.PrecioVentaEUR = .Fields("PrecioVentaEUR")
        mudtProps.PrecioCosteEUR = .Fields("PrecioCosteEUR")
        mudtProps.AlmacenID = IIf(IsNull(.Fields("AlmacenID")), 0, .Fields("AlmacenID"))
        
        .Close
    End With

    Set rsMovimiento = Nothing
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

