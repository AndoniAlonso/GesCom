VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "PedidoCompraArticuloPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext

Private mudtProps As PedidoCompraArticuloProps
Private mudtSave As PedidoCompraArticuloProps

Private Function GetState() As String
    Dim udtData As PedidoCompraArticuloData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As PedidoCompraArticuloData

    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String
    Dim rsItem As Recordset
    Dim strSQL As String

    On Error GoTo errh
    SetState Buffer

    Set rsItem = New Recordset
    strSQL = "SELECT * FROM PedidoCompraArticulos WHERE PedidoCompraArticuloID=" & _
        mudtProps.PedidoCompraArticuloID
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    
    If mudtProps.IsNew Then
        rsItem.AddNew
    Else
        ' Leemos el pedido y lo guardamos para actualizaciones/desactualizaciones
        Leer (mudtProps.PedidoCompraArticuloID)
    End If
    
    With rsItem
    
        If mudtProps.DesactualizarAlta Then Desactualizar
        If mudtProps.DesactualizarAlbaran Then DesactualizarAlbaran
        
        If mudtProps.ActualizarAlta Then Actualizar
        If mudtProps.ActualizarAlbaran Then ActualizarAlbaran
        
        ' Miramos si el pedido esta completo
        If mudtProps.CantidadT36 = mudtProps.ServidoT36 And _
           mudtProps.CantidadT38 = mudtProps.ServidoT38 And _
           mudtProps.CantidadT40 = mudtProps.ServidoT40 And _
           mudtProps.CantidadT42 = mudtProps.ServidoT42 And _
           mudtProps.CantidadT44 = mudtProps.ServidoT44 And _
           mudtProps.CantidadT46 = mudtProps.ServidoT46 And _
           mudtProps.CantidadT48 = mudtProps.ServidoT48 And _
           mudtProps.CantidadT50 = mudtProps.ServidoT50 And _
           mudtProps.CantidadT52 = mudtProps.ServidoT52 And _
           mudtProps.CantidadT54 = mudtProps.ServidoT54 And _
           mudtProps.CantidadT56 = mudtProps.ServidoT56 Then
           mudtProps.Situacion = "C"
        Else
           mudtProps.Situacion = "A"
        End If


        ' Actualizamos el registro
        .Fields("PedidoCompraID") = mudtProps.PedidoCompraID
        .Fields("ArticuloColorID") = mudtProps.ArticuloColorID

        .Fields("Situacion") = mudtProps.Situacion
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
        .Fields("ServidoT36") = mudtProps.ServidoT36
        .Fields("ServidoT38") = mudtProps.ServidoT38
        .Fields("ServidoT40") = mudtProps.ServidoT40
        .Fields("ServidoT42") = mudtProps.ServidoT42
        .Fields("ServidoT44") = mudtProps.ServidoT44
        .Fields("ServidoT46") = mudtProps.ServidoT46
        .Fields("ServidoT48") = mudtProps.ServidoT48
        .Fields("ServidoT50") = mudtProps.ServidoT50
        .Fields("ServidoT52") = mudtProps.ServidoT52
        .Fields("ServidoT54") = mudtProps.ServidoT54
        .Fields("ServidoT56") = mudtProps.ServidoT56
        .Fields("PrecioCompraEUR") = mudtProps.PrecioCompraEUR
        .Fields("Descuento") = mudtProps.Descuento
        .Fields("BrutoEUR") = mudtProps.BrutoEUR
        .Fields("Comision") = mudtProps.Comision
        .Fields("TemporadaID") = mudtProps.TemporadaID
        .Fields("Observaciones") = IIf(Len(mudtProps.Observaciones) = 0, Null, mudtProps.Observaciones)
        .Fields("AlmacenID") = mudtProps.AlmacenID
    
        .Update
    
        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.PedidoCompraArticuloID = .Fields("PedidoCompraArticuloID")
        End If
        .Close
    End With
    
    Set rsItem = Nothing
    Save = GetState
    mobjContext.SetComplete
    
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal PedidoCompraArticuloID As Long)
    Dim cnGescom As Connection

    On Error GoTo errh
  
    If Not mudtProps.IsNew Then
        ' Al borrar una linea de pedidos hay que desactualizar SIEMPRE.
        Leer (PedidoCompraArticuloID)
        Desactualizar
    End If
  
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM PedidoCompraArticulos WHERE PedidoCompraArticuloID=" & _
        PedidoCompraArticuloID

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

Public Function Fetch(ByVal PedidoCompraArticuloID As Long) As String
    Dim rsItem As Recordset
    Dim strSQL As String
    
  On Error GoTo errh
    strSQL = "SELECT * FROM vPedidoCompraArticulos WHERE PedidoCompraArticuloID=" & _
        PedidoCompraArticuloID
    Set rsItem = New Recordset
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsItem
        mudtProps.PedidoCompraArticuloID = .Fields("PedidoCompraArticuloID")
        mudtProps.PedidoCompraID = .Fields("PedidoCompraID")
        mudtProps.ArticuloColorID = .Fields("ArticuloColorID")
        mudtProps.NombreArticuloColor = .Fields("NombreArticuloColor") & vbNullString
        mudtProps.Situacion = .Fields("Situacion") & vbNullString
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
        mudtProps.ServidoT36 = .Fields("ServidoT36")
        mudtProps.ServidoT38 = .Fields("ServidoT38")
        mudtProps.ServidoT40 = .Fields("ServidoT40")
        mudtProps.ServidoT42 = .Fields("ServidoT42")
        mudtProps.ServidoT44 = .Fields("ServidoT44")
        mudtProps.ServidoT46 = .Fields("ServidoT46")
        mudtProps.ServidoT48 = .Fields("ServidoT48")
        mudtProps.ServidoT50 = .Fields("ServidoT50")
        mudtProps.ServidoT52 = .Fields("ServidoT52")
        mudtProps.ServidoT54 = .Fields("ServidoT54")
        mudtProps.ServidoT56 = .Fields("ServidoT56")
        mudtProps.PrecioCompraEUR = .Fields("PrecioCompraEUR")
        mudtProps.Descuento = .Fields("Descuento")
        mudtProps.BrutoEUR = .Fields("BrutoEUR")
        mudtProps.Comision = .Fields("Comision")
        mudtProps.TemporadaID = .Fields("TemporadaID")
        mudtProps.Observaciones = IIf(IsNull(.Fields("Observaciones")), vbNullString, .Fields("Observaciones"))
        mudtProps.AlmacenID = .Fields("AlmacenID")

        .Close
    End With

    Set rsItem = Nothing
    Fetch = GetState
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

' Se ocupa de leer los datos de un item de pedido existente para desactualizar/actualizar
Private Sub Leer(ByVal PedidoCompraArticuloID As Long)
    Dim rsItem As Recordset
    Dim strSQL As String

    strSQL = "SELECT * FROM PedidoCompraArticulos WHERE PedidoCompraArticuloID=" & _
        PedidoCompraArticuloID
    Set rsItem = New Recordset
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsItem
        mudtSave.PedidoCompraArticuloID = .Fields("PedidoCompraArticuloID")
        mudtSave.PedidoCompraID = .Fields("PedidoCompraID")
        mudtSave.ArticuloColorID = .Fields("ArticuloColorID")
        mudtSave.AlmacenID = .Fields("AlmacenID")
        mudtSave.Situacion = .Fields("Situacion") & vbNullString
        mudtSave.CantidadT36 = .Fields("CantidadT36")
        mudtSave.CantidadT38 = .Fields("CantidadT38")
        mudtSave.CantidadT40 = .Fields("CantidadT40")
        mudtSave.CantidadT42 = .Fields("CantidadT42")
        mudtSave.CantidadT44 = .Fields("CantidadT44")
        mudtSave.CantidadT46 = .Fields("CantidadT46")
        mudtSave.CantidadT48 = .Fields("CantidadT48")
        mudtSave.CantidadT50 = .Fields("CantidadT50")
        mudtSave.CantidadT52 = .Fields("CantidadT52")
        mudtSave.CantidadT54 = .Fields("CantidadT54")
        mudtSave.CantidadT56 = .Fields("CantidadT56")
        mudtSave.ServidoT36 = .Fields("ServidoT36")
        mudtSave.ServidoT38 = .Fields("ServidoT38")
        mudtSave.ServidoT40 = .Fields("ServidoT40")
        mudtSave.ServidoT42 = .Fields("ServidoT42")
        mudtSave.ServidoT44 = .Fields("ServidoT44")
        mudtSave.ServidoT46 = .Fields("ServidoT46")
        mudtSave.ServidoT48 = .Fields("ServidoT48")
        mudtSave.ServidoT50 = .Fields("ServidoT50")
        mudtSave.ServidoT52 = .Fields("ServidoT52")
        mudtSave.ServidoT54 = .Fields("ServidoT54")
        mudtSave.ServidoT56 = .Fields("ServidoT56")
        mudtSave.PrecioCompraEUR = .Fields("PrecioCompraEUR")
        mudtSave.Descuento = .Fields("Descuento")
        mudtSave.BrutoEUR = .Fields("BrutoEUR")
        mudtSave.Comision = .Fields("Comision")
        mudtSave.TemporadaID = .Fields("TemporadaID")
        mudtSave.Observaciones = IIf(IsNull(.Fields("Observaciones")), vbNullString, .Fields("Observaciones"))
    
        .Close
    End With

    Set rsItem = Nothing

End Sub

Public Function SaveChild(PedidoCompraID As Long, Buffer As String) As String

    On Error GoTo errh
    SetState Buffer

    With mudtProps
        If .IsDeleted Then
            DeleteObject .PedidoCompraArticuloID
            SaveChild = vbNullString

        ElseIf .IsDirty Or .IsNew Then
            .PedidoCompraID = PedidoCompraID
            SaveChild = Save(GetState)

        Else
            SaveChild = Buffer
            mobjContext.SetComplete
  
        End If

    End With
    Exit Function
    
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

' Se encarga de actualizar el pedido.
Private Sub Actualizar()
    Dim objPersist As MoviArticuloPersist
    Dim udtMovimiento As MoviArticuloProps
    Dim udtMovimientoData As MoviArticuloData
    
    With udtMovimiento
        .ArticuloColorID = mudtProps.ArticuloColorID
        .Tipo = TMMReserva
        .Concepto = "Pedido de Compra"
        .CantidadT36 = mudtProps.CantidadT36
        .CantidadT38 = mudtProps.CantidadT38
        .CantidadT40 = mudtProps.CantidadT40
        .CantidadT42 = mudtProps.CantidadT42
        .CantidadT44 = mudtProps.CantidadT44
        .CantidadT46 = mudtProps.CantidadT46
        .CantidadT48 = mudtProps.CantidadT48
        .CantidadT50 = mudtProps.CantidadT50
        .CantidadT52 = mudtProps.CantidadT52
        .CantidadT54 = mudtProps.CantidadT54
        .CantidadT56 = mudtProps.CantidadT56
        .PrecioEUR = mudtProps.PrecioCompraEUR
        .AlmacenID = mudtProps.AlmacenID
        .IsDeleted = False
        .IsDirty = True
        .IsNew = True
    End With
    
    LSet udtMovimientoData = udtMovimiento
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.MoviArticuloPersist")
    
    objPersist.Save (udtMovimientoData.Buffer)
    
    Set objPersist = Nothing

End Sub

' Actualizar el pedido por la cantidad servida desde un albaran.
Private Sub ActualizarAlbaran()
    Dim objPersist As MoviArticuloPersist
    Dim udtMovimiento As MoviArticuloProps
    Dim udtMovimientoData As MoviArticuloData
    
    With udtMovimiento
        .ArticuloColorID = mudtProps.ArticuloColorID
        .AlmacenID = mudtProps.AlmacenID
        .Tipo = TMMEntrega
        .Concepto = "Entrega de albar�n"
        .CantidadT36 = mudtProps.ServidoT36 - mudtSave.ServidoT36
        .CantidadT38 = mudtProps.ServidoT38 - mudtSave.ServidoT38
        .CantidadT40 = mudtProps.ServidoT40 - mudtSave.ServidoT40
        .CantidadT42 = mudtProps.ServidoT42 - mudtSave.ServidoT42
        .CantidadT44 = mudtProps.ServidoT44 - mudtSave.ServidoT44
        .CantidadT46 = mudtProps.ServidoT46 - mudtSave.ServidoT46
        .CantidadT48 = mudtProps.ServidoT48 - mudtSave.ServidoT48
        .CantidadT50 = mudtProps.ServidoT50 - mudtSave.ServidoT50
        .CantidadT52 = mudtProps.ServidoT52 - mudtSave.ServidoT52
        .CantidadT54 = mudtProps.ServidoT54 - mudtSave.ServidoT54
        .CantidadT56 = mudtProps.ServidoT56 - mudtSave.ServidoT56
        .PrecioEUR = mudtProps.PrecioCompraEUR
        .IsDeleted = False
        .IsDirty = True
        .IsNew = True
    End With
    
    LSet udtMovimientoData = udtMovimiento
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.MoviArticuloPersist")
    
    objPersist.Save (udtMovimientoData.Buffer)
    
    Set objPersist = Nothing
    
End Sub

Private Sub Desactualizar()
    Dim objPersist As MoviArticuloPersist
    Dim udtMovimiento As MoviArticuloProps
    Dim udtMovimientoData As MoviArticuloData
    
    With udtMovimiento
        .ArticuloColorID = mudtSave.ArticuloColorID
        .AlmacenID = mudtSave.AlmacenID
        .Tipo = TMMEntrega
        .Concepto = "Correcci�n pedido de Compra"
        .CantidadT36 = mudtSave.CantidadT36
        .CantidadT38 = mudtSave.CantidadT38
        .CantidadT40 = mudtSave.CantidadT40
        .CantidadT42 = mudtSave.CantidadT42
        .CantidadT44 = mudtSave.CantidadT44
        .CantidadT46 = mudtSave.CantidadT46
        .CantidadT48 = mudtSave.CantidadT48
        .CantidadT50 = mudtSave.CantidadT50
        .CantidadT52 = mudtSave.CantidadT52
        .CantidadT54 = mudtSave.CantidadT54
        .CantidadT56 = mudtSave.CantidadT56
        .PrecioEUR = mudtSave.PrecioCompraEUR
        .IsDeleted = False
        .IsDirty = True
        .IsNew = True
    End With
    
    LSet udtMovimientoData = udtMovimiento
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.MoviArticuloPersist")
    
    objPersist.Save (udtMovimientoData.Buffer)
    
    Set objPersist = Nothing

End Sub

' Se desactualiza la entrega realizada por el albaran.
Private Sub DesactualizarAlbaran()
    Dim objPersist As MoviArticuloPersist
    Dim udtMovimiento As MoviArticuloProps
    Dim udtMovimientoData As MoviArticuloData

    With udtMovimiento
        .ArticuloColorID = mudtSave.ArticuloColorID
        .AlmacenID = mudtSave.AlmacenID
        .Tipo = TMMReserva
        .Concepto = "Correcci�n albar�n de Compra"
        .CantidadT36 = mudtSave.ServidoT36 - mudtProps.ServidoT36
        .CantidadT38 = mudtSave.ServidoT38 - mudtProps.ServidoT38
        .CantidadT40 = mudtSave.ServidoT40 - mudtProps.ServidoT40
        .CantidadT42 = mudtSave.ServidoT42 - mudtProps.ServidoT42
        .CantidadT44 = mudtSave.ServidoT44 - mudtProps.ServidoT44
        .CantidadT46 = mudtSave.ServidoT46 - mudtProps.ServidoT46
        .CantidadT48 = mudtSave.ServidoT48 - mudtProps.ServidoT48
        .CantidadT50 = mudtSave.ServidoT50 - mudtProps.ServidoT50
        .CantidadT52 = mudtSave.ServidoT52 - mudtProps.ServidoT52
        .CantidadT54 = mudtSave.ServidoT54 - mudtProps.ServidoT54
        .CantidadT56 = mudtSave.ServidoT56 - mudtProps.ServidoT56
        .PrecioEUR = mudtSave.PrecioCompraEUR
        .IsDeleted = False
        .IsDirty = True
        .IsNew = True
    End With
    
    LSet udtMovimientoData = udtMovimiento
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.MoviArticuloPersist")
    
    objPersist.Save (udtMovimientoData.Buffer)
    
    Set objPersist = Nothing

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

