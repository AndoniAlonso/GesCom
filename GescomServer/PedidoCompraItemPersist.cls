VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "PedidoCompraItemPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
  
Private mudtProps As PedidoCompraItemProps
Private mudtSave As PedidoCompraItemProps

Private Function GetState() As String
    Dim udtData As PedidoCompraItemData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As PedidoCompraItemData

    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String
    Dim rsItem As Recordset
    Dim strSQL As String

    On Error GoTo errh
    SetState Buffer

    Set rsItem = New Recordset
    strSQL = "SELECT * FROM PedidoCompraItems WHERE PedidoCompraItemID=" & _
        mudtProps.PedidoCompraItemID
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then
        rsItem.AddNew
    Else
        ' Leemos el pedido y lo guardamos para actualizaciones/desactualizaciones
        Leer (mudtProps.PedidoCompraItemID)
    End If

    With rsItem
        If mudtProps.DesactualizarAlta Then Desactualizar
        If mudtProps.DesactualizarAlbaran Then DesactualizarAlbaran
        
        If mudtProps.ActualizarAlta Then Actualizar
        If mudtProps.ActualizarAlbaran Then ActualizarAlbaran
        
        ' Miramos si el pedido esta completo
        If mudtProps.Cantidad = mudtProps.Servido Then
           mudtProps.Situacion = "C"
        Else
           mudtProps.Situacion = "A"
        End If

        .Fields("PedidoCompraID") = mudtProps.PedidoCompraID
        .Fields("MaterialID") = mudtProps.MaterialID
        .Fields("Situacion") = mudtProps.Situacion
        .Fields("Cantidad") = mudtProps.Cantidad
        .Fields("Servido") = mudtProps.Servido
        .Fields("PrecioCostePTA") = mudtProps.PrecioCostePTA
        .Fields("PrecioCosteEUR") = mudtProps.PrecioCosteEUR
        .Fields("Descuento") = mudtProps.Descuento
        .Fields("BrutoPTA") = mudtProps.BrutoPTA
        .Fields("BrutoEUR") = mudtProps.BrutoEUR
        .Fields("Comision") = mudtProps.Comision

        .Update

        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.PedidoCompraItemID = .Fields("PedidoCompraItemID")
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

Public Sub DeleteObject(ByVal PedidoCompraItemID As Long)
    Dim cnGescom As Connection

    On Error GoTo errh
    
    If Not mudtProps.IsNew Then
        ' Al borrar una linea de pedidos hay que desactualizar SIEMPRE.
        Leer (PedidoCompraItemID)
        Desactualizar
    End If
  
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM PedidoCompraItems WHERE PedidoCompraItemID=" & _
        PedidoCompraItemID

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

Public Function Fetch(ByVal PedidoCompraItemID As Long) As String
    Dim rsItem As Recordset
    Dim strSQL As String


    On Error GoTo errh
    strSQL = "SELECT * FROM PedidoCompraItems WHERE PedidoCompraItemID=" & _
        PedidoCompraItemID
    Set rsItem = New Recordset
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsItem
        mudtProps.PedidoCompraItemID = .Fields("PedidoCompraItemID")
        mudtProps.PedidoCompraID = .Fields("PedidoCompraID")
        mudtProps.MaterialID = .Fields("MaterialID")
        mudtProps.Situacion = .Fields("Situacion") & vbNullString
        mudtProps.Cantidad = .Fields("Cantidad")
        mudtProps.Servido = .Fields("Servido")
        mudtProps.PrecioCostePTA = .Fields("PrecioCostePTA")
        mudtProps.PrecioCosteEUR = .Fields("PrecioCosteEUR")
        mudtProps.Descuento = .Fields("Descuento")
        mudtProps.BrutoPTA = .Fields("BrutoPTA")
        mudtProps.BrutoEUR = .Fields("BrutoEUR")
        mudtProps.Comision = .Fields("Comision")

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
Private Sub Leer(ByVal PedidoCompraItemID As Long)
    Dim rsItem As Recordset
    Dim strSQL As String

    strSQL = "SELECT * FROM PedidoCompraItems WHERE PedidoCompraItemID=" & _
        PedidoCompraItemID
    Set rsItem = New Recordset
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsItem
        mudtSave.PedidoCompraItemID = .Fields("PedidoCompraItemID")
        mudtSave.PedidoCompraID = .Fields("PedidoCompraID")
        mudtSave.MaterialID = .Fields("MaterialID")
        mudtSave.Situacion = .Fields("Situacion") & vbNullString
        mudtSave.Cantidad = .Fields("Cantidad")
        mudtSave.Servido = .Fields("Servido")
        mudtSave.PrecioCostePTA = .Fields("PrecioCostePTA")
        mudtSave.PrecioCosteEUR = .Fields("PrecioCosteEUR")
        mudtSave.Descuento = .Fields("Descuento")
        mudtSave.BrutoPTA = .Fields("BrutoPTA")
        mudtSave.BrutoEUR = .Fields("BrutoEUR")
        mudtSave.Comision = .Fields("Comision")
    
        .Close
    End With

    Set rsItem = Nothing

End Sub

Public Function SaveChild(PedidoCompraID As Long, Buffer As String) As String

    On Error GoTo errh
    SetState Buffer

    With mudtProps
        If .IsDeleted Then
            DeleteObject .PedidoCompraItemID
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
    Dim objPersist As MoviMaterialPersist
    Dim udtMovimiento As MoviMaterialProps
    Dim udtMovimientoData As MoviMaterialData
    
    With udtMovimiento
        .MaterialID = mudtProps.MaterialID
        .Tipo = TMMReserva
        .Concepto = "Pedido de compra"
        .Cantidad = mudtProps.Cantidad
        .PrecioEUR = mudtProps.PrecioCosteEUR
        .PrecioPTA = mudtProps.PrecioCostePTA
        .DocumentoID = mudtProps.PedidoCompraItemID
        .TipoDocumento = "P"
        .IsDeleted = False
        .IsDirty = True
        .IsNew = True
    End With
    
    LSet udtMovimientoData = udtMovimiento
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.MoviMaterialPersist")
    
    objPersist.Save (udtMovimientoData.Buffer)
    
    Set objPersist = Nothing

End Sub

' Actualizar el pedido por la cantidad servida desde un albaran.
Private Sub ActualizarAlbaran()
    Dim objPersist As MoviMaterialPersist
    Dim udtMovimiento As MoviMaterialProps
    Dim udtMovimientoData As MoviMaterialData
    
    With udtMovimiento
        .MaterialID = mudtProps.MaterialID
        .Tipo = TMMEntrega
        .Concepto = "Entrega de albarán"
        .Cantidad = mudtProps.Servido - mudtSave.Servido
        .PrecioEUR = mudtProps.PrecioCosteEUR
        .PrecioPTA = mudtProps.PrecioCostePTA
        .DocumentoID = mudtProps.PedidoCompraItemID
        .TipoDocumento = "P"
        .IsDeleted = False
        .IsDirty = True
        .IsNew = True
    End With
    
    LSet udtMovimientoData = udtMovimiento
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.MoviMaterialPersist")
    
    objPersist.Save (udtMovimientoData.Buffer)
    
    Set objPersist = Nothing
    
End Sub

Private Sub Desactualizar()
    Dim objPersist As MoviMaterialPersist
    Dim udtMovimiento As MoviMaterialProps
    Dim udtMovimientoData As MoviMaterialData
    
    With udtMovimiento
        .MaterialID = mudtSave.MaterialID
        .Tipo = TMMEntrega
        .Concepto = "Corrección pedido de compra"
        .Cantidad = mudtSave.Cantidad
        .PrecioEUR = mudtSave.PrecioCosteEUR
        .PrecioPTA = mudtSave.PrecioCostePTA
        .DocumentoID = mudtProps.PedidoCompraItemID
        .TipoDocumento = "P"
        .IsDeleted = False
        .IsDirty = True
        .IsNew = True
    End With
    
    LSet udtMovimientoData = udtMovimiento
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.MoviMaterialPersist")
    
    objPersist.Save (udtMovimientoData.Buffer)
    
    Set objPersist = Nothing

End Sub

' Se desactualiza la entrega realizada por el albaran.
Private Sub DesactualizarAlbaran()
    Dim objPersist As MoviMaterialPersist
    Dim udtMovimiento As MoviMaterialProps
    Dim udtMovimientoData As MoviMaterialData
    

    With udtMovimiento
        .MaterialID = mudtSave.MaterialID
        .Tipo = TMMReserva
        .Concepto = "Corrección albarán de compra"
        .Cantidad = mudtSave.Servido - mudtProps.Servido
        .PrecioEUR = mudtSave.PrecioCosteEUR
        .PrecioPTA = mudtSave.PrecioCostePTA
        .DocumentoID = mudtProps.PedidoCompraItemID
        .TipoDocumento = "P"
        .IsDeleted = False
        .IsDirty = True
        .IsNew = True
    End With
    
    LSet udtMovimientoData = udtMovimiento
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.MoviMaterialPersist")
    
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

