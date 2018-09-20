VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "AlbaranCompraItemPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
      
Implements ObjectControl
  
Private mobjContext As ObjectContext
  
Private mudtProps As AlbaranCompraItemProps
    
Private Function GetState() As String
    Dim udtData As AlbaranCompraItemData
    
    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As AlbaranCompraItemData
    
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String
    Dim rsItem As Recordset
    Dim strSQL As String
    
    On Error GoTo errh
    SetState Buffer
    
    Set rsItem = New Recordset
    strSQL = "SELECT * FROM AlbaranCompraItems WHERE AlbaranCompraItemID=" & _
      mudtProps.AlbaranCompraItemID
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsItem.AddNew
    
    With rsItem
      If mudtProps.DesactualizarAlta Then LeerYDesactualizar (mudtProps.AlbaranCompraItemID)
      .Fields("AlbaranCompraID") = mudtProps.AlbaranCompraID
      ' Solo asigno el valor si viene informado
      .Fields("MaterialID") = IIf(mudtProps.MaterialID, mudtProps.MaterialID, Null)
      ' Solo asigno el valor si viene informado
      .Fields("PedidoCompraItemID") = IIf(mudtProps.PedidoCompraItemID, mudtProps.PedidoCompraItemID, Null)
      .Fields("Situacion") = mudtProps.Situacion
      .Fields("Cantidad") = mudtProps.Cantidad
      .Fields("PrecioCostePTA") = mudtProps.PrecioCostePTA
      .Fields("PrecioCosteEUR") = mudtProps.PrecioCosteEUR
      .Fields("Descuento") = mudtProps.Descuento
      .Fields("BrutoPTA") = mudtProps.BrutoPTA
      .Fields("BrutoEUR") = mudtProps.BrutoEUR
      .Fields("Comision") = mudtProps.Comision
    
      .Update
    
      If mudtProps.IsNew Then
         .Bookmark = .Bookmark
          mudtProps.AlbaranCompraItemID = .Fields("AlbaranCompraItemID")
      End If
      If mudtProps.ActualizarAlta Then Actualizar
      
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

Public Sub DeleteObject(ByVal AlbaranCompraItemID As Long)
    Dim cnGescom As Connection
    
    On Error GoTo errh

    If Not mudtProps.IsNew Then
        ' Al borrar una linea de albaranes hay que desactualizar SIEMPRE.
        LeerYDesactualizar (AlbaranCompraItemID)
    End If

    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    cnGescom.Execute "DELETE FROM AlbaranCompraItems WHERE AlbaranCompraItemID=" & _
        AlbaranCompraItemID
    
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

Public Function Fetch(ByVal AlbaranCompraItemID As Long) As String
    Dim rsItem As Recordset
    Dim strSQL As String
    
    On Error GoTo errh
    strSQL = "SELECT * FROM AlbaranCompraItems WHERE AlbaranCompraItemID=" & _
        AlbaranCompraItemID
    Set rsItem = New Recordset
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    With rsItem
      mudtProps.AlbaranCompraItemID = .Fields("AlbaranCompraItemID")
      mudtProps.AlbaranCompraID = .Fields("AlbaranCompraID")
      mudtProps.MaterialID = IIf(IsNull(.Fields("MaterialID")), 0, .Fields("MaterialID"))
      mudtProps.PedidoCompraItemID = IIf(IsNull(.Fields("PedidoCompraItemID")), 0, .Fields("PedidoCompraItemID"))
      mudtProps.Situacion = .Fields("Situacion") & vbNullString
      mudtProps.Cantidad = .Fields("Cantidad")
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


Public Function SaveChild(AlbaranCompraID As Long, Buffer As String) As String

    On Error GoTo errh
    SetState Buffer
    
    With mudtProps
      If .IsDeleted Then
          DeleteObject .AlbaranCompraItemID
        SaveChild = vbNullString
    
      ElseIf .IsDirty Or .IsNew Then
          .AlbaranCompraID = AlbaranCompraID
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

' Se encarga de actualizar el albarán.
Private Sub Actualizar()
    Dim objPersist As MoviMaterialPersist
    Dim udtMovimiento As MoviMaterialProps
    Dim udtMovimientoData As MoviMaterialData
    Dim objPedidoPersist As PedidoCompraItemPersist
    Dim udtPedido As PedidoCompraItemProps
    Dim udtPedidoData As PedidoCompraItemData
    Dim objAlbaranPersist As AlbaranCompraPersist
    Dim udtAlbaran As AlbaranCompraProps
    Dim udtAlbaranData As AlbaranCompraData
    
    ' Primero actualizamos el stock por la parte del pedido.
    If mudtProps.PedidoCompraItemID Then
        Set objPedidoPersist = mobjContext.CreateInstance("GCServerMTS.PedidoCompraItemPersist")
        udtPedidoData.Buffer = objPedidoPersist.Fetch(mudtProps.PedidoCompraItemID)
        LSet udtPedido = udtPedidoData
        With udtPedido
            .Servido = .Servido + mudtProps.Cantidad
            .IsDeleted = False
            .IsDirty = True
            .IsNew = False
            .ActualizarAlbaran = True
            .ActualizarAlta = False
            .DesactualizarAlbaran = False
            .DesactualizarAlta = False
        End With
        LSet udtPedidoData = udtPedido
        udtPedidoData.Buffer = objPedidoPersist.Save(udtPedidoData.Buffer)
        Set objPedidoPersist = Nothing
    End If
    
    ' Leemos el albaran para obtener el numero (solo tenemos el ID)
    Set objAlbaranPersist = mobjContext.CreateInstance("GCServerMTS.AlbaranCompraPersist")
    udtAlbaranData.Buffer = objAlbaranPersist.Fetch(mudtProps.AlbaranCompraID)
    LSet udtAlbaran = udtAlbaranData
    Set objAlbaranPersist = Nothing
    
    With udtMovimiento
        .MaterialID = mudtProps.MaterialID
        .Tipo = TMMEntrada
        .Concepto = "Albarán de Compra " & CStr(udtAlbaran.Numero)
        .Cantidad = mudtProps.Cantidad
        .PrecioEUR = mudtProps.PrecioCosteEUR
        .PrecioPTA = mudtProps.PrecioCostePTA
        .DocumentoID = mudtProps.AlbaranCompraItemID
        .TipoDocumento = "A"
        .IsDeleted = False
        .IsDirty = True
        .IsNew = True
    End With
    
    LSet udtMovimientoData = udtMovimiento
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.MoviMaterialPersist")
    
    objPersist.Save (udtMovimientoData.Buffer)
    
    Set objPersist = Nothing

End Sub

Private Sub LeerYDesactualizar(ByVal AlbaranCompraItemID As Long)
    Dim rsItem As Recordset
    Dim strSQL As String
    Dim objPersist As MoviMaterialPersist
    Dim udtMovimiento As MoviMaterialProps
    Dim udtMovimientoData As MoviMaterialData
    Dim objPedidoPersist As PedidoCompraItemPersist
    Dim udtPedido As PedidoCompraItemProps
    Dim udtPedidoData As PedidoCompraItemData
    Dim objAlbaranPersist As AlbaranCompraPersist
    Dim udtAlbaran As AlbaranCompraProps
    Dim udtAlbaranData As AlbaranCompraData
    

    Set rsItem = New Recordset
    strSQL = "SELECT * FROM AlbaranCompraItems WHERE AlbaranCompraItemID=" & _
        AlbaranCompraItemID
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    
    If Not IsNull(rsItem.Fields("MaterialID")) Then
        ' Leemos el albaran para obtener el numero (solo tenemos el ID)
        Set objAlbaranPersist = mobjContext.CreateInstance("GCServerMTS.AlbaranCompraPersist")
        udtAlbaranData.Buffer = objAlbaranPersist.Fetch(mudtProps.AlbaranCompraID)
        LSet udtAlbaran = udtAlbaranData
        Set objAlbaranPersist = Nothing
    
        ' Primeramente actualizamos el stock por la parte del albarán.
        With udtMovimiento
            .MaterialID = rsItem.Fields("MaterialID")
            .Tipo = TMMSalida
            .Concepto = "Corrección albarán de Compra " & CStr(udtAlbaran.Numero)
            .Cantidad = rsItem.Fields("Cantidad")
            .PrecioEUR = rsItem.Fields("PrecioCosteEUR")
            .PrecioPTA = rsItem.Fields("PrecioCostePTA")
            .DocumentoID = mudtProps.AlbaranCompraItemID
            .TipoDocumento = "A"
            .IsDeleted = False
            .IsDirty = True
            .IsNew = True
        End With
        
        LSet udtMovimientoData = udtMovimiento
        
        Set objPersist = mobjContext.CreateInstance("GCServerMTS.MoviMaterialPersist")
        
        objPersist.Save (udtMovimientoData.Buffer)
        
        Set objPersist = Nothing
        
        ' Despues desactualizamos el stock por la parte del pedido.
        If mudtProps.PedidoCompraItemID Then
           Set objPedidoPersist = mobjContext.CreateInstance("GCServerMTS.PedidoCompraItemPersist")
           udtPedidoData.Buffer = objPedidoPersist.Fetch(mudtProps.PedidoCompraItemID)
           LSet udtPedido = udtPedidoData
           With udtPedido
                .Servido = .Servido - rsItem.Fields("Cantidad")
                .IsDeleted = False
                .IsDirty = True
                .IsNew = False
                .ActualizarAlbaran = False
                .ActualizarAlta = False
                .DesactualizarAlbaran = True
                .DesactualizarAlta = False
           End With
           LSet udtPedidoData = udtPedido
           udtPedidoData.Buffer = objPedidoPersist.Save(udtPedidoData.Buffer)
           Set objPedidoPersist = Nothing
        End If
    End If
    
    rsItem.Close

    Set rsItem = Nothing

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

