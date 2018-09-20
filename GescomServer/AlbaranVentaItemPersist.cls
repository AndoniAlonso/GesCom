VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "AlbaranVentaItemPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext

Private mudtProps As AlbaranVentaItemProps

Private Function GetState() As String

    Dim udtData As AlbaranVentaItemData
    
    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

    Dim udtData As AlbaranVentaItemData
    
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String

    Dim rsItem As Recordset
    Dim strSQL As String
    
  On Error GoTo errh
    SetState Buffer
    
    Set rsItem = New Recordset
    strSQL = "SELECT * FROM AlbaranVentaItems WHERE AlbaranVentaItemID=" & _
      mudtProps.AlbaranVentaItemID
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsItem.AddNew
    
    With rsItem
        If mudtProps.DesactualizarAlta Then LeerYDesactualizar (mudtProps.AlbaranVentaItemID)
        If mudtProps.ActualizarAlta Then Actualizar
         
        ' Actualizamos el registro
        .Fields("AlbaranVentaID") = mudtProps.AlbaranVentaID
        ' Solo asigno el valor si viene informado
        .Fields("ArticuloColorID") = IIf(mudtProps.ArticuloColorID, mudtProps.ArticuloColorID, Null)
        .Fields("Descripcion") = mudtProps.Descripcion
        ' Solo asigno el valor si viene informado
        .Fields("PedidoVentaItemID") = IIf(mudtProps.PedidoVentaItemID, mudtProps.PedidoVentaItemID, Null)
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
        .Fields("PrecioVentaPTA") = mudtProps.PrecioVentaPTA
        .Fields("PrecioVentaEUR") = mudtProps.PrecioVentaEUR
        .Fields("Descuento") = mudtProps.Descuento
        .Fields("BrutoPTA") = mudtProps.BrutoPTA
        .Fields("BrutoEUR") = mudtProps.BrutoEUR
        .Fields("Comision") = mudtProps.Comision
        .Fields("TemporadaID") = mudtProps.TemporadaID
        .Fields("FacturadoAB") = mudtProps.FacturadoAB
        .Fields("FacturaVentaItemIDA") = IIf(mudtProps.FacturaVentaItemIDA, mudtProps.FacturaVentaItemIDA, Null)
        .Fields("FacturaVentaItemIDB") = IIf(mudtProps.FacturaVentaItemIDB, mudtProps.FacturaVentaItemIDB, Null)
        .Fields("AlmacenID") = IIf(mudtProps.AlmacenID, mudtProps.AlmacenID, Null)
    
        .Update
    
        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.AlbaranVentaItemID = .Fields("AlbaranVentaItemID")
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

Public Sub DeleteObject(ByVal AlbaranVentaItemID As Long)
    Dim cnGescom As Connection
    
    On Error GoTo errh
  
    If Not mudtProps.IsNew Then
        ' Al borrar una linea de albaranes hay que desactualizar SIEMPRE.
        LeerYDesactualizar (AlbaranVentaItemID)
    End If
  
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    cnGescom.Execute "DELETE FROM AlbaranVentaItems WHERE AlbaranVentaItemID=" & _
      AlbaranVentaItemID
    
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

Public Function Fetch(ByVal AlbaranVentaItemID As Long) As String

    Dim rsItem As Recordset
    Dim strSQL As String
    
  On Error GoTo errh
    strSQL = "SELECT * FROM AlbaranVentaItems WHERE AlbaranVentaItemID=" & _
      AlbaranVentaItemID
    Set rsItem = New Recordset
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    With rsItem
      mudtProps.AlbaranVentaItemID = .Fields("AlbaranVentaItemID")
      mudtProps.AlbaranVentaID = .Fields("AlbaranVentaID")
      mudtProps.ArticuloColorID = IIf(IsNull(.Fields("ArticuloColorID")), 0, .Fields("ArticuloColorID"))
      mudtProps.Descripcion = .Fields("Descripcion") & vbNullString
      mudtProps.PedidoVentaItemID = IIf(IsNull(.Fields("PedidoVentaItemID")), 0, .Fields("PedidoVentaItemID"))
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
      mudtProps.PrecioVentaPTA = .Fields("PrecioVentaPTA")
      mudtProps.PrecioVentaEUR = .Fields("PrecioVentaEUR")
      mudtProps.Descuento = .Fields("Descuento")
      mudtProps.BrutoPTA = .Fields("BrutoPTA")
      mudtProps.BrutoEUR = .Fields("BrutoEUR")
      mudtProps.Comision = .Fields("Comision")
      mudtProps.TemporadaID = .Fields("TemporadaID")
      mudtProps.FacturadoAB = .Fields("FacturadoAB")
      mudtProps.FacturaVentaItemIDA = IIf(IsNull(.Fields("FacturaVentaItemIDA")), 0, .Fields("FacturaVentaItemIDA"))
      mudtProps.FacturaVentaItemIDB = IIf(IsNull(.Fields("FacturaVentaItemIDB")), 0, .Fields("FacturaVentaItemIDB"))
      mudtProps.AlmacenID = IIf(IsNull(.Fields("AlmacenID")), 0, .Fields("AlmacenID"))
      
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


Public Function SaveChild(AlbaranVentaID As Long, Buffer As String) As String
    
    On Error GoTo errh
    SetState Buffer
    
    With mudtProps
      If .IsDeleted Then
        DeleteObject .AlbaranVentaItemID
        SaveChild = vbNullString
    
      ElseIf .IsDirty Or .IsNew Then
        .AlbaranVentaID = AlbaranVentaID
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
    Dim objPersist As MoviArticuloPersist
    Dim udtMovimiento As MoviArticuloProps
    Dim udtMovimientoData As MoviArticuloData
    Dim objPedidoPersist As PedidoVentaItemPersist
    Dim udtPedido As PedidoVentaItemProps
    Dim udtPedidoData As PedidoVentaItemData
    
    ' Primero actualizamos el stock por la parte del pedido.
    If mudtProps.PedidoVentaItemID Then
        Set objPedidoPersist = mobjContext.CreateInstance("GCServerMTS.PedidoVentaItemPersist")
        udtPedidoData.Buffer = objPedidoPersist.Fetch(mudtProps.PedidoVentaItemID)
        LSet udtPedido = udtPedidoData
        With udtPedido
            .ServidoT36 = .ServidoT36 + mudtProps.CantidadT36
            .ServidoT38 = .ServidoT38 + mudtProps.CantidadT38
            .ServidoT40 = .ServidoT40 + mudtProps.CantidadT40
            .ServidoT42 = .ServidoT42 + mudtProps.CantidadT42
            .ServidoT44 = .ServidoT44 + mudtProps.CantidadT44
            .ServidoT46 = .ServidoT46 + mudtProps.CantidadT46
            .ServidoT48 = .ServidoT48 + mudtProps.CantidadT48
            .ServidoT50 = .ServidoT50 + mudtProps.CantidadT50
            .ServidoT52 = .ServidoT52 + mudtProps.CantidadT52
            .ServidoT54 = .ServidoT54 + mudtProps.CantidadT54
            .ServidoT56 = .ServidoT56 + mudtProps.CantidadT56
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
    
    With udtMovimiento
        .ArticuloColorID = mudtProps.ArticuloColorID
        .Tipo = TMMSalida
        .Concepto = "Albarán de venta"
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
        .PrecioEUR = mudtProps.PrecioVentaEUR
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
Private Sub LeerYDesactualizar(ByVal AlbaranVentaItemID As Long)

    Dim rsItem As Recordset
    Dim strSQL As String
    Dim objPersist As MoviArticuloPersist
    Dim udtMovimiento As MoviArticuloProps
    Dim udtMovimientoData As MoviArticuloData
    Dim objPedidoPersist As PedidoVentaItemPersist
    Dim udtPedido As PedidoVentaItemProps
    Dim udtPedidoData As PedidoVentaItemData
    

    Set rsItem = New Recordset
    strSQL = "SELECT * FROM AlbaranVentaItems WHERE AlbaranVentaItemID=" & _
        AlbaranVentaItemID
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    
    If Not IsNull(rsItem.Fields("articulocolorid")) Then
        ' Primeramente actualizamos el stock por la parte del albarán.
        With udtMovimiento
            .ArticuloColorID = rsItem.Fields("ArticuloColorID")
            .Tipo = TMMEntrada
            .Concepto = "Corrección albarán de venta"
            .CantidadT36 = rsItem.Fields("CantidadT36")
            .CantidadT38 = rsItem.Fields("CantidadT38")
            .CantidadT40 = rsItem.Fields("CantidadT40")
            .CantidadT42 = rsItem.Fields("CantidadT42")
            .CantidadT44 = rsItem.Fields("CantidadT44")
            .CantidadT46 = rsItem.Fields("CantidadT46")
            .CantidadT48 = rsItem.Fields("CantidadT48")
            .CantidadT50 = rsItem.Fields("CantidadT50")
            .CantidadT52 = rsItem.Fields("CantidadT52")
            .CantidadT54 = rsItem.Fields("CantidadT54")
            .CantidadT56 = rsItem.Fields("CantidadT56")
            .PrecioEUR = rsItem.Fields("PrecioVentaEUR")
            .AlmacenID = rsItem.Fields("AlmacenID")
            .IsDeleted = False
            .IsDirty = True
            .IsNew = True
        End With
        
        LSet udtMovimientoData = udtMovimiento
        
        Set objPersist = mobjContext.CreateInstance("GCServerMTS.MoviArticuloPersist")
        
        objPersist.Save (udtMovimientoData.Buffer)
        
        Set objPersist = Nothing
        
        ' Despues desactualizamos el stock por la parte del pedido.
        If mudtProps.PedidoVentaItemID Then
           Set objPedidoPersist = mobjContext.CreateInstance("GCServerMTS.PedidoVentaItemPersist")
           udtPedidoData.Buffer = objPedidoPersist.Fetch(mudtProps.PedidoVentaItemID)
           LSet udtPedido = udtPedidoData
           With udtPedido
                .ServidoT36 = .ServidoT36 - rsItem.Fields("CantidadT36")
                .ServidoT38 = .ServidoT38 - rsItem.Fields("CantidadT38")
                .ServidoT40 = .ServidoT40 - rsItem.Fields("CantidadT40")
                .ServidoT42 = .ServidoT42 - rsItem.Fields("CantidadT42")
                .ServidoT44 = .ServidoT44 - rsItem.Fields("CantidadT44")
                .ServidoT46 = .ServidoT46 - rsItem.Fields("CantidadT46")
                .ServidoT48 = .ServidoT48 - rsItem.Fields("CantidadT48")
                .ServidoT50 = .ServidoT50 - rsItem.Fields("CantidadT50")
                .ServidoT52 = .ServidoT52 - rsItem.Fields("CantidadT52")
                .ServidoT54 = .ServidoT54 - rsItem.Fields("CantidadT54")
                .ServidoT56 = .ServidoT56 - rsItem.Fields("CantidadT56")
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

