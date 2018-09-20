VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "OrdenCorteItemPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext

Private mudtProps As OrdenCorteItemProps

Private Function GetState() As String
    Dim udtData As OrdenCorteItemData
    
    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As OrdenCorteItemData
    
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String
    Dim rsItem As Recordset
    Dim strSQL As String
    
    On Error GoTo errh
    SetState Buffer
    
    'Cuando es un item de orden de corte nuevo, hay que marcar el pedido.
    If mudtProps.IsNew Then ActualizarPedido
    
    Set rsItem = New Recordset
    strSQL = "SELECT * FROM OrdenCorteItems WHERE OrdenCorteItemID=" & _
      mudtProps.OrdenCorteItemID
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsItem.AddNew
    
    With rsItem
        If mudtProps.Desactualizar Then LeerYDesactualizar (mudtProps.OrdenCorteItemID)
        If mudtProps.Actualizar Then Actualizar
         
        ' Actualizamos el registro
        .Fields("OrdenCorteID") = mudtProps.OrdenCorteID
        .Fields("ArticuloColorID") = mudtProps.ArticuloColorID
        .Fields("ArticuloID") = mudtProps.ArticuloID
        .Fields("PedidoVentaItemID") = mudtProps.PedidoVentaItemID
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
        .Fields("TemporadaID") = mudtProps.TemporadaID
    
        .Update
    
        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.OrdenCorteItemID = .Fields("OrdenCorteItemID")
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

Public Sub DeleteObject(ByVal OrdenCorteItemID As Long)
    Dim cnGescom As Connection
    
    On Error GoTo errh
  
    LeerYDesactualizarPedido (mudtProps.OrdenCorteItemID)
    
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    cnGescom.Execute "DELETE FROM OrdenCorteItems WHERE OrdenCorteItemID=" & _
      OrdenCorteItemID
    
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

Public Function Fetch(ByVal OrdenCorteItemID As Long) As String
    Dim rsItem As Recordset
    Dim strSQL As String
    
    On Error GoTo errh
    strSQL = "SELECT * FROM vOrdenCorteItems WHERE OrdenCorteItemID=" & _
      OrdenCorteItemID
    Set rsItem = New Recordset
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    With rsItem
      mudtProps.OrdenCorteItemID = .Fields("OrdenCorteItemID")
      mudtProps.OrdenCorteID = .Fields("OrdenCorteID")
      mudtProps.ArticuloColorID = .Fields("ArticuloColorID")
      mudtProps.ArticuloID = .Fields("ArticuloID")
      mudtProps.PedidoVentaItemID = .Fields("PedidoVentaItemID")
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
      mudtProps.Numero = .Fields("Numero")
      mudtProps.Descripcion = .Fields("Descripcion") & vbNullString
      mudtProps.Cliente = .Fields("Cliente") & vbNullString
      mudtProps.TemporadaID = .Fields("TemporadaID")
      
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

Public Function SaveChild(OrdenCorteID As Long, Buffer As String) As String
    
    On Error GoTo errh
    SetState Buffer
    
    With mudtProps
      If .IsDeleted Then
        DeleteObject .OrdenCorteItemID
        SaveChild = vbNullString
    
      ElseIf .IsDirty Or .IsNew Then
        .OrdenCorteID = OrdenCorteID
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

' Se encarga de actualizar la orden de corte.
Private Sub Actualizar()
    Dim objArticuloPersist As ArticuloPersist
    Dim udtArticulo As ArticuloProps
    Dim udtArticuloData As ArticuloData
    
    Dim objArticuloColorPersist As ArticuloColorPersist
    Dim udtArticuloColor As ArticuloColorProps
    Dim udtArticuloColorData As ArticuloColorData
    
    Dim objSeriePersist As SeriePersist
    Dim udtSerie As SerieProps
    Dim udtSerieData As SerieData
    
    Dim objMaterialPersist As MaterialPersist
    Dim udtMaterial As MaterialProps
    Dim udtMaterialData As MaterialData
    
    Dim objModeloPersist As ModeloPersist
    Dim udtModelo As ModeloProps
    Dim udtModeloData As ModeloData
    
    Dim objMoviArticuloPersist As MoviArticuloPersist
    Dim udtMoviArticulo As MoviArticuloProps
    Dim udtMoviArticuloData As MoviArticuloData
    
    Dim objMoviMaterialPersist As MoviMaterialPersist
    Dim udtMoviMaterial As MoviMaterialProps
    Dim udtMoviMaterialData As MoviMaterialData
    
    Dim udtEstrModelo As EstrModeloProps
    Dim udtEstrModeloData As EstrModeloData
    
    Dim objEstrModelosPersist As EstrModelosPersist
    'Dim udtEstrModelosData As String
    
    Dim objOrdenCortePersist As OrdenCortePersist
    Dim udtOrdenCorte As OrdenCorteProps
    Dim udtOrdenCorteData As OrdenCorteData

    Dim Cantidad As Long
    Dim objBuffer As Buffer
    Dim lngIndex As Long
    
    Dim objEstrModeloPersist As EstrModeloPersist

    
    mudtProps.Situacion = "F"
    
    ' Leemos el OrdenCorte para obtener el numero (solo tenemos el ID)
    Set objOrdenCortePersist = mobjContext.CreateInstance("GCServerMTS.OrdenCortePersist")
    udtOrdenCorteData.Buffer = objOrdenCortePersist.Fetch(mudtProps.OrdenCorteID)
    LSet udtOrdenCorte = udtOrdenCorteData
    Set objOrdenCortePersist = Nothing

    ' Leemos el articulocolor y el articulo
    Set objArticuloColorPersist = mobjContext.CreateInstance("GCServerMTS.ArticuloColorPersist")
    udtArticuloColorData.Buffer = objArticuloColorPersist.Fetch(mudtProps.ArticuloColorID)
    LSet udtArticuloColor = udtArticuloColorData
    
    Set objArticuloPersist = mobjContext.CreateInstance("GCServerMTS.ArticuloPersist")
    udtArticuloData.Buffer = objArticuloPersist.Fetch(udtArticuloColor.ArticuloID)
    LSet udtArticulo = udtArticuloData
    
    ' Leemos la serie y el modelo
    Set objSeriePersist = mobjContext.CreateInstance("GCServerMTS.SeriePersist")
    udtSerieData.Buffer = objSeriePersist.Fetch(udtArticulo.SerieID)
    LSet udtSerie = udtSerieData
    
    Set objModeloPersist = mobjContext.CreateInstance("GCServerMTS.ModeloPersist")
    udtModeloData.Buffer = objModeloPersist.Fetch(udtArticulo.ModeloID)
    LSet udtModelo = udtModeloData
    
    ' Leemos la tela de la serie
    Set objMaterialPersist = mobjContext.CreateInstance("GCServerMTS.MaterialPersist")
    udtMaterialData.Buffer = objMaterialPersist.Fetch(udtSerie.MaterialID)
    LSet udtMaterial = udtMaterialData
    
    ' Movimiento de entrada de articulos
    With udtMoviArticulo
        .ArticuloColorID = mudtProps.ArticuloColorID
        .Tipo = TMMEntrada
        .Concepto = "Orden de corte " & CStr(udtOrdenCorte.Numero)
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
        .IsDeleted = False
        .IsDirty = True
        .IsNew = True
    End With
    
    LSet udtMoviArticuloData = udtMoviArticulo
    Set objMoviArticuloPersist = mobjContext.CreateInstance("GCServerMTS.MoviArticuloPersist")
    objMoviArticuloPersist.Save (udtMoviArticuloData.Buffer)
    Set objMoviArticuloPersist = Nothing
    
    ' Calculamos la cantidad de articulos total
    Cantidad = mudtProps.CantidadT36 + mudtProps.CantidadT38 + mudtProps.CantidadT40 + _
               mudtProps.CantidadT42 + mudtProps.CantidadT44 + mudtProps.CantidadT46 + _
               mudtProps.CantidadT48 + mudtProps.CantidadT50 + mudtProps.CantidadT52 + _
               mudtProps.CantidadT54 + mudtProps.CantidadT56
               
    ' Movimiento de salida de material de la serie
    ' cantidad = (Nº articulos * cantidad de tela del modelo * anchura tela estandar) _
    '            / anchura de la tela.
    With udtMoviMaterial
        .MaterialID = udtSerie.MaterialID
        .Tipo = TMMSalida
        .Concepto = "Orden de corte " & CStr(udtOrdenCorte.Numero)
        .Cantidad = Round((Cantidad * udtModelo.CantidadTela * dblAnchuraTelaEstandar) / udtMaterial.AnchuraTela, 2)
        .PrecioCosteEUR = udtMaterial.PrecioPonderadoEUR
        .PrecioCostePTA = udtMaterial.PrecioPonderadoPTA
        .DocumentoID = mudtProps.OrdenCorteItemID
        .TipoDocumento = "O"
        .IsDeleted = False
        .IsDirty = True
        .IsNew = True
    End With
    
    LSet udtMoviMaterialData = udtMoviMaterial
    Set objMoviMaterialPersist = mobjContext.CreateInstance("GCServerMTS.MoviMaterialPersist")
    objMoviMaterialPersist.Save (udtMoviMaterialData.Buffer)
    Set objMoviMaterialPersist = Nothing
    
    ' Ahora tratamos la estructura de materiales del modelo
    Set objBuffer = New Buffer
    Set objEstrModelosPersist = mobjContext.CreateInstance("GCServerMTS.EstrModelosPersist")
    objBuffer.SetState objEstrModelosPersist.Fetch(udtModelo.ModeloID)
    
    For lngIndex = 1 To objBuffer.Count
        ' Extraemos la estructura de modelo
        udtEstrModeloData.Buffer = objBuffer.Item(lngIndex)
        LSet udtEstrModelo = udtEstrModeloData
        
        ' Leemos el material de la estructura ...
        Set objMaterialPersist = mobjContext.CreateInstance("GCServerMTS.MaterialPersist")
        udtMaterialData.Buffer = objMaterialPersist.Fetch(udtEstrModelo.MaterialID)
        LSet udtMaterial = udtMaterialData
    
        ' ... y creamos el movimiento correspondiente
        With udtMoviMaterial
            .MaterialID = udtEstrModelo.MaterialID
            .Tipo = TMMSalida
            .Concepto = "Orden de corte " & CStr(udtOrdenCorte.Numero)
            .Cantidad = Round(Cantidad * udtEstrModelo.Cantidad, 2)
            .PrecioCosteEUR = udtMaterial.PrecioPonderadoEUR
            .PrecioCostePTA = udtMaterial.PrecioPonderadoPTA
            .DocumentoID = mudtProps.OrdenCorteItemID
            .TipoDocumento = "O"
            .IsDeleted = False
            .IsDirty = True
            .IsNew = True
        End With
        
        LSet udtMoviMaterialData = udtMoviMaterial
        Set objMoviMaterialPersist = mobjContext.CreateInstance("GCServerMTS.MoviMaterialPersist")
        objMoviMaterialPersist.Save (udtMoviMaterialData.Buffer)
        Set objMoviMaterialPersist = Nothing
        
    Next

    Set objEstrModeloPersist = Nothing
    
    Set objMaterialPersist = Nothing
    
End Sub

Private Sub LeerYDesactualizarPedido(ByVal OrdenCorteItemID As Long)
    Dim rsItem As Recordset
    Dim strSQL As String
    Dim objPedidoPersist As PedidoVentaItemPersist
    Dim udtPedido As PedidoVentaItemProps
    Dim udtPedidoData As PedidoVentaItemData
    
    Set rsItem = New Recordset
    strSQL = "SELECT * FROM OrdenCorteItems WHERE OrdenCorteItemID=" & _
        OrdenCorteItemID
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    
    Set objPedidoPersist = mobjContext.CreateInstance("GCServerMTS.PedidoVentaItemPersist")
    udtPedidoData.Buffer = objPedidoPersist.Fetch(rsItem.Fields("PedidoVentaItemID"))
    LSet udtPedido = udtPedidoData
    With udtPedido
         .SituacionCorte = "A"
         .IsDeleted = False
         .IsDirty = True
         .IsNew = False
         .ActualizarAlbaran = False
         .ActualizarAlta = False
         .DesactualizarAlbaran = False
         .DesactualizarAlta = False
    End With
    LSet udtPedidoData = udtPedido
    udtPedidoData.Buffer = objPedidoPersist.Save(udtPedidoData.Buffer)
    Set objPedidoPersist = Nothing
    
    rsItem.Close

    Set rsItem = Nothing

End Sub

Private Sub ActualizarPedido()
    Dim objPedidoPersist As PedidoVentaItemPersist
    Dim udtPedido As PedidoVentaItemProps
    Dim udtPedidoData As PedidoVentaItemData
    
    Set objPedidoPersist = mobjContext.CreateInstance("GCServerMTS.PedidoVentaItemPersist")
    udtPedidoData.Buffer = objPedidoPersist.Fetch(mudtProps.PedidoVentaItemID)
    LSet udtPedido = udtPedidoData
    With udtPedido
         .SituacionCorte = "C"
         .IsDeleted = False
         .IsDirty = True
         .IsNew = False
         .ActualizarAlbaran = False
         .ActualizarAlta = False
         .DesactualizarAlbaran = False
         .DesactualizarAlta = False
    End With
    LSet udtPedidoData = udtPedido
    udtPedidoData.Buffer = objPedidoPersist.Save(udtPedidoData.Buffer)
    Set objPedidoPersist = Nothing
    
End Sub

Private Sub LeerYDesactualizar(ByVal OrdenCorteItemID As Long)
    Dim rsItem As Recordset
    Dim strSQL As String
    Dim objPersist As MoviArticuloPersist
    Dim udtMovimiento As MoviArticuloProps
    Dim udtMovimientoData As MoviArticuloData
    Dim objPedidoPersist As PedidoVentaItemPersist
    Dim udtPedido As PedidoVentaItemProps
    Dim udtPedidoData As PedidoVentaItemData
    

    mudtProps.Situacion = "C"
    
    Set rsItem = New Recordset
    strSQL = "SELECT * FROM OrdenCorteItems WHERE OrdenCorteItemID=" & _
        OrdenCorteItemID
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

