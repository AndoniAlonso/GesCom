VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "TraspasoItemPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext

Private mudtProps As TraspasoItemProps
Private mudtSave As TraspasoItemProps

Private Function GetState() As String
    Dim udtData As TraspasoItemData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As TraspasoItemData

    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String
    Dim rsItem As Recordset
    Dim strSQL As String

    On Error GoTo errh
    SetState Buffer

    Set rsItem = New Recordset
    strSQL = "SELECT * FROM TraspasoItems WHERE TraspasoItemID=" & _
        mudtProps.TraspasoItemID
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    
    If mudtProps.IsNew Then
        rsItem.AddNew
    End If
    
    With rsItem

        ' Actualizamos el registro
        .Fields("TraspasoID") = mudtProps.TraspasoID
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
        .Fields("Observaciones") = IIf(Len(Trim(mudtProps.Observaciones)) = 0, Null, mudtProps.Observaciones)
        .Update

        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.TraspasoItemID = .Fields("TraspasoItemID")
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

Public Sub DeleteObject(ByVal TraspasoItemID As Long)
    Dim cnGescom As Connection

    On Error GoTo errh
  
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM TraspasoItems WHERE TraspasoItemID=" & _
        TraspasoItemID

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

Public Function Fetch(ByVal TraspasoItemID As Long) As String
    Dim rsItem As Recordset
    Dim strSQL As String

  On Error GoTo errh
    strSQL = "SELECT * FROM TraspasoItems WHERE TraspasoItemID=" & _
        TraspasoItemID
    Set rsItem = New Recordset
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsItem
        mudtProps.TraspasoItemID = .Fields("TraspasoItemID")
        mudtProps.TraspasoID = .Fields("TraspasoID")
        mudtProps.ArticuloColorID = .Fields("ArticuloColorID")
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
        mudtProps.Observaciones = IIf(IsNull(.Fields("Observaciones")), vbNullString, .Fields("Observaciones"))
    
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

Public Function SaveChild(TraspasoID As Long, Buffer As String) As String

    On Error GoTo errh
    SetState Buffer

    With mudtProps
        If .IsDeleted Then
            DeleteObject .TraspasoItemID
            SaveChild = vbNullString

        ElseIf .IsDirty Or .IsNew Then
            .TraspasoID = TraspasoID
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

Public Sub Enviar(ByVal AlmacenTransitoID As Long, ByVal AlmacenOrigenID As Long, ByVal Buffer As String)
    Dim objPersist As MoviArticuloPersist
    Dim udtMovimiento As MoviArticuloProps
    Dim udtMovimientoData As MoviArticuloData

    On Error GoTo errh
    SetState Buffer
    
    Me.Save Buffer
    
    ' Primero realizamos el movimiento de salida del almacen origen
    With udtMovimiento
        .ArticuloColorID = mudtProps.ArticuloColorID
        .Tipo = TMMSalida
        .Concepto = "Salida por traspaso"
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
        .AlmacenID = AlmacenOrigenID
        .IsDeleted = False
        .IsDirty = True
        .IsNew = True
    End With
    
    LSet udtMovimientoData = udtMovimiento
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.MoviArticuloPersist")
    
    objPersist.Save (udtMovimientoData.Buffer)
    
    Set objPersist = Nothing


    ' después realizamos el movimiento de entrada en el almacen en transito
    With udtMovimiento
        .ArticuloColorID = mudtProps.ArticuloColorID
        .Tipo = TMMEntrada
        .Concepto = "Entrada por traspaso en almacen en transito"
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
        .AlmacenID = AlmacenTransitoID
        .IsDeleted = False
        .IsDirty = True
        .IsNew = True
    End With
    
    LSet udtMovimientoData = udtMovimiento
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.MoviArticuloPersist")
    
    objPersist.Save (udtMovimientoData.Buffer)
    
    Set objPersist = Nothing
    
    mobjContext.SetComplete

    Exit Sub
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub


Public Sub Recepcionar(ByVal AlmacenTransitoID As Long, ByVal AlmacenDestinoID As Long, ByVal Buffer As String)
    Dim objPersist As MoviArticuloPersist
    Dim udtMovimiento As MoviArticuloProps
    Dim udtMovimientoData As MoviArticuloData

    On Error GoTo errh
    SetState Buffer
    
    Me.Save Buffer
    
    ' Primero realizamos el movimiento de salida del almacen en transito
    With udtMovimiento
        .ArticuloColorID = mudtProps.ArticuloColorID
        .Tipo = TMMSalida
        .Concepto = "Salida por traspaso del almacen en transito"
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
        .AlmacenID = AlmacenTransitoID
        .IsDeleted = False
        .IsDirty = True
        .IsNew = True
    End With
    
    LSet udtMovimientoData = udtMovimiento
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.MoviArticuloPersist")
    
    objPersist.Save (udtMovimientoData.Buffer)
    
    Set objPersist = Nothing


    ' después realizamos el movimiento de entrada en el almacen destino
    With udtMovimiento
        .ArticuloColorID = mudtProps.ArticuloColorID
        .Tipo = TMMEntrada
        .Concepto = "Entrada por traspaso en almacen destino"
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
        .AlmacenID = AlmacenDestinoID
        .IsDeleted = False
        .IsDirty = True
        .IsNew = True
    End With
    
    LSet udtMovimientoData = udtMovimiento
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.MoviArticuloPersist")
    
    objPersist.Save (udtMovimientoData.Buffer)
    
    Set objPersist = Nothing
    
    mobjContext.SetComplete

    Exit Sub
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Public Sub EnviarYRecepcionar(ByVal AlmacenOrigenID As Long, ByVal AlmacenDestinoID As Long, ByVal Buffer As String)
    Dim objPersist As MoviArticuloPersist
    Dim udtMovimiento As MoviArticuloProps
    Dim udtMovimientoData As MoviArticuloData

    On Error GoTo errh
    SetState Buffer
    
    Me.Save Buffer
    
    ' Primero realizamos el movimiento de salida del almacen origen
    With udtMovimiento
        .ArticuloColorID = mudtProps.ArticuloColorID
        .Tipo = TMMSalida
        .Concepto = "Salida por traspaso desde el almacen origen"
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
        .AlmacenID = AlmacenOrigenID
        .IsDeleted = False
        .IsDirty = True
        .IsNew = True
    End With
    
    LSet udtMovimientoData = udtMovimiento
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.MoviArticuloPersist")
    
    objPersist.Save (udtMovimientoData.Buffer)
    
    Set objPersist = Nothing


    ' después realizamos el movimiento de entrada en el almacen destino
    With udtMovimiento
        .ArticuloColorID = mudtProps.ArticuloColorID
        .Tipo = TMMEntrada
        .Concepto = "Entrada por traspaso en almacen destino"
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
        .AlmacenID = AlmacenDestinoID
        .IsDeleted = False
        .IsDirty = True
        .IsNew = True
    End With
    
    LSet udtMovimientoData = udtMovimiento
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.MoviArticuloPersist")
    
    objPersist.Save (udtMovimientoData.Buffer)
    
    Set objPersist = Nothing
    
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
