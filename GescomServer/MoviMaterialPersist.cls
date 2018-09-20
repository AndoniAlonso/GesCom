VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "MoviMaterialPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
  
Private mudtProps As MoviMaterialProps

Private Function GetState() As String
    Dim udtData As MoviMaterialData
    
    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As MoviMaterialData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByRef Buffer As String) As String
    Dim rsMovimiento As Recordset
    Dim strSQL As String
    Dim objMaterialPersist As MaterialPersist
    Dim udtMaterialData As MaterialData
    Dim udtMaterial As MaterialProps
  

    On Error GoTo errh
    SetState Buffer
  
    Set rsMovimiento = New Recordset
    strSQL = "SELECT * FROM MoviMateriales WHERE MoviMaterialID=" & _
        mudtProps.MoviMaterialID
    rsMovimiento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsMovimiento.AddNew
     
    ' Leemos el Material del movimiento
    Set objMaterialPersist = mobjContext.CreateInstance("GCServerMTS.MaterialPersist")
    udtMaterialData.Buffer = objMaterialPersist.Fetch(mudtProps.MaterialID)
    LSet udtMaterial = udtMaterialData
    
    ' Calculamos el nuevo stock
    udtMaterial.IsDeleted = False
    udtMaterial.IsDirty = True
    udtMaterial.IsNew = False
    
    ' Si no se ha informado del precio del movimiento, tomamos el del material.
    If mudtProps.PrecioPTA = 0 Then
        mudtProps.PrecioPTA = udtMaterial.PrecioPonderadoPTA
    End If
    
    If mudtProps.PrecioEUR = 0 Then
        mudtProps.PrecioEUR = udtMaterial.PrecioPonderadoEUR
    End If
    
    Select Case mudtProps.Tipo
        Case TMMReserva
            udtMaterial.StockPendiente = udtMaterial.StockPendiente + mudtProps.Cantidad
        
        Case TMMEntrega
            udtMaterial.StockPendiente = udtMaterial.StockPendiente - mudtProps.Cantidad
        
        Case TMMEntrada
            ' Se calcula el precio medio ponderado para una entrada de material
            ' Se tiene en cuenta que el stock resultante puede ser cero.
            If udtMaterial.StockActual + mudtProps.Cantidad <> 0 Then
                udtMaterial.PrecioPonderadoPTA = _
                 Round(((udtMaterial.PrecioPonderadoPTA * udtMaterial.StockActual) + _
                     (mudtProps.PrecioPTA * mudtProps.Cantidad)) / _
                     (udtMaterial.StockActual + mudtProps.Cantidad), 0)
                    
                udtMaterial.PrecioPonderadoEUR = _
                 Round(((udtMaterial.PrecioPonderadoEUR * udtMaterial.StockActual) + _
                     (mudtProps.PrecioEUR * mudtProps.Cantidad)) / _
                     (udtMaterial.StockActual + mudtProps.Cantidad), 2)
            Else
                udtMaterial.PrecioPonderadoPTA = 0
                udtMaterial.PrecioPonderadoEUR = 0
            End If
            
            udtMaterial.StockActual = udtMaterial.StockActual + mudtProps.Cantidad
        
        Case TMMSalida
            ' Se calcula el precio medio ponderado para una salida de material
            ' Se tiene en cuenta que el stock resultante puede ser cero.
            If udtMaterial.StockActual - mudtProps.Cantidad <> 0 Then
                udtMaterial.PrecioPonderadoPTA = _
                 Round(((udtMaterial.PrecioPonderadoPTA * udtMaterial.StockActual) - _
                     (mudtProps.PrecioPTA * mudtProps.Cantidad)) / _
                     (udtMaterial.StockActual - mudtProps.Cantidad), 0)
                    
                udtMaterial.PrecioPonderadoEUR = _
                 Round(((udtMaterial.PrecioPonderadoEUR * udtMaterial.StockActual) - _
                     (mudtProps.PrecioEUR * mudtProps.Cantidad)) / _
                     (udtMaterial.StockActual - mudtProps.Cantidad), 2)
            Else
                udtMaterial.PrecioPonderadoPTA = 0
                udtMaterial.PrecioPonderadoEUR = 0
            End If
            udtMaterial.StockActual = udtMaterial.StockActual - mudtProps.Cantidad
        
    End Select
   
    ' grabamos el registro resultante
    LSet udtMaterialData = udtMaterial
    udtMaterialData.Buffer = objMaterialPersist.Save(udtMaterialData.Buffer)
    LSet udtMaterial = udtMaterialData
    
    Set objMaterialPersist = Nothing
    
     
    With rsMovimiento
        'Creo que lo mejor es que obtenga la fecha real aqui.
        mudtProps.Fecha = Now
        .Fields("Fecha") = mudtProps.Fecha
        .Fields("MaterialID") = mudtProps.MaterialID
        .Fields("Tipo") = mudtProps.Tipo
        .Fields("Concepto") = mudtProps.Concepto
        .Fields("Cantidad") = mudtProps.Cantidad
        Select Case mudtProps.Tipo
        Case TMMReserva, TMMEntrega
            .Fields("StockFinal") = udtMaterial.StockPendiente
        
        Case TMMEntrada, TMMSalida
            .Fields("StockFinal") = udtMaterial.StockActual
        
        End Select
        .Fields("PrecioEUR") = mudtProps.PrecioEUR
        .Fields("PrecioPTA") = mudtProps.PrecioPTA
        .Fields("PrecioCosteEUR") = udtMaterial.PrecioCosteEUR
        .Fields("PrecioCostePTA") = udtMaterial.PrecioCostePTA
        .Fields("DocumentoID") = mudtProps.DocumentoID
        .Fields("TipoDocumento") = mudtProps.TipoDocumento
        
        .Update
      
        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.MoviMaterialID = .Fields("MoviMaterialID")
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

Public Sub DeleteObject(ByVal MoviMaterialID As Long)
    Dim cnGescom As Connection


    On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM MoviMateriales WHERE MoviMaterialID=" & _
        MoviMaterialID

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

Public Function Fetch(ByVal MoviMaterialID As Long) As String
    Dim rsMovimiento As Recordset
    Dim strSQL As String
  

    On Error GoTo errh
    strSQL = "SELECT * FROM MoviMateriales WHERE MoviMaterialID=" & _
        MoviMaterialID
    Set rsMovimiento = New Recordset
    rsMovimiento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsMovimiento
        mudtProps.MoviMaterialID = .Fields("MoviMaterialID")
        mudtProps.Fecha = .Fields("Fecha")
        mudtProps.MaterialID = .Fields("MaterialID")
        mudtProps.Tipo = .Fields("Tipo") & vbNullString
        mudtProps.Concepto = .Fields("Concepto") & vbNullString
        mudtProps.Cantidad = .Fields("Cantidad")
        mudtProps.StockFinal = .Fields("StockFinal")
        mudtProps.PrecioEUR = .Fields("PrecioEUR")
        mudtProps.PrecioPTA = .Fields("PrecioPTA")
        mudtProps.PrecioCosteEUR = .Fields("PrecioCosteEUR")
        mudtProps.PrecioCostePTA = .Fields("PrecioCostePTA")
        mudtProps.DocumentoID = .Fields("TipoDocumentoID")
        mudtProps.TipoDocumento = .Fields("TipoDocumento")

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

