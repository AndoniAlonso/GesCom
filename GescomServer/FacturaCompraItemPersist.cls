VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "FacturaCompraItemPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ObjectControl

Private mobjContext As ObjectContext

Private mudtProps As FacturaCompraItemProps

Private Function GetState() As String
    Dim udtData As FacturaCompraItemData
    
    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As FacturaCompraItemData
    
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String
    Dim rsItem As Recordset
    Dim strSQL As String
    
    On Error GoTo errh
    SetState Buffer
    
    Set rsItem = New Recordset
    strSQL = "SELECT * FROM FacturaCompraItems WHERE FacturaCompraItemID=" & _
        mudtProps.FacturaCompraItemID
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsItem.AddNew
    
    With rsItem
        If mudtProps.DesactualizarAlta Then LeerYDesactualizar (mudtProps.FacturaCompraItemID)
        If mudtProps.ActualizarAlta Then Actualizar
         
        .Fields("FacturaCompraID") = mudtProps.FacturaCompraID
        ' Solo asigno el valor si viene informado
        .Fields("MaterialID") = IIf(mudtProps.MaterialID, mudtProps.MaterialID, Null)
        ' Solo asigno el valor si viene informado
        .Fields("AlbaranCompraItemID") = IIf(mudtProps.AlbaranCompraItemID, mudtProps.AlbaranCompraItemID, Null)
        .Fields("SituacionImpresa") = mudtProps.SituacionImpresa
        .Fields("Cantidad") = mudtProps.Cantidad
        .Fields("PrecioCostePTA") = mudtProps.PrecioCostePTA
        .Fields("PrecioCosteEUR") = mudtProps.PrecioCosteEUR
        .Fields("Descuento") = mudtProps.Descuento
        .Fields("BrutoPTA") = mudtProps.BrutoPTA
        .Fields("BrutoEUR") = mudtProps.BrutoEUR
        
        .Update
    
        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.FacturaCompraItemID = .Fields("FacturaCompraItemID")
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

Public Sub DeleteObject(ByVal FacturaCompraItemID As Long)
    Dim cnGescom As Connection
    
    On Error GoTo errh
    
    If Not mudtProps.IsNew Then
        ' Al borrar una linea de facturas hay que desactualizar SIEMPRE.
        LeerYDesactualizar (FacturaCompraItemID)
    End If
  
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    cnGescom.Execute "DELETE FROM FacturaCompraItems WHERE FacturaCompraItemID=" & _
        FacturaCompraItemID
    
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

Public Function Fetch(ByVal FacturaCompraItemID As Long) As String
    Dim rsItem As Recordset
    Dim strSQL As String
    
    On Error GoTo errh
    strSQL = "SELECT * FROM vFacturaCompraItems WHERE FacturaCompraItemID=" & _
        FacturaCompraItemID
    Set rsItem = New Recordset
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    With rsItem
        mudtProps.FacturaCompraItemID = .Fields("FacturaCompraItemID")
        mudtProps.FacturaCompraID = .Fields("FacturaCompraID")
        mudtProps.MaterialID = .Fields("MaterialID")
        mudtProps.NombreMaterial = .Fields("NombreMaterial")
        mudtProps.AlbaranCompraItemID = IIf(IsNull(.Fields("AlbaranCompraItemID")), 0, .Fields("AlbaranCompraItemID"))
        mudtProps.SituacionImpresa = .Fields("SituacionImpresa") & vbNullString
        mudtProps.Cantidad = .Fields("Cantidad")
        mudtProps.PrecioCostePTA = .Fields("PrecioCostePTA")
        mudtProps.PrecioCosteEUR = .Fields("PrecioCosteEUR")
        mudtProps.Descuento = .Fields("Descuento")
        mudtProps.BrutoPTA = .Fields("BrutoPTA")
        mudtProps.BrutoEUR = .Fields("BrutoEUR")
          
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

Public Function SaveChild(FacturaCompraID As Long, Buffer As String) As String

    On Error GoTo errh
    SetState Buffer
    
    With mudtProps
        If .IsDeleted Then
            DeleteObject .FacturaCompraItemID
            SaveChild = vbNullString
    
        ElseIf .IsDirty Or .IsNew Then
            .FacturaCompraID = FacturaCompraID
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

' Se encarga de actualizar la factura.
' - Marcar el albarán relacionado como facturado.
' Más adelante se controlará el precio medio ponderado.
' también se podrá alimentar datos de facturación por cliente.
Private Sub Actualizar()
    Dim objAlbaranPersist As AlbaranCompraItemPersist
    Dim udtAlbaran As AlbaranCompraItemProps
    Dim udtAlbaranData As AlbaranCompraItemData
    
    ' Actualizamos el stock por la parte del Albaran.
    If mudtProps.AlbaranCompraItemID Then
        Set objAlbaranPersist = mobjContext.CreateInstance("GCServerMTS.AlbaranCompraItemPersist")
        udtAlbaranData.Buffer = objAlbaranPersist.Fetch(mudtProps.AlbaranCompraItemID)
        LSet udtAlbaran = udtAlbaranData
        With udtAlbaran
            .Situacion = "F" ' Situacion <- facturado
            .IsDeleted = False
            .IsDirty = True
            .IsNew = False
            .ActualizarAlta = False
            .ActualizarFactura = True
            .DesactualizarAlta = False
            .DesactualizarFactura = False
        End With
        LSet udtAlbaranData = udtAlbaran
        udtAlbaranData.Buffer = objAlbaranPersist.Save(udtAlbaranData.Buffer)
        Set objAlbaranPersist = Nothing
    End If
    
End Sub

Private Sub LeerYDesactualizar(ByVal FacturaCompraItemID As Long)
    Dim rsItem As Recordset
    Dim strSQL As String
    Dim objAlbaranPersist As AlbaranCompraItemPersist
    Dim udtAlbaran As AlbaranCompraItemProps
    Dim udtAlbaranData As AlbaranCompraItemData
    
    Set rsItem = New Recordset
    strSQL = "SELECT * FROM FacturaCompraItems WHERE FacturaCompraItemID=" & _
        FacturaCompraItemID
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    
    ' Desactualizamos por la parte del Albaran si lo hay.
    If mudtProps.AlbaranCompraItemID Then
       Set objAlbaranPersist = mobjContext.CreateInstance("GCServerMTS.AlbaranCompraItemPersist")
       udtAlbaranData.Buffer = objAlbaranPersist.Fetch(mudtProps.AlbaranCompraItemID)
       LSet udtAlbaran = udtAlbaranData
       With udtAlbaran
            .Situacion = "A"    ' Situacion <- No facturado
            .IsDeleted = False
            .IsDirty = True
            .IsNew = False
            .ActualizarFactura = False
            .ActualizarAlta = False
            .DesactualizarFactura = True
            .DesactualizarAlta = False
       End With
       LSet udtAlbaranData = udtAlbaran
       udtAlbaranData.Buffer = objAlbaranPersist.Save(udtAlbaranData.Buffer)
       Set objAlbaranPersist = Nothing
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
