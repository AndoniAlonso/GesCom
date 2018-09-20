VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "FacturaVentaItemPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext

Private mudtProps As FacturaVentaItemProps

Private Function GetState() As String
    Dim udtData As FacturaVentaItemData
    
    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As FacturaVentaItemData
    
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String
    Dim rsItem As Recordset
    Dim strSQL As String
    
    On Error GoTo errh
    SetState Buffer
    
    Set rsItem = New Recordset
    strSQL = "SELECT * FROM FacturaVentaItems WHERE FacturaVentaItemID=" & _
        mudtProps.FacturaVentaItemID
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsItem.AddNew
    
    With rsItem
        If mudtProps.DesactualizarAlta Then LeerYDesactualizar (mudtProps.FacturaVentaItemID)
        If mudtProps.ActualizarAlta Then Actualizar
         
        .Fields("FacturaVentaID") = mudtProps.FacturaVentaID
        ' Solo asigno el valor si viene informado
        .Fields("ArticuloColorID") = IIf(mudtProps.ArticuloColorID, mudtProps.ArticuloColorID, Null)
        .Fields("Descripcion") = mudtProps.Descripcion
        ' Solo asigno el valor si viene informado
        .Fields("AlbaranVentaItemID") = IIf(mudtProps.AlbaranVentaItemID, mudtProps.AlbaranVentaItemID, Null)
        .Fields("SituacionImpresa") = mudtProps.SituacionImpresa
        .Fields("Cantidad") = mudtProps.Cantidad
        .Fields("PrecioVentaPTA") = mudtProps.PrecioVentaPTA
        .Fields("PrecioVentaEUR") = mudtProps.PrecioVentaEUR
        .Fields("Descuento") = mudtProps.Descuento
        .Fields("BrutoPTA") = mudtProps.BrutoPTA
        .Fields("BrutoEUR") = mudtProps.BrutoEUR
        .Fields("Comision") = mudtProps.Comision
        .Fields("ComisionPTA") = mudtProps.ComisionPTA
        .Fields("ComisionEUR") = mudtProps.ComisionEUR
        .Fields("TemporadaID") = mudtProps.TemporadaID
        
        .Update
        
        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.FacturaVentaItemID = .Fields("FacturaVentaItemID")
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

Public Sub DeleteObject(ByVal FacturaVentaItemID As Long)
    Dim cnGescom As Connection
    
    On Error GoTo errh
    
    If Not mudtProps.IsNew Then
        ' Al borrar una linea de facturas hay que desactualizar SIEMPRE.
        LeerYDesactualizar (FacturaVentaItemID)
    End If
  
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    cnGescom.Execute "DELETE FROM FacturaVentaItems WHERE FacturaVentaItemID=" & _
      FacturaVentaItemID
    
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

Public Function Fetch(ByVal FacturaVentaItemID As Long) As String
    Dim rsItem As Recordset
    Dim strSQL As String
    
    On Error GoTo errh
    strSQL = "SELECT * FROM FacturaVentaItems WHERE FacturaVentaItemID=" & _
      FacturaVentaItemID
    Set rsItem = New Recordset
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    With rsItem
      mudtProps.FacturaVentaItemID = .Fields("FacturaVentaItemID")
      mudtProps.FacturaVentaID = .Fields("FacturaVentaID")
      mudtProps.ArticuloColorID = IIf(IsNull(.Fields("ArticuloColorID")), 0, .Fields("ArticuloColorID"))
      mudtProps.Descripcion = .Fields("Descripcion") & vbNullString
      mudtProps.AlbaranVentaItemID = IIf(IsNull(.Fields("AlbaranVentaItemID")), 0, .Fields("AlbaranVentaItemID"))
      mudtProps.SituacionImpresa = .Fields("SituacionImpresa") & vbNullString
      mudtProps.Cantidad = .Fields("Cantidad")
      mudtProps.PrecioVentaPTA = .Fields("PrecioVentaPTA")
      mudtProps.PrecioVentaEUR = .Fields("PrecioVentaEUR")
      mudtProps.Descuento = .Fields("Descuento")
      mudtProps.BrutoPTA = .Fields("BrutoPTA")
      mudtProps.BrutoEUR = .Fields("BrutoEUR")
      mudtProps.Comision = .Fields("Comision")
      mudtProps.ComisionPTA = .Fields("ComisionPTA")
      mudtProps.ComisionEUR = .Fields("ComisionEUR")
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


Public Function SaveChild(FacturaVentaID As Long, Buffer As String) As String

    On Error GoTo errh
    SetState Buffer
    
    With mudtProps
      If .IsDeleted Then
        DeleteObject .FacturaVentaItemID
        SaveChild = vbNullString
    
      ElseIf .IsDirty Or .IsNew Then
        .FacturaVentaID = FacturaVentaID
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
' - Marcar el albarán relacionado (si lo hay) como facturado.
' Más adelante se controlará el precio medio ponderado.
' también se podrá alimentar datos de facturación por cliente.
Private Sub Actualizar()
    Dim objAlbaranPersist As AlbaranVentaItemPersist
    Dim udtAlbaran As AlbaranVentaItemProps
    Dim udtAlbaranData As AlbaranVentaItemData
    
    ' Actualizamos el stock por la parte del Albaran.
    If mudtProps.AlbaranVentaItemID Then
        Set objAlbaranPersist = mobjContext.CreateInstance("GCServerMTS.AlbaranVentaItemPersist")
        udtAlbaranData.Buffer = objAlbaranPersist.Fetch(mudtProps.AlbaranVentaItemID)
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

Private Sub LeerYDesactualizar(ByVal FacturaVentaItemID As Long)
    Dim rsItem As Recordset
    Dim strSQL As String
    Dim objAlbaranPersist As AlbaranVentaItemPersist
    Dim udtAlbaran As AlbaranVentaItemProps
    Dim udtAlbaranData As AlbaranVentaItemData
    
    Set rsItem = New Recordset
    strSQL = "SELECT * FROM FacturaVentaItems WHERE FacturaVentaItemID=" & _
        FacturaVentaItemID
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    
    ' Desactualizamos por la parte del Albaran si lo hay.
    If mudtProps.AlbaranVentaItemID Then
       Set objAlbaranPersist = mobjContext.CreateInstance("GCServerMTS.AlbaranVentaItemPersist")
       udtAlbaranData.Buffer = objAlbaranPersist.Fetch(mudtProps.AlbaranVentaItemID)
       LSet udtAlbaran = udtAlbaranData
       With udtAlbaran
            '.Situacion = "A"    ' Situacion <- No facturado
            If .FacturaVentaItemIDA = FacturaVentaItemID Then .FacturaVentaItemIDA = 0
            If .FacturaVentaItemIDB = FacturaVentaItemID Then .FacturaVentaItemIDB = 0
            
            If .FacturaVentaItemIDA = 0 And .FacturaVentaItemIDB = 0 Then
                .FacturadoAB = False
                .Situacion = "A"    ' Situacion <- No facturado
            End If
            
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

