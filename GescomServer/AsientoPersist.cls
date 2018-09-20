VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "AsientoPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
  
 Public Function Save(Buffer() As Byte) As Byte()
    Dim rsAsiento As Recordset
    Dim strSQL As String
    Dim udtProps As AsientoProps
    Dim udtData As AsientoData
    Dim objPB As PropertyBag
    Dim objPBOut As PropertyBag
    Dim objPersist As ApuntesPersist
    Dim objIVAPersist As IVAApuntesPersist
    

    On Error GoTo errh
    Set objPB = New PropertyBag
    objPB.Contents = Buffer
    udtData.Buffer = objPB.ReadProperty("State")
    LSet udtProps = udtData
    
    Set objPBOut = New PropertyBag
    Set rsAsiento = New Recordset
    
    If Not udtProps.IsDeleted Then
        strSQL = "SELECT * FROM Asientos WHERE AsientoID=" & CStr(udtProps.AsientoID)
        rsAsiento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
        If udtProps.IsNew Then rsAsiento.AddNew

        With rsAsiento
            .Fields("Numero") = udtProps.Numero
            .Fields("Ejercicio") = udtProps.Ejercicio
            .Fields("Concepto") = udtProps.Concepto
            .Fields("FechaAlta") = udtProps.FechaAlta
            .Fields("TemporadaID") = udtProps.TemporadaID
            .Fields("EmpresaID") = udtProps.EmpresaID
            .Fields("Situacion") = udtProps.Situacion
    
            .Update
    
            If udtProps.IsNew Then
                .Bookmark = .Bookmark
                udtProps.AsientoID = .Fields("AsientoID")
            End If
            udtProps.IsNew = False
            udtProps.IsDirty = False
        End With
        
        
        LSet udtData = udtProps
        objPBOut.WriteProperty "State", udtData.Buffer
        
        Set objPersist = mobjContext.CreateInstance("GCServerMTS.ApuntesPersist")
        objPBOut.WriteProperty "Items", objPersist.Save(objPB.ReadProperty("Items"), udtProps.AsientoID)
        Set objPersist = Nothing
        
        Set objIVAPersist = mobjContext.CreateInstance("GCServerMTS.IVAApuntesPersist")
        objPBOut.WriteProperty "IVAItems", objIVAPersist.Save(objPB.ReadProperty("IVAItems"), udtProps.AsientoID)
        Set objPersist = Nothing
        
        rsAsiento.Close
    Else
        DeleteAsiento udtProps.AsientoID
    End If
    
    Set objPB = Nothing
    Set rsAsiento = Nothing
    
    Save = objPBOut.Contents
    Set objPBOut = Nothing
    
    mobjContext.SetComplete
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Private Sub DeleteAsiento(ByVal AsientoID As Long)
    Dim cnGescom As Connection
    Dim objPersist As ApuntesPersist
    Dim objIVAPersist As IVAApuntesPersist

    ' Primero borramos los registros de detalle de estructura de Asiento
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.ApuntesPersist")
    objPersist.DeleteObject AsientoID
    Set objPersist = Nothing
    
    ' También borramos los registros de detalle de IVA de Asiento
    Set objIVAPersist = mobjContext.CreateInstance("GCServerMTS.IVAApuntesPersist")
    objIVAPersist.DeleteObject AsientoID
    Set objIVAPersist = Nothing
    
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"
    cnGescom.Execute "DELETE FROM Asientos WHERE AsientoID=" & CStr(AsientoID)
    cnGescom.Close
    Set cnGescom = Nothing
    
End Sub

'Public Sub DeleteObject(ByVal AsientoID As Long, ByVal ItemsBuffer As String)
'    Dim cnGescom As Connection
'    Dim objPersist As ApuntesPersist
'    Dim strSave As String
'
'    On Error GoTo errh
'
'    ' Primero borramos los registros de detalle de estructura de Asiento
'    Set objPersist = mobjContext.CreateInstance("GCServerMTS.ApuntesPersist")
'    strSave = objPersist.Save(AsientoID, ItemsBuffer)
'    Set objPersist = Nothing
'
'    Set cnGescom = New Connection
'    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"
'
'    cnGescom.Execute "DELETE FROM Asientos WHERE AsientoID=" & _
'        AsientoID
'
'    cnGescom.Close
'    Set cnGescom = Nothing
'
'    mobjContext.SetComplete
'
'    Exit Sub
'
'errh:
'    mobjContext.SetAbort
'    With Err
'        .Raise .Number, .Source, .Description
'    End With
'End Sub
'
Public Function Fetch(ByVal AsientoID As Long) As Byte()
    Dim rsAsiento As Recordset
    Dim strSQL As String
    Dim udtProps As AsientoProps
    Dim udtData As AsientoData
    Dim objPB As PropertyBag
    Dim objAsientoPB As PropertyBag
    Dim lngCount As Long
    Dim objPersist As ApuntesPersist
    Dim objIVAPersist As IVAApuntesPersist

    On Error GoTo errh
    
    Set rsAsiento = New Recordset
    strSQL = "SELECT * FROM Asientos WHERE AsientoID=" & CStr(AsientoID)
    rsAsiento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenForwardOnly, adLockReadOnly
    Set objPB = New PropertyBag
    Do While Not rsAsiento.EOF
        With rsAsiento
            udtProps.AsientoID = .Fields("AsientoID")
            udtProps.Numero = .Fields("Numero")
            udtProps.Ejercicio = .Fields("Ejercicio") & vbNullString
            udtProps.Concepto = .Fields("Concepto") & vbNullString
            udtProps.FechaAlta = .Fields("FechaAlta")
            udtProps.TemporadaID = .Fields("TemporadaID")
            udtProps.EmpresaID = .Fields("EmpresaID")
            udtProps.Situacion = .Fields("Situacion") & vbNullString
            udtProps.IsNew = False
            udtProps.IsDirty = False
            udtProps.IsDeleted = False
        End With
        
        LSet udtData = udtProps
        Set objAsientoPB = New PropertyBag
        With objPB
            .WriteProperty "state", udtData.Buffer
            Set objPersist = CreateObject("GCServerMTS.Apuntespersist")
            .WriteProperty "Items", objPersist.Fetch(udtProps.AsientoID)
            Set objPersist = Nothing
            
            lngCount = lngCount + 1
            objPB.WriteProperty "Item" & CStr(lngCount), .Contents
        
            Set objIVAPersist = CreateObject("GCServerMTS.IVAApuntespersist")
            .WriteProperty "IVAItems", objIVAPersist.Fetch(udtProps.AsientoID)
            Set objIVAPersist = Nothing
            
            lngCount = lngCount + 1
            objPB.WriteProperty "IVAItem" & CStr(lngCount), .Contents
        End With
        Set objAsientoPB = Nothing
        rsAsiento.MoveNext
    Loop
    
    rsAsiento.Close
    Set rsAsiento = Nothing
    
    objPB.WriteProperty "Count", lngCount
    
    
    Fetch = objPB.Contents
    Set objPB = Nothing
    
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


