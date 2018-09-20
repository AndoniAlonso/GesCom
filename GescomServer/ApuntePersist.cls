VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ApuntePersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
  
Public Function Fetch(ByVal AsientoID As Long) As Byte()
    Dim rsApunte As Recordset
    Dim strSQL As String
    Dim udtProps As ApunteProps
    Dim udtData As ApunteData
    Dim objPB As PropertyBag
    Dim lngCount As Long
  

    On Error GoTo errh
    Set rsApunte = New Recordset
    strSQL = "SELECT * FROM Apuntes WHERE AsientoID=" & CStr(AsientoID)
    rsApunte.Open strSQL, DB_CONNECT, adOpenForwardOnly, adLockReadOnly

    Set objPB = New PropertyBag
    
    Do While Not rsApunte.EOF
        With rsApunte
            udtProps.ApunteID = .Fields("ApunteID")
            udtProps.AsientoID = .Fields("AsientoID")
            udtProps.Cuenta = .Fields("Cuenta") & vbNullString
            udtProps.TipoImporte = .Fields("TipoImporte") & vbNullString
            udtProps.ImportePTA = .Fields("ImportePTA")
            udtProps.ImporteEUR = .Fields("ImporteEUR")
            udtProps.Descripcion = .Fields("Descripcion") & vbNullString
            udtProps.Fecha = .Fields("Fecha")
            udtProps.Documento = .Fields("Documento") & vbNullString
            udtProps.TipoApunte = .Fields("TipoApunte") & vbNullString
            udtProps.IsNew = False
            udtProps.IsDirty = False
            udtProps.IsDeleted = False
            
        End With
        
        LSet udtData = udtProps

        ' put task UDT into propertybag
        lngCount = lngCount + 1
        objPB.WriteProperty "Item" & CStr(lngCount), udtData.Buffer
        rsApunte.MoveNext
    Loop
    
    rsApunte.Close
    Set rsApunte = Nothing
    Fetch = GetState
    mobjContext.SetComplete

    ' Almacenar el total de registros en el property bag y devolver su contenido
    ' como resultado de la funcion
    With objPB
        .WriteProperty "Count", lngCount
        Fetch = .Contents
    End With
    Set objPB = Nothing
    
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Function Save(Buffer() As Byte, ByVal AsientoID As Long) As Byte()
    Dim rsApunte As Recordset
    Dim strSQL As String
    Dim udtProps As ApunteProps
    Dim udtData As ApunteData
    Dim objPB As PropertyBag
    Dim objPBOut As PropertyBag
    Dim lngIndex As Long
    Dim lngCount As Long
  

    On Error GoTo errh
    Set objPB = New PropertyBag
    Set objPBOut = New PropertyBag
    
    ' Restaurar el contenido del PB
    objPB.Contents = Buffer
    Set rsApunte = New Recordset
    
   
    For lngIndex = 1 To objPB.ReadProperty("Count")
        ' Ir obteniendo los datos del PB
        udtData.Buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData
    
        ' salvar o borrar el registro
        If Not udtProps.IsDeleted Then
            strSQL = "SELECT * FROM Apuntes WHERE ApunteID=" & CStr(udtProps.ApunteID)
            rsApunte.Open strSQL, DB_CONNECT, adOpenKeyset, adLockOptimistic
            If udtProps.IsNew Then rsApunte.AddNew
            
            With udtProps
                rsApunte.Fields("AsientoID") = .AsientoID
                rsApunte.Fields("Cuenta") = .Cuenta
                rsApunte.Fields("TipoImporte") = .TipoImporte
                rsApunte.Fields("ImportePTA") = .ImportePTA
                rsApunte.Fields("ImporteEUR") = .ImporteEUR
                rsApunte.Fields("Descripcion") = .Descripcion
                rsApunte.Fields("Fecha") = .Fecha
                rsApunte.Fields("Documento") = .Documento
                rsApunte.Fields("TipoApunte") = .TipoApunte
            
                rsApunte.Update
            
                If .IsNew Then
                    rsApunte.Bookmark = rsApunte.Bookmark
                    .ApunteID = rsApunte.Fields("ApunteID")
                End If
                .IsNew = False
                .IsDirty = False
            End With
            LSet udtData = udtProps
            
            ' Escribir los datos resultantes en el PB de vuelta
            lngCount = lngCount + 1
            objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.Buffer
            rsApunte.Close
        Else
            DeleteApunte udtProps.ApunteID
        End If
    Next
    
    ' Escribir el total de registros de salida en el PB de vuelta
    objPBOut.WriteProperty "Count", lngCount
    
    Set objPB = Nothing
    Set rsApunte = Nothing
    
    ' devolver el contenido del PB saliente como resultado de la funcion
    Save = objPBOut.Contents
    Set objPBOut = Nothing
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Private Sub DeleteApunte(ByVal ApunteID As Long)
    Dim cnClient As Connection
    Dim strSQL As String
    
    ' borrar un registro simple
    strSQL = "DELETE FROM Apuntes WHERE ApunteID=" & CStr(ApunteID)
    Set cnClient = New Connection
    cnClient.Open DB_CONNECT
    cnClient.Execute strSQL
    cnClient.Close
    Set cnClient = Nothing
    mobjContext.SetComplete
    Exit Sub

errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Private Sub DeleteObject(ByVal AsientoID As Long)
    Dim cnClient As Connection
    Dim strSQL As String
    
    ' borrar todos los registros de un asiento
    strSQL = "DELETE FROM Apuntes WHERE AsientoID=" & CStr(AsientoID)
    Set cnClient = New Connection
    On Error GoTo errh
    cnClient.Open DB_CONNECT
    cnClient.Execute strSQL
    cnClient.Close
    Set cnClient = Nothing
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

'Private Sub SetState(Buffer As String)
'    Dim udtData As ApunteData
'
'    udtData.Buffer = Buffer
'    LSet mudtProps = udtData
'
'End Sub
'
'Private Function GetState() As String
'    Dim udtData As ApunteData
'
'    LSet udtData = mudtProps
'    GetState = udtData.Buffer
'
'End Function
'

