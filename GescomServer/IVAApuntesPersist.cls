VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "IVAApuntesPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
  
Public Function Fetch(ByVal AsientoID As Long) As Byte()
    Dim rsIVAApunte As Recordset
    Dim strSQL As String
    Dim udtProps As IVAApunteProps
    Dim udtData As IVAApunteData
    Dim objPB As PropertyBag
    Dim lngCount As Long
  

    On Error GoTo errh
    Set rsIVAApunte = New Recordset
    strSQL = "SELECT * FROM IVAApuntes WHERE AsientoID=" & CStr(AsientoID)
    rsIVAApunte.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenForwardOnly, adLockReadOnly

    Set objPB = New PropertyBag
    
    Do While Not rsIVAApunte.EOF
        With rsIVAApunte
            udtProps.IVAApunteID = .Fields("IVAApunteID")
            udtProps.AsientoID = .Fields("AsientoID")
            udtProps.TipoApunte = .Fields("TipoApunte")
            udtProps.NumeroFactura = .Fields("NumeroFactura")
            udtProps.CuentaIVA = .Fields("CuentaIVA")
            udtProps.CuentaTotal = .Fields("CuentaTotal")
            udtProps.CuentaBase = .Fields("CuentaBase")
            udtProps.Titular = .Fields("Titular")
            udtProps.DNINIF = .Fields("DNINIF")
            udtProps.BaseImponible = .Fields("BaseImponible")
            udtProps.Total = .Fields("Total")
            udtProps.IVA = .Fields("IVA")
            udtProps.CuotaIVA = .Fields("CuotaIVA")
            udtProps.RecargoEquivalencia = .Fields("RecargoEquivalencia")
            udtProps.CuotaRecargo = .Fields("CuotaRecargo")
            udtProps.IsNew = False
            udtProps.IsDirty = False
            udtProps.IsDeleted = False
            
        End With
        
        LSet udtData = udtProps

        ' put task UDT into propertybag
        lngCount = lngCount + 1
        objPB.WriteProperty "IVAItem" & CStr(lngCount), udtData.Buffer
        rsIVAApunte.MoveNext
    Loop
    
    rsIVAApunte.Close
    Set rsIVAApunte = Nothing

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
    Dim rsIVAApunte As Recordset
    Dim strSQL As String
    Dim udtProps As IVAApunteProps
    Dim udtData As IVAApunteData
    Dim objPB As PropertyBag
    Dim objPBOut As PropertyBag
    Dim lngIndex As Long
    Dim lngCount As Long
  

    On Error GoTo errh
    Set objPB = New PropertyBag
    Set objPBOut = New PropertyBag
    
    ' Restaurar el contenido del PB
    objPB.Contents = Buffer
    Set rsIVAApunte = New Recordset
    
   
    For lngIndex = 1 To objPB.ReadProperty("Count")
        ' Ir obteniendo los datos del PB
        udtData.Buffer = objPB.ReadProperty("IVAItem" & CStr(lngIndex))
        LSet udtProps = udtData
    
        ' salvar o borrar el registro
        If Not udtProps.IsDeleted Then
            strSQL = "SELECT * FROM IVAApuntes WHERE IVAApunteID=" & CStr(udtProps.IVAApunteID)
            rsIVAApunte.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
            If udtProps.IsNew Then rsIVAApunte.AddNew
            
            With udtProps
                rsIVAApunte.Fields("AsientoID") = AsientoID
                rsIVAApunte.Fields("TipoApunte") = .TipoApunte
                rsIVAApunte.Fields("NumeroFactura") = .NumeroFactura
                rsIVAApunte.Fields("CuentaIVA") = .CuentaIVA
                rsIVAApunte.Fields("CuentaTotal") = .CuentaTotal
                rsIVAApunte.Fields("CuentaBase") = .CuentaBase
                rsIVAApunte.Fields("Titular") = .Titular
                rsIVAApunte.Fields("DNINIF") = .DNINIF
                rsIVAApunte.Fields("BaseImponible") = .BaseImponible
                rsIVAApunte.Fields("Total") = .Total
                rsIVAApunte.Fields("IVA") = .IVA
                rsIVAApunte.Fields("CuotaIVA") = .CuotaIVA
                rsIVAApunte.Fields("RecargoEquivalencia") = .RecargoEquivalencia
                rsIVAApunte.Fields("CuotaRecargo") = .CuotaRecargo
            
                rsIVAApunte.Update
            
                If .IsNew Then
                    rsIVAApunte.Bookmark = rsIVAApunte.Bookmark
                    .IVAApunteID = rsIVAApunte.Fields("IVAApunteID")
                End If
                .IsNew = False
                .IsDirty = False
            End With
            LSet udtData = udtProps
            
            ' Escribir los datos resultantes en el PB de vuelta
            lngCount = lngCount + 1
            objPBOut.WriteProperty "IVAItem" & CStr(lngCount), udtData.Buffer
            rsIVAApunte.Close
        Else
            DeleteIVAApunte udtProps.IVAApunteID
        End If
    Next
    
    ' Escribir el total de registros de salida en el PB de vuelta
    objPBOut.WriteProperty "Count", lngCount
    
    Set objPB = Nothing
    Set rsIVAApunte = Nothing
    
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

Private Sub DeleteIVAApunte(ByVal IVAApunteID As Long)
    Dim cnClient As Connection
    Dim strSQL As String
    
    ' borrar un registro simple
    strSQL = "DELETE FROM IVAApuntes WHERE IVAApunteID=" & CStr(IVAApunteID)
    Set cnClient = New Connection
    cnClient.Open "FILE NAME=" & App.Path & "\Gescom.UDL"
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

Public Sub DeleteObject(ByVal AsientoID As Long)
    Dim cnClient As Connection
    Dim strSQL As String
    
    ' borrar todos los registros de un asiento
    strSQL = "DELETE FROM IVAApuntes WHERE AsientoID=" & CStr(AsientoID)
    Set cnClient = New Connection
    On Error GoTo errh
    cnClient.Open "FILE NAME=" & App.Path & "\Gescom.UDL"
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

