VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "TransportistaPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
  Option Explicit
  
    Implements ObjectControl

    Private mobjContext As ObjectContext
  
    
  Private mudtProps As TransportistaProps

Private Function GetState() As String

    Dim udtData As TransportistaData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

    Dim udtData As TransportistaData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String

    Dim rsTransportista As Recordset
    Dim strSQL As String
    Dim objDireccionPersist As DireccionPersist
    Dim udtDireccionProps As DireccionProps
    

  On Error GoTo errh
    SetState Buffer
  
    Set objDireccionPersist = mobjContext.CreateInstance("GCServerMTS.DireccionPersist")
    mudtProps.Direccion.Buffer = objDireccionPersist.Save(mudtProps.Direccion.Buffer)
    LSet udtDireccionProps = mudtProps.Direccion
    mudtProps.DireccionID = udtDireccionProps.DireccionID
    Set objDireccionPersist = Nothing
    
    Set rsTransportista = New Recordset
    strSQL = "SELECT * FROM Transportistas WHERE TransportistaID=" & _
        mudtProps.TransportistaID
    rsTransportista.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsTransportista.AddNew
  
    With rsTransportista
        .Fields("Nombre") = mudtProps.Nombre
        .Fields("Titular") = mudtProps.Titular
        .Fields("DNINIF") = mudtProps.DNINIF
        .Fields("Contacto") = mudtProps.Contacto
        .Fields("Zona") = mudtProps.Zona
        .Fields("DireccionID") = mudtProps.DireccionID
        .Update
    
        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.TransportistaID = .Fields("TransportistaID")
        End If
        .Close
    End With

    Set rsTransportista = Nothing
    Save = GetState
    mobjContext.SetComplete
 
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal TransportistaID As Long, _
                        ByVal DireccionID As Long)
  
    Dim objDireccionPersist As DireccionPersist
    Dim cnGescom As Connection


  On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"
  
    cnGescom.Execute "DELETE FROM Transportistas WHERE TransportistaID=" & _
        TransportistaID
  
    Set objDireccionPersist = mobjContext.CreateInstance("GCServerMTS.DireccionPersist")
    objDireccionPersist.DeleteObject (DireccionID)
    Set objDireccionPersist = Nothing
  
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

Public Function Fetch(ByVal TransportistaID As Long) As String

    Dim rsTransportista As Recordset
    Dim strSQL As String
    Dim objDireccionPersist As DireccionPersist
  

  On Error GoTo errh
    strSQL = "SELECT * FROM Transportistas WHERE TransportistaID=" & _
        TransportistaID
    Set rsTransportista = New Recordset
    rsTransportista.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsTransportista
        mudtProps.TransportistaID = .Fields("TransportistaID")
        mudtProps.Nombre = .Fields("Nombre") & vbNullString
        mudtProps.Titular = .Fields("Titular") & vbNullString
        mudtProps.DNINIF = .Fields("DNINIF") & vbNullString
        mudtProps.Contacto = .Fields("Contacto") & vbNullString
        mudtProps.Zona = .Fields("Zona") & vbNullString
        mudtProps.DireccionID = .Fields("DireccionID")
        .Close
    End With

    Set rsTransportista = Nothing
    
    Set objDireccionPersist = mobjContext.CreateInstance("GCServerMTS.DireccionPersist")
    mudtProps.Direccion.Buffer = objDireccionPersist.Fetch(mudtProps.DireccionID)
    Set objDireccionPersist = Nothing
    
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


