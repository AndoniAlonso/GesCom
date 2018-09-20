VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "TraspasoPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext

Private mudtProps As TraspasoProps

Private Function GetState() As String
    Dim udtData As TraspasoData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As TraspasoData

    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String, ByRef ItemsBuffer As String) As String
    Dim rsDocumento As Recordset
    Dim strSQL As String
    Dim objPersist As TraspasoItemsPersist
    
    On Error GoTo errh
    
    SetState Buffer
  

    Set rsDocumento = New Recordset
    strSQL = "SELECT * FROM Traspasos WHERE TraspasoID=" & _
        mudtProps.TraspasoID
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsDocumento.AddNew
    
    With rsDocumento
        .Fields("AlmacenOrigenID") = mudtProps.AlmacenOrigenID
        .Fields("AlmacenDestinoID") = mudtProps.AlmacenDestinoID
        .Fields("Situacion") = mudtProps.Situacion
        .Fields("FechaAlta") = mudtProps.FechaAlta
        .Fields("FechaTransito") = IIf(mudtProps.FechaTransito = 0, Null, mudtProps.FechaTransito)
        .Fields("FechaRecepcion") = IIf(mudtProps.FechaRecepcion = 0, Null, mudtProps.FechaRecepcion)
        
        .Update

        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.TraspasoID = .Fields("TraspasoID")
        End If
        .Close
    End With

    Set rsDocumento = Nothing
    Save = GetState
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.TraspasoItemsPersist")
    ItemsBuffer = objPersist.Save(mudtProps.TraspasoID, ItemsBuffer)
    Set objPersist = Nothing
    
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

Public Sub DeleteObject(ByVal TraspasoID As Long, ByVal ItemsBuffer As String)
    Dim cnGescom As Connection
    Dim objPersist As TraspasoItemsPersist
    Dim strSave As String

    On Error GoTo errh
    
    ' Provocamos el borrado de los items de pedido
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.TraspasoItemsPersist")
    strSave = objPersist.Save(TraspasoID, ItemsBuffer)
    Set objPersist = Nothing
    
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"
  
    cnGescom.Execute "DELETE FROM Traspasos WHERE TraspasoID=" & _
        TraspasoID

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

Public Function Fetch(ByVal TraspasoID As Long) As String
    Dim rsDocumento As Recordset
    Dim strSQL As String

    On Error GoTo errh
    strSQL = "SELECT * FROM Traspasos WHERE TraspasoID=" & _
        TraspasoID
    
    Set rsDocumento = New Recordset
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsDocumento
        mudtProps.TraspasoID = .Fields("TraspasoID")
        mudtProps.AlmacenOrigenID = .Fields("AlmacenOrigenID")
        mudtProps.AlmacenDestinoID = .Fields("AlmacenDestinoID")
        mudtProps.Situacion = .Fields("Situacion")
        mudtProps.FechaAlta = .Fields("FechaAlta")
        mudtProps.FechaTransito = IIf(IsNull(.Fields("FechaTransito")), CDate("00:00:00"), .Fields("FechaTransito"))
        mudtProps.FechaRecepcion = IIf(IsNull(.Fields("FechaRecepcion")), CDate("00:00:00"), .Fields("FechaRecepcion"))
    
        .Close
    End With

    Set rsDocumento = Nothing
    
    Fetch = GetState
    mobjContext.SetComplete
    
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With


End Function

Public Function Enviar(ByVal Buffer As String, ByVal ItemsBuffer As String, ByVal AlmacenTransitoID As Long) As String
    Dim rsDocumento As Recordset
    Dim strSQL As String
    Dim objPersist As TraspasoItemsPersist
    
    On Error GoTo errh
    
    SetState Buffer
  

    Set rsDocumento = New Recordset
    strSQL = "SELECT * FROM Traspasos WHERE TraspasoID=" & _
        mudtProps.TraspasoID
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    
    With rsDocumento
        .Fields("Situacion") = mudtProps.Situacion
        .Fields("FechaTransito") = mudtProps.FechaTransito
        
        .Update

        .Close
    End With

    Set rsDocumento = Nothing
    Enviar = GetState
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.TraspasoItemsPersist")
    objPersist.Enviar AlmacenTransitoID, mudtProps.AlmacenOrigenID, ItemsBuffer
    Set objPersist = Nothing
    
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function


Public Function Recepcionar(ByVal Buffer As String, ByVal ItemsBuffer As String, ByVal AlmacenTransitoID As Long) As String
    Dim rsDocumento As Recordset
    Dim strSQL As String
    Dim objPersist As TraspasoItemsPersist
    
    On Error GoTo errh
    
    SetState Buffer
  

    Set rsDocumento = New Recordset
    strSQL = "SELECT * FROM Traspasos WHERE TraspasoID=" & _
        mudtProps.TraspasoID
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    
    With rsDocumento
        .Fields("Situacion") = mudtProps.Situacion
        .Fields("FechaRecepcion") = mudtProps.FechaRecepcion
        
        .Update

        .Close
    End With

    Set rsDocumento = Nothing
    Recepcionar = GetState
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.TraspasoItemsPersist")
    objPersist.Recepcionar AlmacenTransitoID, mudtProps.AlmacenDestinoID, ItemsBuffer
    Set objPersist = Nothing
    
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

Public Function EnviarYRecepcionar(ByVal Buffer As String, ByVal ItemsBuffer As String) As String
    Dim rsDocumento As Recordset
    Dim strSQL As String
    Dim objPersist As TraspasoItemsPersist
    
    On Error GoTo errh
    
    SetState Buffer
  

    Set rsDocumento = New Recordset
    strSQL = "SELECT * FROM Traspasos WHERE TraspasoID=" & _
        mudtProps.TraspasoID
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    
    With rsDocumento
        .Fields("Situacion") = mudtProps.Situacion
        .Fields("FechaRecepcion") = mudtProps.FechaRecepcion
        .Fields("FechaTransito") = mudtProps.FechaTransito
        
        .Update

        .Close
    End With

    Set rsDocumento = Nothing
    EnviarYRecepcionar = GetState
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.TraspasoItemsPersist")
    objPersist.EnviarYRecepcionar mudtProps.AlmacenOrigenID, mudtProps.AlmacenDestinoID, ItemsBuffer
    Set objPersist = Nothing
    
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

