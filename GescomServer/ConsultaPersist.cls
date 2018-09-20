VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "ConsultaPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ObjectControl

Private mobjContext As ObjectContext

Private mudtProps As ConsultaProps

Private Function GetState() As String
    Dim udtData As ConsultaData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As ConsultaData

    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String, ByRef ItemsBuffer As String) As String
    Dim rsDocumento As Recordset
    Dim strSQL As String
    Dim objPersist As ConsultaItemsPersist
    
    On Error GoTo errh
  
    SetState Buffer
      
    Set rsDocumento = New Recordset
    strSQL = "SELECT * FROM Consultas WHERE Nombre = '" & _
        mudtProps.Nombre & "'"
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If rsDocumento.RecordCount = 0 Then
        rsDocumento.AddNew
        mudtProps.IsNew = True
    End If
    'If mudtProps.IsNew Then rsDocumento.AddNew
    
    With rsDocumento
        .Fields("Nombre") = mudtProps.Nombre
        .Fields("Objeto") = mudtProps.Objeto
    
        .Update
    
        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.ConsultaID = .Fields("ConsultaID")
        End If
        .Close
    End With

    Set rsDocumento = Nothing
    Save = GetState
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.ConsultaItemsPersist")
    ItemsBuffer = objPersist.Save(mudtProps.ConsultaID, ItemsBuffer)
    Set objPersist = Nothing
    
    mobjContext.SetComplete
  
  Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

Public Sub DeleteObject(ByVal ConsultaID As Long, ByVal ItemsBuffer As String)
    Dim cnGescom As Connection
    Dim objPersist As ConsultaItemsPersist
    Dim strSave As String

    On Error GoTo errh
    
    ' Provocamos el borrado de los items de consulta
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.ConsultaItemsPersist")
    strSave = objPersist.Save(ConsultaID, ItemsBuffer)
    Set objPersist = Nothing
    
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM Consultas WHERE ConsultaID=" & _
        ConsultaID

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

Public Function Fetch(ByVal ConsultaID As Long) As String
    Dim rsDocumento As Recordset
    Dim strSQL As String
    
    On Error GoTo errh
    strSQL = "SELECT * FROM Consultas WHERE ConsultaID=" & _
        ConsultaID
    Set rsDocumento = New Recordset
    rsDocumento.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    With rsDocumento
        mudtProps.ConsultaID = .Fields("ConsultaID")
        mudtProps.Nombre = .Fields("Nombre") & ""
        mudtProps.Objeto = .Fields("Objeto") & ""
    
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

Private Function ObjectControl_CanBePooled() As Boolean

    ObjectControl_CanBePooled = False
  
End Function

Private Sub ObjectControl_Activate()

    Set mobjContext = GetObjectContext

End Sub

Private Sub ObjectControl_Deactivate()

    Set mobjContext = Nothing

End Sub

