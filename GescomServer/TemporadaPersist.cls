VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "TemporadaPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
  Option Explicit
  
    Implements ObjectControl

    Private mobjContext As ObjectContext
  
  Private mudtProps As TemporadaProps

Private Function GetState() As String

    Dim udtData As TemporadaData
    
    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

    Dim udtData As TemporadaData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String

    Dim rsTemporada As Recordset
    Dim strSQL As String
  

  On Error GoTo errh
    SetState Buffer
  
    Set rsTemporada = New Recordset
    strSQL = "SELECT * FROM Temporadas WHERE TemporadaID=" & _
        mudtProps.TemporadaID
    rsTemporada.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsTemporada.AddNew
  
    With rsTemporada
        .Fields("Nombre") = mudtProps.Nombre
        .Fields("codigo") = mudtProps.Codigo
        .Update

        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.TemporadaID = .Fields("TemporadaID")
        End If
        .Close
    End With

    Set rsTemporada = Nothing
    Save = GetState
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal TemporadaID As Long)
    Dim cnGescom As Connection


  On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    'cnGescom.Execute "DELETE FROM Temporadas WHERE TemporadaID=" & _
    '    TemporadaID
    cnGescom.Execute "spTemporadaDelete " & TemporadaID

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

Public Function Fetch(ByVal TemporadaID As Long) As String

    Dim strSQL As String

    Dim ThisConnection As adodb.Connection
    Dim ThisRecordset As adodb.Recordset
    
  On Error GoTo errh
    Set ThisConnection = New adodb.Connection
    Set ThisRecordset = New adodb.Recordset
    
    ThisConnection.Open "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    strSQL = "spTemporadaGet " & TemporadaID
    
    ThisRecordset.Open strSQL, ThisConnection

    With ThisRecordset
        mudtProps.TemporadaID = .Fields("TemporadaID")
        mudtProps.Codigo = .Fields("codigo") & vbNullString
        mudtProps.Nombre = .Fields("Nombre") & vbNullString
    End With

    ThisConnection.Close
    
    Set ThisConnection = Nothing
    
    Fetch = GetState
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

'Public Function Fetch(ByVal TemporadaID As Long) As String

'    Dim rsTemporada As Recordset
'    Dim strSQL As String
  

'  On Error GoTo errh
'    strSQL = "SELECT * FROM Temporadas WHERE TemporadaID=" & _
'        TemporadaID
'    Set rsTemporada = New Recordset
'    rsTemporada.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

'    With rsTemporada
'        mudtProps.TemporadaID = .Fields("TemporadaID")
'        mudtProps.Codigo = .Fields("codigo") & vbNullString
'        mudtProps.Nombre = .Fields("Nombre") & vbNullString
'        .Close
'    End With

'    Set rsTemporada = Nothing
'    Fetch = GetState
'    mobjContext.SetComplete

'    Exit Function
  
'errh:
'    mobjContext.SetAbort
'    With Err
'        .Raise .Number, .Source, .Description
'    End With
'End Function

Private Function ObjectControl_CanBePooled() As Boolean

  ObjectControl_CanBePooled = False
  
End Function

Private Sub ObjectControl_Activate()

  Set mobjContext = GetObjectContext

End Sub

Private Sub ObjectControl_Deactivate()

  Set mobjContext = Nothing

End Sub


