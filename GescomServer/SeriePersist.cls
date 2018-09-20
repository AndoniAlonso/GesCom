VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "SeriePersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
  
Private mudtProps As SerieProps

Private Function GetState() As String
    Dim udtData As SerieData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As SerieData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String
    Dim rsSerie As Recordset
    Dim strSQL As String
  

    On Error GoTo errh
    SetState Buffer
  
    Set rsSerie = New Recordset
    strSQL = "SELECT * FROM Series WHERE SerieID=" & _
        mudtProps.SerieID
    rsSerie.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsSerie.AddNew
  
    With rsSerie
        .Fields("Nombre") = mudtProps.Nombre
        .Fields("codigo") = mudtProps.Codigo
        .Fields("MaterialID") = mudtProps.MaterialID
        .Fields("TemporadaID") = mudtProps.TemporadaID
        .Update

        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.SerieID = .Fields("SerieID")
        End If
        .Close
    End With

    Set rsSerie = Nothing
    Save = GetState
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal SerieID As Long)
    Dim cnGescom As Connection

    On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM Series WHERE SerieID=" & _
        SerieID

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

Public Function Fetch(ByVal SerieID As Long, Optional Codigo As String, _
                      Optional TemporadaID As Long, Optional Nombre As String) As String
    Dim rsSerie As Recordset
    Dim strSQL As String
  
    On Error GoTo errh
    
    'Lectura por identificador o por codigo
    If SerieID <> 0 Then
        strSQL = "SELECT * FROM Series WHERE SerieID=" & _
            SerieID
    ElseIf Trim(Codigo) <> vbNullString Then
        strSQL = "SELECT * FROM Series WHERE Codigo='" & _
            Codigo & "' AND TemporadaID = " & TemporadaID
    ElseIf Trim(Nombre) <> vbNullString Then
        strSQL = "SELECT * FROM Series WHERE Nombre='" & _
            Nombre & "' AND TemporadaID = " & TemporadaID
    End If
    Set rsSerie = New Recordset
    rsSerie.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsSerie
        mudtProps.SerieID = .Fields("SerieID")
        mudtProps.Codigo = .Fields("codigo") & vbNullString
        mudtProps.Nombre = .Fields("Nombre") & vbNullString
        mudtProps.MaterialID = .Fields("MaterialID")
        mudtProps.TemporadaID = .Fields("TemporadaID")

        .Close
    End With

    Set rsSerie = Nothing
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

