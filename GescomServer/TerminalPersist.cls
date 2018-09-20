VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "TerminalPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
 
Private mudtProps As TerminalProps

Private Function GetState() As String
    Dim udtData As TerminalData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As TerminalData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String
    Dim rsTerminal As Recordset
    Dim strSQL As String
  
    On Error GoTo errh
    SetState Buffer
  
    Set rsTerminal = New Recordset
    strSQL = "SELECT * FROM Terminales WHERE TerminalID=" & _
        mudtProps.TerminalID
    rsTerminal.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsTerminal.AddNew
  
    With rsTerminal
        .Fields("Nombre") = mudtProps.Nombre
        .Fields("CentroGestionID") = mudtProps.CentroGestionID
        .Update

        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.TerminalID = .Fields("TerminalID")
        End If
        .Close
    End With

    Set rsTerminal = Nothing
    Save = GetState
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal TerminalID As Long)
    Dim cnGescom As Connection


    On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM Terminales WHERE TerminalID=" & _
        TerminalID

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

Public Function Fetch(ByVal TerminalID As Long) As String
    Dim rsTerminal As Recordset
    Dim strSQL As String
  
    On Error GoTo errh
    
    strSQL = "SELECT * FROM vTerminales WHERE TerminalID=" & _
        TerminalID
    
    Set rsTerminal = New Recordset
    rsTerminal.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsTerminal
        mudtProps.TerminalID = .Fields("TerminalID")
        mudtProps.Nombre = .Fields("Nombre") & vbNullString
        mudtProps.CentroGestionID = .Fields("CentroGestionID")
        mudtProps.AlmacenID = .Fields("AlmacenID")
        
        .Close
    End With

    Set rsTerminal = Nothing
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

