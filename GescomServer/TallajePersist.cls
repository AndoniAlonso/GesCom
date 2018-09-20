VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "TallajePersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
 
Private mudtProps As TallajeProps

Private Function GetState() As String
    Dim udtData As TallajeData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As TallajeData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String
    Dim rsTallaje As Recordset
    Dim strSQL As String
  
    On Error GoTo errh
    SetState Buffer
  
    Set rsTallaje = New Recordset
    strSQL = "SELECT * FROM Tallajes WHERE TallajeID=" & _
        mudtProps.TallajeID
    rsTallaje.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsTallaje.AddNew
  
    With rsTallaje
            .Fields("Nombre") = mudtProps.Nombre
            .Fields("Estandar") = mudtProps.Estandar
            .Fields("DescripcionT36") = mudtProps.DescripcionT36
            .Fields("DescripcionT38") = mudtProps.DescripcionT38
            .Fields("DescripcionT40") = mudtProps.DescripcionT40
            .Fields("DescripcionT42") = mudtProps.DescripcionT42
            .Fields("DescripcionT44") = mudtProps.DescripcionT44
            .Fields("DescripcionT46") = mudtProps.DescripcionT46
            .Fields("DescripcionT48") = mudtProps.DescripcionT48
            .Fields("DescripcionT50") = mudtProps.DescripcionT50
            .Fields("DescripcionT52") = mudtProps.DescripcionT52
            .Fields("DescripcionT54") = mudtProps.DescripcionT54
            .Fields("DescripcionT56") = mudtProps.DescripcionT56
            .Fields("PermitidoT36") = mudtProps.PermitidoT36
            .Fields("PermitidoT38") = mudtProps.PermitidoT38
            .Fields("PermitidoT40") = mudtProps.PermitidoT40
            .Fields("PermitidoT42") = mudtProps.PermitidoT42
            .Fields("PermitidoT44") = mudtProps.PermitidoT44
            .Fields("PermitidoT46") = mudtProps.PermitidoT46
            .Fields("PermitidoT48") = mudtProps.PermitidoT48
            .Fields("PermitidoT50") = mudtProps.PermitidoT50
            .Fields("PermitidoT52") = mudtProps.PermitidoT52
            .Fields("PermitidoT54") = mudtProps.PermitidoT54
            .Fields("PermitidoT56") = mudtProps.PermitidoT56

        .Update

        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.TallajeID = .Fields("TallajeID")
        End If
        .Close
    End With

    Set rsTallaje = Nothing
    Save = GetState
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal TallajeID As Long)
    Dim cnGescom As Connection


    On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM Tallajes WHERE TallajeID=" & _
        TallajeID

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

Public Function Fetch(ByVal TallajeID As Long) As String
    Dim rsTallaje As Recordset
    Dim strSQL As String
  
    On Error GoTo errh
    
    strSQL = "SELECT * FROM Tallajes WHERE TallajeID=" & _
        TallajeID
    
    Set rsTallaje = New Recordset
    rsTallaje.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsTallaje
            mudtProps.TallajeID = .Fields("TallajeID")
            mudtProps.Nombre = .Fields("Nombre") & ""
            mudtProps.Estandar = .Fields("Estandar")
            mudtProps.DescripcionT36 = .Fields("DescripcionT36") & ""
            mudtProps.DescripcionT38 = .Fields("DescripcionT38") & ""
            mudtProps.DescripcionT40 = .Fields("DescripcionT40") & ""
            mudtProps.DescripcionT42 = .Fields("DescripcionT42") & ""
            mudtProps.DescripcionT44 = .Fields("DescripcionT44") & ""
            mudtProps.DescripcionT46 = .Fields("DescripcionT46") & ""
            mudtProps.DescripcionT48 = .Fields("DescripcionT48") & ""
            mudtProps.DescripcionT50 = .Fields("DescripcionT50") & ""
            mudtProps.DescripcionT52 = .Fields("DescripcionT52") & ""
            mudtProps.DescripcionT54 = .Fields("DescripcionT54") & ""
            mudtProps.DescripcionT56 = .Fields("DescripcionT56") & ""
            mudtProps.PermitidoT36 = .Fields("PermitidoT36")
            mudtProps.PermitidoT38 = .Fields("PermitidoT38")
            mudtProps.PermitidoT40 = .Fields("PermitidoT40")
            mudtProps.PermitidoT42 = .Fields("PermitidoT42")
            mudtProps.PermitidoT44 = .Fields("PermitidoT44")
            mudtProps.PermitidoT46 = .Fields("PermitidoT46")
            mudtProps.PermitidoT48 = .Fields("PermitidoT48")
            mudtProps.PermitidoT50 = .Fields("PermitidoT50")
            mudtProps.PermitidoT52 = .Fields("PermitidoT52")
            mudtProps.PermitidoT54 = .Fields("PermitidoT54")
            mudtProps.PermitidoT56 = .Fields("PermitidoT56")
        
        .Close
    End With

    Set rsTallaje = Nothing
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

