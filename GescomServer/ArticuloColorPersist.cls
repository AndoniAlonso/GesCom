VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "ArticuloColorPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
  
Private mudtProps As ArticuloColorProps

Private Function GetState() As String
    Dim udtData As ArticuloColorData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As ArticuloColorData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByRef Buffer As String) As String
    Dim rsArticuloColor As Recordset
    Dim strSQL As String
  
    On Error GoTo errh
    SetState Buffer
  
    Set rsArticuloColor = New Recordset
    strSQL = "SELECT * FROM ArticuloColores WHERE ArticuloColorID=" & _
        mudtProps.ArticuloColorID
    rsArticuloColor.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsArticuloColor.AddNew
  
    With rsArticuloColor
        .Fields("Nombre") = mudtProps.Nombre
        .Fields("NombreColor") = mudtProps.NombreColor
        .Fields("ArticuloID") = mudtProps.ArticuloID
        .Fields("TemporadaID") = mudtProps.TemporadaID
        .Fields("StockActualT36") = mudtProps.StockActualT36
        .Fields("StockActualT38") = mudtProps.StockActualT38
        .Fields("StockActualT40") = mudtProps.StockActualT40
        .Fields("StockActualT42") = mudtProps.StockActualT42
        .Fields("StockActualT44") = mudtProps.StockActualT44
        .Fields("StockActualT46") = mudtProps.StockActualT46
        .Fields("StockActualT48") = mudtProps.StockActualT48
        .Fields("StockActualT50") = mudtProps.StockActualT50
        .Fields("StockActualT52") = mudtProps.StockActualT52
        .Fields("StockActualT54") = mudtProps.StockActualT54
        .Fields("StockActualT56") = mudtProps.StockActualT56
        .Fields("StockPendienteT36") = mudtProps.StockPendienteT36
        .Fields("StockPendienteT38") = mudtProps.StockPendienteT38
        .Fields("StockPendienteT40") = mudtProps.StockPendienteT40
        .Fields("StockPendienteT42") = mudtProps.StockPendienteT42
        .Fields("StockPendienteT44") = mudtProps.StockPendienteT44
        .Fields("StockPendienteT46") = mudtProps.StockPendienteT46
        .Fields("StockPendienteT48") = mudtProps.StockPendienteT48
        .Fields("StockPendienteT50") = mudtProps.StockPendienteT50
        .Fields("StockPendienteT52") = mudtProps.StockPendienteT52
        .Fields("StockPendienteT54") = mudtProps.StockPendienteT54
        .Fields("StockPendienteT56") = mudtProps.StockPendienteT56
        .Update

        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.ArticuloColorID = .Fields("ArticuloColorID")
        End If
        .Close
    End With
    
    Set rsArticuloColor = Nothing
    Save = GetState
    mobjContext.SetComplete
    
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal ArticuloColorID As Long)
    Dim cnGescom As Connection

    On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM MoviArticulos WHERE ArticuloColorID=" & _
        ArticuloColorID

    cnGescom.Execute "DELETE FROM ArticuloColores WHERE ArticuloColorID=" & _
        ArticuloColorID

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

Public Function Fetch(ByVal ArticuloColorID As Long, Optional Codigo As String, _
                      Optional TemporadaID As Long) As String
    Dim rsArticuloColor As Recordset
    Dim strSQL As String
  
    On Error GoTo errh
    
    'Lectura por identificador o por codigo
    If ArticuloColorID <> 0 Then
        strSQL = "SELECT * FROM vArticuloColores WHERE ArticuloColorID=" & _
            ArticuloColorID
    Else
        strSQL = "SELECT * FROM vArticuloColores WHERE Nombre='" & _
            Codigo & "' AND TemporadaID = " & TemporadaID
    End If
    Set rsArticuloColor = New Recordset
    rsArticuloColor.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsArticuloColor
        mudtProps.ArticuloColorID = .Fields("ArticuloColorID")
        mudtProps.Nombre = .Fields("Nombre") & vbNullString
        mudtProps.NombreColor = .Fields("NombreColor") & vbNullString
        mudtProps.ArticuloID = .Fields("ArticuloID")
        mudtProps.TemporadaID = .Fields("TemporadaID")
        mudtProps.StockActualT36 = .Fields("StockActualT36")
        mudtProps.StockActualT38 = .Fields("StockActualT38")
        mudtProps.StockActualT40 = .Fields("StockActualT40")
        mudtProps.StockActualT42 = .Fields("StockActualT42")
        mudtProps.StockActualT44 = .Fields("StockActualT44")
        mudtProps.StockActualT46 = .Fields("StockActualT46")
        mudtProps.StockActualT48 = .Fields("StockActualT48")
        mudtProps.StockActualT50 = .Fields("StockActualT50")
        mudtProps.StockActualT52 = .Fields("StockActualT52")
        mudtProps.StockActualT54 = .Fields("StockActualT54")
        mudtProps.StockActualT56 = .Fields("StockActualT56")
        mudtProps.StockPendienteT36 = .Fields("StockPendienteT36")
        mudtProps.StockPendienteT38 = .Fields("StockPendienteT38")
        mudtProps.StockPendienteT40 = .Fields("StockPendienteT40")
        mudtProps.StockPendienteT42 = .Fields("StockPendienteT42")
        mudtProps.StockPendienteT44 = .Fields("StockPendienteT44")
        mudtProps.StockPendienteT46 = .Fields("StockPendienteT46")
        mudtProps.StockPendienteT48 = .Fields("StockPendienteT48")
        mudtProps.StockPendienteT50 = .Fields("StockPendienteT50")
        mudtProps.StockPendienteT52 = .Fields("StockPendienteT52")
        mudtProps.StockPendienteT54 = .Fields("StockPendienteT54")
        mudtProps.StockPendienteT56 = .Fields("StockPendienteT56")
        .Close
    End With

    Set rsArticuloColor = Nothing
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

