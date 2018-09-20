VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "ArticuloColorAlmacenPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
  
Private mudtProps As ArticuloColorAlmacenProps

Private Function GetState() As String
    Dim udtData As ArticuloColorAlmacenData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As ArticuloColorAlmacenData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByRef Buffer As String) As String
    Dim rsArticuloColor As Recordset
    Dim strSQL As String
  
    On Error GoTo errh
    SetState Buffer
  
    Set rsArticuloColor = New Recordset
    strSQL = "SELECT * FROM ArticuloColorAlmacen WHERE ArticuloColorID=" & _
        mudtProps.ArticuloColorID & " AND AlmacenID = " & mudtProps.AlmacenID
    rsArticuloColor.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsArticuloColor.AddNew
  
    With rsArticuloColor
        .Fields("ArticuloColorID") = mudtProps.ArticuloColorID
        .Fields("AlmacenID") = IIf(mudtProps.AlmacenID <> 0, mudtProps.AlmacenID, Null)
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
            mudtProps.AlmacenID = .Fields("AlmacenID")
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

Public Sub DeleteObject(ByVal ArticuloColorID As Long, ByVal AlmacenID As Long)
    Dim cnGescom As Connection

    On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM MoviArticulos WHERE ArticuloColorID=" & _
        ArticuloColorID & " AND AlmacenID=" & AlmacenID

    cnGescom.Execute "DELETE FROM ArticuloColorAlmacen WHERE ArticuloColorID=" & _
        ArticuloColorID & " AND AlmacenID=" & AlmacenID

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

Public Function Fetch(ByVal ArticuloColorID As Long, ByVal AlmacenID As Long) As String
    Dim rsArticuloColor As Recordset
    Dim strSQL As String
  
    On Error GoTo errh
    
    strSQL = "SELECT * FROM ArticuloColorAlmacen WHERE ArticuloColorID=" & _
           ArticuloColorID & " AND AlmacenID=" & AlmacenID
    Set rsArticuloColor = New Recordset
    rsArticuloColor.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsArticuloColor
        mudtProps.ArticuloColorID = .Fields("ArticuloColorID")
        mudtProps.AlmacenID = IIf(IsNull(.Fields("AlmacenID")), 0, .Fields("AlmacenID"))
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
        mudtProps.IsNew = False
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

Public Function FetchCreateIfNotExists(ByVal ArticuloColorID As Long, ByVal AlmacenID As Long) As String
    Dim rsArticuloColor As Recordset
    Dim strSQL As String
    Dim udtArticuloColorAlmacenData As ArticuloColorAlmacenData
    Dim udtArticuloColorAlmacen As ArticuloColorAlmacenProps
  
    On Error GoTo errh
    
    strSQL = "SELECT * FROM ArticuloColorAlmacen WHERE ArticuloColorID=" & _
           ArticuloColorID & " AND AlmacenID=" & AlmacenID
    Set rsArticuloColor = New Recordset
    rsArticuloColor.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset ' PONGO EL adOpenKeyset para que funcione el recordcount
    
    If rsArticuloColor.RecordCount > 0 Then
        FetchCreateIfNotExists = Fetch(ArticuloColorID, AlmacenID)
    Else
        
        With udtArticuloColorAlmacen
            .AlmacenID = AlmacenID
            .ArticuloColorID = ArticuloColorID
            .IsDeleted = False
            .IsDirty = True
            .IsNew = True
            .StockActualT36 = 0
            .StockActualT38 = 0
            .StockActualT40 = 0
            .StockActualT42 = 0
            .StockActualT44 = 0
            .StockActualT46 = 0
            .StockActualT48 = 0
            .StockActualT50 = 0
            .StockActualT52 = 0
            .StockActualT54 = 0
            .StockActualT56 = 0
            .StockPendienteT36 = 0
            .StockPendienteT38 = 0
            .StockPendienteT40 = 0
            .StockPendienteT42 = 0
            .StockPendienteT44 = 0
            .StockPendienteT46 = 0
            .StockPendienteT48 = 0
            .StockPendienteT50 = 0
            .StockPendienteT52 = 0
            .StockPendienteT54 = 0
            .StockPendienteT56 = 0
        End With
        LSet udtArticuloColorAlmacenData = udtArticuloColorAlmacen
        FetchCreateIfNotExists = Me.Save(udtArticuloColorAlmacenData.Buffer)
    End If
    
    Set rsArticuloColor = Nothing
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
