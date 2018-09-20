VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "EstrModeloPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
    Implements ObjectControl

    Private mobjContext As ObjectContext
  
    Private mudtProps As EstrModeloProps

Private Sub SetState(Buffer As String)
    Dim udtData As EstrModeloData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Private Function GetState() As String
    Dim udtData As EstrModeloData
  
    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Public Function Fetch(ByVal EstrModeloID As Long) As String
    Dim rsEstrModelo As Recordset
    Dim strSQL As String
  

    On Error GoTo errh
    strSQL = "SELECT * FROM EstrModelos WHERE EstrModeloID=" & EstrModeloID
    Set rsEstrModelo = New Recordset
    rsEstrModelo.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsEstrModelo
        mudtProps.EstrModeloID = .Fields("EstrModeloID")
        mudtProps.ModeloID = .Fields("ModeloID")
        mudtProps.Observaciones = .Fields("Observaciones") & vbNullString
        mudtProps.Cantidad = .Fields("Cantidad")
        mudtProps.MaterialID = .Fields("MaterialID")
        mudtProps.PrecioCostePTA = .Fields("PrecioCostePTA")
        mudtProps.PrecioCosteEUR = .Fields("PrecioCosteEUR")
        mudtProps.PrecioEUR = .Fields("PrecioEUR")
        mudtProps.PrecioPTA = .Fields("PrecioPTA")
    
        rsEstrModelo.Close
    End With

    Set rsEstrModelo = Nothing
    Fetch = GetState
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Function Save(ByVal Buffer As String) As String

    Dim rsEstrModelo As Recordset
    Dim strSQL As String
  

    On Error GoTo errh
    SetState Buffer
    
    strSQL = "SELECT * FROM EstrModelos WHERE EstrModeloID=" & mudtProps.EstrModeloID
    Set rsEstrModelo = New Recordset
    rsEstrModelo.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsEstrModelo.AddNew
  
    With rsEstrModelo
        .Fields("ModeloID") = mudtProps.ModeloID
        .Fields("Cantidad") = mudtProps.Cantidad
        .Fields("MaterialID") = mudtProps.MaterialID
        .Fields("Observaciones") = mudtProps.Observaciones
        .Fields("PrecioCosteEUR") = mudtProps.PrecioCosteEUR
        .Fields("PrecioCostePTA") = mudtProps.PrecioCostePTA
        .Fields("PrecioEUR") = mudtProps.PrecioEUR
        .Fields("PrecioPTA") = mudtProps.PrecioPTA
        
        .Update
    
        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.EstrModeloID = .Fields("EstrModeloID")
        End If
        .Close
    End With
    
    Set rsEstrModelo = Nothing
    Save = GetState
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal EstrModeloID As Long)
    Dim cnGescom As Connection


    On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM EstrModelos WHERE EstrModeloID=" & EstrModeloID

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

Public Function SaveChild(ModeloID As Long, Buffer As String) As String

    On Error GoTo errh
    SetState Buffer

    With mudtProps
        If .IsDeleted Then
            DeleteObject .EstrModeloID
            SaveChild = vbNullString

        ElseIf .IsDirty Or .IsNew Then
            .ModeloID = ModeloID
            SaveChild = Save(GetState)

        Else
            SaveChild = Buffer
            mobjContext.SetComplete

        End If

    End With

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


