VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "PrendaPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
 
Private mudtProps As PrendaProps

Private Function GetState() As String
    Dim udtData As PrendaData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As PrendaData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String
    Dim rsPrenda As Recordset
    Dim strSQL As String
  
    On Error GoTo errh
    SetState Buffer
  
    Set rsPrenda = New Recordset
    strSQL = "SELECT * FROM Prendas WHERE PrendaID=" & _
        mudtProps.PrendaID
    rsPrenda.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsPrenda.AddNew
  
    With rsPrenda
        .Fields("Nombre") = mudtProps.Nombre
        .Fields("codigo") = mudtProps.Codigo
        .Fields("PlanchaPTA") = mudtProps.PlanchaPTA
        .Fields("PlanchaEUR") = mudtProps.PlanchaEUR
        .Fields("TransportePTA") = mudtProps.TransportePTA
        .Fields("TransporteEUR") = mudtProps.TransporteEUR
        .Fields("PerchaPTA") = mudtProps.PerchaPTA
        .Fields("PerchaEUR") = mudtProps.PerchaEUR
        .Fields("CartonPTA") = mudtProps.CartonPTA
        .Fields("CartonEUR") = mudtProps.CartonEUR
        .Fields("EtiquetaPTA") = mudtProps.EtiquetaPTA
        .Fields("EtiquetaEUR") = mudtProps.EtiquetaEUR
        .Fields("Administracion") = mudtProps.Administracion
        .Update

        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.PrendaID = .Fields("PrendaID")
        End If
        .Close
    End With

    Set rsPrenda = Nothing
    Save = GetState
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal PrendaID As Long)
    Dim cnGescom As Connection


    On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM Prendas WHERE PrendaID=" & _
        PrendaID

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

Public Function Fetch(ByVal PrendaID As Long, Optional Codigo As String) As String
    Dim rsPrenda As Recordset
    Dim strSQL As String
  
    On Error GoTo errh
    
    'Lectura por identificador o por codigo
    If PrendaID <> 0 Then
        strSQL = "SELECT * FROM Prendas WHERE PrendaID=" & _
            PrendaID
    Else
        strSQL = "SELECT * FROM Prendas WHERE Codigo='" & _
            Codigo & "'"
    End If
    Set rsPrenda = New Recordset
    rsPrenda.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsPrenda
        mudtProps.PrendaID = .Fields("PrendaID")
        mudtProps.Codigo = .Fields("codigo") & vbNullString
        mudtProps.Nombre = .Fields("Nombre") & vbNullString
        mudtProps.PlanchaPTA = .Fields("PlanchaPTA")
        mudtProps.PlanchaEUR = .Fields("PlanchaEUR")
        mudtProps.TransportePTA = .Fields("TransportePTA")
        mudtProps.TransporteEUR = .Fields("TransporteEUR")
        mudtProps.PerchaPTA = .Fields("PerchaPTA")
        mudtProps.PerchaEUR = .Fields("PerchaEUR")
        mudtProps.CartonPTA = .Fields("CartonPTA")
        mudtProps.CartonEUR = .Fields("CartonEUR")
        mudtProps.EtiquetaPTA = .Fields("EtiquetaPTA")
        mudtProps.EtiquetaEUR = .Fields("EtiquetaEUR")
        mudtProps.Administracion = .Fields("Administracion")
        
        .Close
    End With

    Set rsPrenda = Nothing
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

