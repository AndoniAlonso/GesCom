VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "MaterialPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
  
Private mudtProps As MaterialProps

Private Function GetState() As String
    Dim udtData As MaterialData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As MaterialData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByRef Buffer As String) As String
    Dim rsMaterial As Recordset
    Dim strSQL As String
  
    On Error GoTo errh
    SetState Buffer
  
    Set rsMaterial = New Recordset
    strSQL = "SELECT * FROM Materiales WHERE MaterialID=" & _
        mudtProps.MaterialID
    rsMaterial.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsMaterial.AddNew
  
    With rsMaterial
        .Fields("Nombre") = mudtProps.Nombre
        .Fields("codigo") = mudtProps.Codigo
        .Fields("UnidadMedida") = mudtProps.UnidadMedida
        .Fields("StockActual") = mudtProps.StockActual
        .Fields("StockPendiente") = mudtProps.StockPendiente
        .Fields("StockMinimo") = mudtProps.StockMinimo
        .Fields("StockMaximo") = mudtProps.StockMaximo
        .Fields("PrecioCosteEUR") = mudtProps.PrecioCosteEUR
        .Fields("PrecioCostePTA") = mudtProps.PrecioCostePTA
        .Fields("PrecioPonderadoEUR") = mudtProps.PrecioPonderadoEUR
        .Fields("PrecioPonderadoPTA") = mudtProps.PrecioPonderadoPTA
        .Fields("AnchuraTela") = mudtProps.AnchuraTela
        .Fields("TipoMaterial") = mudtProps.TipoMaterial
        .Fields("FechaAlta") = mudtProps.FechaAlta
        .Fields("Composicion1") = mudtProps.Composicion1
        .Fields("PorcComposicion1") = mudtProps.PorcComposicion1
        .Fields("Composicion2") = mudtProps.Composicion2
        .Fields("PorcComposicion2") = mudtProps.PorcComposicion2
        .Fields("Composicion3") = mudtProps.Composicion3
        .Fields("PorcComposicion3") = mudtProps.PorcComposicion3
        .Fields("Composicion4") = mudtProps.Composicion4
        .Fields("PorcComposicion4") = mudtProps.PorcComposicion4
        .Update

        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.MaterialID = .Fields("MaterialID")
        End If
        .Close
    End With
    Set rsMaterial = Nothing
    Save = GetState
    mobjContext.SetComplete
    
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal MaterialID As Long)
    Dim cnGescom As Connection

    On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM MoviMateriales WHERE MaterialID=" & _
        MaterialID

    cnGescom.Execute "DELETE FROM Materiales WHERE MaterialID=" & _
        MaterialID

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

Public Function Fetch(ByVal MaterialID As Long) As String
    Dim rsMaterial As Recordset
    Dim strSQL As String
  
    On Error GoTo errh
    strSQL = "SELECT * FROM Materiales WHERE MaterialID=" & _
        MaterialID
    Set rsMaterial = New Recordset
    rsMaterial.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    With rsMaterial
        mudtProps.MaterialID = .Fields("MaterialID")
        mudtProps.Codigo = .Fields("Codigo") & vbNullString
        mudtProps.Nombre = .Fields("Nombre") & vbNullString
        mudtProps.UnidadMedida = .Fields("UnidadMedida") & vbNullString
        mudtProps.StockActual = .Fields("StockActual")
        mudtProps.StockPendiente = .Fields("StockPendiente")
        mudtProps.StockMinimo = .Fields("StockMinimo")
        mudtProps.StockMaximo = .Fields("StockMaximo")
        mudtProps.PrecioCosteEUR = .Fields("PrecioCosteEUR")
        mudtProps.PrecioCostePTA = .Fields("PrecioCostePTA")
        mudtProps.PrecioPonderadoEUR = .Fields("PrecioPonderadoEUR")
        mudtProps.PrecioPonderadoPTA = .Fields("PrecioPonderadoPTA")
        mudtProps.AnchuraTela = .Fields("AnchuraTela")
        mudtProps.TipoMaterial = .Fields("TipoMaterial")
        mudtProps.FechaAlta = .Fields("FechaAlta")
        mudtProps.Composicion1 = .Fields("Composicion1") & vbNullString
        mudtProps.PorcComposicion1 = .Fields("PorcComposicion1")
        mudtProps.Composicion2 = .Fields("Composicion2") & vbNullString
        mudtProps.PorcComposicion2 = .Fields("PorcComposicion2")
        mudtProps.Composicion3 = .Fields("Composicion3") & vbNullString
        mudtProps.PorcComposicion3 = .Fields("PorcComposicion3")
        mudtProps.Composicion4 = .Fields("Composicion4") & vbNullString
        mudtProps.PorcComposicion4 = .Fields("PorcComposicion4")
        
        .Close
    End With

    Set rsMaterial = Nothing
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

