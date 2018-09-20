VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "ArticuloPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
  
Private mudtProps As ArticuloProps

Private Function GetState() As String
    Dim udtData As ArticuloData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As ArticuloData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByRef Buffer As String) As String
    Dim rsArticulo As Recordset
    Dim strSQL As String
  
    On Error GoTo errh
    SetState Buffer
  
    Set rsArticulo = New Recordset
    strSQL = "SELECT * FROM Articulos WHERE ArticuloID=" & _
        mudtProps.ArticuloID
    rsArticulo.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsArticulo.AddNew
  
    With rsArticulo
        .Fields("Nombre") = mudtProps.Nombre
        .Fields("StockActual") = mudtProps.StockActual
        .Fields("StockPendiente") = mudtProps.StockPendiente
        .Fields("StockMinimo") = mudtProps.StockMinimo
        .Fields("StockMaximo") = mudtProps.StockMaximo
        .Fields("LoteEconomico") = mudtProps.LoteEconomico
        .Fields("PrecioCosteEUR") = mudtProps.PrecioCosteEUR
        .Fields("PrecioCompraEUR") = mudtProps.PrecioCompraEUR
        .Fields("PrecioVentaEUR") = mudtProps.PrecioVentaEUR
        .Fields("PrecioVentaPublico") = mudtProps.PrecioVentaPublico
        .Fields("PrendaID") = mudtProps.PrendaID
        .Fields("ModeloID") = mudtProps.ModeloID
        .Fields("SerieID") = mudtProps.SerieID
        .Fields("TemporadaID") = mudtProps.TemporadaID
        .Fields("SuReferencia") = IIf(Trim(mudtProps.SuReferencia) = vbNullString, Null, mudtProps.SuReferencia)
        .Fields("ProveedorID") = IIf(mudtProps.ProveedorID = 0, Null, mudtProps.ProveedorID)
        .Fields("TallajeID") = mudtProps.TallajeID
        .Update

        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.ArticuloID = .Fields("ArticuloID")
        End If
        .Close
    End With
    
    Set rsArticulo = Nothing
    Save = GetState
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal ArticuloID As Long)
    Dim cnGescom As Connection

    On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM Articulos WHERE ArticuloID=" & _
        ArticuloID

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

Public Function Fetch(ByVal ArticuloID As Long, Optional Codigo As String, _
                      Optional TemporadaID As Long, Optional ReferenciaProveedor As String, Optional IDProveedor As Long) As String
    Dim rsArticulo As Recordset
    Dim strSQL As String
  
    On Error GoTo errh
    
    'Lectura por identificador o por codigo
    If ArticuloID <> 0 Then
        strSQL = "SELECT * FROM vArticulos WHERE ArticuloID=" & _
            ArticuloID
    ElseIf Trim(Codigo) <> vbNullString Then
        strSQL = "SELECT * FROM vArticulos WHERE Nombre='" & _
            Codigo & "' AND TemporadaID = " & TemporadaID
    ElseIf Trim(ReferenciaProveedor) <> vbNullString Then
        strSQL = "SELECT * FROM vArticulos WHERE SuReferencia='" & _
            ReferenciaProveedor & "' AND TemporadaID = " & TemporadaID
    End If
    Set rsArticulo = New Recordset
    rsArticulo.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsArticulo
        mudtProps.ArticuloID = .Fields("ArticuloID")
        mudtProps.Nombre = .Fields("Nombre") & vbNullString
        mudtProps.StockActual = .Fields("StockActual")
        mudtProps.StockPendiente = .Fields("StockPendiente")
        mudtProps.StockMinimo = .Fields("StockMinimo")
        mudtProps.StockMaximo = .Fields("StockMaximo")
        mudtProps.LoteEconomico = .Fields("LoteEconomico")
        mudtProps.PrecioCosteEUR = .Fields("PrecioCosteEUR")
        mudtProps.PrecioCompraEUR = .Fields("PrecioCompraEUR")
        mudtProps.PrecioVentaEUR = .Fields("PrecioVentaEUR")
        mudtProps.PrecioVentaPublico = .Fields("PrecioVentaPublico")
        mudtProps.PrendaID = .Fields("PrendaID")
        mudtProps.ModeloID = .Fields("ModeloID")
        mudtProps.SerieID = .Fields("SerieID")
        mudtProps.NombrePrenda = .Fields("NombrePrenda") & vbNullString
        mudtProps.NombreModelo = .Fields("NombreModelo") & vbNullString
        mudtProps.NombreSerie = .Fields("NombreSerie") & vbNullString
        mudtProps.TemporadaID = .Fields("TemporadaID")
        mudtProps.SuReferencia = .Fields("SuReferencia") & vbNullString
        mudtProps.ProveedorID = IIf(IsNull(.Fields("ProveedorID")), 0, .Fields("ProveedorID"))
        mudtProps.TallajeID = .Fields("TallajeID")
        
        .Close
    End With

    Set rsArticulo = Nothing
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

