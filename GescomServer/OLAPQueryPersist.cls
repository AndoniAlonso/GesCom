VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "OLAPQueryPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
Private mqryTipoConsulta As Integer
Private mobjBuffer As Buffer
Private mrsQuery As Recordset

Public Function Fetch(ByVal qryTipoConsulta As Integer, Optional ByVal ClausulaWhere As String) As String
    Dim strSQL As String
    Dim strTabla As String
  
    mqryTipoConsulta = qryTipoConsulta
    strTabla = GetQueryName

    On Error GoTo errh
    strSQL = strTabla & IIf(ClausulaWhere = vbNullString, vbNullString, " WHERE " & ClausulaWhere)
  
    Set mrsQuery = New Recordset
    mrsQuery.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    InicializarBuffer

    Do While Not mrsQuery.EOF
        mobjBuffer.Add GetRecordState
        mrsQuery.MoveNext
    Loop

    mrsQuery.Close
    Set mrsQuery = Nothing
    Fetch = mobjBuffer.GetState
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

' Nos va a devolver el nombre de la consulta a realizar.
Private Function GetQueryName() As String

    Select Case mqryTipoConsulta
    Case QRY_PrevEtiqueta
        GetQueryName = "spPrevEtiqueta"
        
    Case QRY_TallajePedido
        GetQueryName = "SELECT * FROM vTallajePedido"
        
    Case Else
        Err.Raise vbObjectError + 1001, "GetQueryName", "Tipo de consulta no válido"
                
    End Select

End Function

' Nos va a devolver el tamaño de un registro de la consulta.
Private Function GetRecordLenght() As Long
    Dim udtPrevEtiqueta As PrevEtiquetaData
    Dim udtTallajePedido As TallajePedidoData

    Select Case mqryTipoConsulta
    Case QRY_PrevEtiqueta
        GetRecordLenght = Len(udtPrevEtiqueta.Buffer)
        
    Case QRY_TallajePedido
        GetRecordLenght = Len(udtTallajePedido.Buffer)
        
    Case Else
        Err.Raise vbObjectError + 1001, "GetRecordLenght", "Tipo de consulta no válido"
                
    End Select
    
End Function

Private Function GetRecordState() As String

    Select Case mqryTipoConsulta
    Case QRY_PrevEtiqueta
        GetRecordState = PrevEtiquetaState
        
    Case QRY_TallajePedido
        GetRecordState = TallajePedidoState
        
    Case Else
        Err.Raise vbObjectError + 1001, "GetRecordState", "Tipo de consulta no válido"
                
    End Select
        
End Function

Private Function PrevEtiquetaState() As String
    Dim udtProps As PrevEtiquetaProps
    Dim udtData As PrevEtiquetaData
    
    With udtProps
        .CodigoSerie = mrsQuery("CodigoSerie") & vbNullString
        .NombreSerie = mrsQuery("NombreSerie") & vbNullString
        .CodigoPrenda = mrsQuery("CodigoPrenda") & vbNullString
        .NombrePrenda = mrsQuery("NombrePrenda") & vbNullString
        .Cantidad = mrsQuery("Cantidad")
    End With
        
    LSet udtData = udtProps
    PrevEtiquetaState = udtData.Buffer
End Function

Private Function TallajePedidoState() As String
    Dim udtProps As TallajePedidoProps
    Dim udtData As TallajePedidoData
    
    With udtProps
        .CodigoModelo = mrsQuery("CodigoModelo") & vbNullString
        .NombreModelo = mrsQuery("NombreModelo") & vbNullString
        .TallaMaxima = mrsQuery("TallaMaxima") & vbNullString
        .TallaMinima = mrsQuery("TallaMinima") & vbNullString
    End With
        
    LSet udtData = udtProps
    TallajePedidoState = udtData.Buffer
End Function

Private Sub InicializarBuffer()

    Set mobjBuffer = New Buffer
    mobjBuffer.Initialize GetRecordLenght, 30

End Sub

Private Function ObjectControl_CanBePooled() As Boolean

    ObjectControl_CanBePooled = False
  
End Function

Private Sub ObjectControl_Activate()

    Set mobjContext = GetObjectContext

End Sub

Private Sub ObjectControl_Deactivate()

    Set mobjContext = Nothing

End Sub

