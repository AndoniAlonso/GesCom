VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OLAPQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolDisplay As Collection
Private mcolColumns As Collection
Private mqryTipoConsulta As Integer
Private mstrClausulaWhere As String

Private Sub Class_Initialize()

    Set mcolDisplay = New Collection
    Set mcolColumns = New Collection

End Sub

Private Sub Class_Terminate()
    
    Set mcolDisplay = Nothing
    Set mcolColumns = Nothing

End Sub

Public Function Count() As Long

    Count = mcolDisplay.Count

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mcolDisplay.[_NewEnum]

End Function

Public Function Item(ByVal Index As Variant) As Collection
Attribute Item.VB_UserMemId = 0

    Set Item = mcolDisplay(Index)

End Function

Public Function ReportColumns() As Collection

    Set ReportColumns = mcolColumns

End Function

Public Sub Load(ByVal qryTipoConsulta As Integer, ByVal strWhere As String)
    Dim objPersist As GCServerMTS.OLAPQueryPersist
  
    mqryTipoConsulta = qryTipoConsulta
    mstrClausulaWhere = strWhere
    Set objPersist = CreateObject("GCServerMTS.OLAPQueryPersist", PERSIST_SERVER)
    SetState (objPersist.Fetch(qryTipoConsulta, strWhere))
    Set objPersist = Nothing
  
End Sub

Public Property Get TipoConsulta() As Integer

    TipoConsulta = mqryTipoConsulta

End Property

Public Property Get ClausulaWhere() As String

    ClausulaWhere = mstrClausulaWhere

End Property

Private Sub SetState(Buffer As String)
    Dim objBuffer As Buffer
    Dim objDisplay As Collection
    Dim lngIndex As Long
  
    Set objBuffer = New Buffer
    With objBuffer
        .SetState Buffer
        For lngIndex = 1 To .Count
            Set objDisplay = New Collection
            SetRecordState .Item(lngIndex)
           
        Next
    End With

    Set objBuffer = Nothing
    
    SetColumn

End Sub

Private Sub SetRecordState(strState As String)

    Select Case mqryTipoConsulta
    Case QRY_PrevEtiqueta
        SetPrevEtiquetaState (strState)
        
    Case QRY_TallajePedido
        SetTallajePedidoState (strState)
        
    Case Else
        Err.Raise vbObjectError + 1001, "SetRecordState", "Tipo de consulta no válido"
                
    End Select
        
End Sub

Private Sub SetPrevEtiquetaState(strState As String)
    Dim udtProps As PrevEtiquetaProps
    Dim udtData As PrevEtiquetaData
    Dim colItem As Collection
    
    udtData.Buffer = strState
    LSet udtProps = udtData
            
    Set colItem = New Collection
    
    With udtProps
        colItem.Add .CodigoSerie
        colItem.Add .NombreSerie
        colItem.Add .CodigoPrenda
        colItem.Add .NombrePrenda
        colItem.Add .Cantidad
    End With
    
    mcolDisplay.Add colItem
    
    Set colItem = Nothing
    
End Sub

Private Sub SetTallajePedidoState(strState As String)
    Dim udtProps As TallajePedidoProps
    Dim udtData As TallajePedidoData
    Dim colItem As Collection
    
    udtData.Buffer = strState
    LSet udtProps = udtData
            
    Set colItem = New Collection
    
    With udtProps
        colItem.Add .CodigoModelo
        colItem.Add .NombreModelo
        colItem.Add .TallaMinima
        colItem.Add .TallaMaxima
    End With
    
    mcolDisplay.Add colItem
    
    Set colItem = Nothing
    
End Sub

Private Sub SetColumn()

    Select Case mqryTipoConsulta
    Case QRY_PrevEtiqueta
        PrevEtiquetaColumn
        
    Case QRY_TallajePedido
        TallajePedidoColumn
        
    Case Else
        Err.Raise vbObjectError + 1001, "SetColumn", "Tipo de consulta no válido"
                
    End Select
        
End Sub
Private Sub PrevEtiquetaColumn()

    mcolColumns.Add "Serie"
    mcolColumns.Add "Nombre"
    mcolColumns.Add "Prenda"
    mcolColumns.Add "Nombre"
    mcolColumns.Add "Cantidad"
    
End Sub

Private Sub TallajePedidoColumn()

    mcolColumns.Add "Modelo"
    mcolColumns.Add "Nombre"
    mcolColumns.Add "Mínima"
    mcolColumns.Add "Máxima"
    
End Sub

Public Property Get Titulo() As String

    Select Case mqryTipoConsulta
    Case QRY_PrevEtiqueta
        Titulo = "Prevision de necesidades de etiquetas"
        
    Case QRY_TallajePedido
        Titulo = "Tallajes máximo y mínimo"
        
    Case Else
        Err.Raise vbObjectError + 1001, "Titulo", "Tipo de consulta no válido"
                
    End Select
        
End Property

