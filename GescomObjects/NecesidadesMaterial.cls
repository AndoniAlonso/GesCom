VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NecesidadesMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mrsRecordset As ADOR.Recordset
Private mlngPedidoInicial As Long
Private mlngPedidoFinal As Long
Private mstrTipo As String * 1
Private mstrSituacion As String * 1
Private mlngTemporadaID As Long
Private mlngSerieID As Long
Private mobjSeries As TextList

Event Valid(IsValid As Boolean)

Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Private Sub Class_Initialize()

    Set mrsRecordset = New ADOR.Recordset

    mstrTipo = "T"
    mstrSituacion = "P"
    mlngPedidoInicial = 0
    mlngPedidoFinal = 0
    mlngTemporadaID = 0
    Set mobjSeries = Nothing
    mlngSerieID = 0

    Set mobjValid = New BrokenRules
    mobjValid.RuleBroken "TemporadaID", True

End Sub

Private Sub Class_Terminate()

    Set mrsRecordset = Nothing
    Set mobjSeries = Nothing
    Set mobjValid = Nothing
    
End Sub

Public Property Get NecesidadesMaterialRecordset() As ADOR.Recordset

    Set NecesidadesMaterialRecordset = mrsRecordset
  
End Property

Public Property Get TemporadaID() As Long

    TemporadaID = mlngTemporadaID
  
End Property

Public Property Let TemporadaID(Value As Long)

    mlngTemporadaID = Value
  
    mobjValid.RuleBroken "TemporadaID", (Value = 0)
  
End Property

Public Property Let Serie(Value As String)

    If mobjSeries Is Nothing Then CargarListaSeries
    
    mlngSerieID = mobjSeries.Key(Value)
  
End Property

Public Property Get Serie() As String

    If mobjSeries Is Nothing Then CargarListaSeries
    
    Serie = mobjSeries.Item(mlngSerieID)

End Property

Public Property Get Series() As TextList

    If mobjSeries Is Nothing Then CargarListaSeries
    
    Set Series = mobjSeries

End Property

Public Property Get SerieID() As Long

    SerieID = mlngSerieID

End Property

Private Sub CargarListaSeries()
    
    Set mobjSeries = New TextList
    mobjSeries.Load "Series", , mlngTemporadaID

End Sub
  
Public Sub TipoTela()

    mstrTipo = "T"
  
End Sub

Public Sub TipoOtros()

    mstrTipo = "O"
  
End Sub

Public Sub SituacionTodos()

    mstrSituacion = "T"
  
End Sub

Public Sub SituacionPendientes()

    mstrSituacion = "P"
  
End Sub

Public Sub SituacionServidos()

    mstrSituacion = "S"
  
End Sub

Public Property Let PedidoInicial(Value As Long)

  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Dato numerico no puede ser negativo"
  
  mobjValid.RuleBroken "PedidoInicial", (Value < 0)
  
  mlngPedidoInicial = Value
  
End Property

Public Property Get PedidoInicial() As Long

    PedidoInicial = mlngPedidoInicial
  
End Property

Public Property Let PedidoFinal(Value As Long)

  If Value < 0 Then _
    Err.Raise vbObjectError + 1001, "Dato numerico no puede ser negativo"
  
  mobjValid.RuleBroken "PedidoFinal", (Value < 0)
  
  mlngPedidoFinal = Value
  
End Property

Public Property Get PedidoFinal() As Long

    PedidoFinal = mlngPedidoFinal
  
End Property

Public Property Get Tipo() As String

    Tipo = mstrTipo
  
End Property

Public Property Get Situacion() As String

    Situacion = mstrSituacion
  
End Property

Public Property Get IsValid() As Boolean

    IsValid = (mobjValid.Count = 0)

End Property

Private Sub mobjValid_BrokenRule()

    RaiseEvent Valid(False)

End Sub

Private Sub mobjValid_NoBrokenRules()

    RaiseEvent Valid(True)

End Sub

Public Sub Load()
    Dim objRecordList As RecordList
    Dim strSQL As String
    Dim strWhere As String
  
    If Not IsValid Then Err.Raise 445
    
    Set objRecordList = New RecordList
    
    strSQL = "Select MaterialID, Codigo, Nombre, AnchuraTela, StockActual, " & _
             "PrecioCostePTA, PrecioCosteEUR, NombreProveedor, NombreSerie, NombreColor,Sum(Cantidad) as Cantidad, Sum(Servido) as Servido from "
    Select Case mstrTipo
    Case "T"  'Telas
        strSQL = strSQL & "vNecesidadesMaterial1"
    Case "O"  'Otros
        strSQL = strSQL & "vNecesidadesMaterial2"
    Case Else
        Err.Raise vbObjectError + 1001, "NecesidadesMaterial", "Tipo de material incorrecto"
    End Select
    
    strWhere = "TemporadaID = " & mlngTemporadaID _
        & IIf(mlngPedidoInicial = 0 And mlngPedidoFinal = 0, vbNullString, " AND Numero BETWEEN " & mlngPedidoInicial & " AND " & mlngPedidoFinal) _
        & IIf(mlngSerieID = 0, vbNullString, " AND SerieID = " & mlngSerieID) _
        & " GROUP BY MaterialID, Codigo, Nombre, AnchuraTela, StockActual, PrecioCostePTA, PrecioCosteEUR, NombreProveedor, NombreSerie, NombreColor" _
        & " ORDER BY NombreProveedor, NombreSerie, NombreColor"
        
    Set mrsRecordset = objRecordList.Load(strSQL, strWhere) _

    Set objRecordList = Nothing
  
End Sub

