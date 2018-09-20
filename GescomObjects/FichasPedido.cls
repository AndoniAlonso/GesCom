VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FichasPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mrsRecordset As ADOR.Recordset
Private mlngPedidoInicial As Long
Private mlngPedidoFinal As Long
Private mstrColor As String * 2
Private mlngSerieID As Long

Event Valid(IsValid As Boolean)

Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Private Sub Class_Initialize()

    Set mrsRecordset = New ADOR.Recordset

    mstrColor = vbNullString
    mlngSerieID = 0
    mlngPedidoInicial = 0
    mlngPedidoFinal = 0

    Set mobjValid = New BrokenRules
    mobjValid.RuleBroken "SerieID", True

End Sub

Private Sub Class_Terminate()

    Set mrsRecordset = Nothing
    Set mobjValid = Nothing
    
End Sub

'Public Function Count() As Long
'
'    Count = mrsRecordset.Count
'
'End Function
'
'Public Function NewEnum() As IUnknown
'
'    Set NewEnum = mrsRecordset.[_NewEnum]
'
'End Function
'
'Public Function Item(ByVal Index As Variant) As FichaPedidoDisplay
'
'    Set Item = mrsRecordset(Index)
'
'End Function

Public Property Get FichasPedidoRecordset() As ADOR.Recordset

    Set FichasPedidoRecordset = mrsRecordset
  
End Property

Public Property Get SerieID() As Long

    SerieID = mlngSerieID
  
End Property

Public Property Let SerieID(Value As Long)

    mlngSerieID = Value
  
    mobjValid.RuleBroken "SerieID", (Value = 0)
  
End Property

Public Property Let Color(Value As String)

    If Len(Value) > Len(mstrColor) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
    'mobjValid.RuleBroken "Color", (Len(trim(Value)) = 0)
  
    mstrColor = Trim(UCase(Value))
  
End Property

Public Property Get Color() As String

    Color = mstrColor
  
End Property

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

Public Property Get IsValid() As Boolean

    IsValid = (mobjValid.Count = 0)

End Property

Public Property Get Cantidad() As Long
    
    Cantidad = 0
    mrsRecordset.MoveFirst
    While Not mrsRecordset.EOF
        Cantidad = Cantidad + mrsRecordset("Cantidad")
        mrsRecordset.MoveNext
    Wend
    mrsRecordset.MoveFirst
  
End Property

Private Sub mobjValid_BrokenRule()

    RaiseEvent Valid(False)

End Sub

Private Sub mobjValid_NoBrokenRules()

    RaiseEvent Valid(True)

End Sub

Public Sub Load()
    Dim objRecordList As RecordList
  
    If Not IsValid Then Err.Raise 445
    
    Set objRecordList = New RecordList
    Set mrsRecordset = objRecordList.Load("SELECT * FROM vFichasPedido", "SerieID = " & mlngSerieID _
        & IIf(Trim(mstrColor) = vbNullString, vbNullString, " AND CodigoColor = '" & Trim(mstrColor) & "'") _
        & IIf(mlngPedidoInicial = 0 And mlngPedidoFinal = 0, vbNullString, " AND Numero BETWEEN " & mlngPedidoInicial & " AND " & mlngPedidoFinal) _
        & " ORDER BY CodigoColor, Numero, NombreModelo")
        
    Set objRecordList = Nothing
  
End Sub

