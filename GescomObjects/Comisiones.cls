VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Comisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mrsRecordset As ADOR.Recordset
Private mdtFechaInicial As Date
Private mdtFechaFinal As Date
'Private mlngTemporadaID As Long
Private mlngEmpresaID As Long
Private mlngRepresentanteID As Long
Private mobjRepresentantes As TextList

Event Valid(IsValid As Boolean)

Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Private Sub Class_Initialize()

    Set mrsRecordset = New ADOR.Recordset

    mdtFechaInicial = Date
    mdtFechaFinal = Date
    'mlngTemporadaID = 0
    mlngEmpresaID = 0
    Set mobjRepresentantes = Nothing
    mlngRepresentanteID = 0

    Set mobjValid = New BrokenRules
    'mobjValid.RuleBroken "TemporadaID", True
    mobjValid.RuleBroken "EmpresaID", True

End Sub

Private Sub Class_Terminate()

    Set mrsRecordset = Nothing
    Set mobjRepresentantes = Nothing
    Set mobjValid = Nothing
    
End Sub

Public Property Get ComisionesRecordset() As ADOR.Recordset

    Set ComisionesRecordset = mrsRecordset
  
End Property

'Public Property Get TemporadaID() As Long
'
'    TemporadaID = mlngTemporadaID
'
'End Property
'
'Public Property Let TemporadaID(Value As Long)
'
'    mlngTemporadaID = Value
'
'    mobjValid.RuleBroken "TemporadaID", (Value = 0)
'
'End Property

Public Property Get EmpresaID() As Long

    EmpresaID = mlngEmpresaID
  
End Property

Public Property Let EmpresaID(Value As Long)

    mlngEmpresaID = Value
  
    mobjValid.RuleBroken "EmpresaID", (Value = 0)
  
End Property

Public Property Let Representante(Value As String)

    If mobjRepresentantes Is Nothing Then CargarListaRepresentantes
    
    mlngRepresentanteID = mobjRepresentantes.Key(Value)
  
End Property

Public Property Get Representante() As String

    If mobjRepresentantes Is Nothing Then CargarListaRepresentantes
    
    Representante = mobjRepresentantes.Item(mlngRepresentanteID)

End Property

Public Property Get Representantes() As TextList

    If mobjRepresentantes Is Nothing Then CargarListaRepresentantes
    
    Set Representantes = mobjRepresentantes

End Property

Public Property Get RepresentanteID() As Long

    RepresentanteID = mlngRepresentanteID

End Property

Private Sub CargarListaRepresentantes()
    
    Set mobjRepresentantes = New TextList
    mobjRepresentantes.Load "Representantes"

End Sub
  
Public Property Let FechaInicial(Value As Date)

  mobjValid.RuleBroken "FechaInicial", Not IsDate(Value)
  
  mdtFechaInicial = Value
  
End Property

Public Property Get FechaInicial() As Date

    FechaInicial = mdtFechaInicial
  
End Property

Public Property Let FechaFinal(Value As Date)

  mobjValid.RuleBroken "FechaFinal", Not IsDate(Value)
  
  mdtFechaFinal = Value
  
End Property

Public Property Get FechaFinal() As Date

    FechaFinal = mdtFechaFinal
  
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
    
    strSQL = "Select * from vComisiones"
    
    'strWhere = "TemporadaID = " & mlngTemporadaID
    strWhere = " EmpresaID = " & mlngEmpresaID _
        & IIf(mdtFechaInicial = 0 And mdtFechaFinal = 0, vbNullString, " AND Fecha BETWEEN '" & mdtFechaInicial & "' AND '" & mdtFechaFinal & "'") _
        & IIf(mlngRepresentanteID = 0, vbNullString, " AND RepresentanteID = " & mlngRepresentanteID) _
        & " ORDER BY NombreRepresentante, Fecha"
        
    Set mrsRecordset = objRecordList.Load(strSQL, strWhere)

    Set objRecordList = Nothing
  
End Sub

