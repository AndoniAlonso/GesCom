VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CobroPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolDisplay As Collection
Private mTemporadaID As Long
Private mEmpresaID As Long
Private mTipo As String
Private mMoneda As String

Private Sub Class_Initialize()

    Set mcolDisplay = New Collection

End Sub

Public Function Count() As Long

    Count = mcolDisplay.Count

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mcolDisplay.[_NewEnum]

End Function

Public Function Item(ByVal Index As Variant) As CobroPagoDisplay
Attribute Item.VB_UserMemId = 0

    Set Item = mcolDisplay(Index)
  
End Function

Public Sub Load(ByVal ClausulaWhere As String)
    Dim objPersist As CobroPagosPersist
  
    Set objPersist = CreateObject("GCServerMTS.CobroPagosPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(mTemporadaID, mEmpresaID, mTipo, ClausulaWhere)
    Set objPersist = Nothing
  
End Sub

Public Property Let Moneda(Value As String)

    If Not EsMonedaValida(Value) Then _
        Err.Raise vbObjectError + 1002, "Moneda debe ser EUR ó PTA"

    mMoneda = UCase$(Value)

End Property

Public Property Get Moneda() As String

    Moneda = Trim$(mMoneda)

End Property

Private Sub SetState(Buffer As String)

    Dim objBuffer As Buffer
    Dim objDisplay As CobroPagoDisplay
    Dim lngIndex As Long
    Dim udtData As CobroPagoData
    Dim udtProps As CobroPagoProps
  
    Set objBuffer = New Buffer
    With objBuffer
        .SetState Buffer
        For lngIndex = 1 To .Count
            Set objDisplay = New CobroPagoDisplay
            udtData.Buffer = .Item(lngIndex)
            LSet udtProps = udtData
    
            With objDisplay
                .CobroPagoID = udtProps.CobroPagoID
                .Vencimiento = udtProps.Vencimiento
                .NombrePersona = Trim$(udtProps.NombrePersona)
                .FormaPago = Trim$(udtProps.FormaPago)
                .NumeroFactura = udtProps.NumeroFactura
                .NumeroGiro = udtProps.NumeroGiro
                .SituacionComercial = udtProps.SituacionComercial
                .SituacionContable = udtProps.SituacionContable
                .Moneda = mMoneda
                .ImporteEUR = udtProps.ImporteEUR
                .ImportePTA = udtProps.ImportePTA
                .FechaEmision = udtProps.FechaEmision
                .FechaDomiciliacion = udtProps.FechaDomiciliacion
                .FechaContable = udtProps.FechaContable
                .NombreBanco = Trim$(udtProps.NombreBanco)
                mcolDisplay.Add objDisplay
                Set objDisplay = Nothing
            End With
        Next
    End With

    Set objBuffer = Nothing

End Sub

Public Property Let TemporadaID(Value As Long)

    mTemporadaID = Value

End Property

Public Property Get TemporadaID() As Long

    TemporadaID = mTemporadaID

End Property

Public Property Let EmpresaID(Value As Long)

    mEmpresaID = Value

End Property

Public Property Get EmpresaID() As Long

    EmpresaID = mEmpresaID

End Property

Public Property Let Tipo(Value As String)

    mTipo = Value
    
End Property

Public Property Get Tipo() As String

    Tipo = mTipo
    
End Property
