VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CobroPagosPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ObjectControl

Private mobjContext As ObjectContext

Private mobjBuffer As Buffer

Public Function Fetch(ByVal TemporadaID As Long, ByVal EmpresaID As Long, ByVal Tipo As String, ByVal ClausulaWhere As String) As String
    Dim rsCobroPago As Recordset
    Dim strSQL As String
    Dim udtData As CobroPagoData
    Dim udtProps As CobroPagoProps
    
    On Error GoTo errh
    If Tipo = "C" Then _
        strSQL = "SELECT * FROM vCobros WHERE Tipo = '" & Tipo & "'"
        
    If Tipo = "P" Then _
        strSQL = "SELECT * FROM vPagos WHERE Tipo = '" & Tipo & "'"
        
    If TemporadaID Then strSQL = strSQL & " AND TemporadaID = " & TemporadaID
    If EmpresaID Then strSQL = strSQL & " AND EmpresaID = " & EmpresaID
          
    ' Aqui se pondrian condiciones de busqueda/filtrado en base a parametros
    If Len(ClausulaWhere) > 0 Then
        strSQL = strSQL & " AND " & ClausulaWhere
    End If
    
    Set rsCobroPago = New Recordset
    rsCobroPago.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    Set mobjBuffer = New Buffer
    mobjBuffer.Initialize Len(udtData.Buffer), 30
    
    Do While Not rsCobroPago.EOF
        With udtProps
            .CobroPagoID = rsCobroPago("CobroPagoID")
            .Vencimiento = rsCobroPago("Vencimiento")
            .NombrePersona = rsCobroPago("NombrePersona") & vbNullString
            .FormaPago = rsCobroPago("FormaPago") & vbNullString
            .NumeroFactura = IIf(IsNull(rsCobroPago("NumeroFactura")), 0, rsCobroPago("NumeroFactura"))
            .NumeroGiro = rsCobroPago("NumeroGiro")
            .SituacionComercial = rsCobroPago("SituacionComercial") & vbNullString
            .SituacionContable = rsCobroPago("SituacionContable") & vbNullString
            .TemporadaID = rsCobroPago("TemporadaID")
            .EmpresaID = rsCobroPago("EmpresaID")
            .ImporteEUR = rsCobroPago("ImporteEUR")
            .ImportePTA = rsCobroPago("ImportePTA")
            .FechaContable = rsCobroPago("FechaContable")
            .FechaDomiciliacion = rsCobroPago("FechaDomiciliacion")
            .FechaEmision = rsCobroPago("FechaEmision")
            .NombreBanco = rsCobroPago("NombreBanco") & vbNullString
        End With
                
        LSet udtData = udtProps
        mobjBuffer.Add udtData.Buffer
        rsCobroPago.MoveNext
    Loop
       
    rsCobroPago.Close
    Set rsCobroPago = Nothing
    Fetch = mobjBuffer.GetState
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
