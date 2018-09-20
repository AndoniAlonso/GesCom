VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "FichasPedidoPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjBuffer As Buffer

Public Function Fetch(ByVal SerieID As Long, ByVal PedidoInicial As Long, ByVal PedidoFinal As Long, _
                              ByVal Color As String) As String
    Dim rsFichaPedido As Recordset
    Dim strSQL As String
    Dim udtData As FichaPedidoData
    Dim udtProps As FichaPedidoProps
  
    strSQL = "SELECT * FROM vFichasPedido WHERE SerieID = " & SerieID _
        & IIf(Color = vbNullString, vbNullString, " AND CodigoColor = '" & Color & "'") _
        & IIf(PedidoInicial = 0 And PedidoFinal = 0, vbNullString, " AND Numero BETWEEN " & PedidoInicial & " AND " & PedidoFinal) _
        & " ORDER BY CodigoColor, Numero, NombreModelo"
  
    Set rsFichaPedido = New Recordset
    rsFichaPedido.Open strSQL, DB_CONNECT
    Set mobjBuffer = New Buffer
    mobjBuffer.Initialize Len(udtData.Buffer), 30
    
    Do While Not rsFichaPedido.EOF
        With udtProps
            .Numero = rsFichaPedido("Numero")
            .Fecha = rsFichaPedido("Fecha")
            .FechaTopeServicio = IIf(IsNull(rsFichaPedido("FechaTopeServicio")), CDate("00:00:00"), rsFichaPedido("FechaTopeServicio"))
            .CodigoColor = rsFichaPedido("CodigoColor")
            .NombreModelo = rsFichaPedido("NombreModelo")
            .Observaciones = rsFichaPedido("Observaciones")
            .NombreCliente = rsFichaPedido("NombreCliente")
            .Cantidad = rsFichaPedido("Cantidad")
            .FechaOrden = IIf(IsNull(rsFichaPedido("FechaOrden")), CDate("00:00:00"), rsFichaPedido("FechaOrden"))
            .NumeroOrden = IIf(IsNull(rsFichaPedido("NumeroOrden")), 0, rsFichaPedido("NumeroOrden"))
        End With
                
        LSet udtData = udtProps
        mobjBuffer.Add udtData.Buffer
        rsFichaPedido.MoveNext
    Loop
       
    rsFichaPedido.Close
    Set rsFichaPedido = Nothing
    Fetch = mobjBuffer.GetState
  
End Function
