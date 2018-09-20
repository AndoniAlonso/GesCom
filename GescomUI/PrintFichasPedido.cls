VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrintFichasPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjFichasPedido As FichasPedido
Private mrsRecordset As ADOR.Recordset
Private mobjSerie As Serie

Private mobjTablaCuerpo As PrintTabla
Private mobjTablaTotal As PrintTabla
Private mobjTablaTitulo As PrintTabla
Private mobjTablaSerie As PrintTabla
Private mobjTablaPie As PrintTabla
Private mintIndex As Integer
Private mlngTotalColor As Long

Public Sub Component(FichasPedidoObject As FichasPedido)

    Set mobjFichasPedido = FichasPedidoObject
    Set mrsRecordset = FichasPedidoObject.FichasPedidoRecordset
    
    Set mobjSerie = New Serie
    mobjSerie.Load mobjFichasPedido.SerieID

End Sub

Private Sub Class_Initialize()
    
' Establecemos orientacion vertical y tamaño de papel A4
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = vbPRORPortrait
    Printer.Copies = 1

    Set mobjTablaCuerpo = New PrintTabla
    With mobjTablaCuerpo
        .NumeroColumnas = 9
        .Dimensiones 0.4, 6.5, 19.2, 15.4
        .LongitudColumna 1.5
        .LongitudColumna 1.5
        .LongitudColumna 5
        .LongitudColumna 1.8
        .LongitudColumna 1.5
        .LongitudColumna 1.2
        .LongitudColumna 1.5
        .LongitudColumna 1.2
        .FuenteCuerpoSize = 7
        .FuenteCabeceraSize = 7
    End With

    Set mobjTablaTotal = New PrintTabla
    With mobjTablaTotal
        .Dimensiones 12.1, 23.2, 7.5, 0.8
        .NumeroColumnas = 2
        .LongitudColumna 4
        .Cabecera = False
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = True
        .GrosorLinea = 10
    End With
    
    Set mobjTablaTitulo = New PrintTabla
    With mobjTablaTitulo
        .Dimensiones 0.4, 2, 19, 1.4
        .Cabecera = False
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = False
        .FuenteCuerpoSize = 16
    End With
    
    Set mobjTablaSerie = New PrintTabla
    With mobjTablaSerie
        .Dimensiones 9.5, 3.2, 10.1, 3
        .NumeroColumnas = 2
        .LongitudColumna 2
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = True
    End With
    
    Set mobjTablaPie = New PrintTabla
    With mobjTablaPie
        .Dimensiones 0.4, 28, 19, 0.6
        .Cabecera = False
        .NumeroColumnas = 2
        .Autosize
        .FuenteCuerpoBold = False
        .FuenteCuerpoItalic = False
        .FuenteCuerpoSize = 8
    End With
    
    mlngTotalColor = 0
    
End Sub

Private Sub Class_Terminate()

    Set mobjFichasPedido = Nothing
    Set mrsRecordset = Nothing
    Set mobjSerie = Nothing
    
    Set mobjTablaCuerpo = Nothing
    Set mobjTablaTotal = Nothing
    Set mobjTablaTitulo = Nothing
    Set mobjTablaSerie = Nothing
    Set mobjTablaPie = Nothing
     
    Printer.EndDoc

End Sub

Private Sub ImprimirNuevaPagina()
    
    ImprimirTotal
    Printer.NewPage
    ImprimirDatosFijos

End Sub

Private Sub ImprimirDatosFijos()
    
    ImprimirCabecera
    ImprimirRejilla
    ImprimirPie
    
End Sub

Private Sub ImprimirCabecera()
    
    ImprimirSerie
    
    mobjTablaTitulo.TextoCuerpo "FICHAS DE PEDIDOS", 1, 1, vbCenter

End Sub

Private Sub ImprimirRejilla()
    
    With mobjTablaCuerpo
        .LineaTodasColumnas
        .LineasBorde
        .LineaCabecera
        .TextoCabecera "NºPedido", 1
        .TextoCabecera "Fecha", 2
        .TextoCabecera "Cliente", 3
        .TextoCabecera "Modelo", 4
        .TextoCabecera "Fecha Tope", 5
        .TextoCabecera "Cantidad", 6
        .TextoCabecera "Fecha Orden", 7
        .TextoCabecera "Nº Orden", 8
        .TextoCabecera "Observaciones", 9
    End With
    
End Sub

Private Sub ImprimirSerie()

    With mobjTablaSerie
        .LineasBorde
        
        .TextoCuerpo "Serie: ", 1, 1
        .TextoCuerpo mobjSerie.Nombre, 2, 1
        .TextoCuerpo "Color: ", 1, 2
        .TextoCuerpo mrsRecordset("CodigoColor"), 2, 2
        
    End With
    
End Sub

Private Sub ImprimirPie()

    mobjTablaPie.LineaCeldaArriba 1, 1
    mobjTablaPie.LineaCeldaArriba 2, 1
    
    mobjTablaPie.TextoCuerpo "Página " & Printer.Page, 1, 1, vbLeftJustify
    
    'Escribimos la fecha y hora de impresion ajustado a la izquierda
    mobjTablaPie.TextoCuerpo Format$(Now, "Long Date") & " a las " & Format$(Now, "short time"), _
                            2, 1, vbRightJustify
    
End Sub

Private Sub ImprimirDetalle()
    Dim LineaCuerpo As Integer
    Dim strColorActual As String
    
    LineaCuerpo = 1
    mrsRecordset.MoveFirst
    strColorActual = mrsRecordset("CodigoColor")
    
    mintIndex = 1
    ImprimirDatosFijos
    While Not mrsRecordset.EOF
    
        If strColorActual <> mrsRecordset("CodigoColor") Then
            ImprimirNuevaPagina
            LineaCuerpo = 1
        End If
        
        With mobjTablaCuerpo
            .TextoCuerpo FormatoCantidad(mrsRecordset("Numero")), 1, LineaCuerpo
            .TextoCuerpo FormatoFecha(mrsRecordset("Fecha")), 2, LineaCuerpo, vbRightJustify
            .TextoCuerpo mrsRecordset("NombreCliente"), 3, LineaCuerpo
            .TextoCuerpo mrsRecordset("NombreModelo"), 4, LineaCuerpo
            .TextoCuerpo IIf(mrsRecordset("FechaTopeServicio") = "0:00:00", _
                             vbNullString, _
                             FormatoFecha(mrsRecordset("FechaTopeServicio"))), _
                         5, LineaCuerpo, vbRightJustify
            .TextoCuerpo FormatoCantidad(mrsRecordset("Cantidad")), 6, LineaCuerpo, vbRightJustify
            If Not IsNull(mrsRecordset("FechaOrden")) Then
                .TextoCuerpo FormatoFecha(mrsRecordset("FechaOrden")), 7, LineaCuerpo, vbRightJustify
            End If
            .TextoCuerpo Format(mrsRecordset("NumeroOrden"), "####"), 8, LineaCuerpo, vbRightJustify
            .TextoCuerpo mrsRecordset("Observaciones"), 9, LineaCuerpo, vbRightJustify
            
        End With
        
        mlngTotalColor = mlngTotalColor + mrsRecordset("Cantidad")
        
        LineaCuerpo = LineaCuerpo + 1
        If LineaCuerpo > mobjTablaCuerpo.NumeroFilasCuerpo Then
            ImprimirNuevaPagina
            LineaCuerpo = 1
        End If
        
        strColorActual = mrsRecordset("CodigoColor")
        
        mrsRecordset.MoveNext
        
    Wend
    
    ImprimirTotal
    
End Sub

Private Sub ImprimirTotal()

    With mobjTablaTotal
        .LineaCelda 2, 1
        .TextoCuerpo "TOTAL Pedidos >>", 1, 1, vbCenter
        .TextoCuerpo FormatoCantidad(CDbl(mlngTotalColor)), 2, 1, vbCenter
        
    End With
    
    mlngTotalColor = 0
    
End Sub


Public Sub PrintObject()

    On Error GoTo ErrorManager
    
    If mrsRecordset.EOF Then _
        Err.Raise vbObjectError + 1001, "Impresion de fichas de pedidos", _
                                        "No existen fichas de pedidos para la serie-color."
    'ImprimirDatosFijos
    ImprimirDetalle
    Exit Sub

ErrorManager:
    Printer.KillDoc
    Err.Raise vbObjectError + 1001, Err.Source, Err.Description
End Sub

Public Property Let PrinterNumber(PrinterNumber As Integer)
    
    Set Printer = Printers(PrinterNumber)
    
End Property

Public Property Let Copies(Copies As Integer)
    
    Printer.Copies = Copies
    
End Property


