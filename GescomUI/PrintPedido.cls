VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrintPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const TwipsPerCm = 567

Private mobjPedidoVenta As PedidoVenta
Private mobjCliente As Cliente
Private mobjEmpresa As Empresa

Private mobjTablaCuerpo As PrintTabla
Private mobjTablaObservaciones As PrintTabla
Private mobjTablaFecha As PrintTabla
Private mobjTablaTotal As PrintTabla
Private mobjTablaCabecera As PrintTabla

Public Sub Component(PedidoVentaObject As PedidoVenta)

    Set mobjPedidoVenta = PedidoVentaObject
    
    Set mobjCliente = New Cliente
    mobjCliente.Load mobjPedidoVenta.ClienteID
    
    Set mobjEmpresa = New Empresa
    mobjEmpresa.Load mobjPedidoVenta.EmpresaID

End Sub

Private Sub Class_Initialize()
    Dim i As Integer
    
' Establecemos orientacion vertical y tamaño de papel A4
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = vbPRORPortrait
    Printer.Copies = 1

    Set mobjTablaCuerpo = New PrintTabla
    With mobjTablaCuerpo
        .NumeroColumnas = 17
        .Dimensiones 0.4, 6, 19.2, 20
        .LongitudColumna 3.35
        .LongitudColumna 1.4
        .LongitudColumna 2.3
        .LongitudColumna 1
        ' Columnas de tallas
        For i = 1 To 11
            .LongitudColumna 0.75
        Next
        
        .LongitudColumna 1
    End With

    Set mobjTablaObservaciones = New PrintTabla
    With mobjTablaObservaciones
        .NumeroColumnas = 1
        .Dimensiones 0.4, 6, 19.2, 20
    End With

    
    Set mobjTablaFecha = New PrintTabla
    With mobjTablaFecha
        .Dimensiones 12.8, 2.5, 6.8, 1.3
        .NumeroColumnas = 2
        .Autosize
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = True
    End With
    
    Set mobjTablaTotal = New PrintTabla
    With mobjTablaTotal
        .Dimensiones 12.2, 27, 7.3, 0.8
        .NumeroColumnas = 2
        .LongitudColumna 4
        .Cabecera = False
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = True
        .GrosorLinea = 2
    End With
    
    Set mobjTablaCabecera = New PrintTabla
    With mobjTablaCabecera
        .Dimensiones 0.4, 4, 19.2, 1.5
        .NumeroColumnas = 3
        .LongitudColumna 6.2
        .LongitudColumna 6.5
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = True
    End With
    
End Sub

Private Sub Class_Terminate()

    Set mobjPedidoVenta = Nothing
    Set mobjCliente = Nothing
    Set mobjEmpresa = Nothing

    Set mobjTablaCuerpo = Nothing
    Set mobjTablaObservaciones = Nothing
    Set mobjTablaFecha = Nothing
    Set mobjTablaTotal = Nothing
    Set mobjTablaCabecera = Nothing
     
    Printer.EndDoc

End Sub

Private Sub ImprimirNuevaPagina()
    
    Printer.NewPage
    ImprimirDatosFijos

End Sub

Private Sub ImprimirDatosFijos()
    
    ImprimirCabecera
    ImprimirRejilla
    
End Sub

Private Sub ImprimirCabecera()
    
    ' Imprimimos el logotipo
    Printer.PaintPicture LoadPicture(App.Path & "\logo.bmp"), _
                         1 * TwipsPerCm, 2 * TwipsPerCm, 7 * TwipsPerCm, 1.5 * TwipsPerCm
                         
    ImprimirEncabezamiento
    ImprimirFechaYNumero
    
End Sub

Private Sub ImprimirRejilla()
    Dim i As Integer
    
    With mobjTablaCuerpo
'        .LineaTodasColumnas

        .LineasBorde
        .LineaCabecera

        .TextoCabecera "REFERENCIA", 2
        '.LineaColumna 4
        For i = 1 To 11
            .TextoCabecera "T" & CStr(34 + i * 2), i + 4
            '.LineaColumna i + 4
        Next
        .TextoCabecera "UNID.", 16, vbCenter
        '.LineaColumna 16
        .TextoCabecera "PRECIO", 17, vbCenter
        
    End With
    
End Sub

Private Sub ImprimirEncabezamiento()

    With mobjTablaCabecera
        .LineasBorde
        .LineaCabecera
        .LineaTodasColumnas
        
        .TextoCabecera "Observaciones", 1, vbCenter
        .TextoCabecera "Cliente", 2
        .TextoCabecera "Representante", 3
        
        .TextoCuerpo mobjPedidoVenta.Observaciones, 1, 1
        .TextoCuerpo mobjPedidoVenta.Cliente, 2, 1
        .TextoCuerpo mobjPedidoVenta.Representante, 3, 1
    End With
    
End Sub

Private Sub ImprimirFechaYNumero()

    With mobjTablaFecha
        .LineasBorde
        .LineaCabecera
        .LineaTodasColumnas
        .TextoCabecera "FECHA", 1, vbCenter
        .TextoCabecera "Nº PEDIDO", 2, vbCenter
        .TextoCuerpo Format(mobjPedidoVenta.Fecha, "dd/mm/yyyy"), 1, 1, vbCenter
        .TextoCuerpo Format(mobjPedidoVenta.Numero, "###,###,###"), 2, 1, vbCenter
        
    End With
    
End Sub

Private Sub ImprimirDetalle()
    Dim objPedidoVentaItem As PedidoVentaItem
    Dim LineaCuerpo As Integer
    Dim intNumeroLinea As Integer
    Dim objArticuloColor As ArticuloColor
    Dim i As Integer
    Dim j As Integer
    
    
    LineaCuerpo = 1
    intNumeroLinea = 1
    For Each objPedidoVentaItem In mobjPedidoVenta.PedidoVentaItems
        Set objArticuloColor = objPedidoVentaItem.objArticuloColor
        
        With mobjTablaCuerpo
            ' OJO: Puerro y de los gordos.
            ' Cambio de criterio de codificación de artículos a partir de la temporada 16 (código 14)
            If mobjPedidoVenta.TemporadaID < 14 Then
                .TextoCuerpo Format(intNumeroLinea, "00 ") & Left(objArticuloColor.objArticulo.NombrePrenda, 10), 1, LineaCuerpo
                .TextoCuerpo objArticuloColor.objArticulo.NombreModelo, 2, LineaCuerpo
                .TextoCuerpo objArticuloColor.objArticulo.NombreSerie, 3, LineaCuerpo
                .TextoCuerpo objArticuloColor.Codigo, 4, LineaCuerpo
            Else
                .TextoCuerpo Format(intNumeroLinea, "00 ") & Left(objArticuloColor.objArticulo.NombrePrenda, 10), 1, LineaCuerpo
                .TextoCuerpo objArticuloColor.objArticulo.objModelo.Codigo, 2, LineaCuerpo
                .TextoCuerpo objArticuloColor.objArticulo.NombreSerie, 3, LineaCuerpo
                .TextoCuerpo objArticuloColor.Codigo, 4, LineaCuerpo
            End If
            .TextoCuerpo Format(objPedidoVentaItem.CantidadT36, "######"), 5, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objPedidoVentaItem.CantidadT38, "######"), 6, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objPedidoVentaItem.CantidadT40, "######"), 7, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objPedidoVentaItem.CantidadT42, "######"), 8, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objPedidoVentaItem.CantidadT44, "######"), 9, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objPedidoVentaItem.CantidadT46, "######"), 10, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objPedidoVentaItem.CantidadT48, "######"), 11, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objPedidoVentaItem.CantidadT50, "######"), 12, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objPedidoVentaItem.CantidadT52, "######"), 13, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objPedidoVentaItem.CantidadT54, "######"), 14, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objPedidoVentaItem.CantidadT56, "######"), 15, LineaCuerpo, vbRightJustify
            .TextoCuerpo objPedidoVentaItem.Cantidad, 16, LineaCuerpo, vbRightJustify
            .TextoCuerpo FormatoMoneda(objPedidoVentaItem.PrecioVenta, mobjPedidoVenta.Moneda, False), 17, LineaCuerpo, vbRightJustify
            
            .LineaCeldaDerecha 4, LineaCuerpo
            For i = 1 To 11
                .LineaCeldaDerecha i + 4, LineaCuerpo
            Next
            .LineaCeldaDerecha 16, LineaCuerpo
            
        End With
        
        Set objArticuloColor = Nothing
        
        intNumeroLinea = intNumeroLinea + 1
        LineaCuerpo = LineaCuerpo + 1
        If LineaCuerpo > mobjTablaCuerpo.NumeroFilasCuerpo Then
            ImprimirNuevaPagina
            LineaCuerpo = 1
        End If
        
        If objPedidoVentaItem.Observaciones <> vbNullString Then
            mobjTablaObservaciones.TextoCuerpo objPedidoVentaItem.Observaciones, 1, LineaCuerpo
            intNumeroLinea = intNumeroLinea + 1
            LineaCuerpo = LineaCuerpo + 1
            If LineaCuerpo > mobjTablaCuerpo.NumeroFilasCuerpo Then
                ImprimirNuevaPagina
                LineaCuerpo = 1
            End If
            
        End If
    Next
    
    'Rellenar las líneas hasta el final
    For j = LineaCuerpo To mobjTablaCuerpo.NumeroFilasCuerpo
        mobjTablaCuerpo.LineaCeldaDerecha 4, j
        For i = 1 To 11
            mobjTablaCuerpo.LineaCeldaDerecha i + 4, j
        Next
        mobjTablaCuerpo.LineaCeldaDerecha 16, j
    Next

    With mobjTablaTotal
        .LineaCelda 2, 1
        .TextoCuerpo "TOTAL Pedido >>", 1, 1, vbCenter
        .TextoCuerpo mobjPedidoVenta.PedidoVentaItems.Cantidad, 2, 1, vbCenter
        
    End With

End Sub

Public Sub PrintObject()

    On Error GoTo ErrorManager
    
    If mobjPedidoVenta.IsNew Then _
        Err.Raise vbObjectError + 1001, "No se puede imprimir, no existe la órden de corte."
    
    ImprimirDatosFijos
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


