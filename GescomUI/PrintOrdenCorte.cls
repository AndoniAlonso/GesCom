VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrintOrdenCorte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjOrdenCorte As OrdenCorte
Private mobjArticulo As Articulo
Private mobjEmpresa As Empresa

Private mobjTablaCuerpo As PrintTabla
Private mobjTablaCuerpo2 As PrintTabla
Private mobjTablaObservaciones As PrintTabla
Private mobjTablaObservaciones2 As PrintTabla
Private mobjTablaFecha As PrintTabla
Private mobjTablaFecha2 As PrintTabla
Private mobjTablaTotal As PrintTabla
Private mobjTablaTotal2 As PrintTabla
Private mobjTablaArticulo As PrintTabla
Private mobjTablaArticulo2 As PrintTabla

Public Sub Component(OrdenCorteObject As OrdenCorte)

    Set mobjOrdenCorte = OrdenCorteObject
    
    Set mobjArticulo = New Articulo
    mobjArticulo.Load mobjOrdenCorte.ArticuloID
    
    Set mobjEmpresa = New Empresa
    mobjEmpresa.Load mobjOrdenCorte.EmpresaID

End Sub

Private Sub Class_Initialize()
    Dim i As Integer
    
' Establecemos orientacion vertical y tamaño de papel A4
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = vbPRORPortrait
    Printer.Copies = 1

    Set mobjTablaCuerpo = New PrintTabla
    With mobjTablaCuerpo
        .NumeroColumnas = 15
        .Dimensiones 0.4, 2.6, 19.2, 10
        .LongitudColumna 3
        ' Columnas de tallas
        For i = 1 To 11
            .LongitudColumna 0.9
        Next
        
        .LongitudColumna 1
        .LongitudColumna 1.5
    End With
    
    Set mobjTablaObservaciones = New PrintTabla
    With mobjTablaObservaciones
        .NumeroColumnas = 1
        .Dimensiones 0.4, 2.6, 19.2, 10
        .Autosize
    End With
    Set mobjTablaCuerpo2 = New PrintTabla
    With mobjTablaCuerpo2
        .NumeroColumnas = 15
        .Dimensiones 0.4, 13.7 + 2.6, 19.2, 10
        .LongitudColumna 3
        ' Columnas de tallas
        For i = 1 To 11
            .LongitudColumna 0.9
        Next
        
        .LongitudColumna 1
        .LongitudColumna 1.5
    End With

    Set mobjTablaObservaciones2 = New PrintTabla
    With mobjTablaObservaciones2
        .NumeroColumnas = 1
        .Dimensiones 0.4, 13.7 + 2.6, 19.2, 10
        .Autosize
    End With

    Set mobjTablaFecha = New PrintTabla
    With mobjTablaFecha
        .Dimensiones 0.4, 1, 7.3, 1.3
        .NumeroColumnas = 2
        .Autosize
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = True
    End With
    
    Set mobjTablaFecha2 = New PrintTabla
    With mobjTablaFecha2
        .Dimensiones 0.4, 13.7 + 1, 7.3, 1.3
        .NumeroColumnas = 2
        .Autosize
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = True
    End With
    
    Set mobjTablaTotal = New PrintTabla
    With mobjTablaTotal
        .Dimensiones 0.4, 12.9, 19.2, 0.8
        .NumeroColumnas = 15
        .LongitudColumna 3
        ' Columnas de tallas
        For i = 1 To 11
            .LongitudColumna 0.9
        Next
        .LongitudColumna 1
        .LongitudColumna 1.5
        .Cabecera = False
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = True
        .GrosorLinea = 10
    End With
    
    Set mobjTablaTotal2 = New PrintTabla
    With mobjTablaTotal2
        .Dimensiones 0.4, 13.7 + 12.9, 7.5, 0.8
        .NumeroColumnas = 15
        .LongitudColumna 3
        ' Columnas de tallas
        For i = 1 To 11
            .LongitudColumna 0.9
        Next
        .LongitudColumna 1
        .LongitudColumna 1.5
        .Cabecera = False
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = True
        .GrosorLinea = 10
    End With
    
    Set mobjTablaArticulo = New PrintTabla
    With mobjTablaArticulo
        .Dimensiones 9.5, 1, 10.1, 1.5
        .NumeroColumnas = 2
        .LongitudColumna 2
        .Cabecera = False
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = True
    End With
    
    Set mobjTablaArticulo2 = New PrintTabla
    With mobjTablaArticulo2
        .Dimensiones 9.5, 13.7 + 1, 10.1, 1.5
        .NumeroColumnas = 2
        .LongitudColumna 2
        .Cabecera = False
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = True
    End With
    
End Sub

Private Sub Class_Terminate()

    Set mobjOrdenCorte = Nothing
    Set mobjArticulo = Nothing
    Set mobjEmpresa = Nothing
    
    Set mobjTablaCuerpo = Nothing
    Set mobjTablaCuerpo2 = Nothing
    Set mobjTablaObservaciones = Nothing
    Set mobjTablaObservaciones2 = Nothing
    Set mobjTablaFecha = Nothing
    Set mobjTablaFecha2 = Nothing
    Set mobjTablaTotal = Nothing
    Set mobjTablaTotal2 = Nothing
    Set mobjTablaArticulo = Nothing
    Set mobjTablaArticulo2 = Nothing
     
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
        
    ImprimirArticulo
    ImprimirFechaYNumero
    
End Sub

Private Sub ImprimirRejilla()
    Dim i As Integer
    
    With mobjTablaCuerpo
        .LineasBorde
        .LineaCabecera
        '.LineaTodasColumnas
        .TextoCabecera "COLOR", 1
        For i = 1 To 11
            .TextoCabecera "T" & CStr(34 + i * 2), i + 1
        Next
        .TextoCabecera "TOTAL", 13, vbCenter
        .TextoCabecera "PEDIDO", 14, vbCenter
        .TextoCabecera "CLIENTE", 15, vbCenter
        
    End With
    
    With mobjTablaCuerpo2
        .LineasBorde
        .LineaCabecera
        '.LineaTodasColumnas
        .TextoCabecera "COLOR", 1
        For i = 1 To 11
            .TextoCabecera "T" & CStr(34 + i * 2), i + 1
        Next
        .TextoCabecera "TOTAL", 13, vbCenter
        .TextoCabecera "PEDIDO", 14, vbCenter
        .TextoCabecera "CLIENTE", 15, vbCenter
        
    End With
    
End Sub

Private Sub ImprimirArticulo()

    With mobjTablaArticulo
        .LineasBorde
        
        .TextoCuerpo "Prenda : ", 1, 1
        .TextoCuerpo mobjArticulo.Prenda, 2, 1
        .TextoCuerpo "Modelo: ", 1, 2
        .TextoCuerpo mobjArticulo.Modelo, 2, 2
        .TextoCuerpo "Serie: ", 1, 3
        .TextoCuerpo mobjArticulo.Serie, 2, 3
        
    End With
    
    With mobjTablaArticulo2
        .LineasBorde
        
        .TextoCuerpo "Prenda : ", 1, 1
        .TextoCuerpo mobjArticulo.Prenda, 2, 1
        .TextoCuerpo "Modelo: ", 1, 2
        .TextoCuerpo mobjArticulo.Modelo, 2, 2
        .TextoCuerpo "Serie: ", 1, 3
        .TextoCuerpo mobjArticulo.Serie, 2, 3
        
    End With
    
End Sub

Private Sub ImprimirFechaYNumero()

    With mobjTablaFecha
        
        .LineasBorde
        .LineaCabecera
        .LineaTodasColumnas
        .TextoCabecera "FECHA", 1, vbCenter
        .TextoCabecera "Nº ORDEN", 2, vbCenter
        .TextoCuerpo Format(mobjOrdenCorte.Fecha, "dd/mm/yyyy"), 1, 1, vbCenter
        .TextoCuerpo Format(mobjOrdenCorte.Numero, "###,###,###"), 2, 1, vbCenter
        
    End With
    
    With mobjTablaFecha2
        
        .LineasBorde
        .LineaCabecera
        .LineaTodasColumnas
        .TextoCabecera "FECHA", 1, vbCenter
        .TextoCabecera "Nº ORDEN", 2, vbCenter
        .TextoCuerpo Format(mobjOrdenCorte.Fecha, "dd/mm/yyyy"), 1, 1, vbCenter
        .TextoCuerpo Format(mobjOrdenCorte.Numero, "###,###,###"), 2, 1, vbCenter
        
    End With
    
End Sub

Private Sub ImprimirDetalle()
    Dim objOrdenCorteItem As OrdenCorteItem
    Dim objArticuloColor As ArticuloColor
    Dim objPedidoVentaItem As PedidoVentaItem
    Dim objPedidoVenta As PedidoVenta
    Dim LineaCuerpo As Integer
    Dim TotalCantidadT36 As Integer
    Dim TotalCantidadT38 As Integer
    Dim TotalCantidadT40 As Integer
    Dim TotalCantidadT42 As Integer
    Dim TotalCantidadT44 As Integer
    Dim TotalCantidadT46 As Integer
    Dim TotalCantidadT48 As Integer
    Dim TotalCantidadT50 As Integer
    Dim TotalCantidadT52 As Integer
    Dim TotalCantidadT54 As Integer
    Dim TotalCantidadT56 As Integer
    Dim i As Integer
    
    TotalCantidadT36 = 0
    TotalCantidadT38 = 0
    TotalCantidadT40 = 0
    TotalCantidadT42 = 0
    TotalCantidadT44 = 0
    TotalCantidadT46 = 0
    TotalCantidadT48 = 0
    TotalCantidadT50 = 0
    TotalCantidadT52 = 0
    TotalCantidadT54 = 0
    TotalCantidadT56 = 0
    
    LineaCuerpo = 0
    For Each objOrdenCorteItem In mobjOrdenCorte.OrdenCorteItems
        LineaCuerpo = LineaCuerpo + 1
        If LineaCuerpo > mobjTablaCuerpo.NumeroFilasCuerpo Then
            ImprimirNuevaPagina
            LineaCuerpo = 1
        End If
        
        Set objArticuloColor = New ArticuloColor
        objArticuloColor.Load objOrdenCorteItem.ArticuloColorID, mobjOrdenCorte.Moneda
        
        Set objPedidoVentaItem = New PedidoVentaItem
        objPedidoVentaItem.Load objOrdenCorteItem.PedidoVentaItemID, mobjOrdenCorte.Moneda
        
        Set objPedidoVenta = New PedidoVenta
        objPedidoVenta.Load objPedidoVentaItem.PedidoVentaID, mobjOrdenCorte.Moneda
        
        'Acumular totales por tallas
        TotalCantidadT36 = TotalCantidadT36 + objOrdenCorteItem.CantidadT36
        TotalCantidadT38 = TotalCantidadT38 + objOrdenCorteItem.CantidadT38
        TotalCantidadT40 = TotalCantidadT40 + objOrdenCorteItem.CantidadT40
        TotalCantidadT42 = TotalCantidadT42 + objOrdenCorteItem.CantidadT42
        TotalCantidadT44 = TotalCantidadT44 + objOrdenCorteItem.CantidadT44
        TotalCantidadT46 = TotalCantidadT46 + objOrdenCorteItem.CantidadT46
        TotalCantidadT48 = TotalCantidadT48 + objOrdenCorteItem.CantidadT48
        TotalCantidadT50 = TotalCantidadT50 + objOrdenCorteItem.CantidadT50
        TotalCantidadT52 = TotalCantidadT52 + objOrdenCorteItem.CantidadT52
        TotalCantidadT54 = TotalCantidadT54 + objOrdenCorteItem.CantidadT54
        TotalCantidadT56 = TotalCantidadT56 + objOrdenCorteItem.CantidadT56
        
        With mobjTablaCuerpo
            .LineaCeldaDerecha 1, LineaCuerpo
            .LineaCeldaDerecha 2, LineaCuerpo
            .LineaCeldaDerecha 3, LineaCuerpo
            .LineaCeldaDerecha 4, LineaCuerpo
            .LineaCeldaDerecha 5, LineaCuerpo
            .LineaCeldaDerecha 6, LineaCuerpo
            .LineaCeldaDerecha 7, LineaCuerpo
            .LineaCeldaDerecha 8, LineaCuerpo
            .LineaCeldaDerecha 9, LineaCuerpo
            .LineaCeldaDerecha 10, LineaCuerpo
            .LineaCeldaDerecha 11, LineaCuerpo
            .LineaCeldaDerecha 12, LineaCuerpo
            .LineaCeldaDerecha 13, LineaCuerpo
            .LineaCeldaDerecha 14, LineaCuerpo
            .TextoCuerpo objArticuloColor.Codigo & "-" & objArticuloColor.NombreColor, 1, LineaCuerpo
            .TextoCuerpo Format(objOrdenCorteItem.CantidadT36, "###"), 2, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objOrdenCorteItem.CantidadT38, "###"), 3, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objOrdenCorteItem.CantidadT40, "###"), 4, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objOrdenCorteItem.CantidadT42, "###"), 5, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objOrdenCorteItem.CantidadT44, "###"), 6, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objOrdenCorteItem.CantidadT46, "###"), 7, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objOrdenCorteItem.CantidadT48, "###"), 8, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objOrdenCorteItem.CantidadT50, "###"), 9, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objOrdenCorteItem.CantidadT52, "###"), 10, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objOrdenCorteItem.CantidadT54, "###"), 11, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objOrdenCorteItem.CantidadT56, "###"), 12, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objOrdenCorteItem.Cantidad, "###"), 13, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objPedidoVenta.Numero, "###,###"), 14, LineaCuerpo, vbRightJustify
            .TextoCuerpo objOrdenCorteItem.Cliente, 15, LineaCuerpo
            
        End With
        
        With mobjTablaCuerpo2
            .LineaCeldaDerecha 1, LineaCuerpo
            .LineaCeldaDerecha 2, LineaCuerpo
            .LineaCeldaDerecha 3, LineaCuerpo
            .LineaCeldaDerecha 4, LineaCuerpo
            .LineaCeldaDerecha 5, LineaCuerpo
            .LineaCeldaDerecha 6, LineaCuerpo
            .LineaCeldaDerecha 7, LineaCuerpo
            .LineaCeldaDerecha 8, LineaCuerpo
            .LineaCeldaDerecha 9, LineaCuerpo
            .LineaCeldaDerecha 10, LineaCuerpo
            .LineaCeldaDerecha 11, LineaCuerpo
            .LineaCeldaDerecha 12, LineaCuerpo
            .LineaCeldaDerecha 13, LineaCuerpo
            .LineaCeldaDerecha 14, LineaCuerpo
            .TextoCuerpo objArticuloColor.Codigo & "-" & objArticuloColor.NombreColor, 1, LineaCuerpo
            .TextoCuerpo Format(objOrdenCorteItem.CantidadT36, "###"), 2, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objOrdenCorteItem.CantidadT38, "###"), 3, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objOrdenCorteItem.CantidadT40, "###"), 4, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objOrdenCorteItem.CantidadT42, "###"), 5, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objOrdenCorteItem.CantidadT44, "###"), 6, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objOrdenCorteItem.CantidadT46, "###"), 7, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objOrdenCorteItem.CantidadT48, "###"), 8, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objOrdenCorteItem.CantidadT50, "###"), 9, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objOrdenCorteItem.CantidadT52, "###"), 10, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objOrdenCorteItem.CantidadT54, "###"), 11, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objOrdenCorteItem.CantidadT56, "###"), 12, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objOrdenCorteItem.Cantidad, "###"), 13, LineaCuerpo, vbRightJustify
            .TextoCuerpo Format(objPedidoVenta.Numero, "###,###"), 14, LineaCuerpo, vbRightJustify
            .TextoCuerpo objOrdenCorteItem.Cliente, 15, LineaCuerpo
            
        End With
        If Trim(objPedidoVentaItem.Observaciones) <> vbNullString Then
        
            LineaCuerpo = LineaCuerpo + 1
            If LineaCuerpo > mobjTablaCuerpo.NumeroFilasCuerpo Then
                ImprimirNuevaPagina
                LineaCuerpo = 1
            End If
            mobjTablaObservaciones.TextoCuerpo objPedidoVentaItem.Observaciones, 1, LineaCuerpo, vbRightJustify
            mobjTablaObservaciones2.TextoCuerpo objPedidoVentaItem.Observaciones, 1, LineaCuerpo, vbRightJustify
        End If

        Set objArticuloColor = Nothing
        Set objPedidoVentaItem = Nothing
        Set objPedidoVenta = Nothing
        
    Next
    ' Imprimir totales de prendas
    With mobjTablaTotal
        .LineaCelda 2, 1
        .LineaCelda 3, 1
        .LineaCelda 4, 1
        .LineaCelda 5, 1
        .LineaCelda 6, 1
        .LineaCelda 7, 1
        .LineaCelda 8, 1
        .LineaCelda 9, 1
        .LineaCelda 10, 1
        .LineaCelda 11, 1
        .LineaCelda 12, 1
        .LineaCelda 13, 1
        .TextoCuerpo "TOTAL Orden >>", 1, 1, vbCenter
        .TextoCuerpo Format(TotalCantidadT36, "###"), 2, 1, vbRightJustify
        .TextoCuerpo Format(TotalCantidadT38, "###"), 3, 1, vbRightJustify
        .TextoCuerpo Format(TotalCantidadT40, "###"), 4, 1, vbRightJustify
        .TextoCuerpo Format(TotalCantidadT42, "###"), 5, 1, vbRightJustify
        .TextoCuerpo Format(TotalCantidadT44, "###"), 6, 1, vbRightJustify
        .TextoCuerpo Format(TotalCantidadT46, "###"), 7, 1, vbRightJustify
        .TextoCuerpo Format(TotalCantidadT48, "###"), 8, 1, vbRightJustify
        .TextoCuerpo Format(TotalCantidadT50, "###"), 9, 1, vbRightJustify
        .TextoCuerpo Format(TotalCantidadT52, "###"), 10, 1, vbRightJustify
        .TextoCuerpo Format(TotalCantidadT54, "###"), 11, 1, vbRightJustify
        .TextoCuerpo Format(TotalCantidadT56, "###"), 12, 1, vbRightJustify
        .TextoCuerpo Format(mobjOrdenCorte.OrdenCorteItems.Cantidad, "###"), 13, 1, vbRightJustify
    End With
    With mobjTablaTotal2
        .LineaCelda 2, 1
        .LineaCelda 3, 1
        .LineaCelda 4, 1
        .LineaCelda 5, 1
        .LineaCelda 6, 1
        .LineaCelda 7, 1
        .LineaCelda 8, 1
        .LineaCelda 9, 1
        .LineaCelda 10, 1
        .LineaCelda 11, 1
        .LineaCelda 12, 1
        .LineaCelda 13, 1
        .TextoCuerpo "TOTAL Orden >>", 1, 1, vbCenter
        .TextoCuerpo Format(TotalCantidadT36, "###"), 2, 1, vbRightJustify
        .TextoCuerpo Format(TotalCantidadT38, "###"), 3, 1, vbRightJustify
        .TextoCuerpo Format(TotalCantidadT40, "###"), 4, 1, vbRightJustify
        .TextoCuerpo Format(TotalCantidadT42, "###"), 5, 1, vbRightJustify
        .TextoCuerpo Format(TotalCantidadT44, "###"), 6, 1, vbRightJustify
        .TextoCuerpo Format(TotalCantidadT46, "###"), 7, 1, vbRightJustify
        .TextoCuerpo Format(TotalCantidadT48, "###"), 8, 1, vbRightJustify
        .TextoCuerpo Format(TotalCantidadT50, "###"), 9, 1, vbRightJustify
        .TextoCuerpo Format(TotalCantidadT52, "###"), 10, 1, vbRightJustify
        .TextoCuerpo Format(TotalCantidadT54, "###"), 11, 1, vbRightJustify
        .TextoCuerpo Format(TotalCantidadT56, "###"), 12, 1, vbRightJustify
        .TextoCuerpo Format(mobjOrdenCorte.OrdenCorteItems.Cantidad, "###"), 13, 1, vbRightJustify
    End With
    
    'Rellenar las líneas hasta el final
    For i = LineaCuerpo + 1 To mobjTablaCuerpo.NumeroFilasCuerpo
        With mobjTablaCuerpo
            .LineaCeldaDerecha 1, i
            .LineaCeldaDerecha 2, i
            .LineaCeldaDerecha 3, i
            .LineaCeldaDerecha 4, i
            .LineaCeldaDerecha 5, i
            .LineaCeldaDerecha 6, i
            .LineaCeldaDerecha 7, i
            .LineaCeldaDerecha 8, i
            .LineaCeldaDerecha 9, i
            .LineaCeldaDerecha 10, i
            .LineaCeldaDerecha 11, i
            .LineaCeldaDerecha 12, i
            .LineaCeldaDerecha 13, i
            .LineaCeldaDerecha 14, i
        End With
        With mobjTablaCuerpo2
            .LineaCeldaDerecha 1, i
            .LineaCeldaDerecha 2, i
            .LineaCeldaDerecha 3, i
            .LineaCeldaDerecha 4, i
            .LineaCeldaDerecha 5, i
            .LineaCeldaDerecha 6, i
            .LineaCeldaDerecha 7, i
            .LineaCeldaDerecha 8, i
            .LineaCeldaDerecha 9, i
            .LineaCeldaDerecha 10, i
            .LineaCeldaDerecha 11, i
            .LineaCeldaDerecha 12, i
            .LineaCeldaDerecha 13, i
            .LineaCeldaDerecha 14, i
        End With
        
    Next
    
End Sub

Public Sub PrintObject()

    On Error GoTo ErrorManager
    
    If mobjOrdenCorte.IsNew Then _
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


