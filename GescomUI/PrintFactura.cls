VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrintFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Indicador de si se debe imprimir el logotipo.
Private mflgPrintLogo As Boolean

Private mobjFacturaVenta As FacturaVenta
Private mobjCliente As Cliente
Private mobjEmpresa As Empresa

Private mobjTablaFecha As PrintTabla
Private mobjTablaDNICIF As PrintTabla
Private mobjTablaCabecera As PrintTabla
Private mobjTablaCuerpo As PrintTabla
Private mobjTablaCuerpoLibre As PrintTabla
Private mobjTablaEmpresa As PrintTabla
Private mobjTablaTotal As PrintTabla
Private mobjTablaImporteEquivalente As PrintTabla
Private mobjTablaFormaPago As PrintTabla
Private mobjTablaPie As PrintTabla
Private mobjTablaDNINIF As PrintTabla

Public Sub Component(FacturaVentaObject As FacturaVenta)

    Set mobjFacturaVenta = FacturaVentaObject
    
    Set mobjCliente = New Cliente
    mobjCliente.Load mobjFacturaVenta.ClienteID
    
    Set mobjEmpresa = New Empresa
    mobjEmpresa.Load mobjFacturaVenta.EmpresaID
    mflgPrintLogo = Not (mobjEmpresa.AnularIVA)

End Sub

Private Sub Class_Initialize()
    
' Establecemos orientacion vertical y tamaño de papel A4
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = vbPRORPortrait
    Printer.Copies = 1

    ' Por defecto imprimimos el logo.
    mflgPrintLogo = True
    
    Set mobjTablaCabecera = New PrintTabla
    With mobjTablaCabecera
        .Dimensiones 10.3, 4.4, 9.3, 3
        .NumeroColumnas = 1
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = True
    End With
    
    Set mobjTablaFecha = New PrintTabla
    With mobjTablaFecha
        .Dimensiones 0.4, 6.1, 7.3, 1.3
        .NumeroColumnas = 2
        .Autosize
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = True
    End With
    
    Set mobjTablaDNICIF = New PrintTabla
    With mobjTablaDNICIF
        .Dimensiones 0.4, 7.5, 5.5, 0.7
        .NumeroColumnas = 2
        .Cabecera = False
        .LongitudColumna 2.5
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = True
    End With
    
    Set mobjTablaCuerpo = New PrintTabla
    With mobjTablaCuerpo
        .Dimensiones 0.4, 8.7, 19.2, 10.3
        .NumeroColumnas = 7
        .Cabecera = True
        .LongitudColumna 2
        .LongitudColumna 4.6
        .LongitudColumna 4.6
        .LongitudColumna 1
        .LongitudColumna 2
        .LongitudColumna 2
        .FuenteCuerpoBold = False
        .FuenteCuerpoItalic = False
    End With
    
    Set mobjTablaCuerpoLibre = New PrintTabla
    With mobjTablaCuerpoLibre
        .Dimensiones 0.4, 8.7, 12.6, 10.3
        .NumeroColumnas = 1
        .Cabecera = True
        .FuenteCuerpoBold = False
        .FuenteCuerpoItalic = False
    End With
    
    Set mobjTablaEmpresa = New PrintTabla
    With mobjTablaEmpresa
        .Dimensiones 2, 2.5, 7, 2
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = True
        .Cabecera = False
    End With
    
    Set mobjTablaTotal = New PrintTabla
    With mobjTablaTotal
        .Dimensiones 12.3, 23.2, 7.3, 0.8
        .NumeroColumnas = 2
        .Cabecera = False
        .LongitudColumna 4.5
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = True
    End With
    
    Set mobjTablaImporteEquivalente = New PrintTabla
    With mobjTablaImporteEquivalente
        .Dimensiones 12.3, 24.1, 7.3, 0.8
        .NumeroColumnas = 2
        .Cabecera = False
        .LongitudColumna 4.5
        .FuenteCuerpoBold = False
        .FuenteCuerpoItalic = False
    End With
    
    Set mobjTablaFormaPago = New PrintTabla
    With mobjTablaFormaPago
        .Dimensiones 0.4, 23.2, 11, 4
        .NumeroColumnas = 2
        .Cabecera = False
        .LongitudColumna 3
        .FuenteCuerpoBold = False
        .FuenteCuerpoItalic = False
    End With
    
    Set mobjTablaPie = New PrintTabla
    With mobjTablaPie
        .Dimensiones 0.4, 19, 19.2, 3
        .NumeroColumnas = 5
        .Cabecera = False
        .LongitudColumna 6
        .LongitudColumna 3
        .LongitudColumna 3.2
        .LongitudColumna 4
        .FuenteCuerpoBold = False
        .FuenteCuerpoItalic = False
    End With
    
    Set mobjTablaDNINIF = New PrintTabla
    With mobjTablaDNINIF
        .Dimensiones 0.4, 28, 19, 0.6
        .Cabecera = False
        .Autosize
        .FuenteCuerpoBold = False
        .FuenteCuerpoItalic = False
        .FuenteCuerpoSize = 6
    End With
    
End Sub

Private Sub Class_Terminate()

    Set mobjFacturaVenta = Nothing
    Set mobjCliente = Nothing
    Set mobjEmpresa = Nothing
     
    Set mobjTablaFecha = Nothing
    Set mobjTablaCabecera = Nothing
    Set mobjTablaDNICIF = Nothing
    Set mobjTablaCuerpo = Nothing
    Set mobjTablaCuerpoLibre = Nothing
    Set mobjTablaEmpresa = Nothing
    Set mobjTablaTotal = Nothing
    Set mobjTablaImporteEquivalente = Nothing
    Set mobjTablaFormaPago = Nothing
    Set mobjTablaPie = Nothing
    Set mobjTablaDNINIF = Nothing
    
    Printer.EndDoc

End Sub

Public Property Let PrinterNumber(PrinterNumber As Integer)
    
    Set Printer = Printers(PrinterNumber)
    
End Property

Public Property Let Copies(Copies As Integer)
    
    Printer.Copies = Copies
    
End Property

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
    If mflgPrintLogo Then
        Printer.PaintPicture LoadPicture(App.Path & "\logo.bmp"), _
                             567, 283, 3997, 850
                             
        ImprimirDireccionEmpresa
    End If
    ImprimirDireccionCliente
    ImprimirFechaYNumero
    
End Sub

Private Sub ImprimirRejilla()
    
    With mobjTablaCuerpo
        .LineaCabecera
        .LineasBorde
        '.LineaTodasColumnas
        .LineaColumna 6
        .LineaColumna 5
        .LineaColumna 4
        .LineaCelda 7, .NumeroFilasCuerpo
        .TextoCabecera "CONCEPTO", 2, vbCenter
        .TextoCabecera "UNIDADES", 5, vbCenter
        .TextoCabecera "PRECIO", 6, vbCenter
        .TextoCabecera "IMPORTE", 7, vbCenter
        .TextoCuerpo "SUMA >>", 6, .NumeroFilasCuerpo, vbCenter
    End With
    
    ' Si la empresa tiene la opcion anular iva, no se imprime la información del NIF
    If Not mobjEmpresa.AnularIVA Then
        mobjTablaDNINIF.TextoCuerpo "HONGO, S.A. - Inscrita en el Registro Mercantil de Guipúzcoa, Libro de Sociedades, " & _
                                    "Tomo 682, Folio 40, Hoja 8349, Inscripción 1ª N.I.F.: A-20096202", _
                                    1, 1, vbCenter
    End If
                                
    
End Sub


Private Sub ImprimirDireccionEmpresa()

    With mobjTablaEmpresa
        .TextoCuerpo "Teléfono: " & mobjEmpresa.Direccion.Telefono1, 1, 1
        .TextoCuerpo "Fax: " & mobjEmpresa.Direccion.Fax, 1, 2
        .TextoCuerpo mobjEmpresa.Direccion.Calle, 1, 3
        .TextoCuerpo mobjEmpresa.Direccion.CodigoPostal & "-" & mobjEmpresa.Direccion.Poblacion & "  (" & mobjEmpresa.Direccion.Provincia & ")", 1, 4

    End With

End Sub
Private Sub ImprimirDireccionCliente()

    With mobjTablaCabecera
        .LineasBorde
        
        .TextoCuerpo mobjCliente.Nombre, 1, 1
        .TextoCuerpo mobjCliente.Titular, 1, 2
        .TextoCuerpo mobjCliente.DireccionFiscal.Calle, 1, 3
        .TextoCuerpo mobjCliente.DireccionFiscal.CodigoPostal & "-" & _
                  mobjCliente.DireccionFiscal.Poblacion & "  (" & _
                  mobjCliente.DireccionFiscal.Provincia & ")", 1, 4
    End With
    
End Sub

Private Sub ImprimirFechaYNumero()

    With mobjTablaFecha
        .LineasBorde
        .LineaCabecera
        .LineaTodasColumnas
        .TextoCabecera "FECHA", 1, vbCenter
        .TextoCabecera "Nº FACTURA", 2, vbCenter
        .TextoCuerpo Format(mobjFacturaVenta.Fecha, "dd/mm/yyyy"), 1, 1, vbCenter
        .TextoCuerpo Format(mobjFacturaVenta.Numero, "###,###,###"), 2, 1, vbCenter
        
    End With
    
    If Not mobjEmpresa.AnularIVA Then
        With mobjTablaDNICIF
            .TextoCuerpo "DNI/CIF :", 1, 1, vbCenter
            .TextoCuerpo mobjCliente.DNINIF, 2, 1, vbCenter
        End With
    End If
    
End Sub

Private Sub ImprimirDetalle()
    Dim objFacturaVentaItem As FacturaVentaItem
    Dim objAlbaranVentaItem As AlbaranVentaItem
    Dim objAlbaranVenta As AlbaranVenta
    Dim LineaCuerpo As Integer
    Dim lngAlbaranVentaID As Long
    Dim objArticuloColor As ArticuloColor
    
    lngAlbaranVentaID = 0
    LineaCuerpo = 1
    For Each objFacturaVentaItem In mobjFacturaVenta.FacturaVentaItems
        If LineaCuerpo + 2 > mobjTablaCuerpo.NumeroFilasCuerpo Then
            ImprimirNuevaPagina
            LineaCuerpo = 1
        End If
        
        ' Si hay albarán se obtiene el número de cabecera
        If objFacturaVentaItem.HayAlbaran Then
            
            Set objAlbaranVentaItem = New AlbaranVentaItem
            ' Leemos la linea de albaran
            objAlbaranVentaItem.Load objFacturaVentaItem.AlbaranVentaItemID
            ' Si el numero de cabecera de albaran es el mismo, no hacemos nada,
            ' sino se lee la cabecera y se escribe el número de documento.
            If lngAlbaranVentaID <> objAlbaranVentaItem.AlbaranVentaID Then
                lngAlbaranVentaID = objAlbaranVentaItem.AlbaranVentaID
                Set objAlbaranVenta = New AlbaranVenta
                objAlbaranVenta.Load objAlbaranVentaItem.AlbaranVentaID
                
                mobjTablaCuerpo.TextoCuerpo "--- Albarán nº " & _
                                            FormatoCantidad(objAlbaranVenta.Numero), 2, LineaCuerpo, vbLeftJustify
                                            
            
                LineaCuerpo = LineaCuerpo + 1
                Set objAlbaranVenta = Nothing
            End If
            
            Set objAlbaranVentaItem = Nothing
            
        End If
        
        Set objArticuloColor = objFacturaVentaItem.objArticuloColor

        With mobjTablaCuerpo
            If objArticuloColor Is Nothing Then
                mobjTablaCuerpoLibre.TextoCuerpo objFacturaVentaItem.Descripcion, 1, LineaCuerpo
            Else
                ' OJO: Puerro y de los gordos.
                ' Cambio de criterio de codificación de artículos a partir de la temporada 16 (código 14)
                If mobjFacturaVenta.TemporadaID < 14 Then
                    .TextoCuerpo objArticuloColor.objArticulo.objPrenda.Nombre, 1, LineaCuerpo
                    .TextoCuerpo objArticuloColor.objArticulo.NombreModelo, 2, LineaCuerpo
                    .TextoCuerpo objArticuloColor.objArticulo.NombreSerie, 3, LineaCuerpo
                    .TextoCuerpo objArticuloColor.Codigo, 4, LineaCuerpo
                Else
                    .TextoCuerpo objArticuloColor.objArticulo.objPrenda.Nombre, 1, LineaCuerpo
                    .TextoCuerpo objArticuloColor.objArticulo.objModelo.Codigo, 2, LineaCuerpo
                    .TextoCuerpo objArticuloColor.objArticulo.NombreSerie, 3, LineaCuerpo
                    .TextoCuerpo objArticuloColor.Codigo, 4, LineaCuerpo
                End If
            End If
            .TextoCuerpo objFacturaVentaItem.Cantidad, 5, LineaCuerpo, vbRightJustify
            .TextoCuerpo FormatoMoneda(objFacturaVentaItem.PrecioVenta, _
                                "EUR", False), 6, LineaCuerpo, vbRightJustify
            .TextoCuerpo FormatoMonedaDocumento(objFacturaVentaItem.BRUTO, _
                                "EUR", False), 7, LineaCuerpo, vbRightJustify
            
            LineaCuerpo = LineaCuerpo + 1
        
        End With
        
        Set objArticuloColor = Nothing


    Next
    
    mobjTablaCuerpo.TextoCuerpo FormatoMonedaDocumento(mobjFacturaVenta.BRUTO, _
                                "EUR", False), 7, _
                                mobjTablaCuerpo.NumeroFilasCuerpo, vbRightJustify
    
End Sub
Private Sub ImprimirPie()
    Dim LineaPie As Integer

    LineaPie = 1
    With mobjTablaPie
        .LineasBorde
        .LineaColumna 3
        .LineaColumna 4
        If mobjFacturaVenta.Descuento <> 0 Then
            .TextoCuerpo "Descuento " & Format(mobjFacturaVenta.DatoComercial.Descuento, "#0.00") & "% :", _
                        4, LineaPie, vbLeftJustify
            .TextoCuerpo FormatoMonedaDocumento(mobjFacturaVenta.Descuento, "EUR"), _
                        5, LineaPie, vbRightJustify
            LineaPie = LineaPie + 1
        End If
        
        If mobjFacturaVenta.Portes <> 0 Then
            .TextoCuerpo "Portes: ", 4, LineaPie, vbLeftJustify
            .TextoCuerpo FormatoMonedaDocumento(mobjFacturaVenta.Portes, "EUR"), _
                        5, LineaPie, vbRightJustify
            LineaPie = LineaPie + 1
        End If
        If mobjFacturaVenta.Embalajes <> 0 Then
            .TextoCuerpo "Embalajes: ", 4, LineaPie, vbLeftJustify
            .TextoCuerpo FormatoMonedaDocumento(mobjFacturaVenta.Embalajes, "EUR"), _
                        5, LineaPie, vbRightJustify
            LineaPie = LineaPie + 1
        End If
        If Not mobjEmpresa.AnularIVA Then
            .TextoCuerpo "Base Imponible: ", 1, LineaPie, vbLeftJustify
            .TextoCuerpo FormatoMonedaDocumento(mobjFacturaVenta.BaseImponible, "EUR"), _
                        2, LineaPie, vbRightJustify
        End If
        If (Not mobjEmpresa.AnularIVA) And mobjFacturaVenta.IVA <> 0 Then
            .TextoCuerpo "I.V.A. " & Format(mobjFacturaVenta.DatoComercial.IVA, "#0.00") & "% : ", _
                        4, LineaPie, vbLeftJustify
            .TextoCuerpo FormatoMonedaDocumento(mobjFacturaVenta.IVA, "EUR"), _
                        5, LineaPie, vbRightJustify
            LineaPie = LineaPie + 1
        End If
        If mobjFacturaVenta.Recargo <> 0 Then
            .TextoCuerpo "Recar.equiv." & Format(mobjFacturaVenta.DatoComercial.RecargoEquivalencia, "#0.00") & "% : ", _
                        4, LineaPie, vbLeftJustify
            .TextoCuerpo FormatoMonedaDocumento(mobjFacturaVenta.Recargo, "EUR"), _
                        5, LineaPie, vbRightJustify
            LineaPie = LineaPie + 1
        End If
    End With

    With mobjTablaTotal
        .LineaCelda 2, 1
        .TextoCuerpo "TOTAL FACTURA   >>", 1, 1, vbCenter
        .TextoCuerpo FormatoMonedaDocumento(mobjFacturaVenta.Neto, "EUR"), 2, 1, vbRightJustify
    End With

    Dim ImporteOtro As Double
    Dim MonedaOtro As String

    ImporteOtro = EUR2PTA(mobjFacturaVenta.Neto)
    MonedaOtro = "PTA"

    With mobjTablaImporteEquivalente
        .TextoCuerpo "Importe equivalente", 1, 1, vbCenter
        .TextoCuerpo FormatoMonedaDocumento(ImporteOtro, MonedaOtro), 2, 1, vbRightJustify
    End With
    
    ' Sólamente imprimir la forma de pago para facturas normales, en el caso
    ' de facturas de abono no se debe imprimir.
    If mobjFacturaVenta.Neto >= 0 Then _
        ImprimirFormaPago
    
    
End Sub

Private Sub ImprimirFormaPago()
    Dim objCobroPago As CobroPago
    Dim LineaFormaPago As Integer
    Dim objFormaPago As formadepago
    
    With mobjTablaFormaPago
        .TextoCuerpo "Forma de pago:", 1, 1, vbLeftJustify
        .TextoCuerpo mobjFacturaVenta.FormaPago, 2, 1, vbLeftJustify
        
        LineaFormaPago = 2
        For Each objCobroPago In mobjFacturaVenta.CobrosPagos
            .TextoCuerpo Format(objCobroPago.Vencimiento, "dd/mm/yyyy") & "    " & _
                        FormatoMonedaDocumento(objCobroPago.Importe, "EUR"), 2, LineaFormaPago, vbLeftJustify
                        
            LineaFormaPago = LineaFormaPago + 1
            If LineaFormaPago > 4 Then
                Exit For
            End If
            
        Next
        ' Si la empresa tiene la opcion anular iva, no se imprime la cuenta bancaria.
        If Not mobjEmpresa.AnularIVA Then
            Set objFormaPago = New formadepago
            objFormaPago.Load mobjFacturaVenta.FormaPagoID
            If Not objFormaPago.Contado Then
                .TextoCuerpo mobjCliente.CuentaBancaria.Entidad & "." & _
                             mobjCliente.CuentaBancaria.Sucursal & "." & _
                             mobjCliente.CuentaBancaria.Control & "." & _
                             mobjCliente.CuentaBancaria.Cuenta, _
                            2, LineaFormaPago, vbLeftJustify
            End If
            Set objFormaPago = Nothing
        End If
    End With
    
End Sub

Public Sub PrintObject()

    If mobjFacturaVenta.IsNew Then _
        Err.Raise vbObjectError + 1001, "No se puede imprimir, no existe la factura."
    
    ImprimirDatosFijos
    ImprimirDetalle
    ImprimirPie
        
End Sub
