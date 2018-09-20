VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrintAlbaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Indicador de si se debe imprimir el logotipo.
Private mflgPrintLogo As Boolean

Private mobjAlbaranVenta As AlbaranVenta
Private mobjCliente As Cliente
Private mobjEmpresa As Empresa

Private mobjTablaCuerpo As PrintTabla
Private mobjTablaFecha As PrintTabla
Private mobjTablaTotal As PrintTabla
Private mobjTablaCabecera As PrintTabla
Private mobjTablaEmpresa As PrintTabla
Private mobjTablaObservaciones As PrintTabla

Public Sub Component(AlbaranVentaObject As AlbaranVenta)

    Set mobjAlbaranVenta = AlbaranVentaObject
    
    Set mobjCliente = New Cliente
    mobjCliente.Load mobjAlbaranVenta.ClienteID
    
    Set mobjEmpresa = New Empresa
    mobjEmpresa.Load mobjAlbaranVenta.EmpresaID
    mflgPrintLogo = Not (mobjEmpresa.AnularIVA)

End Sub

Private Sub Class_Initialize()
    
' Establecemos orientacion vertical y tamaño de papel A4
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = vbPRORPortrait
    Printer.Copies = 1

    ' Por defecto imprimimos el logo.
    mflgPrintLogo = True
    
    Set mobjTablaCuerpo = New PrintTabla
    With mobjTablaCuerpo
        .NumeroColumnas = 18
        .Dimensiones 0.4, 8.7, 19.2, 15
        .LongitudColumna 1.7
        .LongitudColumna 1.5
        .LongitudColumna 1.5
        .LongitudColumna 0.6
        .LongitudColumna 0.8
        .LongitudColumna 0.8
        .LongitudColumna 0.8
        .LongitudColumna 0.8
        .LongitudColumna 0.8
        .LongitudColumna 0.8
        .LongitudColumna 0.8
        .LongitudColumna 0.8
        .LongitudColumna 0.8
        .LongitudColumna 0.8
        .LongitudColumna 0.8
        .LongitudColumna 1.25
        .LongitudColumna 1.5
        ' +2.2 +2.2 +.4 = 5
    End With

    Set mobjTablaCabecera = New PrintTabla
    With mobjTablaCabecera
        .Dimensiones 10.3, 4.4, 9.3, 3
        .NumeroColumnas = 1
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = True
    End With
    
    Set mobjTablaFecha = New PrintTabla
    With mobjTablaFecha
        .Dimensiones 0.4, 6.1, 9, 1.3
        .NumeroColumnas = 3
        .Autosize
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = True
    End With
    
    Set mobjTablaTotal = New PrintTabla
    With mobjTablaTotal
        .Dimensiones 12.2, 24, 7.4, 1.4
        .NumeroColumnas = 2
        .LongitudColumna 4
        .Cabecera = False
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = True
        .GrosorLinea = 2
    End With
    
    Set mobjTablaEmpresa = New PrintTabla
    With mobjTablaEmpresa
        .Dimensiones 2, 2.5, 7, 2.5
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = True
        .Cabecera = False
    End With
    
    Set mobjTablaObservaciones = New PrintTabla
    With mobjTablaObservaciones
        .Dimensiones 0.4, 26, 19.2, 2
        .FuenteCuerpoItalic = True
        .Cabecera = False
    End With
    
End Sub

Private Sub Class_Terminate()

    Set mobjAlbaranVenta = Nothing
    Set mobjCliente = Nothing
    Set mobjEmpresa = Nothing
     
    Set mobjTablaCuerpo = Nothing
    Set mobjTablaFecha = Nothing
    Set mobjTablaTotal = Nothing
    Set mobjTablaCabecera = Nothing
    Set mobjTablaEmpresa = Nothing
    Set mobjTablaObservaciones = Nothing
     
    Printer.EndDoc

End Sub

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
    If mflgPrintLogo Then _
        Printer.PaintPicture LoadPicture(App.Path & "\logo.bmp"), _
                             567, 283, 3997, 850
                         
    ImprimirDireccionEmpresa
    ImprimirDireccionCliente
    ImprimirFechaYNumero
    
End Sub

Private Sub ImprimirRejilla()
    
    With mobjTablaCuerpo
    
        .FuenteCuerpoSize = 8.5

        .LineasBorde
        .LineaCabecera

        .TextoCabecera "CONCEPTO", 2
        .LineaColumna 4
        .TextoCabecera "36", 5
        .TextoCabecera "38", 6
        .TextoCabecera "40", 7
        .TextoCabecera "42", 8
        .TextoCabecera "44", 9
        .TextoCabecera "46", 10
        .TextoCabecera "48", 11
        .TextoCabecera "50", 12
        .TextoCabecera "52", 13
        .TextoCabecera "54", 14
        .TextoCabecera "56", 15
        .TextoCabecera "Uds.", 16, vbCenter
        .LineaColumna 16
        .TextoCabecera "PRECIO", 17, vbCenter
        .LineaColumna 17
        .TextoCabecera "IMPORTE", 18, vbCenter
        
    End With
    
End Sub
    
Private Sub ImprimirDireccionEmpresa()

    With mobjTablaEmpresa
        .TextoCuerpo "Teléfono: " & mobjEmpresa.Direccion.Telefono1, 1, 1
        .TextoCuerpo "Fax: " & mobjEmpresa.Direccion.Fax, 1, 2
        .TextoCuerpo mobjEmpresa.Direccion.Calle, 1, 3
        .TextoCuerpo mobjEmpresa.Direccion.CodigoPostal & "-" & mobjEmpresa.Direccion.Poblacion & "  (" & mobjEmpresa.Direccion.Provincia & ")", 1, 4
        .TextoCuerpo mobjEmpresa.Direccion.EMAIL, 1, 5

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
        .TextoCabecera "ALBARÁN", 2, vbCenter
        .TextoCabecera "Nº BULTOS", 3, vbCenter
        .TextoCuerpo Format(mobjAlbaranVenta.Fecha, "dd/mm/yyyy"), 1, 1, vbCenter
        .TextoCuerpo Format(mobjAlbaranVenta.Numero, "###,###,###"), 2, 1, vbCenter
        .TextoCuerpo Format(mobjAlbaranVenta.Bultos, "###,###"), 3, 1, vbCenter
        
    End With

End Sub

Private Sub ImprimirDetalle()
    Dim objAlbaranVentaItem As AlbaranVentaItem
    Dim LineaCuerpo As Integer
    Dim objArticuloColor As ArticuloColor
    
    
    LineaCuerpo = 1
    For Each objAlbaranVentaItem In mobjAlbaranVenta.AlbaranVentaItems
        Set objArticuloColor = objAlbaranVentaItem.objArticuloColor
        
        With mobjTablaCuerpo
            ' OJO: Puerro y de los gordos.
            ' Cambio de criterio de codificación de artículos a partir de la temporada 16 (código 14)
            If mobjAlbaranVenta.TemporadaID < 14 Then
                .TextoCuerpo objArticuloColor.objArticulo.NombrePrenda, 1, LineaCuerpo
                .TextoCuerpo objArticuloColor.objArticulo.NombreModelo, 2, LineaCuerpo
                .TextoCuerpo objArticuloColor.objArticulo.NombreSerie, 3, LineaCuerpo
                .TextoCuerpo objArticuloColor.Codigo, 4, LineaCuerpo
            Else
                .TextoCuerpo objArticuloColor.objArticulo.NombrePrenda, 1, LineaCuerpo
                .TextoCuerpo objArticuloColor.objArticulo.objModelo.Codigo, 2, LineaCuerpo
                .TextoCuerpo objArticuloColor.objArticulo.NombreSerie, 3, LineaCuerpo
                .TextoCuerpo objArticuloColor.Codigo, 4, LineaCuerpo
            End If
            .TextoCuerpo objAlbaranVentaItem.CantidadT36, 5, LineaCuerpo
            .TextoCuerpo objAlbaranVentaItem.CantidadT38, 6, LineaCuerpo
            .TextoCuerpo objAlbaranVentaItem.CantidadT40, 7, LineaCuerpo
            .TextoCuerpo objAlbaranVentaItem.CantidadT42, 8, LineaCuerpo
            .TextoCuerpo objAlbaranVentaItem.CantidadT44, 9, LineaCuerpo
            .TextoCuerpo objAlbaranVentaItem.CantidadT46, 10, LineaCuerpo
            .TextoCuerpo objAlbaranVentaItem.CantidadT48, 11, LineaCuerpo
            .TextoCuerpo objAlbaranVentaItem.CantidadT50, 12, LineaCuerpo
            .TextoCuerpo objAlbaranVentaItem.CantidadT52, 13, LineaCuerpo
            .TextoCuerpo objAlbaranVentaItem.CantidadT54, 14, LineaCuerpo
            .TextoCuerpo objAlbaranVentaItem.CantidadT56, 15, LineaCuerpo
            .TextoCuerpo objAlbaranVentaItem.Cantidad, 16, LineaCuerpo, vbRightJustify
            .TextoCuerpo FormatoMoneda(objAlbaranVentaItem.PrecioVenta, "EUR", False), 17, LineaCuerpo, vbRightJustify
            .TextoCuerpo FormatoMoneda(objAlbaranVentaItem.BRUTO, "EUR", False), 18, LineaCuerpo, vbRightJustify
            
        End With
        
        Set objArticuloColor = Nothing
        
        LineaCuerpo = LineaCuerpo + 1
        If LineaCuerpo > mobjTablaCuerpo.NumeroFilasCuerpo Then
            ImprimirNuevaPagina
            LineaCuerpo = 1
        End If
        
    Next
    With mobjTablaObservaciones
        .LineasBorde
        .TextoCuerpo "Observaciones: " & mobjAlbaranVenta.Observaciones, 1, 1
    End With
    
    With mobjTablaTotal
        .LineasBorde
        '.LineaTodasColumnas
        '.LineaFilaArriba 2
        '.LineaCelda 2, 1
        .TextoCuerpo "Total Prendas >>", 1, 1, vbCenter
        .TextoCuerpo mobjAlbaranVenta.AlbaranVentaItems.Cantidad, 2, 1, vbCenter
        '.LineaCelda 2, 2
        .TextoCuerpo "TOTAL Albaran >>", 1, 2, vbCenter
        .TextoCuerpo FormatoMoneda(mobjAlbaranVenta.TotalBruto, "EUR"), 2, 2, vbCenter
        
    End With

End Sub

Public Sub PrintObject()

    If mobjAlbaranVenta.IsNew Then _
        Err.Raise vbObjectError + 1001, "No se puede imprimir, no existe la Albaran."
    
    ImprimirDatosFijos
    ImprimirDetalle
        
End Sub

Public Property Let PrinterNumber(PrinterNumber As Integer)
    
    Set Printer = Printers(PrinterNumber)
    
End Property

