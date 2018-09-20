VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrintNecesidadesMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjNecesidadesMaterial As NecesidadesMaterial
Private mrsRecordset As ADOR.Recordset

Private mobjTablaCuerpo As PrintTabla
Private mobjTablaTotal As PrintTabla
Private mobjTablaTitulo As PrintTabla
Private mobjTablaSerie As PrintTabla
Private mobjTablaPie As PrintTabla
Private mdblTotalImporte As Double
Private mdblTotalCantidad As Double

Public Sub Component(NecesidadesMaterialObject As NecesidadesMaterial)

    Set mobjNecesidadesMaterial = NecesidadesMaterialObject
    Set mrsRecordset = NecesidadesMaterialObject.NecesidadesMaterialRecordset
    
End Sub

Private Sub Class_Initialize()
    
' Establecemos orientacion vertical y tamaño de papel A4
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = vbPRORPortrait
    Printer.Copies = 1

    Set mobjTablaCuerpo = New PrintTabla
    With mobjTablaCuerpo
        .NumeroColumnas = 8
        .Dimensiones 0.4, 6.5, 19.2, 15.4
        .LongitudColumna 2
        .LongitudColumna 2
        .LongitudColumna 2
        .LongitudColumna 5
        .LongitudColumna 1.8
        .LongitudColumna 2
        .LongitudColumna 2
        '.LongitudColumna 2
        .FuenteCuerpoSize = 9
        .FuenteCabeceraSize = 9
    End With

    Set mobjTablaTotal = New PrintTabla
    With mobjTablaTotal
        .Dimensiones 9.1, 23.2, 10.5, 0.8
        .NumeroColumnas = 3
        .LongitudColumna 4
        .LongitudColumna 3
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
        .LongitudColumna 3
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
    
End Sub

Private Sub Class_Terminate()

    Set mobjNecesidadesMaterial = Nothing
    Set mrsRecordset = Nothing
    
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
    
    mobjTablaTitulo.TextoCuerpo "NECESIDADES DE MATERIAL", 1, 1, vbCenter

End Sub

Private Sub ImprimirRejilla()
    
    With mobjTablaCuerpo
        .LineaTodasColumnas
        .LineasBorde
        .LineaCabecera
        .TextoCabecera "Serie", 1
        .TextoCabecera "Color", 2
        .TextoCabecera "Código", 3
        .TextoCabecera "Descripción", 4
        .TextoCabecera "Stock", 5, vbRightJustify
        .TextoCabecera "Necesidad", 6, vbRightJustify
        .TextoCabecera "Precio", 7, vbRightJustify
        .TextoCabecera "Importe", 8, vbRightJustify
        '.TextoCabecera "Proveedor", 6, vbCenter
    End With
    
End Sub

Private Sub ImprimirSerie()
    Dim strSituacion As String
    Dim strPedidos As String

    Select Case mobjNecesidadesMaterial.Situacion
    Case "T"
        strSituacion = "Todos los pedidos"
    Case "P"
        strSituacion = "Pendientes de servir"
    Case "S"
        strSituacion = "Pedidos servidos"
    Case Else
        Err.Raise vbObjectError + 1001, "Impresion de necesidades de material", "Situacion de impresion de pedidos erronea: " & mobjNecesidadesMaterial.Situacion
    End Select
    
    If mobjNecesidadesMaterial.PedidoInicial = 0 And mobjNecesidadesMaterial.PedidoFinal = 0 Then
        strPedidos = "Todos los pedidos"
    Else
        strPedidos = "Pedidos entre el " & mobjNecesidadesMaterial.PedidoInicial & " y el " & mobjNecesidadesMaterial.PedidoFinal
    End If
    
    With mobjTablaSerie
        .LineasBorde
        
        .TextoCuerpo "Materiales:", 1, 1
        .TextoCuerpo IIf(mobjNecesidadesMaterial.Tipo = "T", "Telas", "Otros materiales"), 2, 1
        .TextoCuerpo "Situación: ", 1, 2
        .TextoCuerpo strSituacion, 2, 2
        .TextoCuerpo "Pedidos: ", 1, 3
        .TextoCuerpo strPedidos, 2, 3
        
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
    Dim dblNecesidad As Double
    Dim dblImporte As Double
    Dim strSerie As String
    
    mdblTotalImporte = 0
    mdblTotalCantidad = 0
    strSerie = mrsRecordset("NombreSerie")
    LineaCuerpo = 1
    mrsRecordset.MoveFirst
    
    ImprimirDatosFijos
    While Not mrsRecordset.EOF
        If strSerie <> mrsRecordset("NombreSerie") Then
           LineaCuerpo = LineaCuerpo + 1
        End If
    
        With mobjTablaCuerpo
            .TextoCuerpo mrsRecordset("NombreSerie") & "", 1, LineaCuerpo
            .TextoCuerpo mrsRecordset("NombreColor") & "", 2, LineaCuerpo
            .TextoCuerpo mrsRecordset("Codigo") & "", 3, LineaCuerpo
            .TextoCuerpo mrsRecordset("Nombre") & "", 4, LineaCuerpo
            .TextoCuerpo FormatoCantidad(mrsRecordset("StockActual")), 5, LineaCuerpo, vbRightJustify
            Select Case mobjNecesidadesMaterial.Situacion
            Case "T"
                dblNecesidad = mrsRecordset("Cantidad")
            Case "P"
                dblNecesidad = mrsRecordset("Cantidad") - mrsRecordset("Servido")
            Case "S"
                dblNecesidad = mrsRecordset("Servido")
            Case Else
                Err.Raise vbObjectError + 1001, "Impresion de necesidades de material", "Situacion de impresion de pedidos erronea: " & mobjNecesidadesMaterial.Situacion
            End Select
            ' Si son telas --> se tiene en cuenta la anchura de la tela.
            If mobjNecesidadesMaterial.Tipo = "T" Then
                If mrsRecordset("AnchuraTela") = 0 Then
                    dblNecesidad = 0
                Else
                    dblNecesidad = (1.5 * dblNecesidad) / mrsRecordset("AnchuraTela")
                End If
            End If
            
            .TextoCuerpo FormatoCantidad(dblNecesidad), 6, LineaCuerpo, vbRightJustify
            dblImporte = dblNecesidad * mrsRecordset("PrecioCosteEUR")
            mdblTotalImporte = mdblTotalImporte + dblImporte
            mdblTotalCantidad = mdblTotalCantidad + dblNecesidad
            
            .TextoCuerpo FormatoMoneda(mrsRecordset("PrecioCosteEUR"), "EUR", True), 7, LineaCuerpo, vbRightJustify
            .TextoCuerpo FormatoMoneda(dblImporte, "EUR", False), 8, LineaCuerpo, vbRightJustify
            
            ' El proveedor lo dejamos para más tarde
            '.TextoCuerpo mrsRecordset("NombreProveedor") & vbNullString, 6, LineaCuerpo
        End With
        
        
        LineaCuerpo = LineaCuerpo + 1
        If LineaCuerpo >= mobjTablaCuerpo.NumeroFilasCuerpo Then
            ImprimirNuevaPagina
            LineaCuerpo = 1
        End If
        
        strSerie = mrsRecordset("NombreSerie")
        mrsRecordset.MoveNext
        
    Wend
    
    ImprimirTotal
    
End Sub

Private Sub ImprimirTotal()

    With mobjTablaTotal
        .LineaCelda 2, 1
        .LineaCelda 3, 1
        .TextoCuerpo "TOTALES>>", 1, 1, vbCenter
        .TextoCuerpo FormatoCantidad(mdblTotalCantidad), 2, 1, vbCenter
        .TextoCuerpo FormatoMoneda(mdblTotalImporte, "EUR"), 3, 1, vbCenter
        
    End With
    
End Sub


Public Sub PrintObject()

    On Error GoTo ErrorManager
    
    If mrsRecordset.EOF Then _
        Err.Raise vbObjectError + 1001, "Impresion de necesidades de material", _
                                        "No existen necesidades de material para los pedidos seleccionados."
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


