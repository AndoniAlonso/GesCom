VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrintFichaArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const TwipsPerCm = 567

Private mobjArticulo As Articulo
Private mobjMaterialSerie As Material
Private mobjModelo As Modelo
Private mobjPrenda As Prenda

Private mobjTablaCuerpo As PrintTabla
Private mobjTablaTitulo As PrintTabla
Private mobjTablaArticulo As PrintTabla
Private mobjTablaPie As PrintTabla
Private mlngTotalColor As Long

Public Sub Component(objArticulo As Articulo)

    Set mobjArticulo = objArticulo
    Set mobjMaterialSerie = mobjArticulo.objSerie.objMaterial("EUR")
    Set mobjModelo = mobjArticulo.objModelo
    Set mobjPrenda = mobjArticulo.objPrenda
    
End Sub

Private Sub Class_Initialize()
    
' Establecemos orientacion vertical y tamaño de papel A4
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = vbPRORPortrait
    Printer.Copies = 1

    Set mobjTablaCuerpo = New PrintTabla
    With mobjTablaCuerpo
        .NumeroColumnas = 7
        .Dimensiones 0.4, 6.5, 19.2, 15.4
        .LongitudColumna 4
        .LongitudColumna 3.5
        .LongitudColumna 2
        .LongitudColumna 2
        .LongitudColumna 2
        .LongitudColumna 2
        .Cabecera = False
    End With

    Set mobjTablaTitulo = New PrintTabla
    With mobjTablaTitulo
        .Dimensiones 0.4, 4, 10, 1.4
        .Cabecera = False
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = False
        .FuenteCuerpoSize = 16
    End With
    
    Set mobjTablaArticulo = New PrintTabla
    With mobjTablaArticulo
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

    Set mobjArticulo = Nothing
    Set mobjMaterialSerie = Nothing
    Set mobjModelo = Nothing
    Set mobjPrenda = Nothing
    
    Set mobjTablaCuerpo = Nothing
    Set mobjTablaTitulo = Nothing
    Set mobjTablaArticulo = Nothing
    Set mobjTablaPie = Nothing
     
    Printer.EndDoc

End Sub

'Private Sub ImprimirNuevaPagina()
'
'    Printer.NewPage
'    ImprimirDatosFijos
'
'End Sub
'
Private Sub ImprimirDatosFijos()
    
    ImprimirCabecera
    ImprimirRejilla
    ImprimirPie
    
End Sub

Private Sub ImprimirCabecera()
    
    ' Imprimimos el logotipo
    Printer.PaintPicture LoadPicture(App.Path & "\logo.bmp"), _
                         1 * TwipsPerCm, 2 * TwipsPerCm, 7 * TwipsPerCm, 1.5 * TwipsPerCm

    ImprimirArticulo
    
    mobjTablaTitulo.TextoCuerpo "FICHA DE ARTICULO", 1, 1, vbCenter

End Sub

Private Sub ImprimirRejilla()
    
    mobjTablaCuerpo.LineasBorde
    
End Sub

Private Sub ImprimirArticulo()

    With mobjTablaArticulo
        .LineasBorde

        .TextoCuerpo "Prenda: ", 1, 1
        .TextoCuerpo mobjArticulo.NombrePrenda, 2, 1
        .TextoCuerpo "Modelo: ", 1, 2
        .TextoCuerpo mobjArticulo.NombreModelo, 2, 2
        .TextoCuerpo "Serie: ", 1, 3
        .TextoCuerpo mobjArticulo.NombreSerie, 2, 3

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
    Dim dblCostesMaterial As Double
    Dim dblCostesTela As Double
    Dim dblCostesOtrosMateriales As Double
    Dim dblCostesSubtotal As Double
    Dim dblCostesAdministracion As Double
    Dim objEstrModelo As EstrModelo
    
    ' Calculo del coste de tela
    dblCostesTela = (mobjMaterialSerie.PrecioCoste * _
                     mobjModelo.CantidadTela * _
                     mobjMaterialSerie.AnchuraEstandar) / _
                     mobjMaterialSerie.AnchuraTela
    dblCostesTela = Round(dblCostesTela, 2)

    dblCostesOtrosMateriales = mobjModelo.EstrModelos.Precio + _
                               mobjPrenda.Percha + _
                               mobjPrenda.Etiqueta

    dblCostesMaterial = dblCostesTela + dblCostesOtrosMateriales
    dblCostesSubtotal = dblCostesMaterial + _
                        mobjModelo.Taller + _
                        mobjModelo.Corte + _
                        mobjPrenda.Plancha + _
                        mobjPrenda.Transporte + _
                        mobjPrenda.Carton
    dblCostesAdministracion = Round(dblCostesSubtotal * mobjPrenda.administracion / 100, 2)


    LineaCuerpo = 1
    With mobjTablaCuerpo
        .LineaCelda 5, LineaCuerpo
        .TextoCuerpo "COSTES MATERIAL", 1, LineaCuerpo, vbLeftJustify
        .TextoCuerpo FormatoMoneda(dblCostesMaterial, "EUR"), 5, LineaCuerpo, vbRightJustify
        LineaCuerpo = LineaCuerpo + 1
        
        .TextoCuerpo "Tela ", 2, LineaCuerpo, vbLeftJustify
        .TextoCuerpo FormatoMoneda(dblCostesTela, "EUR"), 4, LineaCuerpo, vbRightJustify
        LineaCuerpo = LineaCuerpo + 1
        
        .TextoCuerpo "Otros materiales", 2, LineaCuerpo, vbLeftJustify
        .TextoCuerpo FormatoMoneda(dblCostesOtrosMateriales, "EUR"), 4, LineaCuerpo, vbRightJustify
        LineaCuerpo = LineaCuerpo + 1
        
        For Each objEstrModelo In mobjModelo.EstrModelos
            .TextoCuerpo "-    " & objEstrModelo.Material, 2, LineaCuerpo, vbLeftJustify
            .TextoCuerpo FormatoMoneda(objEstrModelo.Precio, "EUR"), 3, LineaCuerpo, vbRightJustify
            LineaCuerpo = LineaCuerpo + 1
        Next
        .TextoCuerpo "-    Percha", 2, LineaCuerpo, vbLeftJustify
        .TextoCuerpo FormatoMoneda(mobjPrenda.Percha, "EUR"), 3, LineaCuerpo, vbRightJustify
        LineaCuerpo = LineaCuerpo + 1
        
        .TextoCuerpo "-    Etiqueta", 2, LineaCuerpo, vbLeftJustify
        .TextoCuerpo FormatoMoneda(mobjPrenda.Etiqueta, "EUR"), 3, LineaCuerpo, vbRightJustify
        LineaCuerpo = LineaCuerpo + 1
        
        .LineaCelda 5, LineaCuerpo
        .TextoCuerpo "TALLER", 1, LineaCuerpo, vbLeftJustify
        .TextoCuerpo FormatoMoneda(mobjModelo.Taller, "EUR"), 5, LineaCuerpo, vbRightJustify
        LineaCuerpo = LineaCuerpo + 1
        
        .LineaCelda 5, LineaCuerpo
        .TextoCuerpo "CORTE", 1, LineaCuerpo, vbLeftJustify
        .TextoCuerpo FormatoMoneda(mobjModelo.Corte, "EUR"), 5, LineaCuerpo, vbRightJustify
        LineaCuerpo = LineaCuerpo + 1
        
        .LineaCelda 5, LineaCuerpo
        .TextoCuerpo "PLANCHA", 1, LineaCuerpo, vbLeftJustify
        .TextoCuerpo FormatoMoneda(mobjPrenda.Plancha, "EUR"), 5, LineaCuerpo, vbRightJustify
        LineaCuerpo = LineaCuerpo + 1

        .LineaCelda 5, LineaCuerpo
        .TextoCuerpo "TRANSPORTE", 1, LineaCuerpo, vbLeftJustify
        .TextoCuerpo FormatoMoneda(mobjPrenda.Transporte, "EUR"), 5, LineaCuerpo, vbRightJustify
        LineaCuerpo = LineaCuerpo + 1
        
        .LineaCelda 5, LineaCuerpo
        .TextoCuerpo "OTROS", 1, LineaCuerpo, vbLeftJustify
        .TextoCuerpo FormatoMoneda(mobjPrenda.Carton, "EUR"), 5, LineaCuerpo, vbRightJustify
        LineaCuerpo = LineaCuerpo + 2

        .LineaCelda 6, LineaCuerpo
        .TextoCuerpo "SUBTOTAL", 1, LineaCuerpo, vbLeftJustify
        .TextoCuerpo FormatoMoneda(dblCostesSubtotal, "EUR"), 6, LineaCuerpo, vbRightJustify
        LineaCuerpo = LineaCuerpo + 2

        .LineaCelda 6, LineaCuerpo
        .TextoCuerpo "ADMINISTRACION", 1, LineaCuerpo, vbLeftJustify
        .TextoCuerpo FormatoCantidad(mobjPrenda.administracion) & "%", 5, LineaCuerpo, vbRightJustify
        .TextoCuerpo FormatoMoneda(dblCostesAdministracion, "EUR"), 6, LineaCuerpo, vbRightJustify
        LineaCuerpo = LineaCuerpo + 3
        
        .LineaCelda 7, LineaCuerpo
        .TextoCuerpo "COSTE TOTAL", 1, LineaCuerpo, vbLeftJustify
        .TextoCuerpo FormatoMoneda(dblCostesAdministracion + _
                                     dblCostesSubtotal, "EUR"), 7, LineaCuerpo, vbRightJustify
        '.LineaFilaArriba 16

    End With
    
End Sub

Public Sub PrintObject()

    On Error GoTo ErrorManager
    
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


