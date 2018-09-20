VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrintComisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjComisiones As Comisiones
Private mrsRecordset As ADOR.Recordset

Private mobjTablaCuerpo As PrintTabla
Private mobjTablaTotal As PrintTabla
Private mobjTablaTitulo As PrintTabla
Private mobjTablaRepresentante As PrintTabla
Private mobjTablaPie As PrintTabla
Private mdblTotalBruto As Double
Private mdblTotalComision As Double
Private mstrRepresentante As String
Private mstrEmpresa As String

Public Sub Component(ComisionesObject As Comisiones)
    Dim objEmpresa As Empresa
    
    Set mobjComisiones = ComisionesObject
    Set mrsRecordset = ComisionesObject.ComisionesRecordset
    
    Set objEmpresa = New Empresa
    objEmpresa.Load mobjComisiones.EmpresaID
    mstrEmpresa = objEmpresa.Nombre
    Set objEmpresa = Nothing
    
End Sub

Private Sub Class_Initialize()
    
' Establecemos orientacion vertical y tamaño de papel A4
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = vbPRORPortrait
    Printer.Copies = 1

    Set mobjTablaCuerpo = New PrintTabla
    With mobjTablaCuerpo
        .NumeroColumnas = 5
        .Dimensiones 0.4, 6.5, 19.2, 18.4
        .LongitudColumna 2
        .LongitudColumna 2
        .LongitudColumna 8
        .LongitudColumna 4
        .FuenteCuerpoSize = 9
        .FuenteCabeceraSize = 9
    End With

    Set mobjTablaTotal = New PrintTabla
    With mobjTablaTotal
        .Dimensiones 9.1, 26.2, 10.5, 0.8
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
    
    Set mobjTablaRepresentante = New PrintTabla
    With mobjTablaRepresentante
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

    Set mobjComisiones = Nothing
    Set mrsRecordset = Nothing
    
    Set mobjTablaCuerpo = Nothing
    Set mobjTablaTotal = Nothing
    Set mobjTablaTitulo = Nothing
    Set mobjTablaRepresentante = Nothing
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
    
    mobjTablaTitulo.TextoCuerpo "COMISIONES DE REPRESENTANTES", 1, 1, vbCenter

End Sub

Private Sub ImprimirRejilla()
    
    With mobjTablaCuerpo
        .LineaTodasColumnas
        .LineasBorde
        .LineaCabecera
        .TextoCabecera "Factura", 1
        .TextoCabecera "Fecha", 2
        .TextoCabecera "Cliente", 3
        .TextoCabecera "Importe bruto", 4, vbRightJustify
        .TextoCabecera "Importe comision", 5, vbRightJustify
    End With
    
End Sub

Private Sub ImprimirSerie()
    Dim strFacturas As String

    strFacturas = "Del " & mobjComisiones.FechaInicial & " al " & mobjComisiones.FechaFinal
    
    With mobjTablaRepresentante
        .LineasBorde
        
        .TextoCuerpo "Representante:", 1, 1
        .TextoCuerpo mstrRepresentante, 2, 1
        .TextoCuerpo "Empresa:", 1, 2
        .TextoCuerpo mstrEmpresa, 2, 2
        .TextoCuerpo "Facturas: ", 1, 3
        .TextoCuerpo strFacturas, 2, 3
        
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
    
    mdblTotalBruto = 0
    mdblTotalComision = 0
    mstrRepresentante = mrsRecordset("NombreRepresentante")
    LineaCuerpo = 1
    mrsRecordset.MoveFirst
    
    ImprimirDatosFijos
    While Not mrsRecordset.EOF
        If mstrRepresentante <> mrsRecordset("NombreRepresentante") Then
            mstrRepresentante = mrsRecordset("NombreRepresentante")
            ImprimirNuevaPagina
            mdblTotalBruto = 0
            mdblTotalComision = 0
            LineaCuerpo = 1
        End If
    
        With mobjTablaCuerpo
            .TextoCuerpo FormatoCantidad(mrsRecordset("Numero")), 1, LineaCuerpo, vbRightJustify
            .TextoCuerpo mrsRecordset("Fecha"), 2, LineaCuerpo
            .TextoCuerpo mrsRecordset("NombreCliente"), 3, LineaCuerpo
            .TextoCuerpo FormatoMoneda(mrsRecordset("BrutoEUR"), "EUR", False), 4, LineaCuerpo, vbRightJustify
            .TextoCuerpo FormatoMoneda(mrsRecordset("ComisionEUR"), "EUR", False), 5, LineaCuerpo, vbRightJustify
            mdblTotalBruto = mdblTotalBruto + mrsRecordset("BrutoEUR")
            mdblTotalComision = mdblTotalComision + mrsRecordset("ComisionEUR")
            
        End With
        
        
        LineaCuerpo = LineaCuerpo + 1
        If LineaCuerpo >= mobjTablaCuerpo.NumeroFilasCuerpo Then
            ImprimirNuevaPagina
            LineaCuerpo = 1
        End If
        
        mrsRecordset.MoveNext
        
    Wend
    
    ImprimirTotal
    
End Sub

Private Sub ImprimirTotal()

    With mobjTablaTotal
        .LineaCelda 2, 1
        .LineaCelda 3, 1
        .TextoCuerpo "TOTALES>>", 1, 1, vbCenter
        .TextoCuerpo FormatoMoneda(mdblTotalBruto, "EUR"), 2, 1, vbCenter
        .TextoCuerpo FormatoMoneda(mdblTotalComision, "EUR"), 3, 1, vbCenter
        
    End With
    
End Sub


Public Sub PrintObject()

    On Error GoTo ErrorManager
    
    If mrsRecordset.EOF Then _
        Err.Raise vbObjectError + 1001, "Impresion de comisiones", _
                                        "No existen datos de comisiones para las facturas seleccionadas."
    'ImprimirDatosFijos
    ImprimirDetalle
    Exit Sub

ErrorManager:
    Printer.KillDoc
    Err.Raise vbObjectError + 1001, Err.Source, Err.Description
End Sub

