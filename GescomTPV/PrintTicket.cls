VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrintTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Indicador de si se debe imprimir el logotipo.
Private mflgPrintLogo As Boolean

Private mobjAlbaranVenta As AlbaranVenta
Private mobjEmpresa As Empresa

Private Const cnRayaSeparador = "--------------------------------"
Private Const cnEspacioAreaTexto = 23

Public Sub Component(AlbaranVentaObject As AlbaranVenta)

    Set mobjAlbaranVenta = AlbaranVentaObject
    
    Set mobjEmpresa = New Empresa
    mobjEmpresa.Load mobjAlbaranVenta.EmpresaID
    mflgPrintLogo = Not (mobjEmpresa.AnularIVA)

End Sub

Private Sub Class_Initialize()
    
    ' Por defecto imprimimos el logo.
    mflgPrintLogo = True
    Printer.Font.Name = "Courier New"
    
End Sub

Private Sub Class_Terminate()

    Set mobjAlbaranVenta = Nothing
    Set mobjEmpresa = Nothing
     
    Printer.EndDoc

End Sub

Public Property Let Copies(Copies As Integer)
    
    Printer.Copies = Copies
    
End Property

Private Sub ImprimirDatosFijos()
    
    ImprimirCabecera
    
End Sub

Private Sub ImprimirCabecera()
    
    ' Imprimimos el logotipo
    If mflgPrintLogo Then _
        Printer.PaintPicture LoadPicture(App.Path & "\logo.bmp"), _
                              2000, 1, 300, 300
                         
    ImprimirDireccionEmpresa
    ImprimirFechaYNumero
    
End Sub

Private Sub ImprimirDireccionEmpresa()

    Printer.Print mobjEmpresa.Titular   'OJOOJO CONFIRMAR QUE ES LA RAZÓN SOCIAL
    Printer.Print "Telf.: " & mobjEmpresa.Direccion.Telefono1
    Printer.Print mobjEmpresa.Direccion.Calle
    Printer.Print mobjEmpresa.Direccion.CodigoPostal & "-" & mobjEmpresa.Direccion.Poblacion & "  (" & mobjEmpresa.Direccion.Provincia & ")"
    Printer.Print "C.I.F.: " & mobjEmpresa.DNINIF

End Sub

Private Sub ImprimirFechaYNumero()
    
    Printer.Print Format(mobjAlbaranVenta.Fecha, "dd/mm/yyyy HH:MM") & " NUM.:" & Format(mobjAlbaranVenta.Numero, "###,###,###")
    Printer.Print cnRayaSeparador
    
End Sub

Private Sub ImprimirDetalle()
    Dim objAlbaranVentaItem As AlbaranVentaItem
    Dim LineaCuerpo As Integer
'    Dim objArticuloColor As ArticuloColor
    
    
    LineaCuerpo = 1
    For Each objAlbaranVentaItem In mobjAlbaranVenta.AlbaranVentaItems
        
        Printer.Print Left(objAlbaranVentaItem.Descripcion, cnEspacioAreaTexto) & FormatoMoneda(objAlbaranVentaItem.BRUTO, "EUR", False)
        
        LineaCuerpo = LineaCuerpo + 1
    
    Next
    
    Printer.Print cnRayaSeparador
    Printer.Print "    T O T A L...:      " & FormatoMoneda(mobjAlbaranVenta.TotalBruto, "EUR")
    Printer.Print " ENTREGADO......:      " ' OJOOJO PONER LA CANTIDAD ENTREGADA & FormatoMoneda(mobjAlbaranVenta.TotalBruto, "EUR")
    Printer.Print " CAMBIO.........:      " ' OJOOJO PONER EL CAMBIO & FormatoMoneda(mobjAlbaranVenta.TotalBruto, "EUR")
    Printer.Print "  P.V.P    I.V.A.  INCLUIDO"
    Printer.Print vbNullString
    Printer.Print "  GRACIAS POR SU VISITA"

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

