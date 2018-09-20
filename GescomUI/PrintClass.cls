VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrintClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mintAnchuraUtil As Double

Private mintLineaCuerpo As Integer

Private mstrTitulo As String

' Indicador de si se ha informado de las columnas.
Private mflgColumnHeaders As Boolean

' Indicador de si es el primer item para imprimir cabecera...
Private mflgPrimerItem As Boolean

Private mobjColumnHeaders As ColumnHeaders

Private mobjTablaCuerpo As PrintTabla
Private mobjTablaTitulo As PrintTabla
Private mobjTablaPie As PrintTabla

Private Sub Class_Initialize()
    
' Establecemos orientacion vertical y tamaño de papel A4
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = vbPRORPortrait
    Printer.Copies = 1

    mflgColumnHeaders = False
    mflgPrimerItem = True
    
    mintAnchuraUtil = 19
    Set mobjTablaCuerpo = New PrintTabla
    With mobjTablaCuerpo
        .Dimensiones 0.4, 3, mintAnchuraUtil, 25
        .Cabecera = True
        .FuenteCabeceraBold = True
        .FuenteCabeceraItalic = True
    End With
    
    mintLineaCuerpo = 1
    
    Set mobjTablaTitulo = New PrintTabla
    With mobjTablaTitulo
        .Dimensiones 0.4, 2, mintAnchuraUtil, 1.4
        .Cabecera = False
        .FuenteCuerpoBold = True
        .FuenteCuerpoItalic = False
        .FuenteCuerpoSize = 16
    End With
    
    Set mobjTablaPie = New PrintTabla
    With mobjTablaPie
        .Dimensiones 0.4, 28, mintAnchuraUtil, 0.6
        .Cabecera = False
        .NumeroColumnas = 2
        .Autosize
        .FuenteCuerpoBold = False
        .FuenteCuerpoItalic = False
        .FuenteCuerpoSize = 8
    End With
    
End Sub

Private Sub Class_Terminate()
    
    Set mobjTablaCuerpo = Nothing
    Set mobjTablaTitulo = Nothing
    Set mobjTablaPie = Nothing

End Sub

Public Property Let PrinterNumber(PrinterNumber As Integer)
    
    Set Printer = Printers(PrinterNumber)
    
End Property

Public Property Let Orientation(Orientation As Integer)
    
    If Not mflgPrimerItem Then Err.Raise vbObjectError + 1001, "No se puede cambiar la orientación, el listado está en curso."
    
    Select Case Orientation
    Case vbPRORPortrait
        mintAnchuraUtil = 18.5
        mobjTablaCuerpo.Dimensiones 0.4, 3, mintAnchuraUtil, 25
        mobjTablaTitulo.Dimensiones 0.4, 2, mintAnchuraUtil, 1.4
        mobjTablaPie.Dimensiones 0.4, 28, mintAnchuraUtil, 0.6
    Case vbPRORLandscape
        mintAnchuraUtil = 27.3
        mobjTablaCuerpo.Dimensiones 0.4, 3, mintAnchuraUtil, 16.4
        mobjTablaTitulo.Dimensiones 0.4, 2, mintAnchuraUtil, 1.4
        mobjTablaPie.Dimensiones 0.4, 19.2, mintAnchuraUtil, 0.6
    Case Else
        Err.Raise vbObjectError + 1001, "Tipo de orientación erronea, debe ser portrait o landscape."
    End Select
    Printer.Orientation = Orientation

End Property

Public Property Let Copies(Copies As Integer)
    
    If Not mflgPrimerItem Then Err.Raise vbObjectError + 1001, "No se puede cambiar el numero de copias, el listado está en curso."
    
    Printer.Copies = Copies

End Property

Public Property Let Titulo(Titulo As String)
    
    If Not mflgPrimerItem Then Err.Raise vbObjectError + 1001, "No se puede establecer el titulo, el listado está en curso."
    
    mstrTitulo = Titulo

End Property

Public Property Let Columnas(Columnas As ColumnHeaders)
    Dim i As Integer

    If mflgColumnHeaders Then Err.Raise vbObjectError + 1001, "No se pueden establecer los encabezamientos de columnas, ya estan asignados."
    
    Set mobjColumnHeaders = Columnas

    mobjTablaCuerpo.NumeroColumnas = mobjColumnHeaders.Count
    
    For i = 1 To mobjColumnHeaders.Count - 1
        mobjTablaCuerpo.LongitudColumna ((mobjColumnHeaders(i).Width * mintAnchuraUtil) / ColumnHeaderSize)
    Next
    
    mflgColumnHeaders = True
    
End Property

Public Property Let Item(itemList As ListItem)
    Dim i As Integer
    Dim strTexto As String
    
    If mflgPrimerItem Then ImprimirDatosFijos
    
    If mintLineaCuerpo > mobjTablaCuerpo.NumeroFilasCuerpo Then ImprimirNuevaPagina
    
    For i = 1 To ColumnCount
        If i = 1 Then
            strTexto = itemList.Text
        Else
            strTexto = itemList.SubItems(i - 1)
        End If
        mobjTablaCuerpo.TextoCuerpo strTexto, i, mintLineaCuerpo, mobjColumnHeaders(i).Alignment
    Next i
    
    mintLineaCuerpo = mintLineaCuerpo + 1
    
End Property

Private Function ColumnCount() As Integer

    ColumnCount = mobjColumnHeaders.Count
    
End Function

Private Sub ImprimirNuevaPagina()
    
    Printer.NewPage
    ImprimirDatosFijos
    mintLineaCuerpo = 1

End Sub

Private Sub ImprimirDatosFijos()
    
    ImprimirCabecera
    ImprimirColumnas
    ImprimirPie
    
    mflgPrimerItem = False

End Sub

Private Sub ImprimirCabecera()
    
    ' Imprimimos el logotipo
    Printer.PaintPicture LoadPicture(App.Path & "\logo.bmp"), _
                         567, 283, 3997, 850
    
    If Len(mstrTitulo) Then
        mobjTablaTitulo.TextoCuerpo mstrTitulo, 1, 1, vbCenter
    End If

End Sub

Private Sub ImprimirColumnas()
    Dim i As Integer

    For i = 1 To mobjColumnHeaders.Count
        If mobjColumnHeaders(i).Width <> 0 Then _
        mobjTablaCuerpo.TextoCabecera mobjColumnHeaders(i).Text, i, mobjColumnHeaders(i).Alignment
    Next
    
End Sub

Private Sub ImprimirPie()
    
    mobjTablaPie.LineaCeldaArriba 1, 1
    mobjTablaPie.LineaCeldaArriba 2, 1
    
    mobjTablaPie.TextoCuerpo "Página " & Printer.Page, 1, 1, vbLeftJustify
    
    'Escribimos la fecha y hora de impresion ajustado a la izquierda
    mobjTablaPie.TextoCuerpo Format$(Now, "Long Date") & " a las " & Format$(Now, "short time"), _
                            2, 1, vbRightJustify
    
    
End Sub

Private Function ColumnHeaderSize() As Integer
    Dim mobjColumnHeader As ColumnHeader

    ColumnHeaderSize = 0
    For Each mobjColumnHeader In mobjColumnHeaders
        ColumnHeaderSize = ColumnHeaderSize + mobjColumnHeader.Width
    Next

End Function

Public Sub EndDoc()

    If Not mflgColumnHeaders Then Err.Raise vbObjectError + 1001, "No se han establecido los encabezamientos de columnas."
    
    Printer.EndDoc
        
End Sub

