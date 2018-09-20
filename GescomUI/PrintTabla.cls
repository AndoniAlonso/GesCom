VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrintTabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const gcLongitudCabeceraY = 0.6         ' Tamaño de la casilla de encabezamientos de columna Y
Private Const TwipsPerCm = 567
Private Const MargenInterno = 0.1               ' Distancia de los bordes al texto

'Dimensiones
Private mdblPosicionX As Double
Private mdblPosicionY As Double
Private mdblLongitudX As Double
Private mdblLongitudY As Double
        
Private mdblLongitudCabeceraY As Double
        
'Si lleva encabezamiento o no
Private mflgCabecera As Boolean

'Si la fuente del cuerpo está en negrita/cursiva
Private mflgFuenteCuerpoBold As Boolean
Private mflgFuenteCuerpoItalic As Boolean
Private mintFuenteCuerpoSize As Integer

'Si la fuente de la cabecera está en negrita/cursiva
Private mflgFuenteCabeceraBold As Boolean
Private mflgFuenteCabeceraItalic As Boolean
Private mintFuenteCabeceraSize As Integer

'Si el borde se pinta o no
Private mflgBorde As Boolean

'número de columnas de la tabla
Private mintNumeroColumnas As Integer
Private mcolPosicionColumnas As Collection
'Nº de línea de impresión de los datos del cuerpo
Private mintNumeroFila As Integer

Private Sub Class_Initialize()

    mdblPosicionX = 0
    mdblPosicionY = 0
    mdblLongitudX = 0
    mdblLongitudY = 0
    
    mflgCabecera = False
    
    mflgBorde = True
    
    mintNumeroColumnas = 1
    Set mcolPosicionColumnas = New Collection
    
    mintNumeroFila = 1
    
    mflgFuenteCuerpoBold = False
    mflgFuenteCuerpoItalic = False
    mintFuenteCuerpoSize = 10
    
    mflgFuenteCabeceraBold = False
    mflgFuenteCabeceraItalic = False
    mintFuenteCabeceraSize = 9
    
    mdblLongitudCabeceraY = gcLongitudCabeceraY
    
    Printer.DrawWidth = 1
    
End Sub

Public Sub Dimensiones(PosicionX As Double, PosicionY As Double, LongitudX As Double, LongitudY As Double)

    If PosicionX < 0 Or _
       PosicionY < 0 Or _
       LongitudX <= 0 Or _
       LongitudY <= 0 Then _
            Err.Raise vbObjectError + 1001, "Las dimensiones de la tabla deben ser positivas"

    mdblPosicionX = PosicionX
    mdblPosicionY = PosicionY
    mdblLongitudX = LongitudX
    mdblLongitudY = LongitudY

End Sub

Private Sub Class_Terminate()
    
    Set mcolPosicionColumnas = Nothing

End Sub
Public Property Let Cabecera(Value As Boolean)

    mflgCabecera = Value
    mdblLongitudCabeceraY = IIf(mflgCabecera, gcLongitudCabeceraY, 0)
    
End Property

Public Property Let Borde(Value As Boolean)

    mflgBorde = Value
    
End Property

Public Property Let FuenteCuerpoBold(Value As Boolean)

    mflgFuenteCuerpoBold = Value
    
End Property

Public Property Let FuenteCuerpoItalic(Value As Boolean)

    mflgFuenteCuerpoItalic = Value
    
End Property

Public Property Let FuenteCuerpoSize(Value As Integer)

    mintFuenteCuerpoSize = Value
    
End Property

Public Property Let FuenteCabeceraBold(Value As Boolean)

    mflgFuenteCabeceraBold = Value
    
End Property

Public Property Let FuenteCabeceraItalic(Value As Boolean)

    mflgFuenteCabeceraItalic = Value
    
End Property

Public Property Let FuenteCabeceraSize(Value As Integer)

    mintFuenteCabeceraSize = Value
    
End Property

Public Property Let NumeroColumnas(Value As Integer)

    If Value < 1 Then _
        Err.Raise vbObjectError + 1001, "El número de columnas debe ser mayor o igual a 1"
    
    mintNumeroColumnas = Value
    
End Property

Public Sub Autosize()
    Dim i As Integer
    
    For i = 1 To mintNumeroColumnas - 1
        mcolPosicionColumnas.Add mdblLongitudX / mintNumeroColumnas
    Next
End Sub

Public Sub LongitudColumna(Value As Double)

    If mintNumeroColumnas - 1 = mcolPosicionColumnas.Count Then _
        Err.Raise vbObjectError + 1001, "El número de columna debe ser menor que el límite definido"
    
    mcolPosicionColumnas.Add Value
    
End Sub

Private Function ColumnaPosicionIzquierda(NumeroColumna As Integer) As Double
    Dim i As Integer
    
    If NumeroColumna > mintNumeroColumnas Then _
        Err.Raise vbObjectError + 1001, "El número de columna debe ser menor que el límite definido"

    If NumeroColumna = 1 Then
        ColumnaPosicionIzquierda = mdblPosicionX
    Else
        ColumnaPosicionIzquierda = mdblPosicionX + mcolPosicionColumnas(1)
        For i = 2 To NumeroColumna - 1
             ColumnaPosicionIzquierda = ColumnaPosicionIzquierda + mcolPosicionColumnas(i)
        Next
    End If
    
End Function

Private Function ColumnaPosicionDerecha(NumeroColumna As Integer) As Double
    Dim i As Integer
    
    If NumeroColumna > mintNumeroColumnas Then _
        Err.Raise vbObjectError + 1001, "El número de columna debe ser menor que el límite definido"

    If NumeroColumna = mintNumeroColumnas Then
        ColumnaPosicionDerecha = mdblPosicionX + mdblLongitudX
    Else
        ColumnaPosicionDerecha = mdblPosicionX
        For i = 1 To NumeroColumna
             ColumnaPosicionDerecha = ColumnaPosicionDerecha + mcolPosicionColumnas(i)
        Next
    End If
    
End Function

Private Function FilaPosicionArriba(NumeroFila As Integer)

    If NumeroFila > NumeroFilasCuerpo Then _
        Err.Raise vbObjectError + 1001, "El número de fila debe ser menor que el límite definido"
    
    FuenteTextoCuerpo
    FilaPosicionArriba = mdblPosicionY + mdblLongitudCabeceraY + (NumeroFila - 1) * Interlineado / TwipsPerCm + (MargenInterno / 2)
    
End Function

Private Function FilaPosicionAbajo(NumeroFila As Integer)

    If NumeroFila > NumeroFilasCuerpo Then _
        Err.Raise vbObjectError + 1001, "El número de fila debe ser menor que el límite definido"
    
    FuenteTextoCuerpo
    If NumeroFila = NumeroFilasCuerpo Then
       FilaPosicionAbajo = mdblPosicionY + mdblLongitudY
    Else
       FilaPosicionAbajo = mdblPosicionY + mdblLongitudCabeceraY + NumeroFila * Interlineado / TwipsPerCm + (MargenInterno / 2)
    End If
End Function

Public Sub LineasBorde()

    If mflgBorde Then
        PintarCaja mdblPosicionX, mdblPosicionY, _
                   mdblPosicionX + mdblLongitudX, mdblPosicionY + mdblLongitudY

    Else
        Err.Raise vbObjectError + 1001, "No se imprime el borde, es de tipo sin borde"
    End If
    
End Sub

Public Sub LineaCabecera()

    If mdblLongitudY < mdblLongitudCabeceraY Then _
        Err.Raise vbObjectError + 1001, "Espacio insuficiente para la cabecera"

    PintarLinea mdblPosicionX, mdblPosicionY + mdblLongitudCabeceraY, _
                mdblPosicionX + mdblLongitudX, mdblPosicionY + mdblLongitudCabeceraY
    
End Sub

Public Sub LineaMediaVertical()

    PintarLinea (mdblPosicionX + mdblLongitudX) / 2, mdblPosicionY, _
                (mdblPosicionX + mdblLongitudX) / 2, mdblPosicionY + mdblLongitudY

End Sub

Public Sub LineaMediaHorizontal()

    PintarLinea mdblPosicionX, (mdblPosicionY + mdblLongitudY) / 2, _
                mdblPosicionX + mdblLongitudX, (mdblPosicionY + mdblLongitudY) / 2

End Sub

Public Sub LineaDiagonalIzquierda()

    PintarLinea mdblPosicionX, mdblPosicionY, _
                mdblPosicionX + mdblLongitudX, mdblPosicionY + mdblLongitudY

End Sub

Public Sub LineaDiagonalDerecha()

    PintarLinea mdblPosicionX, mdblPosicionY + mdblLongitudY, _
                mdblPosicionX + mdblLongitudX, mdblPosicionY

End Sub

Public Sub TextoCabecera(Texto As String, Columna As Integer, Optional Alineacion As AlignmentConstants = vbLeftJustify)
    
    If Columna > mintNumeroColumnas Then _
        Err.Raise vbObjectError + 1001, "El número de columna debe ser menor que el límite definido"
    
    FuenteTextoCabecera
    Printer.CurrentX = AlinearTextoX(Columna, Texto, Alineacion)
    Printer.CurrentY = TwipsPerCm * (mdblPosicionY + MargenInterno)
    Printer.Print Texto

End Sub

Public Sub TextoCuerpo(Texto As String, Columna As Integer, Fila As Integer, Optional Alineacion As AlignmentConstants = vbLeftJustify)
    Dim strTextoAjustado As String
    Dim dblLongitudHorizontal As Double
    
    If Columna > mintNumeroColumnas Then _
        Err.Raise vbObjectError + 1001, "El número de columna debe ser menor que el límite definido"
    
    If Fila > NumeroFilasCuerpo Then
        Err.Raise vbObjectError + 1001, "El número de línea debe ser menor que el límite"
    End If
    
    dblLongitudHorizontal = ColumnaPosicionDerecha(Columna) - ColumnaPosicionIzquierda(Columna) - (MargenInterno * 2)
    strTextoAjustado = RecortarTexto(Texto, dblLongitudHorizontal)
    
    Printer.CurrentX = AlinearTextoX(Columna, strTextoAjustado, Alineacion)
    Printer.CurrentY = TwipsPerCm * (mdblPosicionY + mdblLongitudCabeceraY + MargenInterno) + Interlineado * (Fila - 1)
    Printer.Print strTextoAjustado
End Sub

Private Function RecortarTexto(Texto As String, Longitud As Double) As String
    Dim TextoAAjustar As String
    
    TextoAAjustar = Trim(Texto)
    
    While Printer.TextWidth(TextoAAjustar) > Longitud * TwipsPerCm And _
          Len(TextoAAjustar) > 0
        TextoAAjustar = Left$(TextoAAjustar, Len(TextoAAjustar) - 1)
    Wend

    RecortarTexto = TextoAAjustar

End Function

Public Sub LineaFilaArriba(Fila As Integer)
    
    If Fila > NumeroFilasCuerpo Then _
        Err.Raise vbObjectError + 1001, "El número de línea debe ser menor que el límite"
    
    FuenteTextoCuerpo
    PintarLinea ColumnaPosicionIzquierda(1), FilaPosicionArriba(Fila), ColumnaPosicionDerecha(mintNumeroColumnas), FilaPosicionArriba(Fila)
    
End Sub

Public Sub LineaCelda(Columna As Integer, Fila As Integer)
    
    If Columna > mintNumeroColumnas Then _
        Err.Raise vbObjectError + 1001, "El número de columna debe ser menor que el límite definido"
    
    If Fila > NumeroFilasCuerpo Then _
        Err.Raise vbObjectError + 1001, "El número de línea debe ser menor que el límite"
    
    FuenteTextoCuerpo
    PintarCaja ColumnaPosicionIzquierda(Columna), FilaPosicionArriba(Fila), ColumnaPosicionDerecha(Columna), FilaPosicionAbajo(Fila)
    
End Sub

Public Sub LineaCeldaDerecha(Columna As Integer, Fila As Integer)
    
    If Columna > mintNumeroColumnas Then _
        Err.Raise vbObjectError + 1001, "El número de columna debe ser menor que el límite definido"
    
    If Fila > NumeroFilasCuerpo Then _
        Err.Raise vbObjectError + 1001, "El número de línea debe ser menor que el límite"
    FuenteTextoCuerpo
    PintarLinea ColumnaPosicionDerecha(Columna), FilaPosicionArriba(Fila), ColumnaPosicionDerecha(Columna), FilaPosicionAbajo(Fila)

End Sub
Public Sub LineaCeldaArriba(Columna As Integer, Fila As Integer)
    
    If Columna > mintNumeroColumnas Then _
        Err.Raise vbObjectError + 1001, "El número de columna debe ser menor que el límite definido"
    
    If Fila > NumeroFilasCuerpo Then _
        Err.Raise vbObjectError + 1001, "El número de línea debe ser menor que el límite"
    
    FuenteTextoCuerpo
    PintarLinea ColumnaPosicionIzquierda(Columna), FilaPosicionArriba(Fila), ColumnaPosicionDerecha(Columna), FilaPosicionArriba(Fila)
    
End Sub

Public Property Get NumeroFilasCuerpo() As Integer
    Dim dblLongitudTextoY As Double
    
    FuenteTextoCuerpo
    
    dblLongitudTextoY = Printer.TextHeight(vbNullString)
    NumeroFilasCuerpo = Int((mdblLongitudY - mdblLongitudCabeceraY - 2 * MargenInterno) * TwipsPerCm / dblLongitudTextoY)
    
End Property

Private Function AlinearTextoX(Columna As Integer, Texto As String, Alineacion As AlignmentConstants) As Double
    Dim dblLongitudTexto As Double
    Dim dblPosicionIzquierda As Double
    Dim dblPosicionDerecha As Double
    
    If Columna > mintNumeroColumnas Then _
        Err.Raise vbObjectError + 1001, "El número de columna debe ser menor que el límite definido"
    
    dblLongitudTexto = Printer.TextWidth(Texto)
    dblPosicionIzquierda = ColumnaPosicionIzquierda(Columna)
    dblPosicionDerecha = ColumnaPosicionDerecha(Columna)
    
    Select Case Alineacion
    Case vbLeftJustify
        AlinearTextoX = (dblPosicionIzquierda + MargenInterno) * TwipsPerCm
    Case vbRightJustify
        AlinearTextoX = (dblPosicionDerecha - MargenInterno) * TwipsPerCm - dblLongitudTexto
    Case vbCenter
        AlinearTextoX = ((dblPosicionIzquierda + dblPosicionDerecha) / 2) * TwipsPerCm - (dblLongitudTexto / 2)
    Case Else
        Err.Raise vbObjectError + 1001, "Valor de alineación incorrecto"
    End Select

End Function

Private Sub PintarLinea(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double)

    Printer.Line (TwipsPerCm * X1, TwipsPerCm * Y1)- _
                 (TwipsPerCm * X2, TwipsPerCm * Y2)

End Sub

Private Sub PintarCaja(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double)

    Printer.Line (TwipsPerCm * X1, TwipsPerCm * Y1)- _
                 (TwipsPerCm * X2, TwipsPerCm * Y2), , B

End Sub

Private Sub FuenteTextoCabecera()
    
    ' Establecemos la fuente de los literales del preimpreso.
    ' Tahoma, 9 puntos, normal
    Printer.Font.Name = "Tahoma"
    Printer.Font.Size = mintFuenteCabeceraSize
    Printer.Font.Bold = mflgFuenteCabeceraBold
    Printer.Font.Italic = mflgFuenteCabeceraItalic

End Sub

Private Sub FuenteTextoCuerpo()
    
    ' Establecemos la fuente del cuerpo del informe.
    ' Tahoma, 10 puntos, normal
    Printer.Font.Name = "Tahoma"
    Printer.Font.Size = mintFuenteCuerpoSize
    Printer.Font.Bold = mflgFuenteCuerpoBold
    Printer.Font.Italic = mflgFuenteCuerpoItalic

End Sub

Public Sub LineaColumna(Columna As Integer)
    
    If Columna = mintNumeroColumnas Then Exit Sub
    
    PintarLinea ColumnaPosicionDerecha(Columna), mdblPosicionY, _
                ColumnaPosicionDerecha(Columna), mdblPosicionY + mdblLongitudY
    
End Sub

Public Sub LineaTodasColumnas()
    Dim i As Integer
    
    For i = 1 To mintNumeroColumnas
        LineaColumna i
    Next

End Sub

Private Function Interlineado() As Double

    Interlineado = Printer.TextHeight(vbNullString)
    
End Function


Public Property Let GrosorLinea(Grosor As Integer)

    Printer.DrawWidth = Grosor
    
End Property
