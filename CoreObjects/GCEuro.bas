Attribute VB_Name = "GCEuro"
Option Explicit

Public Const gcTipoCobro = "C"
Public Const gcTipoPago = "P"

Public Const gcSituacionCobroPendiente = "A"
Public Const gcSituacionCobroPagado = "C"

Public Function PTA2EUR(ImportePTA As Double) As Double
  
    PTA2EUR = Round(ImportePTA / 166.386, 2)
     
End Function

Public Function EUR2PTA(ImporteEUR As Double) As Double
  
    EUR2PTA = Round(ImporteEUR * 166.386, 0)
     
End Function

Public Function EsMonedaValida(Moneda As String) As Boolean
  
    If EsEUR(Moneda) Or EsPTA(Moneda) Then
       EsMonedaValida = True
    Else
       EsMonedaValida = False
    End If
   
End Function

Public Function EsEUR(Moneda As String) As Boolean
  
    EsEUR = (Moneda = "EUR")
   
End Function

Public Function EsPTA(Moneda As String) As Boolean
  
    EsPTA = (Moneda = "PTA")
   
End Function

Public Function FormatoMoneda(Importe As Double, Moneda As String, Optional MostrarMoneda As Boolean = True) As String
    
    If EsEUR(Moneda) Then
        FormatoMoneda = Format(Importe, "###,###,##0.00 ;###,###,##0.00-") & IIf(MostrarMoneda, " €", vbNullString)
        FormatoMoneda = Space(IIf(MostrarMoneda, 16, 14) - Len(FormatoMoneda)) & FormatoMoneda
    ElseIf EsPTA(Moneda) Then
        FormatoMoneda = Format(Importe, "###,###,##0 ;###,###,##0-") & IIf(MostrarMoneda, " Pta", vbNullString)
        FormatoMoneda = Space(IIf(MostrarMoneda, 15, 11) - Len(FormatoMoneda)) & FormatoMoneda
    Else
        Err.Raise vbObjectError + 1001, "El tipo de moneda debe ser EUR o PTA."
    End If

End Function

Public Function FormatoMonedaDocumento(Importe As Double, Moneda As String, Optional MostrarMoneda As Boolean = True) As String
    
    If EsEUR(Moneda) Then
        FormatoMonedaDocumento = Format(Importe, "###,###,##0.00") & IIf(MostrarMoneda, " €", vbNullString)
        FormatoMonedaDocumento = Space(IIf(MostrarMoneda, 16, 14) - Len(FormatoMonedaDocumento)) & FormatoMonedaDocumento
    ElseIf EsPTA(Moneda) Then
        FormatoMonedaDocumento = Format(Importe, "###,###,##0") & IIf(MostrarMoneda, " Pta", vbNullString)
        FormatoMonedaDocumento = Space(IIf(MostrarMoneda, 15, 11) - Len(FormatoMonedaDocumento)) & FormatoMonedaDocumento
    Else
        Err.Raise vbObjectError + 1001, "El tipo de moneda debe ser EUR o PTA."
    End If

End Function


Public Function FormatoCantidad(Cantidad As Double, Optional DosDecimales As Boolean = False) As String

    On Error GoTo ErrorManager

    If DosDecimales Then
        FormatoCantidad = Format(Cantidad, "##,###,##0.00")
        FormatoCantidad = Space(13 - Len(FormatoCantidad)) & FormatoCantidad
    Else
        FormatoCantidad = Format(Cantidad, "###,###,##0")
        FormatoCantidad = Space(11 - Len(FormatoCantidad)) & FormatoCantidad
    End If
    Exit Function

ErrorManager:
    FormatoCantidad = FormatNumber(Cantidad, 2)
End Function

Public Function FormatoFecha(Fecha As Date) As String
    
    FormatoFecha = Format(Fecha, "yyyy/mm/dd")
    
End Function

Public Function CalcularDiasPago(FechaGiro As Date, DiaPago1 As Integer, _
                                DiaPago2 As Integer, DiaPago3 As Integer) As Date
Dim FechaResult As Date

    FechaResult = FechaGiro
    If DiaPago1 <> 0 Then
        If DiaPago1 >= Day(FechaGiro) Then
            FechaResult = CDate(DiaPago1 & "/" & Month(FechaGiro) & "/" & Year(FechaGiro))
        
        ElseIf DiaPago2 <> 0 Then
            If DiaPago2 >= Day(FechaGiro) Then
                FechaResult = CDate(DiaPago2 & "/" & Month(FechaGiro) & "/" & Year(FechaGiro))
                
            ElseIf DiaPago3 <> 0 Then
                If DiaPago3 >= Day(FechaGiro) Then
                    FechaResult = CDate(DiaPago3 & "/" & Month(FechaGiro) & "/" & Year(FechaGiro))
                Else
                    FechaResult = CDate(DiaPago1 & "/" & Month(FechaGiro) & "/" & Year(FechaGiro))
                    FechaResult = DateAdd("m", 1, FechaGiro)
                End If
            Else
                FechaResult = CDate(DiaPago1 & "/" & Month(FechaGiro) & "/" & Year(FechaGiro))
                FechaResult = DateAdd("m", 1, FechaGiro)
            End If
        Else
            FechaResult = CDate(DiaPago1 & "/" & Month(FechaGiro) & "/" & Year(FechaGiro))
            FechaResult = DateAdd("m", 1, FechaGiro)
        End If
        
    End If

    CalcularDiasPago = FechaResult
End Function

Public Function RedondearPVP(dblPrecio As Double) As Double
Dim dblDecimales As Double
Dim dblRedondeo As Double
Dim dblEntera As Double

    dblEntera = Int(dblPrecio)
    dblDecimales = dblPrecio - dblEntera

    Select Case dblDecimales
    Case Is > 0.75
        dblRedondeo = dblEntera + 1
    Case Is > 0.5
        dblRedondeo = dblEntera + 0.75
    Case Is > 0.25
        dblRedondeo = dblEntera + 0.5
    Case Is > 0
        dblRedondeo = dblEntera + 0.25
    Case Else
        dblRedondeo = dblEntera
    End Select
    
    RedondearPVP = dblRedondeo

End Function

' Obtiene una propuesta de código para denominar automáticamente a los modelos y series.
' Incrementa el código en una letra y/o en un numero
Public Function ObtenerSiguienteCodigo(strCodigoActual As String, intLongitudCodigo As Integer) As String
    Dim intCaracterFijo As Integer
    Dim intPrimerCaracter As Integer
    Dim intSegundoCaracter As Integer
    
    Select Case intLongitudCodigo
    Case 2
        intPrimerCaracter = Asc(UCase(Left(strCodigoActual, 1)))
        intSegundoCaracter = Asc(UCase(Mid(strCodigoActual, 2, 1)))
    Case 3
        intCaracterFijo = Asc(UCase(Left(strCodigoActual, 1)))
        intPrimerCaracter = Asc(UCase(Mid(strCodigoActual, 2, 1)))
        intSegundoCaracter = Asc(UCase(Mid(strCodigoActual, 3, 1)))
    Case Else
        Err.Raise vbObjectError + 1001, "ObtenerSiguienteCodigo", "Longitud de código erronea:" & intLongitudCodigo & ". Sólo puede ser 2 ó 3."
    End Select
    
    intSegundoCaracter = intSegundoCaracter + 1
    If intSegundoCaracter > 57 And intSegundoCaracter < 65 Then
        intSegundoCaracter = 65
    End If
    If intSegundoCaracter > 90 Then
        intSegundoCaracter = 48
        intPrimerCaracter = intPrimerCaracter + 1
    End If
    If intPrimerCaracter > 57 And intPrimerCaracter < 65 Then
        intPrimerCaracter = 65
    End If
    If intPrimerCaracter > 90 Then
        intPrimerCaracter = 48
    End If
        
    Select Case intLongitudCodigo
    Case 2
        ObtenerSiguienteCodigo = Chr(intPrimerCaracter) & Chr(intSegundoCaracter)
    Case 3
        ObtenerSiguienteCodigo = Chr(intCaracterFijo) & Chr(intPrimerCaracter) & Chr(intSegundoCaracter)
    Case Else
        Err.Raise vbObjectError + 1001, "ObtenerSiguienteCodigo", "Longitud de código erronea:" & intLongitudCodigo & ". Sólo puede ser 2 ó 3."
    End Select
End Function
