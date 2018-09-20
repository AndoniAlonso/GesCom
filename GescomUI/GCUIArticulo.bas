Attribute VB_Name = "GCUIArticulo"
Option Explicit
Private Const cnSeparadorReferenciaProveedor = "."

Private Function ArticuloBase(CodigoArticulo As String) As String

    If Len(Trim(CodigoArticulo)) <> 8 Then Err.Raise vbObjectError + 1001, "El codigo de articulo no tiene una longitud de 8 caracteres."
    
    ArticuloBase = Mid(Trim(CodigoArticulo), 1, 6)
    
End Function

Public Function ValidarCodigoArticulo(ByVal CodigoArticulo As String, TemporadaID As Long) As Boolean
    Dim Respuesta As VbMsgBoxResult
    Dim objArticulo As Articulo
    Dim frmCrearArticuloEdit As CrearArticuloEdit
    
    ValidarCodigoArticulo = False
    
    Set objArticulo = New Articulo
    
    'Si existe el articulo, ya vemos que sí está validado.
    If objArticulo.ExisteCodigo(ArticuloBase(CodigoArticulo), TemporadaID) Then
        ValidarCodigoArticulo = True
        Set objArticulo = Nothing
        Exit Function
    End If
    
    Respuesta = MostrarMensaje(MSG_CREAR_ARTICULO)
        
    ' Si no se quiere crear el articulo no hacer nada.
    If Respuesta = vbNo Then
        ValidarCodigoArticulo = False
        Set objArticulo = Nothing
        Exit Function
    End If
    
    ' Se intenta crear el articulo, si falla saldrá por error al modulo llamante.
    objArticulo.CreateCodigoArticulo CodigoArticulo, TemporadaID
    
    'Se modifica el precio del articulo.
    Set frmCrearArticuloEdit = New CrearArticuloEdit
    frmCrearArticuloEdit.SetFocusPrecioVenta
    frmCrearArticuloEdit.Component objArticulo
    frmCrearArticuloEdit.Show vbModal
    
    Unload frmCrearArticuloEdit
    
    Set frmCrearArticuloEdit = Nothing
    'objArticulo.BeginEdit
    'objArticulo.PrecioVenta = InputBox("Introducir el precio de venta:", "Modificar el precio de venta del articulo " & CodigoArticulo, _
    '         objArticulo.PrecioVenta)
    'objArticulo.PrecioVentaPublico = objArticulo.CalcularPrecioVentaPublico
    'objArticulo.PrecioVentaPublico = InputBox("Introducir el PVP:", "Modificar el PVP del artículo " & CodigoArticulo, _
    '         objArticulo.PrecioVentaPublico)
    'objArticulo.ApplyEdit
    Set objArticulo = Nothing
    ValidarCodigoArticulo = True

End Function

' La referencia de proveedor tiene la siguiente estructura: TIPOPRENDA.MODELO.SERIE.COLOR(temporada, proveedor).
' La referencia base de proveedor tiene la siguiente estructura: TIPOPRENDA.MODELO.SERIE(temporada, proveedor)
Private Function ArticuloBaseProveedor(ReferenciaProveedor As String) As String
    Dim strResulArr() As String

    If Not EsFormatoArticuloProveedor(ReferenciaProveedor) Then
        Err.Raise vbObjectError + 1001, "ArticuloBaseProveedor", "El código de proveedor '" & ReferenciaProveedor & "' no tiene el formato MODELO.SERIE.COLOR"
        Exit Function
    End If
    
    strResulArr = Split(ReferenciaProveedor, cnSeparadorReferenciaProveedor)
    ArticuloBaseProveedor = strResulArr(0) + cnSeparadorReferenciaProveedor + strResulArr(1) + cnSeparadorReferenciaProveedor + strResulArr(2)
    
End Function

' La referencia de proveedor tiene la siguiente estructura: TIPOPRENDA.MODELO.SERIE.COLOR(temporada, proveedor).
' El color base de proveedor tiene la siguiente estructura: COLOR(temporada, proveedor)
Public Function ColorBaseProveedor(ReferenciaProveedor As String) As String
    Dim strResultado As String

    strResultado = ColorOriginalProveedor(ReferenciaProveedor)
    
    If Len(strResultado) = 3 And IsNumeric(strResultado) Then
        strResultado = CompactaCodigoColor(strResultado)
    End If
    
    ColorBaseProveedor = strResultado
    
End Function


' La referencia de proveedor tiene la siguiente estructura: TIPOPRENDA.MODELO.SERIE.COLOR(temporada, proveedor).
' El color base de proveedor tiene la siguiente estructura: COLOR(temporada, proveedor)
Public Function ColorOriginalProveedor(ReferenciaProveedor As String) As String
    Dim strResulArr() As String
    Dim strResultado As String

    If Not EsFormatoArticuloProveedor(ReferenciaProveedor) Then
        Err.Raise vbObjectError + 1001, "ColorBaseProveedor", "El código de proveedor '" & ReferenciaProveedor & "' no tiene el formato MODELO.SERIE.COLOR"
        Exit Function
    End If
    
    strResulArr = Split(ReferenciaProveedor, cnSeparadorReferenciaProveedor)
    
    strResultado = strResulArr(3)
    
    ' Un código de color no puede tener más de tres caracteres
    If Len(strResultado) > 3 Then
        Err.Raise vbObjectError + 1001, "ColorBaseProveedor", "El código de color no puede tener más de tres caracteres: '" & ReferenciaProveedor & "'."
        Exit Function
    End If
    
    ' Un código de color de más de dos caracteres debe ser numérico
    If Len(strResultado) = 3 And Not IsNumeric(strResultado) Then
        Err.Raise vbObjectError + 1001, "ColorBaseProveedor", "El código de color de tres caracteres debe ser numérico: '" & ReferenciaProveedor & "'."
        Exit Function
    End If

    ColorOriginalProveedor = strResultado
    
End Function

Private Function CompactaCodigoColor(strCodigoColor As String) As String
    Dim intPrimerCaracter As Integer
    Dim intSegundoCaracter As Integer
    Dim strPrimerCaracter As String
    Dim strSegundoCaracter As String
    
    intPrimerCaracter = CInt(strCodigoColor) \ 39  'Se divide entre 39
    strPrimerCaracter = Chr(intPrimerCaracter + 65) ' Se convierte a una letra que se corresponde con la posicion de la letra en el alfabeto
    
    intSegundoCaracter = CInt(strCodigoColor) Mod 39
    Select Case intSegundoCaracter
    Case Is = 0
        strSegundoCaracter = "+"
    Case Is = 1
        strSegundoCaracter = "-"
    Case Is = 2
        strSegundoCaracter = "*"
    Case 3 To 12
        strSegundoCaracter = CStr(intSegundoCaracter - 3)
    Case 13 To 38
        strSegundoCaracter = Chr(52 + intSegundoCaracter)
    Case Else
        Err.Raise vbObjectError + 1001, "GCUIArticulo CompactaCodigoColor", "Error, no se ha podido traducir el segundo caracter de un código color."
    End Select
    
    CompactaCodigoColor = strPrimerCaracter & strSegundoCaracter
End Function

Public Property Get EsFormatoArticuloProveedor(ReferenciaProveedor As String) As Boolean
    
    On Error GoTo ErrorManager

    Dim strResulArr() As String

    strResulArr = Split(ReferenciaProveedor, cnSeparadorReferenciaProveedor)
    
    'Validamos el formato de proveedor si tiene 4 campos separados por "." y el primer campo tiene longitud 1. y el segundo campo tiene longitud >= 3
    EsFormatoArticuloProveedor = (UBound(strResulArr) = 3) And (Len(strResulArr(0)) = 1) And (Len(strResulArr(1)) >= 3)
    Exit Property

ErrorManager:
    EsFormatoArticuloProveedor = False
End Property

' A partir de una referencia de proveedor:
' -Se comprueba a ver si existe ya el artículo en función de esa referencia.
' -
Public Function ValidarArticuloProveedor(ReferenciaProveedor As String, TemporadaID As Long, ProveedorID As Long) As Articulo
    Dim Respuesta As VbMsgBoxResult
    Dim objArticulo As Articulo
    Dim frmCrearArticuloEdit As CrearArticuloEdit
    
    Set ValidarArticuloProveedor = Nothing
    
    Set objArticulo = New Articulo
    If Not EsFormatoArticuloProveedor(ReferenciaProveedor) Then
        Set objArticulo = Nothing
        Set ValidarArticuloProveedor = objArticulo
        Exit Function
    End If
    
    'Si existe el articulo, ya vemos que sí está validado.
    If objArticulo.ExisteArticuloProveedor(ArticuloBaseProveedor(ReferenciaProveedor), TemporadaID, ProveedorID) Then
        CreateCodigoArticuloColorProveedor objArticulo.ArticuloID, objArticulo.Nombre, ReferenciaProveedor, TemporadaID
        Set ValidarArticuloProveedor = objArticulo
        Set objArticulo = Nothing
        Exit Function
    End If
    
    Respuesta = MostrarMensaje(MSG_CREAR_ARTICULO)
        
    ' Si no se quiere crear el articulo no hacer nada.
    If Respuesta = vbNo Then
        Set ValidarArticuloProveedor = objArticulo
        Set objArticulo = Nothing
        Exit Function
    End If
    
    ' Se intenta crear el articulo, si falla saldrá por error al modulo llamante.
    objArticulo.CreateCodigoArticuloproveedor ArticuloBaseProveedor(ReferenciaProveedor), TemporadaID, ProveedorID

    Set ValidarArticuloProveedor = objArticulo
    
'    'Se modifica el precio del articulo.
'    objArticulo.BeginEdit
'    'objArticulo.PrecioVenta = InputBox("Introducir el precio de venta:", "Modificar el precio de venta del articulo " & objArticulo.Nombre, _
'    '         objArticulo.PrecioVenta)
'    objArticulo.PrecioCompra = InputBox("Introducir el precio de COMPRA:", "Modificar el precio de COMPRA del articulo " & ReferenciaProveedor, _
'             objArticulo.PrecioCompra)
'    objArticulo.PrecioVenta = objArticulo.CalcularPrecioVenta
'    objArticulo.PrecioVenta = InputBox("Introducir el PVP:", "Modificar el PVP del artículo " & ReferenciaProveedor, _
'             objArticulo.PrecioVenta)
'    objArticulo.PrecioVentaPublico = objArticulo.PrecioVenta  'objArticulo.CalcularPrecioVentaPublico
'    'objArticulo.PrecioVentaPublico = InputBox("Introducir el PVP:", "Modificar el PVP del artículo " & ReferenciaProveedor, _
'    '         objArticulo.PrecioVentaPublico)
'    objArticulo.ApplyEdit
    'Se modifica el precio del articulo.
    Set frmCrearArticuloEdit = New CrearArticuloEdit
    frmCrearArticuloEdit.SetFocusPrecioCompra
    frmCrearArticuloEdit.Component objArticulo
    frmCrearArticuloEdit.Show vbModal
    
    Unload frmCrearArticuloEdit
    
    Set frmCrearArticuloEdit = Nothing


    ' Creamos el código de artículo
    CreateCodigoArticuloColorProveedor objArticulo.ArticuloID, objArticulo.Nombre, ReferenciaProveedor, TemporadaID
    
    Set ValidarArticuloProveedor = objArticulo
    Set objArticulo = Nothing

End Function

Private Sub CreateCodigoArticuloColorProveedor(ArticuloID As Long, CodigoArticulo As String, ReferenciaProveedor As String, TemporadaID As Long)
    Dim objArticuloColor As ArticuloColor
    Dim strCodigoArticuloColor As String
    Dim strColorBaseProveedor As String
    
    strColorBaseProveedor = ColorBaseProveedor(ReferenciaProveedor)
    strCodigoArticuloColor = CodigoArticulo & strColorBaseProveedor
    
    Set objArticuloColor = New ArticuloColor
    With objArticuloColor
        If Not .ExisteCodigo(strCodigoArticuloColor, TemporadaID) Then
                .BeginEdit
                .TemporadaID = TemporadaID
                .ArticuloID = ArticuloID
                .Codigo = strColorBaseProveedor
                .NombreColor = ColorOriginalProveedor(ReferenciaProveedor)
                .ApplyEdit
        End If
    End With
    Set objArticuloColor = Nothing

End Sub
