VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Articulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As ArticuloProps
Private mudtSave As ArticuloProps

' Aqui declaro el ComboBox de prendas, modelos y series
Private mobjPrendas As TextList
Private mobjModelos As TextList
Private mobjSeries As TextList
Private mobjTallajes As TextList

Event Valid(IsValid As Boolean)

Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean

Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private Const cnSeparadorReferenciaProveedor = "."

Public Sub BeginEdit()

    If mflgEditing Then Err.Raise 445
  
    ' save object state
    LSet mudtSave = mudtprops
    mflgEditing = True

End Sub

Public Sub CancelEdit()

    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False
    mflgDeleted = False
    mflgDirty = False
    
    ' restore object state
    LSet mudtprops = mudtSave
  
End Sub

Public Sub ApplyEdit()

    Dim objPersist As ArticuloPersist
  
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("GCServerMTS.ArticuloPersist", PERSIST_SERVER)

    If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.ArticuloID
        mflgNew = True
        mflgDeleted = False

    ElseIf mflgDirty Or mflgNew Then
        If Not IsValid Then Err.Raise 445
        ' save object to database if appropriate
        
        If PrecioCoste = 0 Or IsNew Then _
            PrecioCoste = CalcularPrecioCoste
            
        'If PrecioCompra = 0 And IsNew Then _
        '    PrecioCompra = PrecioCoste
            
        ' Calculamos el precio de venta cuando el articulo es nuevo o es cero.
        If PrecioVenta = 0 Or IsNew Then _
            PrecioVenta = CalcularPrecioVenta

        ' Calculamos el PVP cuando el articulo es nuevo o es cero.
        If PrecioVentaPublico = 0 Or IsNew Then _
            PrecioVentaPublico = CalcularPrecioVentaPublico

        SetState objPersist.Save(GetState)
        ' save object state
        LSet mudtSave = mudtprops
        mflgNew = False

    End If

    Set objPersist = Nothing
    mflgDirty = False
    mflgEditing = False
  
End Sub

Private Sub Class_Initialize()
  
    mflgNew = True
    Set mobjValid = New BrokenRules
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
    '
    mobjValid.RuleBroken "TemporadaID", True
    mobjValid.RuleBroken "PrendaID", True
    mobjValid.RuleBroken "ModeloID", True
    mobjValid.RuleBroken "SerieID", True
      
    Set mobjPrendas = Nothing
    Set mobjModelos = Nothing
    Set mobjSeries = Nothing
    Set mobjTallajes = Nothing
    
    mudtprops.StockActual = 0
    mudtprops.StockPendiente = 0
    mudtprops.StockMinimo = 0
    mudtprops.StockMaximo = 0
    mudtprops.LoteEconomico = 0
    mudtprops.PrecioCosteEUR = 0
    mudtprops.PrecioCompraEUR = 0
    mudtprops.PrecioVentaEUR = 0
    mudtprops.PrecioVentaPublico = 0
    mudtprops.TemporadaID = 0
    mudtprops.PrendaID = 0
    mudtprops.ModeloID = 0
    mudtprops.SerieID = 0
    mudtprops.SuReferencia = vbNullString
    mudtprops.ProveedorID = 0
    mudtprops.TallajeID = 0
      
End Sub

Private Sub Class_Terminate()
  
    Set mobjValid = Nothing
    Set mobjPrendas = Nothing
    Set mobjModelos = Nothing
    Set mobjSeries = Nothing
    Set mobjTallajes = Nothing
    
End Sub

Public Property Get IsValid() As Boolean

    IsValid = (mobjValid.Count = 0)

End Property

Private Sub mobjValid_BrokenRule()

    RaiseEvent Valid(False)

End Sub

Private Sub mobjValid_NoBrokenRules()

    RaiseEvent Valid(True)

End Sub

Public Sub Load(ArticuloID As Long)
    Dim objPersist As ArticuloPersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.ArticuloPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(ArticuloID)
    Set objPersist = Nothing

    mobjValid.RuleBroken "TemporadaID", False
    mobjValid.RuleBroken "PrendaID", False
    mobjValid.RuleBroken "ModeloID", False
    mobjValid.RuleBroken "SerieID", False
      
    mflgNew = False

End Sub

Public Sub LoadCodigo(Optional Codigo As String, _
                Optional TemporadaID As Long)

    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445

    ' Si el codigo buscado existe se asigna como esté en la tabla.
    If ExisteCodigo(Codigo, TemporadaID) Then
        ' El articulo existe y se ha establecido

    Else
        ' Si no existe hay que dar error
        Err.Raise vbObjectError + 1001, "Error leyendo el articulo " & Codigo, _
                                        "El articulo no existe."
    End If

    ' code to load the object goes here

    mobjValid.RuleBroken "TemporadaID", False
    mobjValid.RuleBroken "PrendaID", False
    mobjValid.RuleBroken "ModeloID", False
    mobjValid.RuleBroken "SerieID", False

    mflgNew = False

End Sub

Public Sub CreateCodigoArticulo(Optional Codigo As String, _
                Optional TemporadaID As Long)
    Dim mintPrendaID As Long
    Dim mintModeloID As Long
    Dim mintSerieID As Long
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' Si el codigo buscado se da error porque se quiere crear.
    If ExisteCodigo(Codigo, TemporadaID) Then
        ' El articulo existe y se ha establecido
            Err.Raise vbObjectError + 1001, "Error creando articulo " & Codigo, _
                                            "El articulo ya existe."

        ' Si no existe hay que crearlo en base a prenda-modelo-serie
    Else
        ' Primero buscamos la prenda.
        mintPrendaID = CodigoPrendaID(Mid(Codigo, 1, 1))
        If mintPrendaID = 0 Then
            Err.Raise vbObjectError + 1001, "Error creando articulo " & Codigo, _
                                            "No existe el tipo de prenda " & Mid(Codigo, 1, 1)
        End If
        
        ' A continuacion buscamos el modelo.
        mintModeloID = CodigoModeloID(Mid(Codigo, 2, 3), TemporadaID)
        If mintModeloID = 0 Then
            Err.Raise vbObjectError + 1001, "Error creando articulo " & Codigo, _
                                            "No existe el Modelo " & Mid(Codigo, 2, 3)
        End If
        
        ' Por ultimo buscamos la serie.
        mintSerieID = CodigoSerieID(Mid(Codigo, 5, 2), TemporadaID)
        If mintSerieID = 0 Then
            Err.Raise vbObjectError + 1001, "Error creando articulo " & Codigo, _
                                            "No existe la Serie " & Mid(Codigo, 5, 2)
        End If
        Me.BeginEdit
        Me.TemporadaID = TemporadaID
        mudtprops.PrendaID = mintPrendaID
        mobjValid.RuleBroken "PrendaID", False
        
        mudtprops.ModeloID = mintModeloID
        mobjValid.RuleBroken "ModeloID", False
        
        mudtprops.SerieID = mintSerieID
        mobjValid.RuleBroken "SerieID", False
        
        Me.AsignarTallajePredeterminado
        
        Me.ApplyEdit
    End If
    
    ' code to load the object goes here

    mobjValid.RuleBroken "TemporadaID", False
    mobjValid.RuleBroken "PrendaID", False
    mobjValid.RuleBroken "ModeloID", False
    mobjValid.RuleBroken "SerieID", False
      
    mflgNew = False

End Sub


Public Sub CreateCodigoArticuloProveedor(ReferenciaProveedor As String, _
                TemporadaID As Long, ProveedorID As Long)
    Dim intPrendaID As Long
    Dim intModeloID As Long
    Dim intSerieID As Long
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' Si existe el codigo buscado se da error porque se quiere crear.
    If ExisteArticuloProveedor(ReferenciaProveedor, TemporadaID, ProveedorID) Then
        ' El articulo existe y se ha establecido
            Err.Raise vbObjectError + 1001, "Error creando articulo " & ReferenciaProveedor, _
                                            "El articulo ya existe."

        ' Si no existe hay que crearlo en base a prenda-modelo-serie
    Else
        
        ' Primero buscamos la prenda.
        intPrendaID = CodigoPrendaID(PrendaBaseProveedor(ReferenciaProveedor))
        If intPrendaID = 0 Then
            Err.Raise vbObjectError + 1001, "Error creando articulo " & ReferenciaProveedor, _
                                            "No existe el tipo de prenda " & PrendaBaseProveedor(ReferenciaProveedor)
        End If

        ' A continuacion buscamos el modelo.
        intModeloID = FindOrCreateCodigoModeloProveedor(ReferenciaProveedor, TemporadaID)
        If intModeloID = 0 Then
            Err.Raise vbObjectError + 1001, "Error creando articulo " & ReferenciaProveedor, _
                                            "No existe el Modelo " & Mid(ReferenciaProveedor, 2, 3)
        End If
        
        ' Por ultimo buscamos la serie.
        intSerieID = FindOrCreateCodigoSerieProveedor(ReferenciaProveedor, TemporadaID)
        If intSerieID = 0 Then
            Err.Raise vbObjectError + 1001, "Error creando articulo " & ReferenciaProveedor, _
                                            "No existe la Serie " & ReferenciaProveedor
        End If
        
        Me.BeginEdit
        mudtprops.TemporadaID = TemporadaID
        mobjValid.RuleBroken "TemporadaID", False
        
        mudtprops.PrendaID = intPrendaID
        mobjValid.RuleBroken "PrendaID", False
        
        mudtprops.ModeloID = intModeloID
        mobjValid.RuleBroken "ModeloID", False
        
        mudtprops.SerieID = intSerieID
        mobjValid.RuleBroken "SerieID", False
        
        mudtprops.ProveedorID = ProveedorID
        
        mudtprops.SuReferencia = ReferenciaProveedor
        
        Me.AsignarTallajePredeterminado
        
        Me.ApplyEdit
    End If
    
End Sub

'Devuelve el identificador de la prenda que queremos comprobar si existe.
Private Function CodigoPrendaID(Codigo As String) As Long
    Dim objMyPrenda As Prenda
    
    On Error GoTo ErrorManager

    Set objMyPrenda = New Prenda
    objMyPrenda.Load 0, "EUR", Codigo
    
    CodigoPrendaID = objMyPrenda.PrendaID
    Set objMyPrenda = Nothing
    
    Exit Function

ErrorManager:
    CodigoPrendaID = 0
    Exit Function
    
End Function

'Devuelve el identificador del Modelo que queremos comprobar si existe.
Private Function CodigoModeloID(Codigo As String, TemporadaID As Long) As Long
    Dim objMyModelo As Modelo
    
    On Error GoTo ErrorManager

    Set objMyModelo = New Modelo
    objMyModelo.Load 0, Codigo, TemporadaID, vbNullString
    
    CodigoModeloID = objMyModelo.ModeloID
    Set objMyModelo = Nothing
    
    Exit Function

ErrorManager:
    CodigoModeloID = 0
    Exit Function
    
End Function

Private Function FindOrCreateCodigoModeloProveedor(ReferenciaProveedor As String, TemporadaID As Long) As Long
    Dim objMyModelo As Modelo
    Dim strNombreModelo As String
    
    On Error GoTo ErrorManager

    strNombreModelo = ModeloBaseProveedor(ReferenciaProveedor)

    Set objMyModelo = New Modelo
    
    On Error Resume Next
    objMyModelo.Load 0, vbNullString, TemporadaID, strNombreModelo
    On Error GoTo ErrorManager
    
    'Comprobar si no existe el modelo buscado
    If objMyModelo.ModeloID = 0 Then
        Set objMyModelo = New Modelo
        objMyModelo.BeginEdit
        objMyModelo.Nombre = strNombreModelo
        objMyModelo.Codigo = objMyModelo.CrearCodigoAutomatico(strNombreModelo, TemporadaID)
        objMyModelo.TemporadaID = TemporadaID
        objMyModelo.ApplyEdit
    End If
    
    FindOrCreateCodigoModeloProveedor = objMyModelo.ModeloID
    
    Set objMyModelo = Nothing
    
    Exit Function

ErrorManager:
    FindOrCreateCodigoModeloProveedor = 0
    Exit Function
    
End Function

Private Function FindOrCreateCodigoSerieProveedor(ReferenciaProveedor As String, TemporadaID As Long) As Long
    Dim objMySerie As Serie
    Dim strNombreSerie As String
    
    On Error GoTo ErrorManager

    strNombreSerie = SerieBaseProveedor(ReferenciaProveedor)

    Set objMySerie = New Serie
    
    On Error Resume Next
    objMySerie.Load 0, vbNullString, TemporadaID, strNombreSerie
    On Error GoTo ErrorManager
    
    'Comprobar si no existe el Serie buscado
    If objMySerie.SerieID = 0 Then
        Set objMySerie = New Serie
        objMySerie.BeginEdit
        objMySerie.Nombre = strNombreSerie
        objMySerie.Codigo = objMySerie.CrearCodigoAutomatico(strNombreSerie, TemporadaID)
        objMySerie.TemporadaID = TemporadaID
        objMySerie.AsignarMaterialPredeterminado
        objMySerie.ApplyEdit
    End If
    
    FindOrCreateCodigoSerieProveedor = objMySerie.SerieID
    
    Set objMySerie = Nothing
    
    Exit Function

ErrorManager:
    FindOrCreateCodigoSerieProveedor = 0
    Exit Function
    
End Function

' La referencia de proveedor tiene la siguiente estructura: TIPOPRENDA.MODELO.SERIE.COLOR(temporada, proveedor).
' El modelo base de proveedor tiene la siguiente estructura: TIPOPRENDA(temporada, proveedor)
Private Function PrendaBaseProveedor(ReferenciaProveedor As String) As String
    Dim strResulArr() As String

    strResulArr = Split(ReferenciaProveedor, cnSeparadorReferenciaProveedor)
    PrendaBaseProveedor = UCase(strResulArr(0))
    
End Function

' La referencia de proveedor tiene la siguiente estructura: TIPOPRENDA.MODELO.SERIE.COLOR(temporada, proveedor).
' El modelo base de proveedor tiene la siguiente estructura: MODELO(temporada, proveedor)
Private Function ModeloBaseProveedor(ReferenciaProveedor As String) As String
    Dim strResulArr() As String

    strResulArr = Split(ReferenciaProveedor, cnSeparadorReferenciaProveedor)
    ModeloBaseProveedor = UCase(strResulArr(1))
    
End Function

' La referencia de proveedor tiene la siguiente estructura: MODELO.SERIE.COLOR(temporada, proveedor).
' Serie base de proveedor tiene la siguiente estructura: Serie(temporada, proveedor)
Private Function SerieBaseProveedor(ReferenciaProveedor As String) As String
    Dim strResulArr() As String

    strResulArr = Split(ReferenciaProveedor, cnSeparadorReferenciaProveedor)
    SerieBaseProveedor = UCase(strResulArr(2))
    
End Function


'Devuelve el identificador de la serie que queremos comprobar si existe.
Private Function CodigoSerieID(Codigo As String, TemporadaID As Long) As Long
    Dim objMySerie As Serie
    
    On Error GoTo ErrorManager

    Set objMySerie = New Serie
    objMySerie.Load 0, Codigo, TemporadaID
    
    CodigoSerieID = objMySerie.SerieID
    Set objMySerie = Nothing
    
    Exit Function

ErrorManager:
    CodigoSerieID = 0
End Function

'Private Function ExisteCodigo(Codigo As String, TemporadaID As Long) As Boolean
Public Property Get ExisteCodigo(Codigo As String, TemporadaID As Long) As Boolean
    Dim objPersist As ArticuloPersist
    
    On Error GoTo ErrorManager

    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.ArticuloPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(0, Codigo, TemporadaID)
    Set objPersist = Nothing
    
    ExisteCodigo = True
    
    Exit Property

ErrorManager:
    ExisteCodigo = False
End Property

Public Property Get ExisteArticuloProveedor(ReferenciaProveedor As String, TemporadaID As Long, ProveedorID As Long) As Boolean
    Dim objPersist As ArticuloPersist
    
    On Error GoTo ErrorManager

    Set objPersist = CreateObject("GCServerMTS.ArticuloPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(0, vbNullString, TemporadaID, ReferenciaProveedor)  ', ProveedorID)
    Set objPersist = Nothing
    
    ExisteArticuloProveedor = True
    
    Exit Property

ErrorManager:
    ExisteArticuloProveedor = False
End Property

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
  
    mflgDeleted = True

End Sub

Public Property Get IsDeleted() As Boolean

    IsDeleted = mflgDeleted

End Property

Public Property Get IsNew() As Boolean

    IsNew = mflgNew

End Property

Public Property Get IsDirty() As Boolean

    IsDirty = mflgDirty

End Property

Public Property Get ArticuloID() As Long

    ArticuloID = mudtprops.ArticuloID
  
End Property

Public Property Let Nombre(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Nombre) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    mobjValid.RuleBroken "Nombre", (Len(Trim(Value)) = 0)

    mudtprops.Nombre = Value
    mflgDirty = True

End Property

Public Property Get Nombre() As String

    Nombre = Trim(mudtprops.Nombre)

End Property

Public Property Let Prenda(Value As String)

    If Not mflgEditing Then Err.Raise 383
    If Not mflgNew Then Err.Raise 383
  
    If mobjPrendas Is Nothing Then CargarListaPrendas
    
    mudtprops.PrendaID = mobjPrendas.Key(Value)
  
    mobjValid.RuleBroken "PrendaID", (mudtprops.PrendaID = 0)
  
    mflgDirty = True

End Property

Public Property Get Prenda() As String

    If mobjPrendas Is Nothing Then CargarListaPrendas
    
    Prenda = mobjPrendas.Item(mudtprops.PrendaID)

End Property

Public Property Get Prendas() As TextList

    If mobjPrendas Is Nothing Then CargarListaPrendas
    
    Set Prendas = mobjPrendas

End Property

Public Property Get PrendaID() As Long

    PrendaID = mudtprops.PrendaID

End Property

Public Property Let Modelo(Value As String)

    If Not mflgEditing Then Err.Raise 383
    If Not mflgNew Then Err.Raise 383
  
    If mobjModelos Is Nothing Then CargarListaModelos
    
    mudtprops.ModeloID = mobjModelos.Key(Value)
  
    mobjValid.RuleBroken "ModeloID", (mudtprops.ModeloID = 0)
  
    mflgDirty = True

End Property

Public Property Get Modelo() As String

    If mobjModelos Is Nothing Then CargarListaModelos
    
    Modelo = mobjModelos.Item(mudtprops.ModeloID)

End Property

Public Property Get Modelos() As TextList

    If mobjModelos Is Nothing Then CargarListaModelos
    
    Set Modelos = mobjModelos

End Property

Public Property Get ModeloID() As Long

    ModeloID = mudtprops.ModeloID

End Property

Public Property Let Serie(Value As String)

    If Not mflgEditing Then Err.Raise 383
    If Not mflgNew Then Err.Raise 383
  
    If mobjSeries Is Nothing Then CargarListaSeries
    
    mudtprops.SerieID = mobjSeries.Key(Value)
  
    mobjValid.RuleBroken "SerieID", (mudtprops.SerieID = 0)
  
    mflgDirty = True

End Property

Public Property Get Serie() As String

    If mobjSeries Is Nothing Then CargarListaSeries
    
    Serie = mobjSeries.Item(mudtprops.SerieID)

End Property

Public Property Get Series() As TextList

    If mobjSeries Is Nothing Then CargarListaSeries
    
    Set Series = mobjSeries

End Property

Public Property Get SerieID() As Long

    SerieID = mudtprops.SerieID

End Property

Public Property Let Tallaje(Value As String)

    If Not mflgEditing Then Err.Raise 383
    'If Not mflgNew Then Err.Raise 383
    ' OJOOJOOJO: para permitir cambiar un tallaje a un artículo quizás habría que comprobar que no tiene
    ' existencias en ninguna de las tallas que no están permitidas en el nuevo tallaje.
    ' Por ejemplo si pasas algo a talla única y ahora tiene existencias en talla 48, que hay que hacer? preguntar?, no dejar?
  
    If mobjTallajes Is Nothing Then CargarListaTallajes
    
    mudtprops.TallajeID = mobjTallajes.Key(Value)
  
    mobjValid.RuleBroken "TallajeID", (mudtprops.TallajeID = 0)
  
    mflgDirty = True

End Property

Public Property Get Tallaje() As String

    If mobjTallajes Is Nothing Then CargarListaTallajes
    
    Tallaje = mobjTallajes.Item(mudtprops.TallajeID)

End Property

Public Property Get Tallajes() As TextList

    If mobjTallajes Is Nothing Then CargarListaTallajes
    
    Set Tallajes = mobjTallajes

End Property

Public Property Get TallajeID() As Long

    TallajeID = mudtprops.TallajeID

End Property

Public Property Get NombrePrenda() As String

    NombrePrenda = Trim(mudtprops.NombrePrenda)

End Property

Public Property Get NombreModelo() As String

    NombreModelo = Trim(mudtprops.NombreModelo)

End Property

Public Property Get NombreSerie() As String

    NombreSerie = Trim(mudtprops.NombreSerie)

End Property

Public Property Get NombreCompleto() As String

    NombreCompleto = NombrePrenda & " " & NombreModelo & " " & NombreSerie

End Property

Public Property Let PrecioCoste(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
  
    mobjValid.RuleBroken "PrecioCoste", (Value < 0)
  
    mudtprops.PrecioCosteEUR = Round(Value, 2)
    
    mflgDirty = True

End Property

Public Property Get PrecioCoste() As Double

    PrecioCoste = mudtprops.PrecioCosteEUR
  
End Property

Public Property Let PrecioCompra(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
  
    mobjValid.RuleBroken "PrecioCompra", (Value < 0)
  
    mudtprops.PrecioCompraEUR = Round(Value, 2)
    
    mflgDirty = True

End Property

Public Property Get PrecioCompra() As Double

    PrecioCompra = mudtprops.PrecioCompraEUR
  
End Property

Public Property Let PrecioVenta(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
  
    mobjValid.RuleBroken "PrecioVenta", (Value < 0)
  
    mudtprops.PrecioVentaEUR = Round(Value, 2)
    
    mflgDirty = True

End Property

Public Property Get PrecioVenta() As Double

    PrecioVenta = mudtprops.PrecioVentaEUR
 
End Property

Public Property Let PrecioVentaPublico(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
  
    mobjValid.RuleBroken "PrecioVentaPublico", (Value < 0)
  
    mudtprops.PrecioVentaPublico = Value
    
    mflgDirty = True

End Property

Public Property Get PrecioVentaPublico() As Double

    PrecioVentaPublico = mudtprops.PrecioVentaPublico
 
End Property

Friend Property Let StockActual(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    mudtprops.StockActual = Round(Value, 2)
    
    mflgDirty = True

End Property

' StockActual es una propiedad de solo lectura. Se mantiene a traves de metodos.
Public Property Get StockActual() As Double

    StockActual = mudtprops.StockActual
  
End Property

Friend Property Let StockPendiente(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    mudtprops.StockPendiente = Round(Value, 2)
    
    mflgDirty = True

End Property

' StockPendiente es una propiedad de solo lectura. Se mantiene a traves de metodos.
Public Property Get StockPendiente() As Double

    StockPendiente = mudtprops.StockPendiente
  
End Property

Public Property Let StockMinimo(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Stock no puede ser negativo"
  
    ' Se valida que no sean valores negativos y que no sea mayor que el stock maximo.
    mobjValid.RuleBroken "StockMinimo", (Value < 0)
    mobjValid.RuleBroken "RangoStocks", (Value > mudtprops.StockMaximo)
  
    mudtprops.StockMinimo = Round(Value, 2)
    
    mflgDirty = True

End Property

Public Property Get StockMinimo() As Double

    StockMinimo = mudtprops.StockMinimo
  
End Property

Public Property Let StockMaximo(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Stock no puede ser negativo"
  
    ' Se valida que no sean valores negativos y que no sea mayor que el stock maximo.
    mobjValid.RuleBroken "StockMaximo", (Value < 0)
    mobjValid.RuleBroken "RangoStocks", (Value < mudtprops.StockMinimo)
  
    mudtprops.StockMaximo = Round(Value, 2)
    
    mflgDirty = True

End Property

Public Property Get StockMaximo() As Double

    StockMaximo = mudtprops.StockMaximo
  
End Property
    
Public Property Let LoteEconomico(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Stock no puede ser negativo"
  
    ' Se valida que no sean valores negativos.
    mobjValid.RuleBroken "LoteEconomico", (Value < 0)
  
    mudtprops.LoteEconomico = Round(Value, 2)
    
    mflgDirty = True

End Property

Public Property Get LoteEconomico() As Double

    LoteEconomico = mudtprops.LoteEconomico
  
End Property

Public Property Get SuReferencia() As String

    SuReferencia = Trim(mudtprops.SuReferencia)
  
End Property

Public Property Let TemporadaID(Value As Long)

    If Not mflgNew Then Err.Raise 383
  
    mobjValid.RuleBroken "TemporadaID", (Value = 0)

    mudtprops.TemporadaID = Value
  
    Set mobjModelos = Nothing
    
    CargarListaModelos
    mudtprops.ModeloID = mobjModelos.Key(mobjModelos.IndexItem(1))
    
    Set mobjSeries = Nothing
    CargarListaSeries
    mudtprops.SerieID = mobjSeries.Key(mobjSeries.IndexItem(1))
    
    mflgDirty = True

End Property

Public Property Get TemporadaID() As Long

    TemporadaID = mudtprops.TemporadaID
  
End Property

Public Property Get ProveedorID() As Long

    ProveedorID = mudtprops.ProveedorID

End Property

Private Sub CargarListaPrendas()

    Set mobjPrendas = New TextList
    mobjPrendas.Load "Prendas"

End Sub
  
Private Sub CargarListaModelos()

    Set mobjModelos = New TextList
    mobjModelos.Load "Modelos", , mudtprops.TemporadaID

End Sub
  
Private Sub CargarListaSeries()
    
    Set mobjSeries = New TextList
    mobjSeries.Load "Series", , mudtprops.TemporadaID

End Sub
  
Private Sub CargarListaTallajes()
    
    Set mobjTallajes = New TextList
    mobjTallajes.Load "Tallajes"

End Sub
  
Public Function CalcularPrecioCoste() As Double
    Dim objPrenda As Prenda
    Dim objModelo As Modelo
    Dim objSerie As Serie
    Dim objMaterial As Material

    Dim CostePrenda As Double
    Dim CosteModelo As Double
    Dim CosteModeloSerie As Double
    Dim CosteArticulo As Double
    
    ' Inicializo objetos
    Set objPrenda = New Prenda
    Set objModelo = New Modelo
    Set objSerie = New Serie
    
    ' Cargo los datos de la base de datos
    objPrenda.Load mudtprops.PrendaID, "EUR"
    objModelo.Load mudtprops.ModeloID, vbNullString, 0, vbNullString
    objSerie.Load mudtprops.SerieID
    
    ' Asigno nombres de variables (no normalizadas).
    mudtprops.NombrePrenda = objPrenda.Nombre
    mudtprops.NombreModelo = objModelo.Nombre
    mudtprops.NombreSerie = objSerie.Nombre
    
    'Calculamos primero los costes fijos
    CostePrenda = objPrenda.PrecioCoste
    CosteModelo = objModelo.PrecioCoste
    
    'Obtenemos despues el coste de modelo y serie
    Set objMaterial = New Material
    objMaterial.Load objSerie.MaterialID, "EUR"
    
    ' Coste de la tela de la serie depende de la anchura de la tela
    ' A menos anchura más coste.
    If objMaterial.AnchuraTela <> 0 Then
        CosteModeloSerie = (objModelo.CantidadTela * objMaterial.PrecioCoste * dblAnchuraTelaEstandar) _
                        / objMaterial.AnchuraTela
    Else
        CosteModeloSerie = (objModelo.CantidadTela * objMaterial.PrecioCoste * dblAnchuraTelaEstandar) _
                        / objMaterial.AnchuraEstandar
    End If
                    
    Set objMaterial = Nothing
    
    CosteArticulo = CostePrenda + CosteModelo + CosteModeloSerie
    CosteArticulo = CosteArticulo + (CosteArticulo * objPrenda.Administracion) / 100
    
    ' Asigno el nombre al articulo (aprovecho que tengo todos los objetos cargados)
    If mflgNew Then _
       mudtprops.Nombre = objPrenda.Codigo & objModelo.Codigo & objSerie.Codigo
  
    ' Termino los objetos
    Set objPrenda = Nothing
    Set objModelo = Nothing
    Set objSerie = Nothing
    
    CalcularPrecioCoste = Round(CosteArticulo, 2)
 
End Function

Public Function CalcularPrecioVenta() As Double
    Dim objModelo As Modelo
    
    Set objModelo = New Modelo

    objModelo.Load mudtprops.ModeloID, vbNullString, 0, vbNullString
    
    CalcularPrecioVenta = Round(mudtprops.PrecioCosteEUR * (100 + objModelo.Beneficio) / 100, 2)
    
    Set objModelo = Nothing
End Function

Public Function CalcularPrecioVentaPublico() As Double
    Dim objModelo As Modelo
    
    Set objModelo = New Modelo

    objModelo.Load mudtprops.ModeloID, vbNullString, 0, vbNullString
    
    CalcularPrecioVentaPublico = RedondearPVP(Round(mudtprops.PrecioVentaEUR * (100 + objModelo.BeneficioPVP) / 100, 2))
    
    Set objModelo = Nothing
End Function

Public Property Get objModelo() As Modelo
    Dim MyModelo As Modelo
    
    If mudtprops.ModeloID = 0 Then Exit Sub
    
    Set MyModelo = New Modelo
        
    MyModelo.Load mudtprops.ModeloID, vbNullString, 0, vbNullString
    
    Set objModelo = MyModelo
    
    Set MyModelo = Nothing
    
End Property

Public Property Get objPrenda() As Prenda
    Dim MyPrenda As Prenda
    
    If mudtprops.PrendaID = 0 Then Exit Property
    
    Set MyPrenda = New Prenda
        
    MyPrenda.Load mudtprops.PrendaID, "EUR"
    
    Set objPrenda = MyPrenda
    
    Set MyPrenda = Nothing
    
End Property

Public Property Get objSerie() As Serie
    Dim MySerie As Serie
    
    If mudtprops.SerieID = 0 Then Exit Property
    
    Set MySerie = New Serie
        
    MySerie.Load mudtprops.SerieID
    
    Set objSerie = MySerie
    
    Set MySerie = Nothing
    
End Property

Public Sub AsignarTallajePredeterminado()
    Dim objParametroAplicacion As ParametroAplicacion
    Dim lngTallajeID As Long
    
    Set objParametroAplicacion = New ParametroAplicacion
    lngTallajeID = objParametroAplicacion.TallajePredeterminado
    CargarListaTallajes
    Me.Tallaje = mobjTallajes.Item(lngTallajeID)
    Set objParametroAplicacion = Nothing

End Sub

Private Function GetState() As String
    Dim udtData As ArticuloData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

    Dim udtData As ArticuloData
  
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub

