VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Etiqueta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Valid(IsValid As Boolean)

' Aqui declaro el ComboBox de la lista de ArticuloColor.
Private mobjArticuloColores As TextList

Private mudtprops As EtiquetaProps
Private mudtSave As EtiquetaProps
Private mudtChild As EtiquetaProps
Private mobjParent As Etiquetas
Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean
Private mflgChild As Boolean
Private mflgChildEditing As Boolean
Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Sub BeginEdit()

  If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
  If mflgEditing Then Err.Raise 445
  
  ' save object state
  LSet mudtSave = mudtprops
  mflgEditing = True
  
End Sub

Public Sub CancelEdit()

  If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
  If Not mflgEditing Then Err.Raise 445
  
  mflgEditing = False
  mflgDeleted = False
  mflgDirty = False
  ' restore object state
  LSet mudtprops = mudtSave
  
End Sub

Public Sub ApplyEdit()

  If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
  If Not mflgChild And mflgNew Then Err.Raise 445
  If Not mflgEditing Then Err.Raise 445

  If Not mflgChild Then
    If mflgDeleted And Not mflgNew Then
      ' code to delete the object's data goes here
      mflgNew = True
      mflgDeleted = False

    ElseIf mflgDirty Or mflgNew Then
      If Not IsValid Then Err.Raise 445
      ' save object state
      LSet mudtSave = mudtprops
      mflgNew = False

    End If
    
    mflgDirty = False

  ElseIf mflgNew Then
    If Not mobjParent Is Nothing Then
      mobjParent.AddEtiqueta Me
      Set mobjParent = Nothing
    End If
  
  End If
    
  mflgEditing = False
  
End Sub

Private Sub Class_Initialize()

  mflgNew = True
  Set mobjValid = New BrokenRules
  
  ' No inicializamos la lista porque depende de la temporada en la que estemos.
  Set mobjArticuloColores = Nothing
  
  ' if we know any rules are broken on startup
  ' then add code here to initialize the list
  ' of broken rules
  mobjValid.RuleBroken "ArticuloColorID", True
  mobjValid.RuleBroken "TemporadaID", True
  mobjValid.RuleBroken "Cantidad", True
  
  With mudtprops
    .CantidadT36 = 0
    .CantidadT38 = 0
    .CantidadT40 = 0
    .CantidadT42 = 0
    .CantidadT44 = 0
    .CantidadT46 = 0
    .CantidadT48 = 0
    .CantidadT50 = 0
    .CantidadT52 = 0
    .CantidadT54 = 0
    .CantidadT56 = 0
    .TemporadaID = 0
    .TallajeID = 0
  End With
  
End Sub

Private Sub Class_Terminate()

  Set mobjValid = Nothing
  
  Set mobjArticuloColores = Nothing
  
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

Public Sub Delete()

  If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
  If Not mflgEditing Then Err.Raise 445
  
  mflgDeleted = True
  mflgDirty = True
  
End Sub

Public Property Get IsDeleted() As Boolean

  IsDeleted = mflgDeleted

End Property

Public Property Get IsNew() As Boolean

  IsNew = mflgNew

End Property

Public Property Get IsDirty() As Boolean

  IsDirty = mflgDirty Or mflgNew

End Property

Friend Sub SetAsChild(objEtiquetas As Etiquetas)

  If mflgNew Then Set mobjParent = objEtiquetas
  mflgChild = True

End Sub

Friend Sub ChildBeginEdit()

  If mflgChildEditing Then Err.Raise 445
  
  ' save object state
  LSet mudtChild = mudtprops
  mflgChildEditing = True

End Sub

Friend Sub ChildCancelEdit()

  If Not mflgChildEditing Then Err.Raise 445
  
  mflgChildEditing = False
  mflgDeleted = False
  ' restore object state
  LSet mudtprops = mudtChild

End Sub

Friend Sub ChildApplyEdit()

  If Not mflgChildEditing Then Err.Raise 445
  If Not IsValid Then Err.Raise 445

  mflgChildEditing = False

End Sub

Public Property Get ArticuloColorID() As Long

    ArticuloColorID = mudtprops.ArticuloColorID
  
End Property

Public Property Let ArticuloColorID(Value As Long)

    mudtprops.ArticuloColorID = Value
    If mobjArticuloColores Is Nothing Then CargarListaArticuloColores
    ArticuloColor = mobjArticuloColores.Item(Value)
    

End Property

Public Property Let ArticuloColor(Value As String)
    Dim objArticuloColor As ArticuloColor
    Dim objArticulo As Articulo
    Dim objSerie As Serie
    Dim objMaterial As Material
    Dim objProveedor As Proveedor
    Dim objTallaje As Tallaje

    If Not mflgEditing Then Err.Raise 383
    
    If mobjArticuloColores Is Nothing Then CargarListaArticuloColores
    
    ' Si el articulo introducido es correcto, entonces se carga la serie
    If mobjArticuloColores.Key(Value) Then
        Set objArticuloColor = New ArticuloColor
        
        objArticuloColor.Load mobjArticuloColores.Key(Value), "EUR"
        
        Set objSerie = New Serie
        Set objArticulo = objArticuloColor.objArticulo
        objSerie.Load objArticulo.SerieID
        
        Set objMaterial = objSerie.objMaterial("EUR")
        mudtprops.PorcComposicion1 = objMaterial.PorcComposicion1
        mudtprops.PorcComposicion2 = objMaterial.PorcComposicion2
        mudtprops.PorcComposicion3 = objMaterial.PorcComposicion3
        mudtprops.PorcComposicion4 = objMaterial.PorcComposicion4
        mudtprops.Composicion1 = objMaterial.Composicion1
        mudtprops.Composicion2 = objMaterial.Composicion2
        mudtprops.Composicion3 = objMaterial.Composicion3
        mudtprops.Composicion4 = objMaterial.Composicion4
        mudtprops.NombrePrenda = objArticulo.NombrePrenda
        mudtprops.NombreModelo = objArticulo.NombreModelo
        mudtprops.NombreSerie = objArticulo.NombreSerie
        mudtprops.NombreColor = objArticuloColor.NombreColor
        mudtprops.PrecioVentaPublico = objArticulo.PrecioVentaPublico
        
        ' Si existe proveedor hay que cargar su código corto.
        If objArticulo.ProveedorID <> 0 Then
            Set objProveedor = New Proveedor
            objProveedor.Load objArticulo.ProveedorID
            mudtprops.CodigoProveedor = objProveedor.Codigo
            Set objProveedor = Nothing
        End If
        
        mudtprops.TallajeID = objArticulo.TallajeID
        
        Set objSerie = Nothing
        Set objMaterial = Nothing
        Set objArticulo = Nothing
        Set objArticuloColor = Nothing
    End If
        
    mudtprops.ArticuloColorID = mobjArticuloColores.Key(Value)
    mobjValid.RuleBroken "ArticuloColorID", (mobjArticuloColores.Key(Value) = 0)
    
    mflgDirty = True
    
End Property

Public Property Get ArticuloColor() As String

  If mobjArticuloColores Is Nothing Then CargarListaArticuloColores
  
  ArticuloColor = mobjArticuloColores.Item(mudtprops.ArticuloColorID)

End Property

Public Property Get ArticuloColores() As TextList

  If mobjArticuloColores Is Nothing Then CargarListaArticuloColores
  
  Set ArticuloColores = mobjArticuloColores

End Property

Public Property Let CantidadT36(Value As Double)

  ValidarCantidadTalla Value, "CantidadT36"
  
  ' Las cantidades son enteras
  mudtprops.CantidadT36 = Round(Value, 0)
  
  ' Que el total de cantidades por talla no sea cero
  ValidarCantidad
  
End Property

Public Property Get CantidadT36() As Double

  CantidadT36 = mudtprops.CantidadT36

End Property

Public Property Let CantidadT38(Value As Double)

     ValidarCantidadTalla Value, "CantidadT38"
  
  ' Las cantidades son enteras
  mudtprops.CantidadT38 = Round(Value, 0)
  
  ' Que el total de cantidades por talla no sea cero
  ValidarCantidad
  
End Property

Public Property Get CantidadT38() As Double

  CantidadT38 = mudtprops.CantidadT38

End Property

Public Property Let CantidadT40(Value As Double)

     ValidarCantidadTalla Value, "CantidadT40"
  
  ' Las cantidades son enteras
  mudtprops.CantidadT40 = Round(Value, 0)
  
  ' Que el total de cantidades por talla no sea cero
  ValidarCantidad
  
End Property

Public Property Get CantidadT40() As Double

  CantidadT40 = mudtprops.CantidadT40

End Property

Public Property Let CantidadT42(Value As Double)

     ValidarCantidadTalla Value, "CantidadT42"
  
  ' Las cantidades son enteras
  mudtprops.CantidadT42 = Round(Value, 0)
  
  ' Que el total de cantidades por talla no sea cero
  ValidarCantidad
  
End Property

Public Property Get CantidadT42() As Double

  CantidadT42 = mudtprops.CantidadT42

End Property

Public Property Let CantidadT44(Value As Double)

     ValidarCantidadTalla Value, "CantidadT44"
  
  ' Las cantidades son enteras
  mudtprops.CantidadT44 = Round(Value, 0)
  
  ' Que el total de cantidades por talla no sea cero
  ValidarCantidad
  
End Property

Public Property Get CantidadT44() As Double

  CantidadT44 = mudtprops.CantidadT44

End Property

Public Property Let CantidadT46(Value As Double)

     ValidarCantidadTalla Value, "CantidadT46"
  
  ' Las cantidades son enteras
  mudtprops.CantidadT46 = Round(Value, 0)
  
  ' Que el total de cantidades por talla no sea cero
  ValidarCantidad
  
End Property

Public Property Get CantidadT46() As Double

  CantidadT46 = mudtprops.CantidadT46

End Property

Public Property Let CantidadT48(Value As Double)

     ValidarCantidadTalla Value, "CantidadT48"
  
  ' Las cantidades son enteras
  mudtprops.CantidadT48 = Round(Value, 0)
  
  ' Que el total de cantidades por talla no sea cero
  ValidarCantidad
  
End Property

Public Property Get CantidadT48() As Double

  CantidadT48 = mudtprops.CantidadT48

End Property

Public Property Let CantidadT50(Value As Double)

     ValidarCantidadTalla Value, "CantidadT50"
  
  ' Las cantidades son enteras
  mudtprops.CantidadT50 = Round(Value, 0)
  
  ' Que el total de cantidades por talla no sea cero
  ValidarCantidad
  
End Property

Public Property Get CantidadT50() As Double

  CantidadT50 = mudtprops.CantidadT50

End Property

Public Property Let CantidadT52(Value As Double)

     ValidarCantidadTalla Value, "CantidadT52"
  
  ' Las cantidades son enteras
  mudtprops.CantidadT52 = Round(Value, 0)
  
  ' Que el total de cantidades por talla no sea cero
  ValidarCantidad
  
End Property

Public Property Get CantidadT52() As Double

  CantidadT52 = mudtprops.CantidadT52

End Property

Public Property Let CantidadT54(Value As Double)

     ValidarCantidadTalla Value, "CantidadT54"
  
  ' Las cantidades son enteras
  mudtprops.CantidadT54 = Round(Value, 0)
  
  ' Que el total de cantidades por talla no sea cero
  ValidarCantidad
  
End Property

Public Property Get CantidadT54() As Double

  CantidadT54 = mudtprops.CantidadT54

End Property

Public Property Let CantidadT56(Value As Double)

     ValidarCantidadTalla Value, "CantidadT56"
  
  ' Las cantidades son enteras
  mudtprops.CantidadT56 = Round(Value, 0)
  
  ' Que el total de cantidades por talla no sea cero
  ValidarCantidad
  
End Property

Public Property Get CantidadT56() As Double

  CantidadT56 = mudtprops.CantidadT56

End Property

Public Property Get Cantidad() As Double
  With mudtprops
  Cantidad = .CantidadT36 + .CantidadT38 + .CantidadT40 + .CantidadT42 + _
             .CantidadT44 + .CantidadT46 + .CantidadT48 + .CantidadT50 + _
             .CantidadT52 + .CantidadT54 + .CantidadT56
  End With

End Property

Public Property Let TemporadaID(Value As Long)

  If Not mflgEditing Then Err.Raise 383
  
  mobjValid.RuleBroken "TemporadaID", (Value = 0)
  
  mudtprops.TemporadaID = Value
  
  mflgDirty = True

End Property

Public Property Let Composicion1(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Trim(Value)) > Len(mudtprops.Composicion1) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    mobjValid.RuleBroken "Composicion1", ((Len(Trim(Value)) = 0) And (mudtprops.PorcComposicion1 <> 0))

    mudtprops.Composicion1 = UCase$(Trim(Value))
    mflgDirty = True

End Property

Public Property Get Composicion1() As String

    Composicion1 = mudtprops.Composicion1

End Property

Public Property Let PorcComposicion1(Value As Double)

    If Not mflgEditing Then Err.Raise 383
  
    If Value > 100 Or Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
    ' pongo la asignacion del valor antes porque entra en el testeo de reglas
    mudtprops.PorcComposicion1 = Round(Value, 2)
    ' se rompe una regla si es mayor de cien, si es menor de cero o si la composicion no esta informada.
    ' tampoco se permite que haya un total mayor de cien%.
    mobjValid.RuleBroken "PorcComposicion1", _
        ((Value < 0) Or (Value > 100) Or (Len(Trim(mudtprops.Composicion1)) = 0) Or _
        ((mudtprops.PorcComposicion1 + mudtprops.PorcComposicion2 + mudtprops.PorcComposicion3 + mudtprops.PorcComposicion4) > 100) _
        )

    mflgDirty = True

End Property

Public Property Get PorcComposicion1() As Double

    PorcComposicion1 = mudtprops.PorcComposicion1

End Property

Public Property Let Composicion2(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Trim(Value)) > Len(mudtprops.Composicion2) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
  
    mobjValid.RuleBroken "Composicion2", ((Len(Trim(Value)) = 0) And (mudtprops.PorcComposicion2 <> 0))

    mudtprops.Composicion2 = UCase$(Trim(Value))
    mflgDirty = True

End Property

Public Property Get Composicion2() As String

    Composicion2 = mudtprops.Composicion2

End Property

Public Property Let PorcComposicion2(Value As Double)

    If Not mflgEditing Then Err.Raise 383
  
    If Value > 100 Or Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
    ' pongo la asignacion del valor antes porque entra en el testeo de reglas
    mudtprops.PorcComposicion2 = Round(Value, 2)
    ' se rompe una regla si es mayor de cien, si es menor de cero o si la composicion no esta informada.
    ' tampoco se permite que haya un total mayor de cien%.
    mobjValid.RuleBroken "PorcComposicion2", _
        ((Value < 0) Or (Value > 100) Or (Len(Trim(mudtprops.Composicion2)) = 0) Or _
        ((mudtprops.PorcComposicion1 + mudtprops.PorcComposicion2 + mudtprops.PorcComposicion3 + mudtprops.PorcComposicion4) > 100) _
        )

    mflgDirty = True

End Property

Public Property Get PorcComposicion2() As Double

    PorcComposicion2 = mudtprops.PorcComposicion2

End Property

Public Property Let Composicion3(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Trim(Value)) > Len(mudtprops.Composicion3) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
  
    mobjValid.RuleBroken "Composicion3", ((Len(Trim(Value)) = 0) And (mudtprops.PorcComposicion3 <> 0))

    mudtprops.Composicion3 = UCase$(Trim(Value))
    mflgDirty = True

End Property

Public Property Get Composicion3() As String

    Composicion3 = mudtprops.Composicion3

End Property

Public Property Let PorcComposicion3(Value As Double)

    If Not mflgEditing Then Err.Raise 383
  
    If Value > 100 Or Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
    ' pongo la asignacion del valor antes porque entra en el testeo de reglas
    mudtprops.PorcComposicion3 = Round(Value, 2)
      
    ' se rompe una regla si es mayor de cien, si es menor de cero o si la composicion no esta informada.
    ' tampoco se permite que haya un total mayor de cien%.
    mobjValid.RuleBroken "PorcComposicion3", _
        ((Value < 0) Or (Value > 100) Or (Len(Trim(mudtprops.Composicion3)) = 0) Or _
        ((mudtprops.PorcComposicion1 + mudtprops.PorcComposicion2 + mudtprops.PorcComposicion3 + mudtprops.PorcComposicion4) > 100) _
        )
    
    mflgDirty = True

End Property

Public Property Get PorcComposicion3() As Double

    PorcComposicion3 = mudtprops.PorcComposicion3

End Property

Public Property Let Composicion4(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Trim(Value)) > Len(mudtprops.Composicion4) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
  
    mobjValid.RuleBroken "Composicion4", ((Len(Trim(Value)) = 0) And (mudtprops.PorcComposicion4 <> 0))

    mudtprops.Composicion4 = UCase$(Trim(Value))
    mflgDirty = True

End Property

Public Property Get Composicion4() As String

    Composicion4 = mudtprops.Composicion4

End Property

Public Property Let PorcComposicion4(Value As Double)

    If Not mflgEditing Then Err.Raise 383
  
    If Value > 100 Or Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
    ' pongo la asignacion del valor antes porque entra en el testeo de reglas
    mudtprops.PorcComposicion4 = Round(Value, 2)
    ' se rompe una regla si es mayor de cien, si es menor de cero o si la composicion no esta informada.
    ' tampoco se permite que haya un total mayor de cien%.
    mobjValid.RuleBroken "PorcComposicion4", _
        ((Value < 0) Or (Value > 100) Or (Len(Trim(mudtprops.Composicion4)) = 0) Or _
        ((mudtprops.PorcComposicion1 + mudtprops.PorcComposicion2 + mudtprops.PorcComposicion3 + mudtprops.PorcComposicion4) > 100) _
        )

    mflgDirty = True

End Property

Public Property Get PorcComposicion4() As Double

    PorcComposicion4 = mudtprops.PorcComposicion4

End Property

Public Property Get NombrePrenda() As String

    NombrePrenda = mudtprops.NombrePrenda

End Property

Public Property Get NombreModelo() As String

    NombreModelo = mudtprops.NombreModelo

End Property

Public Property Get NombreSerie() As String

    NombreSerie = mudtprops.NombreSerie

End Property

Public Property Get NombreColor() As String

    NombreColor = mudtprops.NombreColor

End Property

Public Property Let PrecioVentaPublico(Value As Double)

    If Not mflgEditing Then Err.Raise 383
  
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
    ' pongo la asignacion del valor antes porque entra en el testeo de reglas
    mudtprops.PrecioVentaPublico = Round(Value, 2)

    mflgDirty = True

End Property

Public Property Get PrecioVentaPublico() As Double

    PrecioVentaPublico = mudtprops.PrecioVentaPublico

End Property

Public Property Get CodigoProveedor() As String

    CodigoProveedor = mudtprops.CodigoProveedor

End Property

Private Sub ValidarCantidadTalla(CantidadTalla As Double, _
                            Propiedad As String)

  If Not mflgEditing Then Err.Raise 383
  
  If CantidadTalla < 0 Then _
    Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
  ' No se permiten Cantidades negativas
  mobjValid.RuleBroken Propiedad, (CantidadTalla < 0)
End Sub
  
Private Sub ValidarCantidad()
  ' Se valida la cantidad total de la linea de Albaranes
  mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
  mflgDirty = True
End Sub
  
Private Sub CargarListaArticuloColores()

  Set mobjArticuloColores = New TextList
  mobjArticuloColores.Load "vArticuloColores", 0, mudtprops.TemporadaID
  
End Sub

Public Property Get CantidadTalla(i As Integer) As Double
    
    Select Case i
    Case 1
        CantidadTalla = mudtprops.CantidadT36
    Case 2
        CantidadTalla = mudtprops.CantidadT38
    Case 3
        CantidadTalla = mudtprops.CantidadT40
    Case 4
        CantidadTalla = mudtprops.CantidadT42
    Case 5
        CantidadTalla = mudtprops.CantidadT44
    Case 6
        CantidadTalla = mudtprops.CantidadT46
    Case 7
        CantidadTalla = mudtprops.CantidadT48
    Case 8
        CantidadTalla = mudtprops.CantidadT50
    Case 9
        CantidadTalla = mudtprops.CantidadT52
    Case 10
        CantidadTalla = mudtprops.CantidadT54
    Case 11
        CantidadTalla = mudtprops.CantidadT56
    Case Else
        Err.Raise vbObjectError + 1001, "Etiqueta[CantidadTalla]", "Número de talla incorrecto"
    End Select
End Property

Public Property Get TallajeID() As Long

    TallajeID = mudtprops.TallajeID

End Property
