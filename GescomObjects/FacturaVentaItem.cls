VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FacturaVentaItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Valid(IsValid As Boolean)

Private mobjArticuloColor As ArticuloColor

Private mstrCodigoArticuloColor As String

Private mudtprops As FacturaVentaItemProps
Private mudtSave As FacturaVentaItemProps
Private mudtChild As FacturaVentaItemProps
Private mobjParent As FacturaVentaItems
Private mobjAlbaranVentaItem As AlbaranVentaItem
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
    
    ' Si hay Albaran lo cargamos para posteriores validaciones
    If HayAlbaran Then
        Set mobjAlbaranVentaItem = New AlbaranVentaItem
        mobjAlbaranVentaItem.Load mudtprops.AlbaranVentaItemID
    End If
    
End Sub

Public Sub CancelEdit()
    
    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445
    
    mflgEditing = False
    mflgDeleted = False
    mflgDirty = False
    ' restore object state
    LSet mudtprops = mudtSave
    
    ' Eliminamos el Albaran de memoria que habiamos cargado para validaciones.
    If HayAlbaran Then Set mobjAlbaranVentaItem = Nothing

End Sub

Public Sub ApplyEdit()

    Dim objPersist As FacturaVentaItemPersist
    
    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgChild And mflgNew Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445
    
    Set objPersist = CreateObject("GCServerMTS.FacturaVentaItemPersist", PERSIST_SERVER)
    
    If Not mflgChild Then
      If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.FacturaVentaItemID
        mflgNew = True
        mflgDeleted = False
    
      ElseIf mflgDirty Or mflgNew Then
        If Not IsValid Then Err.Raise 445
        ' save object to database if appropriate
        Actualizar
        SetState objPersist.Save(GetState)
        ' save object state
        LSet mudtSave = mudtprops
        mflgNew = False
    
      End If
      
      mflgDirty = False
    
    ElseIf mflgNew Then
      If Not mobjParent Is Nothing Then
        mobjParent.AddFacturaVentaItem Me
        Set mobjParent = Nothing
      End If
    
    
    End If
      
    Set objPersist = Nothing
    mflgEditing = False
    
    ' Eliminamos el Albaran de memoria que habiamos cargado para validaciones.
    If HayAlbaran Then Set mobjAlbaranVentaItem = Nothing
    
End Sub

Private Sub Class_Initialize()

    mflgNew = True
    Set mobjValid = New BrokenRules
    
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
    mobjValid.RuleBroken "Descripcion", True
    mobjValid.RuleBroken "TemporadaID", True
    'mobjValid.RuleBroken "Cantidad", True
    
    With mudtprops
        .SituacionImpresa = "A"    ' Marco la situacion del Factura como "A"--> alta,
                            ' pendiente de codificar las posibles situaciones.
        .Descripcion = Trim(vbNullString)
        .Cantidad = 0
        .PrecioVentaPTA = 0
        .PrecioVentaEUR = 0
        .Descuento = 0
        .BrutoPTA = 0
        .BrutoEUR = 0
        .ComisionPTA = 0
        .ComisionEUR = 0
        .Comision = 0
        .TemporadaID = 0
        .ActualizarAlta = False
        .DesactualizarAlta = False
    End With
    
End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing

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

Public Sub Load(FacturaVentaItemID As Long)
    Dim objPersist As FacturaVentaItemPersist
    
    If mflgChild Then Err.Raise 445
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
    
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.FacturaVentaItemPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(FacturaVentaItemID)
    Set objPersist = Nothing
    
    mobjValid.RuleBroken "ArticuloColorID", False
    mobjValid.RuleBroken "Descripcion", False
    mobjValid.RuleBroken "TemporadaID", False
    'mobjValid.RuleBroken "Cantidad", False
    
    mflgNew = False
    
End Sub

Public Sub Delete()

    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445
    
'    If Contabilizado Then _
'       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha contabilizado la factura"
      
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

Friend Sub SetAsChild(objFacturaVentaItems As FacturaVentaItems)

    If mflgNew Then Set mobjParent = objFacturaVentaItems
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

Public Property Get FacturaVentaItemID() As Long

    FacturaVentaItemID = mudtprops.FacturaVentaItemID

End Property

Public Property Let AlbaranVentaItemID(Value As Long)

    mudtprops.AlbaranVentaItemID = Value

End Property

Public Property Get AlbaranVentaItemID() As Long
  
    AlbaranVentaItemID = mudtprops.AlbaranVentaItemID

End Property

Public Property Let TemporadaID(Value As Long)

    If Not mflgEditing Then Err.Raise 383
    
    mobjValid.RuleBroken "TemporadaID", (Value = 0)
    
    mudtprops.TemporadaID = Value
    
    mflgDirty = True

End Property

Public Property Get SituacionImpresa() As String
    
    Select Case mudtprops.SituacionImpresa
    Case "A"
       SituacionImpresa = "Generada"
    Case "C"
       SituacionImpresa = "Contabilizada"
    Case Else
      Err.Raise vbObjectError + 1001, "Situacion Impresa de Factura erronea"
    End Select

End Property

Public Property Get ArticuloColorID() As Long

    ArticuloColorID = mudtprops.ArticuloColorID
  
End Property
Public Property Let ArticuloColorID(Value As Long)

    mudtprops.ArticuloColorID = Value

End Property

Public Property Let CodigoArticuloColor(Value As String)

    If Not mflgEditing Then Err.Raise 383
    
    If HayAlbaran Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar, está relacionado con un Albaran"
    
'    If Contabilizado Then _
'       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha contabilizado la factura"
    
    If Len(Trim(Value)) > 8 Then Exit Property
    
    mstrCodigoArticuloColor = UCase(Trim(Value))
     
    If Len(mstrCodigoArticuloColor) < 8 Then Exit Property
    
    Set mobjArticuloColor = New ArticuloColor
    
    mobjArticuloColor.LoadCodigo Trim(mstrCodigoArticuloColor), mudtprops.TemporadaID
    
    mobjValid.RuleBroken "ArticuloColorID", False
    
    PrecioVenta = mobjArticuloColor.objArticulo.PrecioVenta
    Descripcion = mobjArticuloColor.objArticulo.NombreCompleto & " " & mobjArticuloColor.NombreColor
    mudtprops.Descripcion = Descripcion
    
    mudtprops.ArticuloColorID = mobjArticuloColor.ArticuloColorID
    
    mflgDirty = True
    
End Property

Public Property Get CodigoArticuloColor() As String
    
    CodigoArticuloColor = mstrCodigoArticuloColor
    
End Property

Public Property Get objArticuloColor() As ArticuloColor
    
    If mudtprops.ArticuloColorID = 0 Then
        Set objArticuloColor = Nothing
        Exit Property
    End If
    
    If mobjArticuloColor Is Nothing Then _
        Set mobjArticuloColor = New ArticuloColor
        
    mobjArticuloColor.Load mudtprops.ArticuloColorID, "EUR"
    Set objArticuloColor = mobjArticuloColor

End Property

Public Property Let Descripcion(Value As String)

    If Not mflgEditing Then Err.Raise 383
    
    If Len(Value) > Len(mudtprops.Descripcion) Then _
      Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
      
    ' Hay que asignar una descripcion si no hemos seleccionado un articulo
    mobjValid.RuleBroken "Descripcion", (Len(Trim(Value)) = 0 And mudtprops.ArticuloColorID = 0)
    
    mudtprops.Descripcion = Trim(Value)
    mflgDirty = True

End Property

Public Property Get Descripcion() As String

    Descripcion = mudtprops.Descripcion

End Property

Public Property Let Cantidad(Value As Double)
    If Not mflgEditing Then Err.Raise 383
    
'    If Cantidad < 0 Then _
'      Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
    If HayAlbaran Then
       ValidarCantidad Value, mudtSave.Cantidad, _
                    mobjAlbaranVentaItem.Cantidad
    End If
    
'    ' No se permiten Cantidades negativas
'    mobjValid.RuleBroken "Cantidad", (Cantidad < 0)
    
    ' Las cantidades son enteras
    mudtprops.Cantidad = Round(Value, 0)
    
    ' Que el total de cantidades por talla no sea cero
    ' Se valida la cantidad total de la linea de Facturas
    'mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    CalcularBruto
    mflgDirty = True
    
End Property

Private Sub ValidarCantidad(Cantidad As Double, Inicial As Double, Pendiente As Double)

  ' Si esta relacionado con un Albaran se mira que no se pase sobre el Albaran original.
  ' Al ser nuevo se compara con el valor que se pretende asignar.
  ' Si no es nuevo se compara ademas con el inicial.
   If mflgNew Then
      If Cantidad <> Pendiente Then
         Err.Raise vbObjectError + 1001, "Cantidad no puede ser distinta a la del Albaran original"
      End If
   Else
      If Cantidad - Inicial <> Pendiente Then
         Err.Raise vbObjectError + 1001, "Cantidad no puede ser distinta a del Albaran original"
      End If
   End If
        
End Sub
  
Public Property Get Cantidad() As Double

    Cantidad = mudtprops.Cantidad

End Property

Public Property Get Impreso() As Boolean

    Impreso = (mudtprops.SituacionImpresa = "I")
  
End Property

Public Property Get HayAlbaran() As Boolean
  
    HayAlbaran = (mudtprops.AlbaranVentaItemID <> 0)
  
End Property

Public Property Let Descuento(Value As Double)
    
    If Not mflgEditing Then Err.Raise 383
    
    mobjValid.RuleBroken "Descuento", (Value < 0 Or Value > 100)
    
    mudtprops.Descuento = Round(Value, 2)
    
    CalcularBruto
        
    mflgDirty = True

End Property

Public Property Get Descuento() As Double

    Descuento = mudtprops.Descuento

End Property

Public Property Let Comision(Value As Double)
    
    If Not mflgEditing Then Err.Raise 383
    
    mobjValid.RuleBroken "Comision", (Value < 0)
    
    mudtprops.Comision = Round(Value, 2)
    
    CalcularComision
    
    mflgDirty = True

End Property

Public Property Get Comision() As Double

    Comision = mudtprops.Comision

End Property

Public Property Let ImporteComision(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
      
'    If Value < 0 Then _
'       Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
'
'    mobjValid.RuleBroken "ImporteComision", (Value < 0)
    
    mudtprops.ComisionEUR = Round(Value, 2)
    mudtprops.ComisionPTA = EUR2PTA(Value)
    
    mflgDirty = True

End Property

Public Property Get ImporteComision() As Double

    ImporteComision = mudtprops.ComisionEUR
  
End Property

Public Property Let PrecioVenta(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
      
    If Value < 0 Then _
       Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
    
    mobjValid.RuleBroken "PrecioVenta", (Value < 0)
    
    mudtprops.PrecioVentaEUR = Round(Value, 2)
    mudtprops.PrecioVentaPTA = EUR2PTA(Value)
    
    CalcularBruto
      
    mflgDirty = True

End Property

Public Property Get PrecioVenta() As Double

    PrecioVenta = mudtprops.PrecioVentaEUR
  
End Property

Public Property Let Bruto(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
      
'    If Value < 0 Then _
'       Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
'
'    mobjValid.RuleBroken "Bruto", (Value < 0)
    
    mudtprops.BrutoEUR = Round(Value, 2)
    mudtprops.BrutoPTA = EUR2PTA(Value)
    
    mflgDirty = True
    
End Property

Public Property Get Bruto() As Double

    Bruto = mudtprops.BrutoEUR
    
End Property

Friend Function GetState() As String
    Dim udtData As FacturaVentaItemData
  
    With mudtprops
      .IsNew = mflgNew
      .IsDeleted = mflgDeleted
      .IsDirty = mflgDirty
    End With
    
    LSet udtData = mudtprops
    GetState = udtData.Buffer
    'mflgNew = False

End Function

Friend Sub SetState(Buffer As String)
    Dim udtData As FacturaVentaItemData
    
    udtData.Buffer = Buffer
    LSet mudtprops = udtData
    mflgNew = False
    mobjValid.RuleBroken "ArticuloColorID", False
    mobjValid.RuleBroken "Descripcion", False
    mobjValid.RuleBroken "TemporadaID", False
    'mobjValid.RuleBroken "Cantidad", False
    
End Sub

Private Sub CalcularBruto()
    Dim curImporteSinDescuento As Currency
    
    curImporteSinDescuento = PrecioVenta * Cantidad
    
    Bruto = curImporteSinDescuento - (curImporteSinDescuento * Descuento / 100)
End Sub

Private Sub CalcularComision()
    ImporteComision = Round(Bruto * Comision / 100, 2)
End Sub

' Indica si hay que Desactualizar o no
' Hay que hacerlo en caso de que no sea nuevo, haya cambiado el articulo,
' o hayan cambiado las cantidades
Private Function HayQueDesactualizar() As Boolean

    If Not mflgDirty Then
       HayQueDesactualizar = False
       Exit Function
    End If
    
    If mflgNew Or (mudtprops.ArticuloColorID = 0) Then
       HayQueDesactualizar = False
       Exit Function
    End If
       
    If mflgDeleted Then
       HayQueDesactualizar = True
       Exit Function
    End If
    
    If mudtSave.ArticuloColorID = mudtprops.ArticuloColorID And _
       mudtSave.Cantidad = mudtprops.Cantidad Then
       HayQueDesactualizar = False
       Exit Function
    End If
    
    HayQueDesactualizar = True
    
End Function

' Indica si hay que Actualizar o no
' Hay que hacerlo en caso de que sea nuevo, haya cambiado el articulo,
' o hayan cambiado las cantidades
Private Function HayQueActualizar() As Boolean

    If Not mflgDirty Then
       HayQueActualizar = False
       Exit Function
    End If
    
    If mflgDeleted Then
       HayQueActualizar = False
       Exit Function
    End If
    
    If mudtprops.ArticuloColorID = 0 Then
       HayQueActualizar = False
       Exit Function
    End If
    
    If mflgNew Then
       HayQueActualizar = True
       Exit Function
    End If
    
    If mudtSave.ArticuloColorID = mudtprops.ArticuloColorID And _
       mudtSave.Cantidad = mudtprops.Cantidad Then
       HayQueActualizar = False
       Exit Function
    End If
    
    HayQueActualizar = True
    
End Function

Public Sub Actualizar()

    mudtprops.ActualizarAlta = HayQueActualizar
    mudtprops.DesactualizarAlta = HayQueDesactualizar
   
End Sub

' Este procedimiento crea los datos de un Factura a partir de los de un Albaran.
Public Sub FacturaDesdeAlbaran(AlbaranItemID As Long)

    Set mobjAlbaranVentaItem = New AlbaranVentaItem
    
    mobjAlbaranVentaItem.Load AlbaranItemID
      
'    ArticuloColor = mobjAlbaranVentaItem.ArticuloColor
    If mobjAlbaranVentaItem.ArticuloColorID <> 0 Then
        mudtprops.TemporadaID = mobjAlbaranVentaItem.objArticuloColor.TemporadaID
        CodigoArticuloColor = mobjAlbaranVentaItem.CodigoArticuloColor
    Else
        mudtprops.ArticuloColorID = 0
    End If
    mudtprops.TemporadaID = mobjAlbaranVentaItem.TemporadaID
    mudtprops.AlbaranVentaItemID = mobjAlbaranVentaItem.AlbaranVentaItemID
    'mudtProps.ArticuloColorID = mobjAlbaranventaitem.ArticuloColorID
    ' aqui hay que cargar el nombre del articulo (o quizas nada).
    Descripcion = mobjAlbaranVentaItem.Descripcion
    Cantidad = mobjAlbaranVentaItem.Cantidad
    
    PrecioVenta = mobjAlbaranVentaItem.PrecioVenta
    Descuento = mobjAlbaranVentaItem.Descuento
    Bruto = mobjAlbaranVentaItem.Bruto
    
    Set mobjAlbaranVentaItem = Nothing

End Sub

