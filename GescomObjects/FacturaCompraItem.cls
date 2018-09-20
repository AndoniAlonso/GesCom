VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FacturaCompraItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Valid(IsValid As Boolean)

' Aqui declaro el ComboBox de la lista de materiales.
Private mobjMateriales As TextList

Private mMoneda As String

Private mudtprops As FacturaCompraItemProps
Private mudtSave As FacturaCompraItemProps
Private mudtChild As FacturaCompraItemProps
Private mobjParent As FacturaCompraItems
Private mobjAlbaranCompraItem As AlbaranCompraItem

Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean
Private mflgChild As Boolean
Private mflgChildEditing As Boolean

Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Sub BeginEdit(Moneda As String)

    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If mflgEditing Then Err.Raise 445
  
    mMoneda = Moneda
    ' save object state
    LSet mudtSave = mudtprops
    mflgEditing = True
  
    ' Si hay Albaran lo cargamos para posteriores validaciones
    If HayAlbaran Then
        Set mobjAlbaranCompraItem = New AlbaranCompraItem
        mobjAlbaranCompraItem.Load mudtprops.AlbaranCompraItemID, mMoneda, ALBARANCOMPRAITEM_MATERIAL ' OJOOJO: AQUI HABRA QUE TENER EN CUENTA EL TIPO DE ARTICULOS
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
    If HayAlbaran Then Set mobjAlbaranCompraItem = Nothing

End Sub

Public Sub ApplyEdit()
    Dim objPersist As FacturaCompraItemPersist
  
    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgChild And mflgNew Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("GCServerMTS.FacturaCompraItemPersist", PERSIST_SERVER)
  
    If Not mflgChild Then
        If mflgDeleted And Not mflgNew Then
            ' code to delete the object's data goes here
            objPersist.DeleteObject mudtprops.FacturaCompraItemID
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
            mobjParent.AddFacturaCompraItem Me
            Set mobjParent = Nothing
        End If

    End If
    
    Set objPersist = Nothing
    mflgEditing = False
  
    ' Eliminamos el Albaran de memoria que habiamos cargado para validaciones.
    If HayAlbaran Then Set mobjAlbaranCompraItem = Nothing

End Sub

Private Sub Class_Initialize()

    mflgNew = True
    Set mobjValid = New BrokenRules
  
    'Set mobjMateriales = New TextList
    'mobjMateriales.Load "vNombreMateriales"
    'mudtprops.MaterialID = mobjMateriales.Key(mobjMateriales.IndexItem(1))
    mudtprops.MaterialID = 0
  
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
    'mobjValid.RuleBroken "Cantidad", True
    mobjValid.RuleBroken "MaterialID", True
  
    With mudtprops
        .SituacionImpresa = "A"    ' Marco la situacion del Factura como "A"--> alta,
                                    ' pendiente de codificar las posibles situaciones.
        .Cantidad = 0
        .PrecioCostePTA = 0
        .PrecioCosteEUR = 0
        .Descuento = 0
        .BrutoPTA = 0
        .BrutoEUR = 0
    End With
  
End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing
  
    Set mobjMateriales = Nothing

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

Public Sub Load(FacturaCompraItemID As Long, Moneda As String)
    Dim objPersist As FacturaCompraItemPersist
  
    If mflgChild Then Err.Raise 445
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    mMoneda = Moneda
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.FacturaCompraItemPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(FacturaCompraItemID)
    Set objPersist = Nothing
  
    mobjValid.RuleBroken "MaterialID", False
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

Friend Sub SetAsChild(objFacturaCompraItems As FacturaCompraItems, Moneda As String)

    mMoneda = Moneda
    If mflgNew Then Set mobjParent = objFacturaCompraItems
    mflgChild = True

End Sub

Friend Sub ChildBeginEdit(Moneda As String)

    If mflgChildEditing Then Err.Raise 445
  
    mMoneda = Moneda
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

Public Property Get FacturaCompraItemID() As Long

    FacturaCompraItemID = mudtprops.FacturaCompraItemID

End Property

Public Property Let AlbaranCompraItemID(Value As Long)

    mudtprops.AlbaranCompraItemID = Value

End Property

Public Property Get AlbaranCompraItemID() As Long

    AlbaranCompraItemID = mudtprops.AlbaranCompraItemID

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

Public Property Get MaterialID() As Long

    MaterialID = mudtprops.MaterialID
  
End Property

Public Property Let MaterialID(Value As Long)

    mudtprops.MaterialID = Value

End Property

Public Property Let Material(Value As String)
    Dim objMaterial As Material

    If Not mflgEditing Then Err.Raise 383
  
    If HayAlbaran Then _
        Err.Raise vbObjectError + 1001, "No se puede modificar, está relacionado con un Albaran"
  
'    If Contabilizado Then _
'        Err.Raise vbObjectError + 1001, "No se puede modificar, ya se han Contabilizado entregas"
  
    If mobjMateriales Is Nothing Then CargarListaMateriales
    
    If mobjMateriales.Key(Value) = 0 Then Exit Property
  
    Set objMaterial = New Material
  
    objMaterial.Load mobjMateriales.Key(Value), mMoneda
  
    PrecioCoste = objMaterial.PrecioCoste
  
    mudtprops.MaterialID = objMaterial.MaterialID
    mudtprops.NombreMaterial = objMaterial.Nombre
    
    mobjValid.RuleBroken "MaterialID", mudtprops.MaterialID = 0
  
    mflgDirty = True

End Property

Public Property Get Material() As String

    'Material = mobjMateriales.Item(mudtprops.MaterialID)
    Material = mudtprops.NombreMaterial

End Property

Public Property Get Materiales() As TextList

    If mobjMateriales Is Nothing Then CargarListaMateriales

    Set Materiales = mobjMateriales

End Property

Public Property Let Cantidad(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
  
'    If Cantidad < 0 Then _
'        Err.Raise vbObjectError + 1001, "Valor numerico erróneo"

    If HayAlbaran Then
        ValidarCantidadTalla Value, mudtSave.Cantidad, _
            mobjAlbaranCompraItem.Cantidad
    End If
  
    ' No se permiten Cantidades negativas
'    mobjValid.RuleBroken "Cantidad", (Cantidad < 0)
  
    mudtprops.Cantidad = Round(Value, 2)
  
    ' Que el total de cantidades por talla no sea cero
    ' Se valida la cantidad total de la linea de Facturas
   ' mobjValid.RuleBroken "Cantidad", (Cantidad <= 0)
    CalcularBruto
    mflgDirty = True
  
End Property

Private Sub ValidarCantidadTalla(Cantidad As Double, Inicial As Double, Pendiente As Double)

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
  
    HayAlbaran = (mudtprops.AlbaranCompraItemID <> 0)
  
End Property

Public Property Let Descuento(Value As Double)

    If Not mflgEditing Then Err.Raise 383
  
    mobjValid.RuleBroken "Descuento", (Value < 0)
  
    mudtprops.Descuento = Round(Value, 2)
  
    mflgDirty = True

End Property

Public Property Get Descuento() As Double

    Descuento = mudtprops.Descuento

End Property

Public Property Let PrecioCoste(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
'    If Value < 0 Then _
'        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
'
'    mobjValid.RuleBroken "PrecioCoste", (Value < 0)
'
    If EsEUR(mMoneda) Then
        mudtprops.PrecioCosteEUR = Round(Value, 3)
        mudtprops.PrecioCostePTA = EUR2PTA(Value)
    ElseIf EsPTA(mMoneda) Then
        mudtprops.PrecioCostePTA = Round(Value, 0)
        mudtprops.PrecioCosteEUR = PTA2EUR(Value)
    Else
        mudtprops.PrecioCostePTA = 0
        mudtprops.PrecioCosteEUR = 0
    End If
  
    CalcularBruto
    
    mflgDirty = True

End Property

Public Property Get PrecioCoste() As Double

    If EsEUR(mMoneda) Then
        PrecioCoste = mudtprops.PrecioCosteEUR
    ElseIf EsPTA(mMoneda) Then
        PrecioCoste = mudtprops.PrecioCostePTA
    Else
        PrecioCoste = 0
    End If
  
End Property

Public Property Let Bruto(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
'    If Value < 0 Then _
'        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
'
'    mobjValid.RuleBroken "Bruto", (Value < 0)
'
    If EsEUR(mMoneda) Then
        mudtprops.BrutoEUR = Round(Value, 2)
        mudtprops.BrutoPTA = EUR2PTA(Value)
    ElseIf EsPTA(mMoneda) Then
        mudtprops.BrutoPTA = Round(Value, 0)
        mudtprops.BrutoEUR = PTA2EUR(Value)
    Else
        mudtprops.BrutoPTA = 0
        mudtprops.BrutoEUR = 0
    End If
        
    mflgDirty = True

End Property

Public Property Get Bruto() As Double

    If EsEUR(mMoneda) Then
        Bruto = mudtprops.BrutoEUR
    ElseIf EsPTA(mMoneda) Then
        Bruto = mudtprops.BrutoPTA
    Else
        Bruto = 0
    End If
  
End Property

Friend Function GetState() As String

    Dim udtData As FacturaCompraItemData
  
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

    Dim udtData As FacturaCompraItemData

    udtData.Buffer = Buffer
    LSet mudtprops = udtData
    mflgNew = False
    mobjValid.RuleBroken "MaterialID", False
    'mobjValid.RuleBroken "Cantidad", False

End Sub

Private Sub CalcularBruto()
    
    Bruto = PrecioCoste * Cantidad
    
End Sub

' Indica si hay que Desactualizar o no
' Hay que hacerlo en caso de que no sea nuevo, haya cambiado el articulo,
' o hayan cambiado las cantidades
Private Function HayQueDesactualizar() As Boolean

    If Not mflgDirty Then
       HayQueDesactualizar = False
       Exit Function
    End If
    
    If mflgNew Or (mudtprops.MaterialID = 0) Then
        HayQueDesactualizar = False
        Exit Function
    End If
     
    If mflgDeleted Then
        HayQueDesactualizar = True
        Exit Function
    End If
      
    If mudtSave.MaterialID = mudtprops.MaterialID And _
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
  
    If mudtprops.MaterialID = 0 Then
       HayQueActualizar = False
       Exit Function
    End If
    
    If mflgNew Then
       HayQueActualizar = True
       Exit Function
    End If
    
    If mudtSave.MaterialID = mudtprops.MaterialID And _
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
    Dim objAlbaranCompraItemMaterial As AlbaranCompraItemMaterial

    Set mobjAlbaranCompraItem = New AlbaranCompraItem
   
    mobjAlbaranCompraItem.Load AlbaranItemID, mMoneda, ALBARANCOMPRAITEM_MATERIAL ' OJOOJO: HABRA QUE TENER EN CUENTA EL TIPO ARTICULO
    
    Set objAlbaranCompraItemMaterial = mobjAlbaranCompraItem
     
    Material = objAlbaranCompraItemMaterial.Material
    mudtprops.AlbaranCompraItemID = mobjAlbaranCompraItem.AlbaranCompraItemID
    Cantidad = mobjAlbaranCompraItem.Cantidad
   
    PrecioCoste = mobjAlbaranCompraItem.PrecioCoste
    Descuento = mobjAlbaranCompraItem.Descuento
    Bruto = mobjAlbaranCompraItem.Bruto
   
    Set mobjAlbaranCompraItem = Nothing

End Sub

Private Sub CargarListaMateriales()
    
    Set mobjMateriales = New TextList
    mobjMateriales.Load "vNombreMateriales"

End Sub
  

