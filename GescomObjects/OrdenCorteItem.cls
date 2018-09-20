VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrdenCorteItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Valid(IsValid As Boolean)

Private mMoneda As String

Private mudtprops As OrdenCorteItemProps
Private mudtSave As OrdenCorteItemProps
Private mudtChild As OrdenCorteItemProps
Private mobjParent As OrdenCorteItems
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
    Dim objPersist As OrdenCorteItemPersist
    
    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgChild And mflgNew Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445
    
    If OrdenCortada Then _
       Err.Raise vbObjectError + 1001, "Modificar orden de corte", "No se puede modificar, ya se ha actualizado la orden"
    
    Set objPersist = CreateObject("GCServerMTS.OrdenCorteItemPersist", PERSIST_SERVER)
    
    If Not mflgChild Then
      If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.OrdenCorteItemID
        mflgNew = True
        mflgDeleted = False
    
      ElseIf mflgDirty Or mflgNew Then
        If Not IsValid Then Err.Raise 445
        ' save object to database if appropriate
        SetState objPersist.Save(GetState)
        ' save object state
        LSet mudtSave = mudtprops
        mflgNew = False
    
      End If
      
      mflgDirty = False
    
    ElseIf mflgNew Then
      If Not mobjParent Is Nothing Then
        mobjParent.AddOrdenCorteItem Me
        Set mobjParent = Nothing
      End If
    
    
    End If
      
    Set objPersist = Nothing
    mflgEditing = False
    
End Sub

Private Sub Class_Initialize()

    mflgNew = True
    Set mobjValid = New BrokenRules
    
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
    mobjValid.RuleBroken "PedidoVentaItemID", True
    
    With mudtprops
        .Situacion = "A"    ' Marco la situacion de la orden de corte como "A"--> alta,
                            ' pendiente de codificar las posibles situaciones.
        .Descripcion = Trim(vbNullString)
        .ArticuloColorID = 0
        .ArticuloID = 0
        .PedidoVentaItemID = 0
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
        .Actualizar = False
        .Desactualizar = False
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

Public Sub Load(OrdenCorteItemID As Long, Moneda As String)
    Dim objPersist As OrdenCorteItemPersist
    
    If mflgChild Then Err.Raise 445
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
    
    mMoneda = Moneda
    
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.OrdenCorteItemPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(OrdenCorteItemID)
    Set objPersist = Nothing
  
    mobjValid.RuleBroken "PedidoVentaItemID", True

    mflgNew = False

End Sub

Public Sub Delete()

    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445
    
    If OrdenCortada Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha actualizado la orden"
    
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

Friend Sub SetAsChild(objOrdenCorteItems As OrdenCorteItems, Moneda As String)

    mMoneda = Moneda
    If mflgNew Then Set mobjParent = objOrdenCorteItems
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

Public Property Get OrdenCorteItemID() As Long

    OrdenCorteItemID = mudtprops.OrdenCorteItemID

End Property

Public Property Let PedidoVentaItemID(Value As Long)

    mudtprops.PedidoVentaItemID = Value

End Property

Public Property Get PedidoVentaItemID() As Long

    PedidoVentaItemID = mudtprops.PedidoVentaItemID
    
    mobjValid.RuleBroken "PedidoVentaItemID", (mudtprops.PedidoVentaItemID = 0)

End Property

Public Property Get Situacion() As String
    
    Select Case mudtprops.Situacion
    Case "A"
       Situacion = "Orden Creada"
    Case "C"
       Situacion = "Actualizada"
    Case Else
      Err.Raise vbObjectError + 1001, "Situacion de orden de corte"
    End Select

End Property

Public Property Get ArticuloColorID() As Long

    ArticuloColorID = mudtprops.ArticuloColorID
  
End Property

Public Property Get ArticuloID() As Long

    ArticuloID = mudtprops.ArticuloID
  
End Property

Public Property Get CantidadT36() As Double

    CantidadT36 = mudtprops.CantidadT36

End Property

Public Property Get CantidadT38() As Double

    CantidadT38 = mudtprops.CantidadT38

End Property

Public Property Get CantidadT40() As Double

    CantidadT40 = mudtprops.CantidadT40

End Property

Public Property Get CantidadT42() As Double

    CantidadT42 = mudtprops.CantidadT42

End Property

Public Property Get CantidadT44() As Double

    CantidadT44 = mudtprops.CantidadT44

End Property

Public Property Get CantidadT46() As Double

    CantidadT46 = mudtprops.CantidadT46

End Property

Public Property Get CantidadT48() As Double

    CantidadT48 = mudtprops.CantidadT48

End Property

Public Property Get CantidadT50() As Double

    CantidadT50 = mudtprops.CantidadT50

End Property

Public Property Get CantidadT52() As Double

    CantidadT52 = mudtprops.CantidadT52

End Property

Public Property Get CantidadT54() As Double

    CantidadT54 = mudtprops.CantidadT54

End Property

Public Property Get CantidadT56() As Double

    CantidadT56 = mudtprops.CantidadT56

End Property

Public Property Get Descripcion() As String

    Descripcion = Trim(mudtprops.Descripcion)

End Property

Public Property Get Numero() As Long

    Numero = mudtprops.Numero

End Property

Public Property Get Cliente() As String

    Cliente = Trim(mudtprops.Cliente)

End Property

Public Property Get Cantidad() As Double
    
    With mudtprops
    Cantidad = .CantidadT36 + .CantidadT38 + .CantidadT40 + .CantidadT42 + _
               .CantidadT44 + .CantidadT46 + .CantidadT48 + .CantidadT50 + _
               .CantidadT52 + .CantidadT54 + .CantidadT56
    End With

End Property

Public Property Get OrdenCortada() As Boolean
  
  OrdenCortada = (mudtprops.Situacion = "F")
  
End Property

Friend Function GetState() As String
    Dim udtData As OrdenCorteItemData
    
    With mudtprops
        .IsNew = mflgNew
        .IsDeleted = mflgDeleted
        .IsDirty = mflgDirty
    End With
    
    LSet udtData = mudtprops
    GetState = udtData.Buffer
    mflgNew = False

End Function

Friend Sub SetState(Buffer As String)
    Dim udtData As OrdenCorteItemData
    
    udtData.Buffer = Buffer
    LSet mudtprops = udtData
    mflgNew = False
    mobjValid.RuleBroken "PedidoVentaItemID", True

End Sub

Public Sub Actualizar()
    
    If Not mflgChildEditing Then Err.Raise 383

    If Not OrdenCortada Then
        mudtprops.Actualizar = True
        mudtprops.Desactualizar = False
    End If
        
    mflgDirty = True
    
End Sub

Public Sub Desactualizar()

    If Not mflgChildEditing Then Err.Raise 383

    If Not OrdenCortada Then
        mudtprops.Actualizar = False
        mudtprops.Desactualizar = True
    End If
   
    mflgDirty = False
    
End Sub

' Este procedimiento crea los datos de una orden de corte a partir de los de un pedido.
' Hay que asegurarse de que el pedido tiene
Public Function OrdenDesdePedido(PedidoItemID As Long) As Long
    Dim objPedidoVentaItem As PedidoVentaItem
    Dim objPedidoVenta As PedidoVenta
    Dim objArticuloColor As ArticuloColor

    If Not mflgEditing Then Err.Raise 383
    
    Set objPedidoVentaItem = New PedidoVentaItem
    Set objPedidoVenta = New PedidoVenta
    Set objArticuloColor = New ArticuloColor
    
    objPedidoVentaItem.Load PedidoItemID, mMoneda
      
    objPedidoVenta.Load objPedidoVentaItem.PedidoVentaID, mMoneda
    
    With mudtprops
        .PedidoVentaItemID = objPedidoVentaItem.PedidoVentaItemID
        .ArticuloColorID = objPedidoVentaItem.ArticuloColorID
        objArticuloColor.Load .ArticuloColorID, mMoneda
        .ArticuloID = objArticuloColor.ObjArticulo.ArticuloID
        .Descripcion = Trim(objArticuloColor.ObjArticulo.NombreCompleto & " " & objArticuloColor.NombreColor)
        
        .CantidadT36 = objPedidoVentaItem.CantidadT36
        .CantidadT38 = objPedidoVentaItem.CantidadT38
        .CantidadT40 = objPedidoVentaItem.CantidadT40
        .CantidadT42 = objPedidoVentaItem.CantidadT42
        .CantidadT44 = objPedidoVentaItem.CantidadT44
        .CantidadT46 = objPedidoVentaItem.CantidadT46
        .CantidadT48 = objPedidoVentaItem.CantidadT48
        .CantidadT50 = objPedidoVentaItem.CantidadT50
        .CantidadT52 = objPedidoVentaItem.CantidadT52
        .CantidadT54 = objPedidoVentaItem.CantidadT54
        .CantidadT56 = objPedidoVentaItem.CantidadT56
        .Numero = objPedidoVenta.Numero
        .Cliente = objPedidoVenta.Cliente
    End With
    
    Set objPedidoVenta = Nothing
    Set objPedidoVentaItem = Nothing
    Set objArticuloColor = Nothing
    
End Function

