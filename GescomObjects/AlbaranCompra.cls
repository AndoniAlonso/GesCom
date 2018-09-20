VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AlbaranCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As AlbaranCompraProps
Private mudtSave As AlbaranCompraProps

Private mMoneda As String

Private mobjAlbaranCompraItems As AlbaranCompraItems

' Aqui declaro el objeto 'child' dato comercial
Private mobjDatoComercial As DatoComercial

' Aqui declaro los ComboBox de la lista de proveedores, bancos, transportistas y formas de pago.
Private mobjProveedores As TextList
Private mobjTransportistas As TextList

Event Valid(IsValid As Boolean)

Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean

Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Sub BeginEdit(Moneda As String)

    If mflgEditing Then Err.Raise 445
    
    ' save object state
    LSet mudtSave = mudtprops
    
    mMoneda = Moneda
    mobjAlbaranCompraItems.BeginEdit Moneda
    mobjDatoComercial.BeginEdit
    
    mflgEditing = True

End Sub

Public Sub CancelEdit()

    If Not mflgEditing Then Err.Raise 445
    
    mflgEditing = False
    mflgDeleted = False
    mflgDirty = False
    
    ' restore object state
    LSet mudtprops = mudtSave
    mobjAlbaranCompraItems.CancelEdit
    mobjDatoComercial.CancelEdit
  
End Sub

Public Sub ApplyEdit()
    Dim objPersist As AlbaranCompraPersist
    Dim strItems As String
    Dim strArticulos As String
  
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("GCServerMTS.AlbaranCompraPersist", PERSIST_SERVER)

    TotalBruto = mobjAlbaranCompraItems.Bruto
    
    If mflgDeleted And Not mflgNew Then
        mobjAlbaranCompraItems.Actualizar
        
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.AlbaranCompraID, mudtprops.DatoComercialID, mobjAlbaranCompraItems.GetMateriales, mobjAlbaranCompraItems.GetArticulos

        mobjDatoComercial.Delete
        mobjDatoComercial.ApplyEdit
    
        mflgNew = True
        mflgDeleted = False

    ElseIf IsDirty Or mobjAlbaranCompraItems.IsDirty Or mflgNew Then
        If Not IsValid Then Err.Raise 445
            mudtprops.DatoComercial.Buffer = mobjDatoComercial.GetState
            mobjAlbaranCompraItems.Actualizar
           
            strItems = mobjAlbaranCompraItems.GetMateriales
            strArticulos = mobjAlbaranCompraItems.GetArticulos
            SetState objPersist.Save(GetState, strItems, strArticulos)
            mobjAlbaranCompraItems.SetMateriales (strItems)
            mobjAlbaranCompraItems.SetArticulos (strArticulos)
            
            mobjDatoComercial.SetState (mudtprops.DatoComercial.Buffer)
            mobjDatoComercial.ApplyEdit
            mudtprops.DatoComercialID = mobjDatoComercial.DatoComercialID
            
            ' save object state
            LSet mudtSave = mudtprops
            mflgNew = False
            
    Else
        mobjDatoComercial.ApplyEdit
    
    End If

    mobjAlbaranCompraItems.ApplyEdit  'mudtprops.AlbaranCompraID
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
  
    Set mobjDatoComercial = New DatoComercial
      
    mudtprops.TemporadaID = 0
    mudtprops.EmpresaID = 0
    mudtprops.Fecha = Date
    mudtprops.Numero = 0
    mudtprops.NuestraReferencia = vbNullString
    mudtprops.SuReferencia = vbNullString
    mudtprops.Observaciones = vbNullString
    mudtprops.PortesEUR = 0
    mudtprops.PortesPTA = 0
    mudtprops.EmbalajesEUR = 0
    mudtprops.EmbalajesPTA = 0
    mudtprops.TotalBrutoEUR = 0
    mudtprops.TotalBrutoPTA = 0
      
    Set mobjProveedores = New TextList
    mobjProveedores.Load "Proveedores"
    mudtprops.ProveedorID = mobjProveedores.Key(mobjProveedores.IndexItem(1))
  
    Set mobjTransportistas = New TextList
    mobjTransportistas.Load "Transportistas"
    mudtprops.TransportistaID = mobjTransportistas.Key(mobjTransportistas.IndexItem(1))
  
    mobjValid.RuleBroken "TemporadaID", True
    mobjValid.RuleBroken "EmpresaID", True
    mobjValid.RuleBroken "ProveedorID", True
    'mobjValid.RuleBroken "TransportistaID", True
        
    Set mobjAlbaranCompraItems = New AlbaranCompraItems
    
End Sub
Private Sub Class_Terminate()

    Set mobjValid = Nothing
  
    Set mobjDatoComercial = Nothing
      
    Set mobjProveedores = Nothing
  
    Set mobjTransportistas = Nothing
        
    Set mobjAlbaranCompraItems = Nothing
    
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

Public Sub Load(AlbaranCompraID As Long, Moneda As String)

    Dim objPersist As AlbaranCompraPersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.AlbaranCompraPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(AlbaranCompraID)
    Set objPersist = Nothing

    mobjDatoComercial.SetState (mudtprops.DatoComercial.Buffer)
  
    mobjValid.RuleBroken "TemporadaID", False
    mobjValid.RuleBroken "EmpresaID", False
    mobjValid.RuleBroken "ProveedorID", False
    'mobjValid.RuleBroken "TransportistaID", False
      
    mMoneda = Moneda
    mobjAlbaranCompraItems.Load AlbaranCompraID, Moneda
  
    mflgNew = False

End Sub

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
  
    mflgDeleted = True
    mflgDirty = True

    mobjAlbaranCompraItems.Delete
    mobjDatoComercial.Delete

End Sub

Public Property Get IsDeleted() As Boolean

    IsDeleted = mflgDeleted

End Property

Public Property Get IsNew() As Boolean

    IsNew = mflgNew

End Property

Public Property Get IsDirty() As Boolean

    IsDirty = mflgDirty Or mobjDatoComercial.IsDirty

End Property

Public Property Get AlbaranCompraID() As Long

    AlbaranCompraID = mudtprops.AlbaranCompraID
  
End Property

Public Property Let TemporadaID(Value As Long)

    If Not mflgNew Then Err.Raise 383
  
    mobjValid.RuleBroken "TemporadaID", (Value = 0)

    mudtprops.TemporadaID = Value
    mflgDirty = True

End Property

Public Property Get TemporadaID() As Long

    TemporadaID = mudtprops.TemporadaID
  
End Property

Public Property Let EmpresaID(Value As Long)

    If Not mflgNew Then Err.Raise 383
  
    mobjValid.RuleBroken "EmpresaID", (Value = 0)

    mudtprops.EmpresaID = Value
    mflgDirty = True

End Property

Public Property Get EmpresaID() As Long

    EmpresaID = mudtprops.EmpresaID
  
End Property

Public Property Get Moneda() As String

    Moneda = Trim$(mMoneda)

End Property

Public Property Let Portes(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
  
    mobjValid.RuleBroken "Portes", (Value < 0)
  
    If EsEUR(mMoneda) Then
        mudtprops.PortesEUR = Round(Value, 2)
        mudtprops.PortesPTA = EUR2PTA(Value)
    ElseIf EsPTA(mMoneda) Then
        mudtprops.PortesPTA = Round(Value, 0)
        mudtprops.PortesEUR = PTA2EUR(Value)
    Else
        mudtprops.PortesPTA = 0
        mudtprops.PortesEUR = 0
    End If
        
    mflgDirty = True

End Property

Public Property Get Portes() As Double

    If EsEUR(mMoneda) Then
        Portes = mudtprops.PortesEUR
    ElseIf EsPTA(mMoneda) Then
        Portes = mudtprops.PortesPTA
    Else
        Portes = 0
    End If
     
End Property

Public Property Let Embalajes(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
        
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
  
    mobjValid.RuleBroken "Embalajes", (Value < 0)
      
    If EsEUR(mMoneda) Then
        mudtprops.EmbalajesEUR = Round(Value, 2)
        mudtprops.EmbalajesPTA = EUR2PTA(Value)
    ElseIf EsPTA(mMoneda) Then
        mudtprops.EmbalajesPTA = Round(Value, 0)
        mudtprops.EmbalajesEUR = PTA2EUR(Value)
    Else
        mudtprops.EmbalajesPTA = 0
        mudtprops.EmbalajesEUR = 0
    End If
        
    mflgDirty = True
    
End Property

Public Property Get Embalajes() As Double

    If EsEUR(mMoneda) Then
        Embalajes = mudtprops.EmbalajesEUR
    ElseIf EsPTA(mMoneda) Then
        Embalajes = mudtprops.EmbalajesPTA
    Else
        Embalajes = 0
    End If
    
End Property

Public Property Get Numero() As Long

    Numero = mudtprops.Numero
  
End Property

Public Property Let Numero(Value As Long)

    If Not mflgEditing Then Err.Raise 383
  
    If Value <= 0 Then _
        Err.Raise vbObjectError + 1001, "Dato numerico no puede ser negativo"
  
    mobjValid.RuleBroken "Numero", (Value <= 0)
      
    mudtprops.Numero = Value
      
    mflgDirty = True

End Property

Public Property Get ProveedorID() As Long

    ProveedorID = mudtprops.ProveedorID
  
End Property

' Esta funcion lee el proveedor para asignar sus propiedades por defecto
Public Property Let Proveedor(Value As String)

    Dim objProveedor As Proveedor

    If Not mflgEditing Then Err.Raise 383
  
    mobjValid.RuleBroken "ProveedorID", (mobjProveedores.Key(Value) = 0)
    
    'Salimos si queremos asignar algo invalido
    If mobjProveedores.Key(Value) = 0 Then Exit Property
     
    'Salimos si queremos asignar el proveedor que ya estaba
    If mobjProveedores.Key(Value) = mudtprops.ProveedorID Then Exit Property
     
    ' Devolvemos error si ya ha habido entregas y queremos cambiar de cliente.
    If mobjAlbaranCompraItems.Facturado Then _
        Err.Raise vbObjectError + 1001, "No se puede cambiar el cliente, ya se han realizado entregas"
     
    ' Devolvemos error si se han asignado a partir de pedidos (por lo tanto pertenecen a un cliente).
    If mobjAlbaranCompraItems.HayPedido Then _
        Err.Raise vbObjectError + 1001, "No se puede cambiar el cliente, ya se han asignado pedidos pendientes"
     
    ' Cambiamos de proveedor
    Set objProveedor = New Proveedor
  
    objProveedor.Load mobjProveedores.Key(Value)
  
    mudtprops.TransportistaID = objProveedor.TransportistaID
    With objProveedor.DatoComercial
        mobjDatoComercial.ChildBeginEdit
        mobjDatoComercial.Descuento = .Descuento
        mobjDatoComercial.RecargoEquivalencia = .RecargoEquivalencia
        mobjDatoComercial.IVA = .IVA
        mobjDatoComercial.ChildApplyEdit
    End With
      
    mudtprops.ProveedorID = objProveedor.ProveedorID
      
    Set objProveedor = Nothing
      
    mflgDirty = True

End Property

Public Property Get Proveedor() As String

    Proveedor = mobjProveedores.Item(mudtprops.ProveedorID)

End Property

Public Property Get Proveedores() As TextList

    Set Proveedores = mobjProveedores

End Property

Public Property Let Fecha(Value As Variant)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Trim$(Value)) = 0 Then
        mudtprops.Fecha = vbNullString
        mobjValid.RuleBroken "Fecha", True

    ElseIf IsDate(Value) Then
        mudtprops.Fecha = CDate(Value)
        mobjValid.RuleBroken "Fecha", False

    Else
        mobjValid.RuleBroken "Fecha", True
    
    End If

    mflgDirty = True

End Property

Public Property Get Fecha() As Variant

    With mudtprops
        Fecha = IIf(.Fecha = 0, vbNullString, .Fecha)
    End With

End Property

Public Property Let NuestraReferencia(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.NuestraReferencia) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mudtprops.NuestraReferencia = Trim$(Value)
    mflgDirty = True

End Property

Public Property Get NuestraReferencia() As String

    NuestraReferencia = Trim$(mudtprops.NuestraReferencia)

End Property

Public Property Let SuReferencia(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.SuReferencia) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mudtprops.SuReferencia = Trim$(Value)
    mflgDirty = True

End Property

Public Property Get SuReferencia() As String

    SuReferencia = Trim$(mudtprops.SuReferencia)

End Property

Public Property Let Observaciones(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Observaciones) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mudtprops.Observaciones = Trim$(Value)
    mflgDirty = True

End Property

Public Property Get Observaciones() As String

    Observaciones = Trim$(mudtprops.Observaciones)

End Property


Public Property Let TotalBruto(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
        
'    If Value < 0 Then _
'        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
      
'    mobjValid.RuleBroken "TotalBruto", (Value < 0)
      
    If EsEUR(mMoneda) Then
        mudtprops.TotalBrutoEUR = Round(Value, 2)
        mudtprops.TotalBrutoPTA = EUR2PTA(Value)
    ElseIf EsPTA(mMoneda) Then
        mudtprops.TotalBrutoPTA = Round(Value, 0)
        mudtprops.TotalBrutoEUR = PTA2EUR(Value)
    Else
        mudtprops.TotalBrutoPTA = 0
        mudtprops.TotalBrutoEUR = 0
    End If
        
    mflgDirty = True

End Property

Public Property Get TotalBruto() As Double

    TotalBruto = mobjAlbaranCompraItems.Bruto
    
End Property

Public Property Get Cantidad() As Double

    Cantidad = mobjAlbaranCompraItems.Cantidad
    
End Property

Public Property Get TransportistaID() As Long

    TransportistaID = mudtprops.TransportistaID
  
End Property

Public Property Let Transportista(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    mudtprops.TransportistaID = mobjTransportistas.Key(Value)
  
    'mobjValid.RuleBroken "TransportistaID", (mudtprops.TransportistaID = 0)
  
    mflgDirty = True

End Property

Public Property Get Transportista() As String

    Transportista = mobjTransportistas.Item(mudtprops.TransportistaID)

End Property

Public Property Get Transportistas() As TextList

    Set Transportistas = mobjTransportistas

End Property

Public Property Get AlbaranCompraItems() As AlbaranCompraItems

    Set AlbaranCompraItems = mobjAlbaranCompraItems

End Property

Public Property Let DatoComercialID(Value As Long)
  
    If Not mflgEditing Then Err.Raise 383
      
    mudtprops.DatoComercialID = Value
    mflgDirty = True

End Property

Public Property Get DatoComercialID() As Long

    DatoComercialID = mudtprops.DatoComercialID
  
End Property

Public Property Get DatoComercial() As DatoComercial
  
    Set DatoComercial = mobjDatoComercial
  
End Property

Private Function GetState() As String

    Dim udtData As AlbaranCompraData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With
    
    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

    Dim udtData As AlbaranCompraData
  
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub
