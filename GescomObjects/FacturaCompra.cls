VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FacturaCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As FacturaCompraProps
Private mudtSave As FacturaCompraProps

Private mMoneda As String

Private mobjFacturaCompraItems As FacturaCompraItems

Private mobjCobrosPagos As CobrosPagos

' Aqui declaro el objeto 'child' dato comercial
Private mobjDatoComercial As DatoComercial

' Aqui declaro los ComboBox de la lista de proveedores,
' bancos, transportistas, formas de pago, y medios de pago.
Private mobjProveedores As TextList
Private mobjBancos As TextList
Private mobjTransportistas As TextList
Private mobjFormasPago As TextList
Private mobjMediosPago As TextList

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
    mobjFacturaCompraItems.BeginEdit Moneda
    mobjCobrosPagos.BeginEdit Moneda
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
    mobjFacturaCompraItems.CancelEdit
    mobjCobrosPagos.CancelEdit
    mobjDatoComercial.CancelEdit
  
End Sub

Public Sub ApplyEdit()
    Dim objPersist As FacturaCompraPersist
    Dim strItems As String
    Dim strCobrosPagos As String
  
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("GCServerMTS.FacturaCompraPersist", PERSIST_SERVER)

    If mflgDeleted And Not mflgNew Then
        mobjFacturaCompraItems.Actualizar
        
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.FacturaCompraID, mudtprops.DatoComercialID, mobjFacturaCompraItems.GetState, mobjCobrosPagos.GetState

        mobjDatoComercial.Delete
        mobjDatoComercial.ApplyEdit
    
        mflgNew = True
        mflgDeleted = False

    ElseIf IsDirty Or mflgNew Then
            If Not IsValid Then Err.Raise 445
            mudtprops.DatoComercial.Buffer = mobjDatoComercial.GetState
            
            mobjFacturaCompraItems.Actualizar
            
            strItems = mobjFacturaCompraItems.GetState
            strCobrosPagos = mobjCobrosPagos.GetState
            SetState objPersist.Save(GetState, strItems, strCobrosPagos)
            mobjFacturaCompraItems.SetState (strItems)
            mobjCobrosPagos.SetState (strCobrosPagos)
            
            mobjDatoComercial.SetState (mudtprops.DatoComercial.Buffer)
            mobjDatoComercial.ApplyEdit
            mudtprops.DatoComercialID = mobjDatoComercial.DatoComercialID
 
            ' save object state
            LSet mudtSave = mudtprops
            mflgNew = False
    Else
        mobjDatoComercial.ApplyEdit

    End If

    mobjFacturaCompraItems.ApplyEdit '  mudtprops.FacturaCompraID
    mobjCobrosPagos.ApplyEdit
    Set objPersist = Nothing
    mflgDirty = False
    mflgEditing = False

End Sub

Private Sub Class_Initialize()
    Dim objParametro As Parametro

    mflgNew = True
    Set mobjValid = New BrokenRules
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
  
    Set mobjDatoComercial = New DatoComercial
  
    mudtprops.TemporadaID = 0
    mudtprops.EmpresaID = 0
    mudtprops.Fecha = Date
    mudtprops.FechaContable = Date
    mudtprops.Numero = 0
    mudtprops.Sufijo = vbNullString
    mudtprops.NuestraReferencia = vbNullString
    mudtprops.SuReferencia = vbNullString
    mudtprops.Observaciones = vbNullString
    mudtprops.SituacionContable = "A"    ' Marco la situacion de la Factura como "A"--> alta,
                                         ' pendiente de codificar las posibles situaciones.
    mudtprops.BrutoEUR = 0
    mudtprops.BrutoPTA = 0
    mudtprops.DescuentoEUR = 0
    mudtprops.DescuentoPTA = 0
    mudtprops.PortesEUR = 0
    mudtprops.PortesPTA = 0
    mudtprops.EmbalajesEUR = 0
    mudtprops.EmbalajesPTA = 0
    mudtprops.BaseImponibleEUR = 0
    mudtprops.BaseImponiblePTA = 0
    mudtprops.IVAEUR = 0
    mudtprops.IVAPTA = 0
    mudtprops.RecargoEUR = 0
    mudtprops.RecargoPTA = 0
    mudtprops.NetoEUR = 0
    mudtprops.NetoPTA = 0
        
    Set mobjProveedores = New TextList
    mobjProveedores.Load "Proveedores"
    mudtprops.ProveedorID = mobjProveedores.Key(mobjProveedores.IndexItem(1))
  
    Set mobjBancos = New TextList
    mobjBancos.Load "vBancos"
    mudtprops.BancoID = mobjBancos.Key(mobjBancos.IndexItem(1))
  
    Set mobjTransportistas = New TextList
    mobjTransportistas.Load "Transportistas"
    mudtprops.TransportistaID = mobjTransportistas.Key(mobjTransportistas.IndexItem(1))
  
    Set mobjFormasPago = New TextList
    mobjFormasPago.Load "FormasDePago"
    mudtprops.FormaPagoID = mobjFormasPago.Key(mobjFormasPago.IndexItem(1))
  
    Set mobjMediosPago = New TextList
    mobjMediosPago.Load "MediosPago"
    mudtprops.MedioPagoID = mobjMediosPago.Key(mobjMediosPago.IndexItem(1))
  
    mobjValid.RuleBroken "TemporadaID", True
    mobjValid.RuleBroken "EmpresaID", True
    mobjValid.RuleBroken "ProveedorID", True
    'mobjValid.RuleBroken "BancoID", True
    'mobjValid.RuleBroken "TransportistaID", True
    mobjValid.RuleBroken "FormaPagoID", True
    mobjValid.RuleBroken "MedioPagoID", True
    
    Set mobjFacturaCompraItems = New FacturaCompraItems
    Set mobjCobrosPagos = New CobrosPagos
    
    Set objParametro = New Parametro
    objParametro.Load
    mudtprops.Sufijo = objParametro.Sufijo
    Set objParametro = Nothing
    
End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing
    Set mobjDatoComercial = Nothing
    Set mobjProveedores = Nothing
    Set mobjBancos = Nothing
    Set mobjTransportistas = Nothing
    Set mobjFormasPago = Nothing
    Set mobjMediosPago = Nothing
    Set mobjFacturaCompraItems = Nothing
    Set mobjCobrosPagos = Nothing
    
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

Public Sub Load(FacturaCompraID As Long, Moneda As String)

    Dim objPersist As FacturaCompraPersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.FacturaCompraPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(FacturaCompraID)
    Set objPersist = Nothing

    mobjDatoComercial.SetState (mudtprops.DatoComercial.Buffer)
  
    mobjValid.RuleBroken "TemporadaID", False
    mobjValid.RuleBroken "EmpresaID", False
    mobjValid.RuleBroken "ProveedorID", False
    'mobjValid.RuleBroken "BancoID", False
    'mobjValid.RuleBroken "TransportistaID", False
    mobjValid.RuleBroken "FormaPagoID", False
    mobjValid.RuleBroken "MedioPagoID", False
  
    mMoneda = Moneda
    mobjFacturaCompraItems.Load FacturaCompraID, Moneda
    mobjCobrosPagos.Load FacturaCompraID, "P", Moneda
  
    mflgNew = False

End Sub

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
  
    If Contabilizado Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar, ya se ha contabilizado la factura"
    
    mflgDeleted = True
    mflgDirty = True

    mobjFacturaCompraItems.Delete
    mobjCobrosPagos.Delete
    mobjDatoComercial.Delete

End Sub

Public Property Get IsDeleted() As Boolean

    IsDeleted = mflgDeleted

End Property

Public Property Get IsNew() As Boolean

    IsNew = mflgNew

End Property

Public Property Get IsDirty() As Boolean

    IsDirty = mflgDirty Or mobjDatoComercial.IsDirty Or mobjFacturaCompraItems.IsDirty Or mobjCobrosPagos.IsDirty

End Property

Public Property Get FacturaCompraID() As Long

    FacturaCompraID = mudtprops.FacturaCompraID
  
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

    Moneda = Trim(mMoneda)

End Property

Public Property Let TotalBruto(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
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
  
    CalcularDescuento
    
    mflgDirty = True

End Property

Public Property Get TotalBruto() As Double

    If EsEUR(mMoneda) Then
        TotalBruto = mudtprops.BrutoEUR
    ElseIf EsPTA(mMoneda) Then
        TotalBruto = mudtprops.BrutoPTA
    Else
        TotalBruto = 0
    End If
    
End Property

Public Property Let Descuento(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If EsEUR(mMoneda) Then
        mudtprops.DescuentoEUR = Round(Value, 2)
        mudtprops.DescuentoPTA = EUR2PTA(Value)
    ElseIf EsPTA(mMoneda) Then
        mudtprops.DescuentoPTA = Round(Value, 0)
        mudtprops.DescuentoEUR = PTA2EUR(Value)
    Else
        mudtprops.DescuentoPTA = 0
        mudtprops.DescuentoEUR = 0
    End If
    
    CalcularBaseImponible
    
    mflgDirty = True

End Property

Public Property Get Descuento() As Double

    If EsEUR(mMoneda) Then
        Descuento = mudtprops.DescuentoEUR
    ElseIf EsPTA(mMoneda) Then
        Descuento = mudtprops.DescuentoPTA
    Else
        Descuento = 0
    End If
 
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
    
    CalcularBaseImponible
  
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
    
    CalcularBaseImponible
  
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

Public Property Let BaseImponible(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If EsEUR(mMoneda) Then
        mudtprops.BaseImponibleEUR = Round(Value, 2)
        mudtprops.BaseImponiblePTA = EUR2PTA(Value)
    ElseIf EsPTA(mMoneda) Then
        mudtprops.BaseImponiblePTA = Round(Value, 0)
        mudtprops.BaseImponibleEUR = PTA2EUR(Value)
    Else
        mudtprops.BaseImponiblePTA = 0
        mudtprops.BaseImponibleEUR = 0
    End If
    
    CalcularIVA
    CalcularRecargo
  
    mflgDirty = True

End Property

Public Property Get BaseImponible() As Double

    If EsEUR(mMoneda) Then
        BaseImponible = mudtprops.BaseImponibleEUR
    ElseIf EsPTA(mMoneda) Then
        BaseImponible = mudtprops.BaseImponiblePTA
    Else
        BaseImponible = 0
    End If
 
End Property

Public Property Let IVA(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If EsEUR(mMoneda) Then
        mudtprops.IVAEUR = Round(Value, 2)
        mudtprops.IVAPTA = EUR2PTA(Value)
    ElseIf EsPTA(mMoneda) Then
        mudtprops.IVAPTA = Round(Value, 0)
        mudtprops.IVAEUR = PTA2EUR(Value)
    Else
        mudtprops.IVAPTA = 0
        mudtprops.IVAEUR = 0
    End If
    
    CalcularNeto
  
    mflgDirty = True

End Property

Public Property Get IVA() As Double

    If EsEUR(mMoneda) Then
        IVA = mudtprops.IVAEUR
    ElseIf EsPTA(mMoneda) Then
        IVA = mudtprops.IVAPTA
    Else
        IVA = 0
    End If
     
End Property

Public Property Let Recargo(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If EsEUR(mMoneda) Then
        mudtprops.RecargoEUR = Round(Value, 2)
        mudtprops.RecargoPTA = EUR2PTA(Value)
    ElseIf EsPTA(mMoneda) Then
        mudtprops.RecargoPTA = Round(Value, 0)
        mudtprops.RecargoEUR = PTA2EUR(Value)
    Else
        mudtprops.RecargoPTA = 0
        mudtprops.RecargoEUR = 0
    End If
        
    CalcularNeto
  
    mflgDirty = True

End Property

Public Property Get Recargo() As Double

    If EsEUR(mMoneda) Then
        Recargo = mudtprops.RecargoEUR
    ElseIf EsPTA(mMoneda) Then
        Recargo = mudtprops.RecargoPTA
    Else
        Recargo = 0
    End If
     
End Property

Public Property Let Neto(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    If EsEUR(mMoneda) Then
        mudtprops.NetoEUR = Round(Value, 2)
        mudtprops.NetoPTA = EUR2PTA(Value)
    ElseIf EsPTA(mMoneda) Then
        mudtprops.NetoPTA = Round(Value, 0)
        mudtprops.NetoEUR = PTA2EUR(Value)
    Else
        mudtprops.NetoPTA = 0
        mudtprops.NetoEUR = 0
    End If
        
    mflgDirty = True

End Property

Public Property Get Neto() As Double

    If EsEUR(mMoneda) Then
        Neto = mudtprops.NetoEUR
    ElseIf EsPTA(mMoneda) Then
        Neto = mudtprops.NetoPTA
    Else
        Neto = 0
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

Public Property Get Sufijo() As String

    Sufijo = mudtprops.Sufijo
  
End Property

Public Property Let Sufijo(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Sufijo) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"

    mudtprops.Sufijo = Value
  
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
     
    ' Devolvemos error si ya se ha contabilizado y queremos cambiar de cliente.
    If Me.Contabilizado Then _
        Err.Raise vbObjectError + 1001, "No se puede cambiar el proveedor, ya se ha contabilizado"
     
    ' Devolvemos error si ya se han contabilizado los cobros y queremos cambiar de cliente.
    If mobjCobrosPagos.Contabilizado Then _
        Err.Raise vbObjectError + 1001, "No se puede cambiar el proveedor, ya se han contabilizado los cobros"
     
    ' Devolvemos error si se han asignado a partir de Albaranes (por lo tanto pertenecen a un cliente).
    If mobjFacturaCompraItems.HayAlbaran Then _
        Err.Raise vbObjectError + 1001, "No se puede cambiar el proveedor, ya se han asignado Albaranes pendientes"
     
    ' Cambiamos de proveedor
    Set objProveedor = New Proveedor
  
    objProveedor.Load mobjProveedores.Key(Value)
  
    mudtprops.BancoID = objProveedor.BancoID
    mudtprops.TransportistaID = objProveedor.TransportistaID
    mudtprops.FormaPagoID = objProveedor.FormaPagoID
    mudtprops.MedioPagoID = objProveedor.MedioPagoID
    
    'mobjValid.RuleBroken "BancoID", False
    'mobjValid.RuleBroken "TransportistaID", False
    mobjValid.RuleBroken "FormaPagoID", False
    mobjValid.RuleBroken "MedioPagoID", False
    
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
  
    If Len(Trim(Value)) = 0 Then
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

Public Property Let FechaContable(Value As Variant)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Trim(Value)) = 0 Then
        mudtprops.FechaContable = vbNullString
        mobjValid.RuleBroken "FechaContable", True

    ElseIf IsDate(Value) Then
        mudtprops.FechaContable = CDate(Value)
        mobjValid.RuleBroken "FechaContable", False

    Else
        mobjValid.RuleBroken "FechaContable", True

    End If

    mflgDirty = True

End Property

Public Property Get FechaContable() As Variant

    With mudtprops
        FechaContable = IIf(.FechaContable = 0, vbNullString, .FechaContable)
    End With

End Property

Public Property Let NuestraReferencia(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.NuestraReferencia) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mudtprops.NuestraReferencia = Trim(Value)
    mflgDirty = True

End Property

Public Property Get NuestraReferencia() As String

    NuestraReferencia = Trim(mudtprops.NuestraReferencia)

End Property

Public Property Let SuReferencia(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.SuReferencia) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mudtprops.SuReferencia = Trim(Value)
    mflgDirty = True

End Property

Public Property Get SuReferencia() As String

    SuReferencia = Trim(mudtprops.SuReferencia)

End Property

Public Property Let Observaciones(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Observaciones) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mudtprops.Observaciones = Trim(Value)
    mflgDirty = True

End Property

Public Property Get Observaciones() As String

    Observaciones = Trim(mudtprops.Observaciones)

End Property

Public Property Get SituacionContable() As String
  
    Select Case mudtprops.SituacionContable
        Case "A"
            SituacionContable = "Generada"
        Case "C"
            SituacionContable = "Contabilizada"
        Case Else
            Err.Raise vbObjectError + 1001, "Situacion contable de Factura erronea"
    End Select

End Property

Public Property Get Contabilizado() As Boolean
  
    Contabilizado = (mudtprops.SituacionContable = "C")
  
End Property

Public Sub Contabilizar(Optional ForzarContabilidad As Boolean = False)
    Dim objAsiento As Asiento
    Dim objProveedor As Proveedor
    Dim strDescripcion As String
    Dim objPersist As FacturaCompraPersist
    Dim strCuentaBase As String
    Dim strTitular As String
    Dim dblImporteProveedor As Double
    Dim strCuentaIVA As String
    Dim strCuentaIVAExtranjero As String
    
    'If Not mflgEditing Then Err.Raise 383
    If mflgEditing Then Err.Raise 383
    
    ' Realizamos la comprobación de si ya está contabilizada, y damos error si no se ha
    ' marcada la contabilidad como forzada.
    If Contabilizado And _
       Not ForzarContabilidad Then _
        Err.Raise vbObjectError + 1010, "La Factura ya esta marcada como contabilizada"
     
    mudtprops.SituacionContable = "C"
    
    ' Aquí por fin contabilizo
    Set objAsiento = New Asiento
    Set objProveedor = New Proveedor
    
    If Me.Neto = 0 Then Exit Sub
    
    objProveedor.Load Me.ProveedorID
    
    With objAsiento
        If Me.Neto > 0 Then
            strDescripcion = "SU FRA. Nº:" & Me.SuReferencia & _
                             "(" & Trim(CStr(Me.Numero)) & Trim(Me.Sufijo) & ")"
        Else
            strDescripcion = "DEVOLUCION SU FRA. Nº:" & Me.SuReferencia & _
                             "(" & Trim(CStr(Me.Numero)) & Trim(Me.Sufijo) & ")"
        End If
        
        objAsiento.NuevoAsiento Me.Numero, CStr(Year(Me.FechaContable)), "Factura compra nº " & Me.SuReferencia & " (" & Trim(CStr(Me.Numero)) & Trim(Me.Sufijo) & ")", _
                                Me.TemporadaID, Me.EmpresaID
                                
        ' Si el proveedor es extranjero, hay que generar:
        '  - el importe del proveedor sin el iva soportado y
        '  - un apunte de iva repercutido que equilibra el asiento.
        If objProveedor.Extranjero Then
            dblImporteProveedor = Me.Neto - Me.IVA
            strCuentaIVA = "47210000" & Format(Me.DatoComercial.IVA, "00")
        Else
            strCuentaIVA = "47200000" & Format(Me.DatoComercial.IVA, "00") 'La cuenta de IVA en funcion del tipo de IVA
            dblImporteProveedor = Me.Neto
            
        End If
        .NuevoApunte objProveedor.CuentaContable, "Haber", _
                               dblImporteProveedor, strDescripcion, _
                               Me.FechaContable, Me.SuReferencia 'Trim(CStr(Me.Numero)) & Trim(Me.Sufijo)
        
        .NuevoApunte "7650000001", "Haber", _
                               Me.Descuento, strDescripcion, _
                               Me.FechaContable, Me.SuReferencia 'Trim(CStr(Me.Numero)) & Trim(Me.Sufijo)
                               
        
        .NuevoApunte strCuentaIVA, "Debe", _
                               Me.IVA, strDescripcion, _
                               Me.FechaContable, Me.SuReferencia 'Trim(CStr(Me.Numero)) & Trim(Me.Sufijo)
        
        If Me.Neto > 0 Then
            strCuentaBase = objProveedor.CuentaContrapartida
        Else
            ' Cambiar la cuenta contable de devoluciones de compras de proveedores extranjeros.
            If objProveedor.Extranjero Then
                strCuentaBase = "6081000004"
            Else
                strCuentaBase = "6081000001"   '"6080000001" Devoluciones de materias primas, y no de mercaderias
            End If
        End If
        ' Se contabilizan tambien los portes en la cuenta de compras
        .NuevoApunte strCuentaBase, "Debe", _
                            Me.TotalBruto + Me.Portes, strDescripcion, _
                            Me.FechaContable, Me.SuReferencia 'Trim(CStr(Me.Numero)) & Trim(Me.Sufijo)
        
        ' Apuntes de IVA.
        If Me.IVA <> 0 Or Me.Recargo <> 0 Then
            strTitular = IIf(objProveedor.Titular = vbNullString, objProveedor.Nombre, objProveedor.Titular)
            .NuevoIVAApunte "S", Me.SuReferencia, strCuentaIVA, strCuentaBase, objProveedor.CuentaContable, _
                          strTitular, objProveedor.DNINIF, Me.BaseImponible, _
                        Me.Neto, Me.DatoComercial.IVA, Me.IVA, Me.DatoComercial.RecargoEquivalencia, Me.Recargo
'            .NuevoIVAApunte "S", Trim(CStr(Me.Numero)) & Trim(Me.Sufijo), strCuentaIVA, strCuentaBase, objProveedor.CuentaContable, _
'                          strTitular, objProveedor.DNINIF, Me.BaseImponible, _
'                        Me.Neto, Me.DatoComercial.IVA, Me.IVA, Me.DatoComercial.RecargoEquivalencia, Me.Recargo
        End If
        
        If objProveedor.Extranjero Then
            strCuentaIVAExtranjero = "47710000" & Format(Me.DatoComercial.IVA, "00")
            .NuevoApunte strCuentaIVAExtranjero, "Haber", _
                                   Me.IVA, strDescripcion, _
                                   Me.FechaContable, Me.SuReferencia 'Trim(CStr(Me.Numero)) & Trim(Me.Sufijo)
            
            ' Apuntes de IVA.
            If Me.IVA <> 0 Or Me.Recargo <> 0 Then
                strTitular = IIf(objProveedor.Titular = vbNullString, objProveedor.Nombre, objProveedor.Titular)
                .NuevoIVAApunte "R", Me.SuReferencia, strCuentaIVAExtranjero, strCuentaBase, objProveedor.CuentaContable, _
                              strTitular, objProveedor.DNINIF, Me.BaseImponible, _
                            Me.Neto, Me.DatoComercial.IVA, Me.IVA, Me.DatoComercial.RecargoEquivalencia, Me.Recargo
'                .NuevoIVAApunte "S", Trim(CStr(Me.Numero)) & Trim(Me.Sufijo), strCuentaIVAExtranjero, strCuentaBase, objProveedor.CuentaContable, _
'                              strTitular, objProveedor.DNINIF, Me.BaseImponible, _
'                            Me.Neto, Me.DatoComercial.IVA, Me.IVA, Me.DatoComercial.RecargoEquivalencia, Me.Recargo
            End If
        End If
        
' Se contabilizan tambien los portes en la cuenta de compras
'        .NuevoApunte "6000000002", "Debe", _
'                               Me.Portes, strDescripcion, _
'                               Me.FechaContable, Me.SuReferencia 'Trim(CStr(Me.Numero)) & Trim(Me.Sufijo)
        .NuevoApunte "6000000003", "Debe", _
                               Me.Embalajes, strDescripcion, _
                               Me.FechaContable, Me.SuReferencia 'Trim(CStr(Me.Numero)) & Trim(Me.Sufijo)
                               
        .CuadrarAsiento objProveedor.CuentaContable, Me.FechaContable, Me.SuReferencia 'Trim(CStr(Me.Numero)) & Trim(Me.Sufijo)
        
        
    End With
    
    Set objPersist = CreateObject("GCServerMTS.FacturaCompraPersist", PERSIST_SERVER)
    objPersist.Contabilizar mudtprops.FacturaCompraID, _
                            mudtprops.SituacionContable, _
                            objAsiento.GetSuperState
    Set objPersist = Nothing

    Set objAsiento = Nothing
    Set objProveedor = Nothing
    
End Sub

Public Sub DesContabilizar()
    
    If Not mflgEditing Then Err.Raise 383
    
    ' Realizamos la comprobación de si no está contabilizada, y damos error si no se ha
    ' marcada la contabilidad como forzada.
    If Not Contabilizado Then _
        Err.Raise vbObjectError + 1010, "La Factura no esta marcada como contabilizada"
     
    mudtprops.SituacionContable = "A"
    mflgDirty = True
    
End Sub

Public Property Get BancoID() As Long

    BancoID = mudtprops.BancoID
  
End Property

Public Property Let Banco(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    mudtprops.BancoID = mobjBancos.Key(Value)
  
    'mobjValid.RuleBroken "BancoID", (mudtprops.BancoID = 0)
  
    mflgDirty = True

End Property

Public Property Get Banco() As String

    Banco = mobjBancos.Item(mudtprops.BancoID)

End Property

Public Property Get Bancos() As TextList

    Set Bancos = mobjBancos

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

Public Property Get FormaPagoID() As Long

    FormaPagoID = mudtprops.FormaPagoID
  
End Property

Public Property Let FormaPago(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    mudtprops.FormaPagoID = mobjFormasPago.Key(Value)
  
    mobjValid.RuleBroken "FormaPagoID", (mudtprops.FormaPagoID = 0)
  
    mflgDirty = True

End Property

Public Property Get FormaPago() As String

    FormaPago = mobjFormasPago.Item(mudtprops.FormaPagoID)

End Property

Public Property Get FormasPago() As TextList

    Set FormasPago = mobjFormasPago

End Property

Public Property Get MedioPagoID() As Long

    MedioPagoID = mudtprops.MedioPagoID
  
End Property

Public Property Let MedioPago(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    mudtprops.MedioPagoID = mobjMediosPago.Key(Value)
  
    mobjValid.RuleBroken "MedioPagoID", (mudtprops.MedioPagoID = 0)
  
    mflgDirty = True

End Property

Public Property Get MedioPago() As String

    MedioPago = mobjMediosPago.Item(mudtprops.MedioPagoID)

End Property

Public Property Get MediosPago() As TextList

    Set MediosPago = mobjMediosPago

End Property

Public Property Get FacturaCompraItems() As FacturaCompraItems

    Set FacturaCompraItems = mobjFacturaCompraItems

End Property

Public Property Get CobrosPagos() As CobrosPagos

    Set CobrosPagos = mobjCobrosPagos

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

Public Sub CalcularBruto()
    
    TotalBruto = mobjFacturaCompraItems.Bruto
    
End Sub

Public Sub CalcularDescuento()
   
    Descuento = TotalBruto * mobjDatoComercial.Descuento / 100
    
End Sub

Public Sub CalcularBaseImponible()
   
    BaseImponible = (TotalBruto - Descuento + Portes + Embalajes)
    
End Sub

Public Sub CalcularRecargo()
    
    Recargo = BaseImponible * mobjDatoComercial.RecargoEquivalencia / 100
    
End Sub

Public Sub CalcularIVA()
   
    IVA = BaseImponible * mobjDatoComercial.IVA / 100
    
End Sub

Public Sub CalcularNeto()
    
    Neto = BaseImponible + Recargo + IVA
    
End Sub

Public Sub CrearPagos()
    Dim objFormaPago As FormaDePago
    Dim i As Integer
    Dim FechaGiro As Date
    Dim ImporteGiro As Double
    Dim ImporteAcumulado As Double
    Dim objCobroPago As CobroPago
    Dim objProveedor As Proveedor
    Dim NetoFactura As Double
     
    Set objFormaPago = New FormaDePago
    objFormaPago.Load mudtprops.FormaPagoID
  
    ' Eliminamos todos los giros que pudiera haber anteriormente.
    mobjCobrosPagos.Delete
    
    FechaGiro = DateAdd("m", objFormaPago.MesesPrimerGiro, mudtprops.Fecha)
    
    ' Si el proveedor es extranjero, hay que generar el vencimiento sin tener en cuenta el IVA ni el RE
    Set objProveedor = New Proveedor
    objProveedor.Load Me.ProveedorID
    If objProveedor.Extranjero Then
        ImporteGiro = (Neto - IVA) / objFormaPago.Giros
        NetoFactura = Neto - IVA
    Else
        ImporteGiro = Neto / objFormaPago.Giros
        NetoFactura = Neto
    End If
    
    ImporteAcumulado = 0
    
    For i = 1 To objFormaPago.Giros
        Set objCobroPago = mobjCobrosPagos.Add
       
        With objCobroPago
            .BeginEdit mMoneda
            .Tipo = "P"
            .Persona = Proveedor
            .FormaPago = FormaPago
            .MedioPago = MedioPago
            .NumeroGiro = i
            
            If i = objFormaPago.Giros Then
                ImporteGiro = NetoFactura - ImporteAcumulado
            End If
          
            ' Se asigna el importe y se devuelve por si las reglas de negocio han
            ' cambiado el valor.
            .Importe = ImporteGiro
            ImporteGiro = .Importe
          
            ImporteAcumulado = ImporteAcumulado + ImporteGiro
          
            ' Se asigna el vencimiento y se devuelve por si las reglas de negocio han
            ' cambiado el valor.
            .Vencimiento = FechaGiro
            FechaGiro = .Vencimiento
            FechaGiro = DateAdd("m", objFormaPago.MesesEntreGiros, FechaGiro)

            .FechaEmision = Fecha
            ' TODO: Ver si asignar el banco o hacerlo al domiciliarlo.
            .Banco = objProveedor.Banco
          
            .TemporadaID = TemporadaID
            .EmpresaID = EmpresaID
          
            .ApplyEdit
          
        End With
       
        Set objCobroPago = Nothing
    Next i
    
    Set objProveedor = Nothing
    
    Set objFormaPago = Nothing
  
End Sub

Private Function GetState() As String
    Dim udtData As FacturaCompraData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As FacturaCompraData
  
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub
