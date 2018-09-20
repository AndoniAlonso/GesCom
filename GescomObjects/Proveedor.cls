VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Proveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As ProveedorProps
Private mudtSave As ProveedorProps

' Aqui declaro los objetos 'child' Direccion,
' Direccion,
' cuenta bancaria y dato comercial
Private mobjDireccion As Direccion
Private mobjCuentaBancaria As CuentaBancaria
Private mobjDatoComercial As DatoComercial

' Aqui declaro los ComboBox de las listas de bancos,
' representantes, formas de pago, medios de pago y tipo de proveedor.
Private mobjTransportistas As TextList
Private mobjBancos As TextList
Private mobjFormasDePago As TextList
Private mobjMediosPago As TextList
Private mobjTiposProveedor As TextList

Event Valid(IsValid As Boolean)

Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean

Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Sub BeginEdit()

    If mflgEditing Then Err.Raise 445
  
    ' save object state
    LSet mudtSave = mudtprops
  
    mobjDireccion.BeginEdit
    mobjCuentaBancaria.BeginEdit
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
  
    mobjDireccion.CancelEdit
    mobjCuentaBancaria.CancelEdit
    mobjDatoComercial.CancelEdit
  
End Sub

Public Sub ApplyEdit()
    Dim objPersist As ProveedorPersist
  
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("GCServerMTS.ProveedorPersist", PERSIST_SERVER)

    If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.ProveedorID, mudtprops.DireccionID, _
                                mudtprops.CuentaBancariaID, mudtprops.DatoComercialID

        ' tratar con los datos de los hijos
        mobjDireccion.Delete
        mobjDireccion.ApplyEdit
        mobjCuentaBancaria.Delete
        mobjCuentaBancaria.ApplyEdit
        mobjDatoComercial.Delete
        mobjDatoComercial.ApplyEdit
        
        mflgNew = True
        mflgDeleted = False

    ElseIf IsDirty Or mflgNew Then
        If Not IsValid Then Err.Raise 445
        
        mudtprops.Direccion.Buffer = mobjDireccion.GetState
        mudtprops.CuentaBancaria.Buffer = mobjCuentaBancaria.GetState
        mudtprops.DatoComercial.Buffer = mobjDatoComercial.GetState
        
        ' save object to database if appropriate
        SetState objPersist.Save(GetState)
        
        
        ' tratar con los objetos hijos primero (en este caso)
        mobjDireccion.SetState (mudtprops.Direccion.Buffer)
        mobjDireccion.ApplyEdit
        mudtprops.DireccionID = mobjDireccion.DireccionID
        
        mobjCuentaBancaria.SetState (mudtprops.CuentaBancaria.Buffer)
        mobjCuentaBancaria.ApplyEdit
        mudtprops.CuentaBancariaID = mobjCuentaBancaria.CuentaBancariaID
        
        mobjDatoComercial.SetState (mudtprops.DatoComercial.Buffer)
        mobjDatoComercial.ApplyEdit
        mudtprops.DatoComercialID = mobjDatoComercial.DatoComercialID
     
        ' save object state
        LSet mudtSave = mudtprops
        mflgNew = False
    Else
        mobjDireccion.ApplyEdit
        mobjCuentaBancaria.ApplyEdit
        mobjDatoComercial.ApplyEdit
    
    End If
    
    Set objPersist = Nothing
    mflgDirty = False
    mflgEditing = False

End Sub

Private Sub Class_Initialize()

    mflgNew = True
    
    Set mobjValid = New BrokenRules
 
    Set mobjDireccion = New Direccion
    Set mobjCuentaBancaria = New CuentaBancaria
    Set mobjDatoComercial = New DatoComercial
  
    Set mobjTransportistas = New TextList
    mobjTransportistas.Load "Transportistas"
    mudtprops.TransportistaID = mobjTransportistas.Key(mobjTransportistas.IndexItem(1))
    
    Set mobjBancos = New TextList
    mobjBancos.Load "vBancos"
    mudtprops.BancoID = mobjBancos.Key(mobjBancos.IndexItem(1))
      
    Set mobjFormasDePago = New TextList
    mobjFormasDePago.Load "FormasDePago"
    mudtprops.FormaPagoID = mobjFormasDePago.Key(mobjFormasDePago.IndexItem(1))
      
    Set mobjMediosPago = New TextList
    mobjMediosPago.Load "MediosPago"
    mudtprops.MedioPagoID = mobjMediosPago.Key(mobjMediosPago.IndexItem(1))
  
    Set mobjTiposProveedor = New TextList
    mobjTiposProveedor.LoadEnum "enuTipoProveedor"
    mudtprops.TipoProveedor = mobjTiposProveedor.Key(mobjTiposProveedor.IndexItem(1))
  
    mobjValid.RuleBroken "Nombre", True
    'mobjValid.RuleBroken "TransportistaID", True
    'mobjValid.RuleBroken "BancoID", True
    mobjValid.RuleBroken "FormaPagoID", True
   
    mudtprops.Titular = vbNullString
    mudtprops.Contacto = vbNullString
    mudtprops.DNINIF = vbNullString
    mudtprops.CuentaContable = "4000000000"
    mudtprops.CuentaContrapartida = "6010000001"
    mudtprops.Codigo = vbNullString
    mudtprops.TipoProveedor = "0"
    
End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing
    Set mobjDireccion = Nothing
    Set mobjCuentaBancaria = Nothing
    Set mobjDatoComercial = Nothing
    Set mobjTransportistas = Nothing
    Set mobjBancos = Nothing
    Set mobjFormasDePago = Nothing
    Set mobjMediosPago = Nothing
    Set mobjTiposProveedor = Nothing
    
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

Public Sub Load(ProveedorID As Long)
    Dim objPersist As ProveedorPersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.ProveedorPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(ProveedorID)
    Set objPersist = Nothing

    mobjValid.RuleBroken "Nombre", False
    'mobjValid.RuleBroken "TransportistaID", False
    mobjValid.RuleBroken "BancoID", False
    mobjValid.RuleBroken "FormaPagoID", False
         
    mobjDireccion.SetState (mudtprops.Direccion.Buffer)
    mobjCuentaBancaria.SetState (mudtprops.CuentaBancaria.Buffer)
    mobjDatoComercial.SetState (mudtprops.DatoComercial.Buffer)
  
    mflgNew = False

End Sub

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
  
    mflgDeleted = True
  
    mobjDireccion.Delete
    mobjCuentaBancaria.Delete
    mobjDatoComercial.Delete

End Sub

Public Property Get IsDeleted() As Boolean

    IsDeleted = mflgDeleted

End Property

Public Property Get IsNew() As Boolean

    IsNew = mflgNew

End Property

Public Property Get IsDirty() As Boolean

    IsDirty = mflgDirty Or mobjDireccion.IsDirty Or _
        mobjCuentaBancaria.IsDirty Or _
        mobjDatoComercial.IsDirty

End Property

Public Property Get ProveedorID() As Long

    ProveedorID = mudtprops.ProveedorID
  
End Property

Public Property Let Nombre(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Nombre) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    ' Regla de longitud de Nombre si es distinto de cero-->bien
    mobjValid.RuleBroken "Nombre", (Len(Trim(Value)) = 0)

    mudtprops.Nombre = Trim(Value)
    mflgDirty = True

End Property

Public Property Get Nombre() As String

    Nombre = Trim(mudtprops.Nombre)

End Property

Public Property Let Codigo(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Codigo) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mudtprops.Codigo = Trim(Value)
    mflgDirty = True

End Property

Public Property Get Codigo() As String

    Codigo = Trim(mudtprops.Codigo)

End Property

Public Property Let Titular(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Titular) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mudtprops.Titular = Trim(Value)
    mflgDirty = True

End Property

Public Property Get Titular() As String

    Titular = Trim(mudtprops.Titular)

End Property

Public Property Let Contacto(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Contacto) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mudtprops.Contacto = Trim(Value)
    mflgDirty = True

End Property

Public Property Get Contacto() As String

    Contacto = Trim(mudtprops.Contacto)

End Property

Public Property Let DNINIF(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.DNINIF) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"

    mudtprops.DNINIF = UCase$(Trim(Value))
    mflgDirty = True

End Property

Public Property Get DNINIF() As String

    DNINIF = Trim(mudtprops.DNINIF)

End Property

Public Property Let DireccionID(Value As Long)
  
    If Not mflgEditing Then Err.Raise 383
  
    mudtprops.DireccionID = Value
    mflgDirty = True

End Property

Public Property Get DireccionID() As Long

    DireccionID = mudtprops.DireccionID
  
End Property

Public Property Get Direccion() As Direccion

    Set Direccion = mobjDireccion
  
End Property

Public Property Let TransportistaID(Value As Long)

    If Not mflgEditing Then Err.Raise 383
    
    mudtprops.TransportistaID = Value
    mflgDirty = True
    
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

Public Property Get BancoID() As Long

    BancoID = mudtprops.BancoID
  
End Property

Public Property Let BancoID(Value As Long)

    If Not mflgEditing Then Err.Raise 383
  
    mudtprops.BancoID = mobjBancos.Key(Value)
  
    'mobjValid.RuleBroken "BancoID", (mudtprops.BancoID = 0)
  
    mflgDirty = True

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

Public Property Let CuentaBancariaID(Value As Long)
  
    If Not mflgEditing Then Err.Raise 383
  
    mudtprops.CuentaBancariaID = Value
    mflgDirty = True

End Property

Public Property Get CuentaBancariaID() As Long

    CuentaBancariaID = mudtprops.CuentaBancariaID
  
End Property

Public Property Get CuentaBancaria() As CuentaBancaria

    Set CuentaBancaria = mobjCuentaBancaria
  
End Property

Public Property Get FormaPagoID() As Long

    FormaPagoID = mudtprops.FormaPagoID
  
End Property

Public Property Let FormaDePago(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    mudtprops.FormaPagoID = mobjFormasDePago.Key(Value)
  
    mobjValid.RuleBroken "FormaPagoID", (mudtprops.FormaPagoID = 0)
  
    mflgDirty = True

End Property

Public Property Get FormaDePago() As String

    FormaDePago = mobjFormasDePago.Item(mudtprops.FormaPagoID)

End Property

Public Property Get FormasDePago() As TextList

    Set FormasDePago = mobjFormasDePago

End Property

Public Property Get MedioPagoID() As Long

    MedioPagoID = mudtprops.MedioPagoID
  
End Property

Public Property Let MedioPago(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    mudtprops.MedioPagoID = mobjMediosPago.Key(Value)
  
    mflgDirty = True

End Property

Public Property Get MedioPago() As String

    MedioPago = mobjMediosPago.Item(mudtprops.MedioPagoID)

End Property

Public Property Get MediosPago() As TextList

    Set MediosPago = mobjMediosPago

End Property

Public Property Get TipoProveedorID() As Long

    TipoProveedorID = mudtprops.TipoProveedor
  
End Property

Public Property Let TipoProveedor(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    mudtprops.TipoProveedor = mobjTiposProveedor.Key(Value)
  
    mflgDirty = True

End Property

Public Property Get TipoProveedor() As String

    TipoProveedor = mobjTiposProveedor.Item(mudtprops.TipoProveedor)

End Property

Public Property Get TiposProveedor() As TextList

    Set TiposProveedor = mobjTiposProveedor

End Property

Public Property Get CuentaContable() As String

    CuentaContable = Trim(mudtprops.CuentaContable)

End Property

Public Property Let CuentaContable(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.CuentaContable) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
    ' Las reglas de validez son:
    '     - longitud 10   y
    '     - todos numericos
    mobjValid.RuleBroken "CuentaContable", _
        ((Len(Trim(Value)) < Len(mudtprops.CuentaContable)) Or (Not IsNumeric(Value)))

    mudtprops.CuentaContable = Trim(Value)
  
    mflgDirty = True

End Property

Public Property Get CuentaContrapartida() As String

    CuentaContrapartida = Trim(mudtprops.CuentaContrapartida)

End Property

Public Property Let CuentaContrapartida(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.CuentaContrapartida) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
    ' Las reglas de validez son:
    '     - longitud 10   y
    '     - todos numericos
    mobjValid.RuleBroken "CuentaContrapartida", _
        ((Len(Trim(Value)) < Len(mudtprops.CuentaContrapartida)) Or (Not IsNumeric(Value)))

    mudtprops.CuentaContrapartida = Trim(Value)
  
    mflgDirty = True

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

Public Property Get Extranjero() As Boolean

    Extranjero = (mobjDireccion.Pais <> vbNullString)

End Property

Private Function GetState() As String
    Dim udtData As ProveedorData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As ProveedorData
  
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub

Public Sub ExportarContawin(strServidorContawin As String, lngEmpresaID As Long)
    Dim strEmpresaContawin As String
    Dim objEmpresa As Empresa
    Dim Empresa  As Object 'cwoEmpresa
    Dim cnnEmpresa As Connection
    Dim cnnPrincipal As Connection
    Dim rsEmpresas As Recordset
    
    ' desde aqui
    Dim clsPlan As Object 'ContawinDll.cwoPlanDeCuenta
    Dim bolResultado As Boolean
    
    If mflgEditing Then Err.Raise 383
     
    Set objEmpresa = New Empresa
    objEmpresa.Load lngEmpresaID
    strEmpresaContawin = objEmpresa.EmpresaContawin
    Set objEmpresa = Nothing
    
    Set Empresa = CreateObject("ContawinDLL.cwoEmpresa")  'New cwoEmpresa
    Set rsEmpresas = New Recordset
    Set cnnEmpresa = New Connection
    Set cnnPrincipal = New Connection
    
    
    cnnEmpresa.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & strEmpresaContawin & ";Data Source=" & strServidorContawin & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False"
    cnnEmpresa.Open
    
    cnnPrincipal.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Contawin;Data Source=" & strServidorContawin & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False"
    cnnPrincipal.Open
    
    If Not Empresa.AbrirEmpresas(cnnPrincipal, rsEmpresas) Then _
        Err.Raise 1001 + vbObjectError, "Exportación Contawin", "Error abriendo empresas Contawin."
    
    Set clsPlan = CreateObject("ContaWindll.cwoPlanDeCuenta")
    bolResultado = clsPlan.CuentaPlan(cnnEmpresa, Me.CuentaContable, Me.Nombre)
    If Not bolResultado Then _
        Err.Raise 1001 + vbObjectError, "Exportación Contawin", "Error creando cuenta '" & Me.Nombre & "' en el plan de cuentas" & vbCrLf & clsPlan.InfoDeError
    
    bolResultado = clsPlan.DatoSupletorio(cnnEmpresa, Me.CuentaContable, Me.Titular, Me.Nombre, _
                            Me.Direccion.Calle, Me.Direccion.CodigoPostal, Me.Direccion.Poblacion, _
                            Me.Direccion.Provincia, Me.Direccion.Pais, Me.DNINIF, _
                            Me.Direccion.Telefono1, Me.Direccion.Telefono2, Me.Direccion.Fax, _
                            Me.Direccion.EMAIL, , Me.Contacto)
    If Not bolResultado Then _
        Err.Raise 1001 + vbObjectError, "Exportación Contawin", "Error creando datos supletorios de '" & Me.Nombre & "' en el plan de cuentas. " & vbCrLf & clsPlan.InfoDeError
        
    Set Empresa = Nothing
    
    rsEmpresas.Close
    Set rsEmpresas = Nothing
    
    cnnEmpresa.Close
    Set cnnEmpresa = Nothing
    
    cnnPrincipal.Close
    Set cnnPrincipal = Nothing

End Sub


