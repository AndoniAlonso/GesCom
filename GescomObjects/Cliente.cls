VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As ClienteProps
Private mudtSave As ClienteProps

' Aqui declaro los objetos 'child' DireccionFiscal,
' DireccionEntrega,
' cuenta bancaria y dato comercial
Private mobjDireccionFiscal As Direccion
Private mobjDireccionEntrega As Direccion
Private mobjCuentaBancaria As CuentaBancaria
Private mobjDatoComercial As DatoComercial
Private mobjDatoComercialB As DatoComercial

' Aqui declaro los ComboBox de las listas de Transportistas,
' representantes y formas de pago.
Private mobjTransportistas As TextList
Private mobjRepresentantes As TextList
Private mobjFormasPago As TextList

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
  
    mobjDireccionFiscal.BeginEdit
    mobjDireccionEntrega.BeginEdit
    mobjCuentaBancaria.BeginEdit
    mobjDatoComercial.BeginEdit
    mobjDatoComercialB.BeginEdit
    mflgEditing = True

End Sub

Public Sub CancelEdit()

    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False
    mflgDeleted = False
    mflgDirty = False
    
    ' restore object state
    LSet mudtprops = mudtSave
  
    mobjDireccionFiscal.CancelEdit
    mobjDireccionEntrega.CancelEdit
    mobjCuentaBancaria.CancelEdit
    mobjDatoComercial.CancelEdit
    mobjDatoComercialB.CancelEdit
  
End Sub

Public Sub ApplyEdit()

    Dim objPersist As ClientePersist
  
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("GCServerMTS.ClientePersist", PERSIST_SERVER)

    If mflgDeleted And Not mflgNew Then
           
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.ClienteID, mudtprops.DireccionFiscalID, _
                                mudtprops.DireccionEntregaID, mudtprops.CuentaBancariaID, _
                                mudtprops.DatoComercialID, mudtprops.DatoComercialBID
        
        ' tratar con los datos de direcciones
        mobjDireccionFiscal.Delete
        mobjDireccionFiscal.ApplyEdit
        mobjDireccionEntrega.Delete
        mobjDireccionEntrega.ApplyEdit
        mobjCuentaBancaria.Delete
        mobjCuentaBancaria.ApplyEdit
        mobjDatoComercial.Delete
        mobjDatoComercial.ApplyEdit
        mobjDatoComercialB.Delete
        mobjDatoComercialB.ApplyEdit
        
        mflgNew = True
        mflgDeleted = False

    ElseIf IsDirty Or mflgNew Then
        If Not IsValid Then Err.Raise 445
        
        mudtprops.DireccionFiscal.Buffer = mobjDireccionFiscal.GetState
        mudtprops.DireccionEntrega.Buffer = mobjDireccionEntrega.GetState
        mudtprops.CuentaBancaria.Buffer = mobjCuentaBancaria.GetState
        mudtprops.DatoComercial.Buffer = mobjDatoComercial.GetState
        mudtprops.DatoComercialB.Buffer = mobjDatoComercialB.GetState
        
        ' save object to database if appropriate
        SetState objPersist.Save(GetState)
        
        mobjDireccionFiscal.SetState (mudtprops.DireccionFiscal.Buffer)
        mobjDireccionFiscal.ApplyEdit
        mudtprops.DireccionFiscalID = mobjDireccionFiscal.DireccionID
        
        mobjDireccionEntrega.SetState (mudtprops.DireccionEntrega.Buffer)
        mobjDireccionEntrega.ApplyEdit
        mudtprops.DireccionEntregaID = mobjDireccionEntrega.DireccionID
        
        mobjCuentaBancaria.SetState (mudtprops.CuentaBancaria.Buffer)
        mobjCuentaBancaria.ApplyEdit
        mudtprops.CuentaBancariaID = mobjCuentaBancaria.CuentaBancariaID
        
        mobjDatoComercial.SetState (mudtprops.DatoComercial.Buffer)
        mobjDatoComercial.ApplyEdit
        mudtprops.DatoComercialID = mobjDatoComercial.DatoComercialID
    
        mobjDatoComercialB.SetState (mudtprops.DatoComercialB.Buffer)
        mobjDatoComercialB.ApplyEdit
        mudtprops.DatoComercialBID = mobjDatoComercialB.DatoComercialID
    
        
        ' save object state
        LSet mudtSave = mudtprops
        mflgNew = False
    Else
        mobjDireccionFiscal.ApplyEdit
        mobjDireccionEntrega.ApplyEdit
        mobjCuentaBancaria.ApplyEdit
        mobjDatoComercial.ApplyEdit
        mobjDatoComercialB.ApplyEdit
    
    End If
    
    Set objPersist = Nothing
    mflgDirty = False
    mflgEditing = False

End Sub

Private Sub Class_Initialize()

    mflgNew = True
    
    Set mobjValid = New BrokenRules
 
    Set mobjDireccionFiscal = New Direccion
    Set mobjDireccionEntrega = New Direccion
    Set mobjCuentaBancaria = New CuentaBancaria
    Set mobjDatoComercial = New DatoComercial
    Set mobjDatoComercialB = New DatoComercial
  
    Set mobjTransportistas = New TextList
    mobjTransportistas.Load "Transportistas"
    mudtprops.TransportistaID = mobjTransportistas.Key(mobjTransportistas.IndexItem(1))
    
    Set mobjRepresentantes = New TextList
    mobjRepresentantes.Load "Representantes"
    mudtprops.RepresentanteID = mobjRepresentantes.Key(mobjRepresentantes.IndexItem(1))
      
    Set mobjFormasPago = New TextList
    mobjFormasPago.Load "FormasDePago"
    mudtprops.FormaPagoID = mobjFormasPago.Key(mobjFormasPago.IndexItem(1))
      
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
    '
    mobjValid.RuleBroken "Nombre", True
    mobjValid.RuleBroken "TransportistaID", True
    mobjValid.RuleBroken "RepresentanteID", True
    mobjValid.RuleBroken "FormaPagoID", True
   
    mudtprops.Titular = vbNullString
    mudtprops.Contacto = vbNullString
    mudtprops.DNINIF = vbNullString
    mudtprops.CuentaContable = "4300000000"
    
    mudtprops.DiaPago1 = 0
    mudtprops.DiaPago2 = 0
    mudtprops.DiaPago3 = 0
    mudtprops.PorcFacturacionAB = 100
    
End Sub

Private Sub Class_Terminate()

    Set mobjValid = New BrokenRules
 
    Set mobjDireccionFiscal = Nothing
    Set mobjDireccionEntrega = Nothing
    Set mobjCuentaBancaria = Nothing
    Set mobjDatoComercial = Nothing
    Set mobjDatoComercialB = Nothing
  
    Set mobjTransportistas = Nothing
    Set mobjRepresentantes = Nothing
    Set mobjFormasPago = Nothing
    
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

Public Sub Load(ClienteID As Long)

    Dim objPersist As ClientePersist
    
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.ClientePersist", PERSIST_SERVER)
    SetState objPersist.Fetch(ClienteID)
    Set objPersist = Nothing

    mobjValid.RuleBroken "Nombre", False
    mobjValid.RuleBroken "TransportistaID", False
    mobjValid.RuleBroken "RepresentanteID", False
    mobjValid.RuleBroken "FormaPagoID", False
    mobjValid.RuleBroken "CuentaContable", False
      
    mobjDireccionFiscal.SetState (mudtprops.DireccionFiscal.Buffer)
    mobjDireccionEntrega.SetState (mudtprops.DireccionEntrega.Buffer)
    mobjCuentaBancaria.SetState (mudtprops.CuentaBancaria.Buffer)
    mobjDatoComercial.SetState (mudtprops.DatoComercial.Buffer)
    mobjDatoComercialB.SetState (mudtprops.DatoComercialB.Buffer)
  
    mflgNew = False

End Sub

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
  
    mflgDeleted = True
  
    mobjDireccionFiscal.Delete
    mobjDireccionEntrega.Delete
    mobjCuentaBancaria.Delete
    mobjDatoComercial.Delete
    mobjDatoComercialB.Delete

End Sub

Public Property Get IsDeleted() As Boolean

    IsDeleted = mflgDeleted

End Property

Public Property Get IsNew() As Boolean

    IsNew = mflgNew

End Property

Public Property Get IsDirty() As Boolean

    IsDirty = mflgDirty Or mobjDireccionFiscal.IsDirty Or _
        mobjDireccionEntrega.IsDirty Or mobjCuentaBancaria.IsDirty Or _
        mobjDatoComercial.IsDirty Or _
        mobjDatoComercialB.IsDirty

End Property

Public Property Get ClienteID() As Long

    ClienteID = mudtprops.ClienteID
  
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

Public Property Let DireccionFiscalID(Value As Long)
  
    If Not mflgEditing Then Err.Raise 383
  
    mudtprops.DireccionFiscalID = Value
    mflgDirty = True

End Property

Public Property Get DireccionFiscalID() As Long

    DireccionFiscalID = mudtprops.DireccionFiscalID
  
End Property

Public Property Get DireccionFiscal() As Direccion

    Set DireccionFiscal = mobjDireccionFiscal
  
End Property

Public Property Let DireccionEntregaID(Value As Long)
  
    If Not mflgEditing Then Err.Raise 383
  
    mudtprops.DireccionEntregaID = Value
    mflgDirty = True

End Property

Public Property Get DireccionEntregaID() As Long

    DireccionEntregaID = mudtprops.DireccionEntregaID
  
End Property

Public Property Get DireccionEntrega() As Direccion

    Set DireccionEntrega = mobjDireccionEntrega
  
End Property

Public Property Get TransportistaID() As Long

    TransportistaID = mudtprops.TransportistaID
  
End Property

Public Property Let Transportista(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    mudtprops.TransportistaID = mobjTransportistas.Key(Value)
  
    mobjValid.RuleBroken "TransportistaID", (mudtprops.TransportistaID = 0)
  
    mflgDirty = True

End Property

Public Property Get Transportista() As String

    Transportista = mobjTransportistas.Item(mudtprops.TransportistaID)

End Property

Public Property Get Transportistas() As TextList

    Set Transportistas = mobjTransportistas

End Property

Public Property Get RepresentanteID() As Long

    RepresentanteID = mudtprops.RepresentanteID
  
End Property

Public Property Let Representante(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    mudtprops.RepresentanteID = mobjRepresentantes.Key(Value)
  
    mobjValid.RuleBroken "RepresentanteID", (mudtprops.RepresentanteID = 0)
  
    mflgDirty = True

End Property

Public Property Get Representante() As String

    Representante = mobjRepresentantes.Item(mudtprops.RepresentanteID)

End Property

Public Property Get Representantes() As TextList

    Set Representantes = mobjRepresentantes

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

Public Property Get CuentaContable() As String

    CuentaContable = Trim(mudtprops.CuentaContable)

End Property

Public Property Let CuentaContable(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.CuentaContable) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
    ' Las reglas de validez son:
    '     - todos numericos
    '     - longitud exacta a 10
    mobjValid.RuleBroken "CuentaContable", _
        ((Len(Trim(Value)) < Len(mudtprops.CuentaContable)) Or (Not IsNumeric(Value)))
  
    mudtprops.CuentaContable = Trim(Value)
  
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

Public Property Get DatoComercialBID() As Long

    DatoComercialBID = mudtprops.DatoComercialBID
  
End Property

Public Property Get DatoComercialB() As DatoComercial
    
    Set DatoComercialB = mobjDatoComercialB
  
End Property

Public Property Let DiaPago1(Value As Integer)

    If Not mflgEditing Then Err.Raise 383
  
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
    mudtprops.DiaPago1 = Value
  
    mobjValid.RuleBroken "DiaPago", ReglasDiasPago

    mflgDirty = True

End Property

Private Function ReglasDiasPago() As Boolean
    ReglasDiasPago = False
    
    If mudtprops.DiaPago1 < 0 Or mudtprops.DiaPago1 > 28 Then
        ReglasDiasPago = True
    End If
    If mudtprops.DiaPago2 < 0 Or mudtprops.DiaPago1 > 28 Then
        ReglasDiasPago = True
    End If
    If mudtprops.DiaPago3 < 0 Or mudtprops.DiaPago1 > 28 Then
        ReglasDiasPago = True
    End If
    If mudtprops.DiaPago1 = 0 And mudtprops.DiaPago2 <> 0 Then
        ReglasDiasPago = True
    End If
    If mudtprops.DiaPago2 = 0 And mudtprops.DiaPago3 <> 0 Then
        ReglasDiasPago = True
    End If
    If mudtprops.DiaPago2 <> 0 And mudtprops.DiaPago1 <> 0 And mudtprops.DiaPago2 < mudtprops.DiaPago1 Then
        ReglasDiasPago = True
    End If
    If mudtprops.DiaPago3 <> 0 And mudtprops.DiaPago2 <> 0 And mudtprops.DiaPago3 < mudtprops.DiaPago2 Then
        ReglasDiasPago = True
    End If
        
End Function

Public Property Get DiaPago1() As Integer

    DiaPago1 = mudtprops.DiaPago1
  
End Property

Public Property Let DiaPago2(Value As Integer)

    If Not mflgEditing Then Err.Raise 383
  
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
    mudtprops.DiaPago2 = Value
  
    mobjValid.RuleBroken "DiaPago", ReglasDiasPago

    mflgDirty = True

End Property

Public Property Get DiaPago2() As Integer

    DiaPago2 = mudtprops.DiaPago2
  
End Property

Public Property Let DiaPago3(Value As Integer)

    If Not mflgEditing Then Err.Raise 383
  
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
    mudtprops.DiaPago3 = Value
  
    mobjValid.RuleBroken "DiaPago", ReglasDiasPago

    mflgDirty = True

End Property

Public Property Get DiaPago3() As Integer

    DiaPago3 = mudtprops.DiaPago3
  
End Property

Public Property Let PorcFacturacionAB(Value As Integer)

    If Not mflgEditing Then Err.Raise 383
  
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
    
    If Value > 100 Then _
        Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
    mudtprops.PorcFacturacionAB = Value
  
    mflgDirty = True

End Property

Public Property Get PorcFacturacionAB() As Integer

    PorcFacturacionAB = mudtprops.PorcFacturacionAB
  
End Property

Private Function GetState() As String

    Dim udtData As ClienteData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

    Dim udtData As ClienteData
  
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
                            Me.DireccionFiscal.Calle, Me.DireccionFiscal.CodigoPostal, Me.DireccionFiscal.Poblacion, _
                            Me.DireccionFiscal.Provincia, Me.DireccionFiscal.Pais, Me.DNINIF, _
                            Me.DireccionFiscal.Telefono1, Me.DireccionFiscal.Telefono2, Me.DireccionFiscal.Fax, _
                            Me.DireccionFiscal.EMAIL, , Me.Contacto)
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




