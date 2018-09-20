VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Banco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As BancoProps
Private mudtSave As BancoProps

' Aqui declaro los objetos 'child' Direccion,
' y cuenta bancaria
Private mobjDireccion As Direccion
Private mobjCuentaBancaria As CuentaBancaria

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
      
End Sub

Public Sub ApplyEdit()

    Dim objPersist As BancoPersist
  
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("GCServerMTS.BancoPersist", PERSIST_SERVER)

    If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.BancoID, mudtprops.DireccionID, _
                                mudtprops.CuentaBancariaID
        
        ' tratar con los datos de los hijos
        mobjDireccion.Delete
        mobjDireccion.ApplyEdit
        mobjCuentaBancaria.Delete
        mobjCuentaBancaria.ApplyEdit
        
        mflgNew = True
        mflgDeleted = False

    ElseIf IsDirty Or mflgNew Then
        If Not IsValid Then Err.Raise 445
        
        mudtprops.Direccion.Buffer = mobjDireccion.GetState
        mudtprops.CuentaBancaria.Buffer = mobjCuentaBancaria.GetState
        
        ' save object to database if appropriate
        SetState objPersist.Save(GetState)
        
        ' salvar el objeto hijo primero (en este caso)
        mobjDireccion.SetState (mudtprops.Direccion.Buffer)
        mobjDireccion.ApplyEdit
        mudtprops.DireccionID = mobjDireccion.DireccionID
        
        mobjCuentaBancaria.SetState (mudtprops.CuentaBancaria.Buffer)
        mobjCuentaBancaria.ApplyEdit
        mudtprops.CuentaBancariaID = mobjCuentaBancaria.CuentaBancariaID
        
        ' save object state
        LSet mudtSave = mudtprops
        mflgNew = False
    Else
        mobjDireccion.ApplyEdit
        mobjCuentaBancaria.ApplyEdit
        
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
    
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
    '
    mobjValid.RuleBroken "EmpresaID", True
    mobjValid.RuleBroken "CuentaContable", True
    
    mudtprops.Contacto = vbNullString
    mudtprops.SufijoNIF = "000"
    
End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing
 
    Set mobjDireccion = Nothing
    Set mobjCuentaBancaria = Nothing
    
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

Public Sub Load(BancoID As Long)

    Dim objPersist As BancoPersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.BancoPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(BancoID)
    Set objPersist = Nothing

    mobjDireccion.SetState (mudtprops.Direccion.Buffer)
    mobjCuentaBancaria.SetState (mudtprops.CuentaBancaria.Buffer)
    
    mobjValid.RuleBroken "EmpresaID", False
    mobjValid.RuleBroken "CuentaContable", False
    
    mflgNew = False

End Sub

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
  
    mflgDeleted = True
  
    mobjDireccion.Delete
    mobjCuentaBancaria.Delete
    
End Sub

Public Property Get IsDeleted() As Boolean

    IsDeleted = mflgDeleted

End Property

Public Property Get IsNew() As Boolean

    IsNew = mflgNew

End Property

Public Property Get IsDirty() As Boolean

    IsDirty = mflgDirty Or mobjDireccion.IsDirty Or _
        mobjCuentaBancaria.IsDirty

End Property

Public Property Get BancoID() As Long

    BancoID = mudtprops.BancoID
  
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

Public Property Let SufijoNIF(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.SufijoNIF) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mudtprops.SufijoNIF = Trim(Value)
    mflgDirty = True

End Property

Public Property Get SufijoNIF() As String

    SufijoNIF = Trim(mudtprops.SufijoNIF)

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

Public Property Get EmpresaID() As Long

    EmpresaID = mudtprops.EmpresaID
  
End Property

Public Property Let EmpresaID(Value As Long)

    If Not mflgEditing Then Err.Raise 383
  
    mobjValid.RuleBroken "EmpresaID", (Value = 0)
  
    mudtprops.EmpresaID = Value
  
    mflgDirty = True

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

Private Function GetState() As String

    Dim udtData As BancoData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

    Dim udtData As BancoData
  
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub

