VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Representante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As RepresentanteProps
Private mudtSave As RepresentanteProps

' Aqui declaro el objeto 'child' Direccion
Private mobjDireccion As Direccion

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
  
End Sub

Public Sub ApplyEdit()

    Dim objPersist As RepresentantePersist
  
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("GCServerMTS.RepresentantePersist", PERSIST_SERVER)

    If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.RepresentanteID, mudtprops.DireccionID
        
        ' tratar con los datos de direcciones
        mobjDireccion.Delete
        mobjDireccion.ApplyEdit
        
        mflgNew = True
        mflgDeleted = False

    ElseIf IsDirty Or mflgNew Then
        If Not IsValid Then Err.Raise 445
        
        mudtprops.Direccion.Buffer = mobjDireccion.GetState
        
        ' save object to database if appropriate
        SetState objPersist.Save(GetState)
        
        ' salvar el objeto hijo primero (en este caso)
        mobjDireccion.SetState (mudtprops.Direccion.Buffer)
        mobjDireccion.ApplyEdit
        mudtprops.DireccionID = mobjDireccion.DireccionID
        
        ' save object state
        LSet mudtSave = mudtprops
        mflgNew = False
    Else
        mobjDireccion.ApplyEdit

    End If
  
    Set objPersist = Nothing
    mflgDirty = False
    mflgEditing = False

End Sub

Private Sub Class_Initialize()

    mflgNew = True
    Set mobjValid = New BrokenRules
 
    Set mobjDireccion = New Direccion
  
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
    '
    mobjValid.RuleBroken "Nombre", True
    mobjValid.RuleBroken "Comision", True
    mobjValid.RuleBroken "IRPF", True
    mobjValid.RuleBroken "IVA", True
  
End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing
    Set mobjDireccion = Nothing
  
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

Public Sub Load(RepresentanteID As Long)

    Dim objPersist As RepresentantePersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.RepresentantePersist", PERSIST_SERVER)
    SetState objPersist.Fetch(RepresentanteID)
    Set objPersist = Nothing

    mobjValid.RuleBroken "Nombre", False
    mobjValid.RuleBroken "Comision", False
    mobjValid.RuleBroken "IRPF", False
    mobjValid.RuleBroken "IVA", False
    
    mobjDireccion.SetState (mudtprops.Direccion.Buffer)

    mflgNew = False

End Sub

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
  
    mflgDeleted = True
  
    mobjDireccion.Delete

End Sub

Public Property Get IsDeleted() As Boolean

    IsDeleted = mflgDeleted

End Property

Public Property Get IsNew() As Boolean

    IsNew = mflgNew

End Property

Public Property Get IsDirty() As Boolean

    IsDirty = mflgDirty Or mobjDireccion.IsDirty

End Property

Public Property Get RepresentanteID() As Long

    RepresentanteID = mudtprops.RepresentanteID
  
End Property

Public Property Let Nombre(Value As String)

    If Not mflgEditing Then Err.Raise 383
        
    If Len(Value) > Len(mudtprops.Nombre) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    ' Regla de longitud de Entidad si es distinto de 0 --> bien
    mobjValid.RuleBroken "Nombre", (Len(Trim(Value)) = 0)

    mudtprops.Nombre = Trim(Value)
    mflgDirty = True

End Property

Public Property Get Nombre() As String

    Nombre = Trim(mudtprops.Nombre)

End Property

Public Property Let DNINIF(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.DNINIF) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mudtprops.DNINIF = Trim(Value)
    mflgDirty = True

End Property

Public Property Get DNINIF() As String

    DNINIF = Trim(mudtprops.DNINIF)

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

Public Property Let Zona(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Zona) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mudtprops.Zona = Trim(Value)
    mflgDirty = True

End Property

Public Property Get Zona() As String

    Zona = Trim(mudtprops.Zona)

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

Public Property Let Comision(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Not IsNumeric(Value) Then _
        Err.Raise vbObjectError + 1001, "Valor no num�rico"
        
    If Len(Value) > Len(mudtprops.Comision) Then _
        Err.Raise vbObjectError + 1001, "Valor num�rico muy largo"
    
    mobjValid.RuleBroken "Comision", (Value < 0)
    
    mudtprops.Comision = Value
    mflgDirty = True

End Property

Public Property Get Comision() As Double

    Comision = mudtprops.Comision
  
End Property

Public Property Let IRPF(Value As Double)
    
    If Not mflgEditing Then Err.Raise 383
    
    If Not IsNumeric(Value) Then _
        Err.Raise vbObjectError + 1001, "Valor no num�rico"
        
    If Len(Value) > Len(mudtprops.IRPF) Then _
        Err.Raise vbObjectError + 1001, "Valor num�rico muy largo"
        
    mobjValid.RuleBroken "IRPF", (Value < 0)
    
    mudtprops.IRPF = Value
    mflgDirty = True

End Property

Public Property Get IRPF() As Double

    IRPF = mudtprops.IRPF
  
End Property

Public Property Let IVA(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Not IsNumeric(Value) Then _
        Err.Raise vbObjectError + 1001, "Valor no num�rico"
    
    If Len(Value) > Len(mudtprops.IVA) Then _
        Err.Raise vbObjectError + 1001, "Valor num�rico muy largo"
    
    mobjValid.RuleBroken "IVA", (Value < 0)
    
    mudtprops.IVA = Value
    mflgDirty = True
    
End Property

Public Property Get IVA() As Double

    IVA = mudtprops.IVA
    
End Property

Public Property Let CuentaContable(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.CuentaContable) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
    ' Las reglas de validez son:
    '     - longitud nula   o
    '     - todos numericos
    mobjValid.RuleBroken "CuentaContable", _
        Not ((Len(Trim(Value)) = 0) Or _
               (Len(Value) = Len(mudtprops.CuentaContable) _
               And _
               IsNumeric(Value)) _
            )
  
    mudtprops.CuentaContable = Trim(Value)
  
    mflgDirty = True

End Property

Public Property Get CuentaContable() As String

    CuentaContable = Trim(mudtprops.CuentaContable)
    
End Property

Private Function GetState() As String

    Dim udtdata As RepresentanteData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With

    LSet udtdata = mudtprops
    GetState = udtdata.Buffer

End Function

Private Sub SetState(Buffer As String)

    Dim udtdata As RepresentanteData
  
    udtdata.Buffer = Buffer
    LSet mudtprops = udtdata

End Sub
