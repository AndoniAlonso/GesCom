VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CuentaBancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Esta clase es de tipo child puro, siempre es un
'child y tendra un unico padre
Event Valid(IsValid As Boolean)

Private mudtprops As CuentaBancariaProps
Private mudtSave As CuentaBancariaProps
Private mudtChild As CuentaBancariaProps

Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean
Private mflgChildEditing As Boolean

Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Public Sub ChildApplyEdit()

    If Not mflgChildEditing Then Err.Raise 445
  
    mflgChildEditing = False

End Sub

Public Sub ApplyEdit()

'    Dim objPersist As CuentaBancariaPersist
  
    If Not mflgEditing Then Err.Raise 445

'    Set objPersist = CreateObject("GCServerMTS.CuentaBancariaPersist", PERSIST_SERVER)
  
    If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
'        objPersist.DeleteObject mudtProps.CuentaBancariaID
        mflgNew = True
        mflgDeleted = False

    ElseIf mflgDirty Or mflgNew Then
        If Not IsValid Then Err.Raise 445
        
        ' save object to database if appropriate
'       SetState objPersist.Save(GetState)
        
        ' save object state
        LSet mudtSave = mudtprops
        mflgNew = False

    End If
    
'    Set objPersist = Nothing
    mflgEditing = False
    mflgDirty = False

End Sub

Private Sub Class_Initialize()

    mflgNew = True
    
    Set mobjValid = New BrokenRules
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
  
    mudtprops.Entidad = 0
    mudtprops.Sucursal = 0
    mudtprops.Control = 0
    mudtprops.Cuenta = 0
    mudtprops.NombreEntidad = vbNullString
    mudtprops.NombreSucursal = vbNullString
  
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

Public Sub Load(CuentaBancariaID As Long)

    Dim objPersist As CuentaBancariaPersist
  
    If mflgChildEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
    
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.CuentaBancariaPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(CuentaBancariaID)
    Set objPersist = Nothing
  
    mflgNew = False

End Sub

Public Sub Delete()

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

Public Sub ChildBeginEdit()

    If mflgChildEditing Then Err.Raise 445
  
    ' save object state
    LSet mudtChild = mudtprops
    mflgChildEditing = True

End Sub

Public Sub ChildCancelEdit()

    If Not mflgChildEditing Then Err.Raise 445
  
    mflgChildEditing = False
    mflgDeleted = False
  
    ' restore object state
    LSet mudtprops = mudtChild

End Sub

Public Sub BeginEdit()

    If mflgEditing Then Err.Raise 445
  
    ' save object state
    LSet mudtSave = mudtprops
    mflgEditing = True

End Sub

Public Sub CancelEdit()

    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False
    mflgDeleted = False
    
    ' restore object state
    LSet mudtprops = mudtSave

End Sub

Public Property Get CuentaBancariaID() As Long

    CuentaBancariaID = mudtprops.CuentaBancariaID

End Property

Public Property Get Entidad() As String

    Entidad = Trim(mudtprops.Entidad)

End Property

Public Property Let Entidad(Value As String)

    If Not mflgChildEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Entidad) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
    ' Las reglas de validez son:
    '     - longitud nula   o
    '     - todos numericos
    mobjValid.RuleBroken "Entidad", _
        ((Len(Trim(Value)) <> 0) And (Format$(Value, "0000") <> Value))

    mudtprops.Entidad = Trim(Value)
  
    mflgDirty = True

End Property

Public Property Get Sucursal() As String

    Sucursal = Trim(mudtprops.Sucursal)

End Property

Public Property Let Sucursal(Value As String)

    If Not mflgChildEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Sucursal) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
    ' Las reglas de validez son:
    '     - longitud nula   o
    '     - todos numericos
    mobjValid.RuleBroken "Sucursal", _
        ((Len(Trim(Value)) <> 0) And (Format$(Value, "0000") <> Value))
  
    mudtprops.Sucursal = Trim(Value)
  
    mflgDirty = True

End Property

Public Property Get Control() As String

    Control = Trim(mudtprops.Control)

End Property

Public Property Let Control(Value As String)

    If Not mflgChildEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Control) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
    ' Las reglas de validez son:
    '     - longitud nula   o
    '     - todos numericos
    mobjValid.RuleBroken "Control", _
        ((Len(Trim(Value)) <> 0) And (Format$(Value, "00") <> Value))
  
    mudtprops.Control = Trim(Value)
  
    mflgDirty = True

End Property

Public Property Get Cuenta() As String

    Cuenta = Trim(mudtprops.Cuenta)

End Property

Public Property Let Cuenta(Value As String)

    If Not mflgChildEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Cuenta) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
    ' Las reglas de validez son:
    '     - longitud nula   o
    '     - todos numericos
    mobjValid.RuleBroken "Cuenta", _
        ((Len(Trim(Value)) <> 0) And (Format$(Value, "0000000000") <> Value))
  
    mudtprops.Cuenta = Trim(Value)
  
    mflgDirty = True

End Property

Public Property Get NombreEntidad() As String

    NombreEntidad = Trim(mudtprops.NombreEntidad)

End Property

Public Property Let NombreEntidad(Value As String)

    If Not mflgChildEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.NombreEntidad) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
    mobjValid.RuleBroken "NombreEntidad", (Len(Trim(Value)) = 0)
  
    mudtprops.NombreEntidad = Trim(Value)
  
    mflgDirty = True

End Property

Public Property Get NombreSucursal() As String

    NombreSucursal = Trim(mudtprops.NombreSucursal)

End Property

Public Property Let NombreSucursal(Value As String)

    If Not mflgChildEditing Then Err.Raise 383
  
    mudtprops.NombreSucursal = Trim(Value)
  
    mflgDirty = True

End Property


Public Property Get CuentaBancariaText() As String
    
    With mudtprops
        CuentaBancariaText = Format$(.Entidad, "0000") & "." & Format$(.Sucursal, "0000") & _
            "." & Format$(.Control, "00") & "." & Format$(.Cuenta, "0000000000") & _
            vbCrLf & Trim(.NombreEntidad) & vbCrLf & Trim(.NombreSucursal)
    End With
    
End Property

Friend Function GetState() As String

    Dim udtdata As CuentaBancariaData
  
    With mudtprops
        .IsNew = mflgNew
        .IsDeleted = mflgDeleted
        .IsDirty = mflgDirty
    End With

    LSet udtdata = mudtprops
    GetState = udtdata.Buffer
    mflgNew = False

End Function

Friend Sub SetState(Buffer As String)

    Dim udtdata As CuentaBancariaData

    udtdata.Buffer = Buffer
    LSet mudtprops = udtdata
    mflgNew = False

End Sub
