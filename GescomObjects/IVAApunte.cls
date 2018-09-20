VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVAApunte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Valid(IsValid As Boolean)

Private mudtprops As IVAApunteProps
Private mudtSave As IVAApunteProps
Private mudtChild As IVAApunteProps
Private mobjParent As IVAApuntes

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
    
    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgChild And mflgNew Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445
    
    If Not mflgChild Then
'        If mflgDeleted And Not mflgNew Then
'            ' code to delete the object's data goes here
'            objPersist.DeleteObject mudtprops.IVAApunteID
'            mflgNew = True
'            mflgDeleted = False
'
'        ElseIf mflgDirty Or mflgNew Then
'            If Not IsValid Then Err.Raise 445
'            ' save object to database if appropriate
'            SetState objPersist.Save(GetState)
'            ' save object state
'            LSet mudtSave = mudtprops
'            mflgNew = False
'
'        End If
'
'        mflgDirty = False
'
    ElseIf mflgNew Then
        If Not mobjParent Is Nothing Then
            mobjParent.AddIVAApunte Me
            Set mobjParent = Nothing
        End If

    End If
    
    mflgEditing = False

End Sub

Private Sub Class_Initialize()

    mflgNew = True
    Set mobjValid = New BrokenRules
  
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
    '
    mobjValid.RuleBroken "TipoApunte", True
    mobjValid.RuleBroken "NumeroFactura", True
    mobjValid.RuleBroken "CuentaIVA", True
    mobjValid.RuleBroken "CuentaTotal", True
    mobjValid.RuleBroken "CuentaBase", True
    mobjValid.RuleBroken "DNINIF", True
    
    With mudtprops
        .IVAApunteID = 0
        .AsientoID = 0
        .TipoApunte = vbNullString
        .NumeroFactura = vbNullString
        .CuentaIVA = vbNullString
        .CuentaTotal = vbNullString
        .CuentaBase = vbNullString
        .DNINIF = vbNullString
        .BaseImponible = 0
        .Total = 0
        .IVA = 0
        .CuotaIVA = 0
        .RecargoEquivalencia = 0
        .CuotaRecargo = 0
            
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

Public Sub Delete()

    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
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

Friend Sub SetAsChild(objIVAApuntes As IVAApuntes)

    If mflgNew Then Set mobjParent = objIVAApuntes
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

Public Property Get IVAApunteID() As Long

    IVAApunteID = mudtprops.IVAApunteID

End Property

Public Property Get AsientoID() As Long

    AsientoID = mudtprops.AsientoID

End Property

Public Property Get TipoApunteText() As String

    Select Case mudtprops.TipoApunte
    Case "S"
        TipoApunte = "Soportado"
    Case "R"
        TipoApunte = "Repercutido"
    Case Else
        Err.Raise vbObjectError + 1001, "Error en tipo de Apunte IVA solo puede ser Soportado/Repercutido."
    End Select

End Property

Public Property Get TipoApunte() As String

    TipoApunte = mudtprops.TipoApunte

End Property

Public Property Let TipoApunte(Value As String)

    If Not mflgEditing Then Err.Raise 383
    
    Select Case Value
    Case "S"
        mudtprops.TipoApunte = Value
    Case "R"
        mudtprops.TipoApunte = Value
    Case Else
  ' Solo puede ser Soportado o Repercutido
        mobjValid.RuleBroken "TipoApunte", Not (Value = "S" Or _
                                            Value = "R")
    End Select
    
    mflgDirty = True

End Property

Public Property Get NumeroFactura() As String
  
    NumeroFactura = Trim(mudtprops.NumeroFactura)

End Property

Public Property Let NumeroFactura(Value As String)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Len(Value) > Len(mudtprops.NumeroFactura) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mudtprops.NumeroFactura = Trim(Value)
    mflgDirty = True
    mobjValid.RuleBroken "NumeroFactura", (Len(Trim(Value)) = 0)

End Property

Public Property Get CuentaIVA() As String

    CuentaIVA = Trim(mudtprops.CuentaIVA)

End Property

Public Property Let CuentaIVA(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.CuentaIVA) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
    ' Las reglas de validez son:
    '     - todos numericos
    '     - longitud exacta a 10
    mobjValid.RuleBroken "CuentaIVA", _
        ((Len(Trim(Value)) < Len(mudtprops.CuentaIVA)) Or (Not IsNumeric(Value)))
  
    mudtprops.CuentaIVA = Trim(Value)
  
    mflgDirty = True

End Property

Public Property Get CuentaTotal() As String

    CuentaTotal = Trim(mudtprops.CuentaTotal)

End Property

Public Property Let CuentaTotal(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.CuentaTotal) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
    ' Las reglas de validez son:
    '     - todos numericos
    '     - longitud exacta a 10
    mobjValid.RuleBroken "CuentaTotal", _
        ((Len(Trim(Value)) < Len(mudtprops.CuentaTotal)) Or (Not IsNumeric(Value)))
  
    mudtprops.CuentaTotal = Trim(Value)
  
    mflgDirty = True

End Property

Public Property Get CuentaBase() As String

    CuentaBase = Trim(mudtprops.CuentaBase)

End Property

Public Property Let CuentaBase(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.CuentaBase) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
     
    ' Las reglas de validez son:
    '     - todos numericos
    '     - longitud exacta a 10
    mobjValid.RuleBroken "CuentaBase", _
        ((Len(Trim(Value)) < Len(mudtprops.CuentaBase)) Or (Not IsNumeric(Value)))
  
    mudtprops.CuentaBase = Trim(Value)
  
    mflgDirty = True

End Property

Public Property Get Titular() As String
  
    Titular = Trim(mudtprops.Titular)

End Property

Public Property Let Titular(Value As String)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Len(Value) > Len(mudtprops.Titular) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mudtprops.Titular = Trim(Value)
    mflgDirty = True

End Property

Public Property Get DNINIF() As String
  
    DNINIF = Trim(mudtprops.DNINIF)

End Property

Public Property Let DNINIF(Value As String)
  
    If Not mflgEditing Then Err.Raise 383
    
    If Len(Value) > Len(mudtprops.DNINIF) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    mobjValid.RuleBroken "DNINIF", (Len(Trim(Value)) = 0)
    
    mudtprops.DNINIF = Trim(Value)
    mflgDirty = True

End Property

Public Property Let BaseImponible(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    mudtprops.BaseImponible = Round(Value, 2)
    
    mflgDirty = True

End Property

Public Property Get BaseImponible() As Double

    BaseImponible = mudtprops.BaseImponible
  
End Property

Public Property Let Total(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    mudtprops.Total = Round(Value, 2)
    
    mflgDirty = True

End Property

Public Property Get Total() As Double

    Total = mudtprops.Total
  
End Property

Public Property Let IVA(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    mudtprops.IVA = Round(Value, 2)
    
    mflgDirty = True

End Property

Public Property Get IVA() As Double

    IVA = mudtprops.IVA
  
End Property

Public Property Let CuotaIVA(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    mudtprops.CuotaIVA = Round(Value, 2)
    
    mflgDirty = True

End Property

Public Property Get CuotaIVA() As Double

    CuotaIVA = mudtprops.CuotaIVA
  
End Property

Public Property Let RecargoEquivalencia(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    mudtprops.RecargoEquivalencia = Round(Value, 2)
    
    mflgDirty = True

End Property

Public Property Get RecargoEquivalencia() As Double

    RecargoEquivalencia = mudtprops.RecargoEquivalencia
  
End Property

Public Property Let CuotaRecargo(Value As Double)
  
    If Not mflgEditing Then Err.Raise 383
    
    mudtprops.CuotaRecargo = Round(Value, 2)
    
    mflgDirty = True

End Property

Public Property Get CuotaRecargo() As Double

    CuotaRecargo = mudtprops.CuotaRecargo
  
End Property

Friend Function GetSuperState() As String
    
    GetSuperState = GetState

End Function

Friend Sub SetSuperState(Buffer As String)
    
    SetState Buffer

End Sub

Private Function GetState() As String
    Dim udtData As IVAApunteData
  
    With mudtprops
        .IsNew = mflgNew
        .IsDeleted = mflgDeleted
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer
    mflgNew = False

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As IVAApunteData

    udtData.Buffer = Buffer
    LSet mudtprops = udtData
    mflgNew = False
    mobjValid.RuleBroken "TipoApunte", False
    mobjValid.RuleBroken "NumeroFactura", False
    mobjValid.RuleBroken "CuentaIVA", False
    mobjValid.RuleBroken "CuentaTotal", False
    mobjValid.RuleBroken "CuentaBase", False
    mobjValid.RuleBroken "DNINIF", False
    
End Sub

