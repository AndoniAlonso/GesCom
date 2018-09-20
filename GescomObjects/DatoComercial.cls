VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DatoComercial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Esta clase es de tipo child puro, siempre es un
'child y tendra un unico padre
Event Valid(IsValid As Boolean)

Private mudtprops As DatoComercialProps
Private mudtSave As DatoComercialProps
Private mudtChild As DatoComercialProps

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

    If Not mflgEditing Then Err.Raise 445

    If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
        mflgNew = True
        mflgDeleted = False

    ElseIf mflgDirty Or mflgNew Then
        If Not IsValid Then Err.Raise 445
        
        ' save object state
        LSet mudtSave = mudtprops
        mflgNew = False

    End If
    
    mflgEditing = False
    mflgDirty = False

End Sub

Private Sub Class_Initialize()

    mflgNew = True
    
    Set mobjValid = New BrokenRules
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
    ' En principio no hay ninguna regla rota, es decir el formato es libre
  
    mudtprops.Descuento = 0
    mudtprops.RecargoEquivalencia = 0
    mudtprops.IVA = 0
  
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

Public Sub Load(DatoComercialID As Long)
    Dim objPersist As DatoComercialPersist
  
    If mflgChildEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.DatoComercialPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(DatoComercialID)
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

Public Property Get DatoComercialID() As Long

    DatoComercialID = mudtprops.DatoComercialID

End Property

Public Property Get Descuento() As Double

    Descuento = mudtprops.Descuento

End Property

Public Property Let Descuento(Value As Double)

    If Not mflgChildEditing Then Err.Raise 383
  
    mobjValid.RuleBroken "Descuento", (Value < 0)
  
    mudtprops.Descuento = Value
  
    mflgDirty = True

End Property

Public Property Get RecargoEquivalencia() As Double

    RecargoEquivalencia = mudtprops.RecargoEquivalencia

End Property

Public Property Let RecargoEquivalencia(Value As Double)

    If Not mflgChildEditing Then Err.Raise 383
  
    mobjValid.RuleBroken "RecargoEquivalencia", (Value < 0)
  
    mudtprops.RecargoEquivalencia = Value
  
    mflgDirty = True

End Property

Public Property Get IVA() As Double

    IVA = mudtprops.IVA

End Property

Public Property Let IVA(Value As Double)

    If Not mflgChildEditing Then Err.Raise 383
  
    mobjValid.RuleBroken "IVA", (Value < 0)
  
    mudtprops.IVA = Value
  
    mflgDirty = True

End Property

Public Property Get DatoComercialText() As String
    Dim Texto As String
   
    Texto = "Descuento:" & Format$(mudtprops.Descuento, "00.00") & "%"
    
    Texto = Texto & vbCrLf & _
        "Recargo Equivalencia:" & Format$(mudtprops.RecargoEquivalencia, "00.00") & "%"
    
    Texto = Texto & vbCrLf & _
        "I.V.A.:" & Format$(mudtprops.IVA, "00.00") & "%"
    
    DatoComercialText = Texto
    
End Property

Friend Function GetState() As String
    Dim udtData As DatoComercialData
  
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
    Dim udtData As DatoComercialData

    udtData.Buffer = Buffer
    LSet mudtprops = udtData
    mflgNew = False

End Sub

' Esta funcion "encapsula" en un único string el estado de un objeto (lo serializa)
Public Function GetSuperState() As Byte()
    Dim objPB As PropertyBag
    
    If Not IsValid Then Err.Raise 445
    
    Set objPB = New PropertyBag
    
    With objPB
        .WriteProperty "State", GetState                            'Almacenar el estado del padre
        .WriteProperty "New", mflgNew
        .WriteProperty "Deleted", mflgDeleted
        .WriteProperty "Dirty", mflgDirty
        .WriteProperty "Editing", mflgEditing
        .WriteProperty "ChildEditing", mflgChildEditing
        GetSuperState = .Contents
    End With
    
    Set objPB = Nothing
End Function

' Esta funcion "deserializa" un objeto, a partir de un churro carga el estado.
Public Function SetSuperState(Buffer() As Byte)
    Dim objPB As PropertyBag
    
    Set objPB = New PropertyBag
    With objPB
        ' restaurar los contenidos del propertybag
        .Contents = Buffer
        SetState .ReadProperty("State")
        mflgNew = .ReadProperty("New")
        mflgDeleted = .ReadProperty("Deleted")
        mflgDirty = .ReadProperty("Dirty")
        mflgEditing = .ReadProperty("Editing")
        mflgChildEditing = .ReadProperty("ChildEditing")
    End With
   
    Set objPB = Nothing
End Function



