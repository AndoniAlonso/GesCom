VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EstrModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Valid(IsValid As Boolean)

' Aqui declaro el ComboBox de la lista de materiales.
Private mobjMateriales As TextList

Private mudtprops As EstrModeloProps
Private mudtSave As EstrModeloProps
Private mudtChild As EstrModeloProps
Private mobjParent As EstrModelos

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

    Dim objPersist As EstrModeloPersist
  
    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgChild And mflgNew Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("GCServerMTS.EstrModeloPersist", PERSIST_SERVER)
  
    If Not mflgChild Then
        If mflgDeleted And Not mflgNew Then
            ' code to delete the object's data goes here
            objPersist.DeleteObject mudtprops.EstrModeloID
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
            mobjParent.AddEstrModelo Me
            Set mobjParent = Nothing
        End If

    End If
    
    Set objPersist = Nothing
    mflgEditing = False

End Sub

Private Sub Class_Initialize()

    mflgNew = True
    Set mobjValid = New BrokenRules
  
    Set mobjMateriales = New TextList
    mobjMateriales.Load "vNombreMateriales"
    mudtprops.MaterialID = mobjMateriales.Key(mobjMateriales.IndexItem(1))
      
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
    '
    ' mobjValid.RuleBroken "RuleName", True
    mobjValid.RuleBroken "MaterialID", True
    mobjValid.RuleBroken "Cantidad", True
    
    With mudtprops
        .Cantidad = 0
        .PrecioCosteEUR = 0
        .PrecioCostePTA = 0
        .PrecioEUR = 0
        .PrecioPTA = 0
    End With
    
End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing
    Set mobjMateriales = Nothing
    
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

Public Sub Load(EstrModeloID As Long)

    Dim objPersist As EstrModeloPersist
  
    If mflgChild Then Err.Raise 445
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.EstrModeloPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(EstrModeloID)
    Set objPersist = Nothing
    
    mflgNew = False

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

Friend Sub SetAsChild(objEstrModelos As EstrModelos)

    If mflgNew Then Set mobjParent = objEstrModelos
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

Public Property Get EstrModeloID() As Long

    EstrModeloID = mudtprops.EstrModeloID

End Property

Public Property Get MaterialID() As Long

    MaterialID = mudtprops.MaterialID
  
End Property

Public Property Let Material(Value As String)

    Dim objMaterial As Material
    
    If Not mflgEditing Then Err.Raise 383
  
    mobjValid.RuleBroken "MaterialID", (mobjMateriales.Key(Value) = 0)
    
    If mobjMateriales.Key(Value) = 0 Then Exit Property
    
    Set objMaterial = New Material
    
    objMaterial.Load mobjMateriales.Key(Value), "EUR"
    
    PrecioCoste = objMaterial.PrecioCoste
    
    mudtprops.MaterialID = objMaterial.MaterialID
    
    mflgDirty = True
    
End Property

Public Property Get Material() As String

    Material = mobjMateriales.Item(mudtprops.MaterialID)

End Property

Public Property Get Materiales() As TextList

    Set Materiales = mobjMateriales

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

Public Property Let Cantidad(Value As Double)

    If Not mflgEditing Then Err.Raise 383
  
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Valor numerico erróneo"
  
    ' No se permiten cantidades cero
    mobjValid.RuleBroken "Cantidad", (Value <= 0)
  
    mudtprops.Cantidad = Round(Value, 2)
    CalcularPrecio
    mflgDirty = True

End Property

Public Property Get Cantidad() As Double

    Cantidad = mudtprops.Cantidad

End Property

Public Property Let PrecioCoste(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
  
    mobjValid.RuleBroken "PrecioCoste", (Value < 0)
  
    mudtprops.PrecioCosteEUR = Round(Value, 3)
    mudtprops.PrecioCostePTA = EUR2PTA(Value)
    
    CalcularPrecio
    
    mflgDirty = True
    
End Property

Public Property Get PrecioCoste() As Double

    PrecioCoste = mudtprops.PrecioCosteEUR
  
End Property

Public Property Let Precio(Value As Double)

    If Not mflgEditing Then Err.Raise 383
    
    If Value < 0 Then _
        Err.Raise vbObjectError + 1001, "Importe no puede ser negativo"
        
    mobjValid.RuleBroken "Precio", (Value < 0)
    
    mudtprops.PrecioEUR = Round(Value, 2)
    mudtprops.PrecioPTA = EUR2PTA(Value)
    
    mflgDirty = True
    
End Property

Public Property Get Precio() As Double

    Precio = mudtprops.PrecioEUR
    
End Property

Friend Function GetState() As String

    Dim udtData As EstrModeloData
  
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

    Dim udtData As EstrModeloData

    udtData.Buffer = Buffer
    LSet mudtprops = udtData
    mflgNew = False
    mobjValid.RuleBroken "MaterialID", False
    mobjValid.RuleBroken "Cantidad", False

End Sub

Private Sub CalcularPrecio()

    Precio = PrecioCoste * Cantidad
    
End Sub
