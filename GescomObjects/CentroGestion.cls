VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CentroGestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As CentroGestionProps
Private mudtSave As CentroGestionProps

' Aqui declaro los objetos 'child' Direccion,
' DireccionEntrega,
' cuenta bancaria y dato comercial
Private mobjDireccion As Direccion

' Aqui declaro los ComboBox de las listas de Almacenes,
' representantes y formas de pago.
'Private mobjAlmacenes As TextList

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

    Dim objPersist As CentroGestionPersist
  
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("GCServerMTS.CentroGestionPersist", PERSIST_SERVER)

    If mflgDeleted And Not mflgNew Then
           
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.CentroGestionID, mudtprops.DireccionID
        
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
  
'    Set mobjAlmacenes = New TextList
'    mobjAlmacenes.Load "Almacenes"
'    mudtprops.AlmacenID = mobjAlmacenes.Key(mobjAlmacenes.IndexItem(1))
'
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
    '
    mobjValid.RuleBroken "Nombre", True
    mobjValid.RuleBroken "EmpresaID", True
   
    mudtprops.SedeCentral = False
    mudtprops.Nombre = vbNullString
    mudtprops.EmpresaID = 0
    
End Sub

Private Sub Class_Terminate()

    Set mobjValid = New BrokenRules
 
    Set mobjDireccion = Nothing
'    Set mobjAlmacenes = Nothing
    
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

Public Sub Load(CentroGestionID As Long)

    Dim objPersist As CentroGestionPersist
    
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.CentroGestionPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(CentroGestionID)
    Set objPersist = Nothing

    mobjValid.RuleBroken "Nombre", False
    mobjValid.RuleBroken "EmpresaID", False
      
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

Public Property Get CentroGestionID() As Long

    CentroGestionID = mudtprops.CentroGestionID
  
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

'Public Property Get AlmacenID() As Long
'
'    AlmacenID = mudtprops.AlmacenID
'
'End Property
'
'Public Property Let Almacen(Value As String)
'
'    If Not mflgEditing Then Err.Raise 383
'
'    mudtprops.AlmacenID = mobjAlmacenes.Key(Value)
'
'    mobjValid.RuleBroken "AlmacenID", (mudtprops.AlmacenID = 0)
'
'    mflgDirty = True
'
'End Property
'
'Public Property Get Almacen() As String
'
'    Almacen = mobjAlmacenes.Item(mudtprops.AlmacenID)
'
'End Property
'
'Public Property Get Almacenes() As TextList
'
'    Set Almacenes = mobjAlmacenes
'
'End Property
'
Public Property Let EmpresaID(Value As Long)
  
    If Not mflgEditing Then Err.Raise 383
  
    mudtprops.EmpresaID = Value
    mobjValid.RuleBroken "EmpresaID", (mudtprops.EmpresaID = 0)
    mflgDirty = True

End Property

Public Property Get EmpresaID() As Long

    EmpresaID = mudtprops.EmpresaID
  
End Property

Private Function GetState() As String

    Dim udtData As CentroGestionData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

    Dim udtData As CentroGestionData
  
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub



