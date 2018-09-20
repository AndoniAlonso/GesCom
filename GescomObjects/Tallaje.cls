VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tallaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As TallajeProps
Private mudtSave As TallajeProps

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
  
    mflgEditing = True

End Sub

Public Sub CancelEdit()

    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False
    mflgDeleted = False
    mflgDirty = False
    
    ' restore object state
    LSet mudtprops = mudtSave
  
End Sub

Public Sub ApplyEdit()
    Dim objPersist As TallajePersist
  
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("GCServerMTS.TallajePersist", PERSIST_SERVER)

    If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.TallajeID

        mflgNew = True
        mflgDeleted = False

    ElseIf IsDirty Or mflgNew Then
        If Not IsValid Then Err.Raise 445
        
        ' save object to database if appropriate
        SetState objPersist.Save(GetState)
        
        ' save object state
        LSet mudtSave = mudtprops
        mflgNew = False
    Else
    
    End If
    
    Set objPersist = Nothing
    mflgDirty = False
    mflgEditing = False

End Sub

Private Sub Class_Initialize()

    mflgNew = True
    
    Set mobjValid = New BrokenRules
 
    mobjValid.RuleBroken "Nombre", True
   
    mudtprops.Nombre = vbNullString
    mudtprops.Estandar = False
    mudtprops.DescripcionT36 = vbNullString
    mudtprops.DescripcionT38 = vbNullString
    mudtprops.DescripcionT40 = vbNullString
    mudtprops.DescripcionT42 = vbNullString
    mudtprops.DescripcionT44 = vbNullString
    mudtprops.DescripcionT46 = vbNullString
    mudtprops.DescripcionT48 = vbNullString
    mudtprops.DescripcionT50 = vbNullString
    mudtprops.DescripcionT52 = vbNullString
    mudtprops.DescripcionT54 = vbNullString
    mudtprops.DescripcionT56 = vbNullString
    mudtprops.PermitidoT36 = False
    mudtprops.PermitidoT38 = False
    mudtprops.PermitidoT40 = False
    mudtprops.PermitidoT42 = False
    mudtprops.PermitidoT44 = False
    mudtprops.PermitidoT46 = False
    mudtprops.PermitidoT48 = False
    mudtprops.PermitidoT50 = False
    mudtprops.PermitidoT52 = False
    mudtprops.PermitidoT54 = False
    mudtprops.PermitidoT56 = False
    
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

Public Sub Load(ByVal TallajeID As Integer)
    Dim objPersist As TallajePersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.TallajePersist", PERSIST_SERVER)
    SetState objPersist.Fetch(TallajeID)
    Set objPersist = Nothing

    mobjValid.RuleBroken "Nombre", False

    mflgNew = False

End Sub

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
  
    mflgDeleted = True
  
End Sub

Public Property Get IsDeleted() As Boolean

    IsDeleted = mflgDeleted

End Property

Public Property Get IsNew() As Boolean

    IsNew = mflgNew

End Property

Public Property Get IsDirty() As Boolean

    IsDirty = mflgDirty

End Property


   Public Property Get TallajeID() As Long
        TallajeID = mudtprops.TallajeID
    End Property

    Public Property Get Nombre() As String
        Nombre = mudtprops.Nombre
    End Property
    
    Public Property Let Nombre(ByVal Value As String)
        If Not mflgEditing Then Err.Raise 383
        
        If Len(Value) > Len(Nombre) Then _
            Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
        
        mobjValid.RuleBroken "Nombre", Len(Trim(Value)) = 0

        mudtprops.Nombre = Trim(Value)
        
        mflgDirty = True
    End Property
    
    Public Property Get Estandar() As Boolean
        Estandar = mudtprops.Estandar
    End Property
    
    Public Property Let Estandar(ByVal Value As Boolean)
        If Not mflgEditing Then Err.Raise 383
        
        mudtprops.Estandar = Value
        
        mflgDirty = True
    End Property
    
    Public Property Get DescripcionT36() As String
        DescripcionT36 = mudtprops.DescripcionT36
    End Property
    
    Public Property Let DescripcionT36(ByVal Value As String)
        If Not mflgEditing Then Err.Raise 383
        
        If Len(Value) > Len(DescripcionT36) Then _
            Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
        
        mudtprops.DescripcionT36 = Value
        
        mflgDirty = True
    End Property
    
    Public Property Get DescripcionT38() As String
        DescripcionT38 = mudtprops.DescripcionT38
    End Property
    
    Public Property Let DescripcionT38(ByVal Value As String)
        If Not mflgEditing Then Err.Raise 383
        
        If Len(Value) > Len(DescripcionT38) Then _
            Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
        
        mudtprops.DescripcionT38 = Value
        
        mflgDirty = True
    End Property
    
    Public Property Get DescripcionT40() As String
        DescripcionT40 = mudtprops.DescripcionT40
    End Property
    
    Public Property Let DescripcionT40(ByVal Value As String)
        If Not mflgEditing Then Err.Raise 383
        
        If Len(Value) > Len(DescripcionT40) Then _
            Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
        
        mudtprops.DescripcionT40 = Value
        
        mflgDirty = True
    End Property
    
    Public Property Get DescripcionT42() As String
        DescripcionT42 = mudtprops.DescripcionT42
    End Property
    
    Public Property Let DescripcionT42(ByVal Value As String)
        If Not mflgEditing Then Err.Raise 383
        
        If Len(Value) > Len(DescripcionT42) Then _
            Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
        
        mudtprops.DescripcionT42 = Value
        
        mflgDirty = True
    End Property
    
    Public Property Get DescripcionT44() As String
        DescripcionT44 = mudtprops.DescripcionT44
    End Property
    
    Public Property Let DescripcionT44(ByVal Value As String)
        If Not mflgEditing Then Err.Raise 383
        
        If Len(Value) > Len(DescripcionT44) Then _
            Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
        
        mudtprops.DescripcionT44 = Value
        
        mflgDirty = True
    End Property
    
    Public Property Get DescripcionT46() As String
        DescripcionT46 = mudtprops.DescripcionT46
    End Property
    
    Public Property Let DescripcionT46(ByVal Value As String)
        If Not mflgEditing Then Err.Raise 383
        
        If Len(Value) > Len(DescripcionT46) Then _
            Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
        
        mudtprops.DescripcionT46 = Value
        
        mflgDirty = True
    End Property
    
    Public Property Get DescripcionT48() As String
        DescripcionT48 = mudtprops.DescripcionT48
    End Property
    
    Public Property Let DescripcionT48(ByVal Value As String)
        If Not mflgEditing Then Err.Raise 383
        
        If Len(Value) > Len(DescripcionT48) Then _
            Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
        
        mudtprops.DescripcionT48 = Value
        
        mflgDirty = True
    End Property
    
    Public Property Get DescripcionT50() As String
        DescripcionT50 = mudtprops.DescripcionT50
    End Property
    
    Public Property Let DescripcionT50(ByVal Value As String)
        If Not mflgEditing Then Err.Raise 383
        
        If Len(Value) > Len(DescripcionT50) Then _
            Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
        
        mudtprops.DescripcionT50 = Value
        
        mflgDirty = True
    End Property
    
    Public Property Get DescripcionT52() As String
        DescripcionT52 = mudtprops.DescripcionT52
    End Property
    
    Public Property Let DescripcionT52(ByVal Value As String)
        If Not mflgEditing Then Err.Raise 383
        
        If Len(Value) > Len(DescripcionT52) Then _
            Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
        
        mudtprops.DescripcionT52 = Value
        
        mflgDirty = True
    End Property
    
    Public Property Get DescripcionT54() As String
        DescripcionT54 = mudtprops.DescripcionT54
    End Property
    
    Public Property Let DescripcionT54(ByVal Value As String)
        If Not mflgEditing Then Err.Raise 383
        
        If Len(Value) > Len(DescripcionT54) Then _
            Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
        
        mudtprops.DescripcionT54 = Value
        
        mflgDirty = True
    End Property
    
    Public Property Get DescripcionT56() As String
        DescripcionT56 = mudtprops.DescripcionT56
    End Property
    
    Public Property Let DescripcionT56(ByVal Value As String)
        If Not mflgEditing Then Err.Raise 383
        
        If Len(Value) > Len(DescripcionT56) Then _
            Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
        
        mudtprops.DescripcionT56 = Value
        
        mflgDirty = True
    End Property
    
    Public Property Get PermitidoT36() As Boolean
        PermitidoT36 = mudtprops.PermitidoT36
    End Property
    
    Public Property Let PermitidoT36(ByVal Value As Boolean)
        If Not mflgEditing Then Err.Raise 383
        
        mudtprops.PermitidoT36 = Value
        
        mflgDirty = True
    End Property
    
    Public Property Get PermitidoT38() As Boolean
        PermitidoT38 = mudtprops.PermitidoT38
    End Property
    
    Public Property Let PermitidoT38(ByVal Value As Boolean)
        If Not mflgEditing Then Err.Raise 383
        
        mudtprops.PermitidoT38 = Value
        
        mflgDirty = True
    End Property
    
    Public Property Get PermitidoT40() As Boolean
        PermitidoT40 = mudtprops.PermitidoT40
    End Property
    
    Public Property Let PermitidoT40(ByVal Value As Boolean)
        If Not mflgEditing Then Err.Raise 383
        
        mudtprops.PermitidoT40 = Value
        
        mflgDirty = True
    End Property
    
    Public Property Get PermitidoT42() As Boolean
        PermitidoT42 = mudtprops.PermitidoT42
    End Property
    
    Public Property Let PermitidoT42(ByVal Value As Boolean)
        If Not mflgEditing Then Err.Raise 383
        
        mudtprops.PermitidoT42 = Value
        
        mflgDirty = True
    End Property
    
    Public Property Get PermitidoT44() As Boolean
        PermitidoT44 = mudtprops.PermitidoT44
    End Property
    
    Public Property Let PermitidoT44(ByVal Value As Boolean)
        If Not mflgEditing Then Err.Raise 383
        
        mudtprops.PermitidoT44 = Value
        
        mflgDirty = True
    End Property
    
    Public Property Get PermitidoT46() As Boolean
        PermitidoT46 = mudtprops.PermitidoT46
    End Property
    
    Public Property Let PermitidoT46(ByVal Value As Boolean)
        If Not mflgEditing Then Err.Raise 383
        
        mudtprops.PermitidoT46 = Value
        
        mflgDirty = True
    End Property
    
    Public Property Get PermitidoT48() As Boolean
        PermitidoT48 = mudtprops.PermitidoT48
    End Property
    
    Public Property Let PermitidoT48(ByVal Value As Boolean)
        If Not mflgEditing Then Err.Raise 383
        
        mudtprops.PermitidoT48 = Value
        
        mflgDirty = True
    End Property
    
    Public Property Get PermitidoT50() As Boolean
        PermitidoT50 = mudtprops.PermitidoT50
    End Property
    
    Public Property Let PermitidoT50(ByVal Value As Boolean)
        If Not mflgEditing Then Err.Raise 383
        
        mudtprops.PermitidoT50 = Value
        
        mflgDirty = True
    End Property
    
    Public Property Get PermitidoT52() As Boolean
        PermitidoT52 = mudtprops.PermitidoT52
    End Property
    
    Public Property Let PermitidoT52(ByVal Value As Boolean)
        If Not mflgEditing Then Err.Raise 383
        
        mudtprops.PermitidoT52 = Value
        
        mflgDirty = True
    End Property
    
    Public Property Get PermitidoT54() As Boolean
        PermitidoT54 = mudtprops.PermitidoT54
    End Property
    
    Public Property Let PermitidoT54(ByVal Value As Boolean)
        If Not mflgEditing Then Err.Raise 383
        
        mudtprops.PermitidoT54 = Value
        
        mflgDirty = True
    End Property
    
    Public Property Get PermitidoT56() As Boolean
        PermitidoT56 = mudtprops.PermitidoT56
    End Property
    
    Public Property Let PermitidoT56(ByVal Value As Boolean)
        If Not mflgEditing Then Err.Raise 383
        
        mudtprops.PermitidoT56 = Value
        
        mflgDirty = True
    End Property
    
Private Function GetState() As String
    Dim udtData As TallajeData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As TallajeData
  
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub
