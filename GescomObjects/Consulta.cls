VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As ConsultaProps
Private mudtSave As ConsultaProps

Private mobjConsultaItems As ConsultaItems

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
  
    mobjConsultaItems.BeginEdit
  
    mflgEditing = True

End Sub

Public Sub CancelEdit()

    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False
    mflgDeleted = False
    mflgDirty = False
    
    ' restore object state
    LSet mudtprops = mudtSave
    mobjConsultaItems.CancelEdit
  
End Sub

Public Sub SaveConsulta()
    Dim objPersist As ConsultaPersist
    Dim strItems As String
    
    If mflgEditing Then Err.Raise 445
    
    Set objPersist = CreateObject("GCServerMTS.ConsultaPersist", PERSIST_SERVER)

    strItems = mobjConsultaItems.GetState
    SetState objPersist.Save(GetState, strItems)
    mobjConsultaItems.SetState (strItems)

    Set objPersist = Nothing
    
End Sub

Public Sub ApplyEdit()
    Dim objPersist As ConsultaPersist
  
    If Not mflgEditing Then Err.Raise 445

    
    If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
        Set objPersist = CreateObject("GCServerMTS.ConsultaPersist", PERSIST_SERVER)
        objPersist.DeleteObject mudtprops.ConsultaID, mobjConsultaItems.GetState
        Set objPersist = Nothing
        
        mflgNew = True
        mflgDeleted = False

    ElseIf mflgDirty Or mflgNew Then
        If Not IsValid Then Err.Raise 445
        ' save object to database if appropriate
        ' SetState objPersist.Save(GetState)
        ' save object state
        LSet mudtSave = mudtprops
        mflgNew = False

    End If

    mobjConsultaItems.ApplyEdit 'mudtprops.ConsultaID
    mflgDirty = False
    mflgEditing = False

End Sub

Private Sub Class_Initialize()

    mflgNew = True
    Set mobjValid = New BrokenRules
    ' if we know any rules are broken on startup
    ' then add code here to initialize the list
    ' of broken rules
    '
  
    Set mobjConsultaItems = New ConsultaItems
      
End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing
    Set mobjConsultaItems = Nothing
      
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

Public Sub Load(ConsultaID As Long)
    Dim objPersist As ConsultaPersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.ConsultaPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(ConsultaID)
    Set objPersist = Nothing

    mobjConsultaItems.Load ConsultaID
    mobjConsultaItems.ConsultaCampos mudtprops.Objeto
  
    mflgNew = False

End Sub

Public Sub Delete()

    If Not mflgEditing Then Err.Raise 445
  
    mflgDeleted = True
    mflgDirty = True

    mobjConsultaItems.Delete

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

Public Property Get ConsultaID() As Long

    ConsultaID = mudtprops.ConsultaID
  
End Property

Public Property Let Nombre(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Nombre) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
  
    'mobjValid.RuleBroken "Nombre", (Len(trim(Value)) = 0)

    mudtprops.Nombre = Trim(Value)
    mflgDirty = True

End Property

Public Property Get Nombre() As String

    Nombre = Trim(mudtprops.Nombre)

End Property

Public Property Let Objeto(Value As String)

    If Not mflgNew Then Err.Raise 383
    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Objeto) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
  
    mobjValid.RuleBroken "Objeto", (Len(Trim(Value)) <> Len(mudtprops.Objeto))

    mudtprops.Objeto = UCase$(Trim(Value))
    mflgDirty = True

End Property

Public Property Get Objeto() As String

    Objeto = Trim(mudtprops.Objeto)

End Property

Public Property Get ConsultaItems() As ConsultaItems

    Set ConsultaItems = mobjConsultaItems

End Property

Public Sub ConsultaCampos(Value As String)
    
    mudtprops.Objeto = Value
    mobjConsultaItems.ConsultaCampos (Value)
    
End Sub

Public Property Get ClausulaWhere() As String
    
    ClausulaWhere = mobjConsultaItems.ClausulaWhere

End Property

Private Function GetState() As String

    Dim udtData As ConsultaData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

    Dim udtData As ConsultaData
  
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub
