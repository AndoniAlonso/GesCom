VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Terminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private mudtprops As TerminalProps
Private mudtSave As TerminalProps

' Aqui declaro los ComboBox de la lista de clientes, representantes, transportistas y formas de pago.
Private mobjCentrosGestion As TextList

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
    Dim objPersist As TerminalPersist
  
    If Not mflgEditing Then Err.Raise 445

    Set objPersist = CreateObject("GCServerMTS.TerminalPersist", PERSIST_SERVER)

    If mflgDeleted And Not mflgNew Then
        ' code to delete the object's data goes here
        objPersist.DeleteObject mudtprops.TerminalID
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

    Set objPersist = Nothing
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
    mobjValid.RuleBroken "Nombre", True
    mobjValid.RuleBroken "CentroGestionID", True
  
    mudtprops.Nombre = vbNullString
    mudtprops.CentroGestionID = 0
      
    Set mobjCentrosGestion = New TextList
    mobjCentrosGestion.Load "CentrosGestion"
    mudtprops.CentroGestionID = mobjCentrosGestion.Key(mobjCentrosGestion.IndexItem(1))

End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing
    Set mobjCentrosGestion = Nothing

      
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

Public Sub Load(TerminalID As Long)

    Dim objPersist As TerminalPersist
  
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.TerminalPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(TerminalID)
    Set objPersist = Nothing

    mobjValid.RuleBroken "Nombre", False
    mobjValid.RuleBroken "CentroGestionID", False
  
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

Public Property Get TerminalID() As Long

    TerminalID = mudtprops.TerminalID
  
End Property

Public Property Let Nombre(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If Len(Value) > Len(mudtprops.Nombre) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    mobjValid.RuleBroken "Nombre", (Len(Trim(Value)) = 0)

    mudtprops.Nombre = Trim(Value)
    mflgDirty = True

End Property

Public Property Get Nombre() As String

    Nombre = Trim(mudtprops.Nombre)

End Property

Public Property Get CentroGestionID() As Long

    CentroGestionID = mudtprops.CentroGestionID
      
End Property

Public Property Let CentroGestion(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    mudtprops.CentroGestionID = mobjCentrosGestion.Key(Value)
  
    mobjValid.RuleBroken "CentroGestionID", (mudtprops.CentroGestionID = 0)
    
    mflgDirty = True

End Property

Public Property Get CentroGestion() As String

    CentroGestion = mobjCentrosGestion.Item(mudtprops.CentroGestionID)

End Property

Public Property Get CentrosGestion() As TextList

    Set CentrosGestion = mobjCentrosGestion

End Property

Public Property Get AlmacenID() As Long

    AlmacenID = mudtprops.AlmacenID
  
End Property

Private Function GetState() As String

    Dim udtData As TerminalData

    With mudtprops
        .IsDeleted = mflgDeleted
        .IsNew = mflgNew
        .IsDirty = mflgDirty
    End With

    LSet udtData = mudtprops
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

    Dim udtData As TerminalData
  
    udtData.Buffer = Buffer
    LSet mudtprops = udtData

End Sub
