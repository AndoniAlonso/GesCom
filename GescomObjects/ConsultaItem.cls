VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConsultaItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Valid(IsValid As Boolean)

' Aqui declaro el ComboBox de la lista de Operadores.
Private mobjOperadores As TextList

Private mobjCampos As TextList

Private mobjConsultaCampo As ConsultaCampo

Private mudtprops As ConsultaItemProps
Private mudtSave As ConsultaItemProps
Private mudtChild As ConsultaItemProps
Private mobjParent As ConsultaItems
Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean
Private mflgChild As Boolean
Private mflgChildEditing As Boolean
Private mflgDosValores As Boolean
Private mstrObjeto As String

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

    '  Dim objPersist As ConsultaItemPersist
  
    If mflgChild Then If Not mflgChildEditing Then Err.Raise 445
    If Not mflgChild And mflgNew Then Err.Raise 445
    If Not mflgEditing Then Err.Raise 445

    '  Set objPersist = CreateObject("GCServerMTS.ConsultaItemPersist", PERSIST_SERVER)
  
    If Not mflgChild Then
        If mflgDeleted And Not mflgNew Then
            ' code to delete the object's data goes here
            ' objPersist.DeleteObject mudtProps.ConsultaItemID
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
    
        mflgDirty = False

    ElseIf mflgNew Then
        If Not mobjParent Is Nothing Then
            mobjParent.AddConsultaItem Me
            Set mobjParent = Nothing
        End If
    End If
    
    mflgEditing = False

End Sub

Private Sub Class_Initialize()

    mflgNew = True
    mflgDosValores = False
    Set mobjValid = New BrokenRules
  
    'Set mobjOperadores = New TextList
    'mobjOperadores.Load "Operadores"
    'mudtProps.OperadorID = mobjOperadores.Key(mobjOperadores.IndexItem(1))
    mudtprops.OperadorID = 0
  
    'Set mobjCampos = New TextList
      
    mobjValid.RuleBroken "Campo", True
    mobjValid.RuleBroken "Operador", True
    mobjValid.RuleBroken "Valor1", True
      
    'Set mobjConsultaCampo = New ConsultaCampo
    Set mobjConsultaCampo = Nothing
    
    mudtprops.Campo = vbNullString
    mudtprops.Alias = vbNullString
    
    mudtprops.Valor1 = vbNullString
    mudtprops.Valor2 = vbNullString
  
End Sub

Private Sub Class_Terminate()

    Set mobjValid = Nothing
  
    Set mobjOperadores = Nothing
    Set mobjCampos = Nothing
    Set mobjConsultaCampo = Nothing
    
End Sub

'Public Sub ConsultaCampos(Value As String)
'
'    'mobjCampos.ConsultaCampos trim(Value)
'    'mudtProps.Alias = mobjCampos.IndexItem(1)
'    'mudtProps.Campo = mobjConsultaCampo.NombreCampo
'End Sub

Public Sub Objeto(Value As String)

    mstrObjeto = Trim(Value)
    
End Sub

Public Property Get IsValid() As Boolean

    IsValid = (mobjValid.Count = 0)

End Property

Public Property Get DosValores() As Boolean

    DosValores = mflgDosValores

End Property

Private Sub mobjValid_BrokenRule()

    RaiseEvent Valid(False)

End Sub

Private Sub mobjValid_NoBrokenRules()

    RaiseEvent Valid(True)

End Sub

Public Sub Load(ConsultaItemID As Long)
    Dim objPersist As ConsultaItemPersist
  
    If mflgChild Then Err.Raise 445
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
  
    ' code to load the object goes here
    Set objPersist = CreateObject("GCServerMTS.ConsultaItemPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(ConsultaItemID)
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

Friend Sub SetAsChild(objConsultaItems As ConsultaItems)

    If mflgNew Then Set mobjParent = objConsultaItems
    mflgChild = True

End Sub

Friend Sub ChildBeginEdit()

    'If mflgChildEditing Then Err.Raise 445
  
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

Public Property Get ConsultaItemID() As Long

    ConsultaItemID = mudtprops.ConsultaItemID

End Property

Public Property Get OperadorID() As Long

    OperadorID = mudtprops.OperadorID
  
End Property

Private Sub CargarOperadores()

    Set mobjOperadores = New TextList
    mobjOperadores.Load "Operadores"

End Sub

Private Sub CargarConsultaCampo()

    Set mobjConsultaCampo = New ConsultaCampo
  
    If mobjCampos Is Nothing Then CargarCampos
    
    If mobjCampos.Key(Trim(mudtprops.Alias)) <> 0 Then
        mobjConsultaCampo.Load mobjCampos.Key(Trim(mudtprops.Alias))
    End If
    
End Sub

Public Property Let Operador(Value As String)

    If Not mflgEditing Then Err.Raise 383
    
    If mobjOperadores Is Nothing Then CargarOperadores
    If mobjConsultaCampo Is Nothing Then CargarConsultaCampo
  
    mudtprops.OperadorID = mobjOperadores.Key(Value)
  
    mobjValid.RuleBroken "Operador", (mudtprops.OperadorID = 0)
  
    ' Cuando seleccionamos el valor "entre", hay que comprobar que el valor no es vacio ni menor
    ' que el valor menor
    ' No tiene sentido utilizar los operadores "Empieza" ni "contiene" en campos de fecha.
    Select Case Value
    Case "Entre"
        mflgDosValores = True
        Valor2 = Valor1
        mobjValid.RuleBroken "Valor2", (Valor2 < Valor1) Or (Len(Valor2) = 0)
    Case "Contiene"
        mobjValid.RuleBroken "Operador", mobjConsultaCampo.IsDate
        mflgDosValores = False
    Case "Empieza"
        mobjValid.RuleBroken "Operador", mobjConsultaCampo.IsDate
        mflgDosValores = False
    Case Else
        mflgDosValores = False
    End Select
  
    mflgDirty = True

End Property

Public Property Get Operador() As String

    If mobjOperadores Is Nothing Then CargarOperadores
  
    Operador = mobjOperadores.Item(mudtprops.OperadorID)
  
End Property

Public Property Get Operadores() As TextList

   If mobjOperadores Is Nothing Then CargarOperadores
   
   Set Operadores = mobjOperadores

End Property

Private Sub CargarCampos()

    Set mobjCampos = New TextList
    mobjCampos.ConsultaCampos mstrObjeto
    If Trim(mudtprops.Alias) = vbNullString Then
        mudtprops.Alias = mobjCampos.IndexItem(1)
        Set mobjConsultaCampo = Nothing
    End If
    'mudtProps.Campo = mobjCampos.IndexItem(1)

End Sub

Public Property Let Campo(Value As String)

    If Not mflgEditing Then Err.Raise 383
  
    If mobjCampos Is Nothing Then CargarCampos
    If mobjConsultaCampo Is Nothing Then CargarConsultaCampo
    
    mobjValid.RuleBroken "Campo", mobjCampos.Key(Value) = 0
    
    'Salimos si queremos asignar algo invalido
    If mobjCampos.Key(Value) = 0 Then Exit Property
  
    mudtprops.Alias = Trim(Value)
  
    mobjConsultaCampo.Load (mobjCampos.Key(Trim(Value)))
    mudtprops.Campo = mobjConsultaCampo.NombreCampo
    
    If mobjConsultaCampo.IsDate Then
       mudtprops.Valor1 = Date
       mudtprops.Valor2 = Date
    End If
  
    mflgDirty = True
End Property

Public Property Get ConsultaCampo() As ConsultaCampo

    If mobjConsultaCampo Is Nothing Then CargarConsultaCampo
    
    Set ConsultaCampo = mobjConsultaCampo

End Property

Public Property Get Campo() As String

    If mobjCampos Is Nothing Then CargarCampos
    
    Campo = Trim(mudtprops.Alias)

End Property

Public Property Get Campos() As TextList

    If mobjCampos Is Nothing Then CargarCampos
    
    Set Campos = mobjCampos

End Property

Public Property Let Valor1(Value As String)
    
    If Not mflgEditing Then Err.Raise 383
  
    If mobjConsultaCampo Is Nothing Then CargarConsultaCampo
    
    Value = Replace(Value, "'", vbNullString)
    
    If Len(Value) > Len(mudtprops.Valor1) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    If mobjConsultaCampo.IsDate And Not IsDate(Value) Then _
        Err.Raise vbObjectError + 1001, "No es un dato de tipo fecha"
    
    mudtprops.Valor1 = Trim(Value)
  
    If mobjConsultaCampo.IsDate Then
       mobjValid.RuleBroken "Valor1", Not IsDate(Value)
    Else
       mobjValid.RuleBroken "Valor1", Len(Trim(Value)) = 0
    End If
   
    mflgDirty = True

End Property

Public Property Get Valor1() As String

    Valor1 = Trim(mudtprops.Valor1)

End Property

Public Property Let Valor2(Value As String)
Dim dtValor1 As Date
Dim dtValor2 As Date
    
    If Not mflgEditing Then Err.Raise 383
  
    If mobjConsultaCampo Is Nothing Then CargarConsultaCampo
    
    Value = Replace(Value, "'", vbNullString)
    
    If Len(Value) > Len(mudtprops.Valor2) Then _
        Err.Raise vbObjectError + 1001, "Valor de cadena muy largo"
    
    If mobjConsultaCampo.IsDate And Not IsDate(Value) Then _
       Err.Raise vbObjectError + 1001, "No es un dato de tipo fecha"
  
    mudtprops.Valor2 = Trim(Value)
    If mobjConsultaCampo.IsDate Then
        dtValor1 = CDate(Valor1)
        dtValor2 = CDate(Valor2)
        mobjValid.RuleBroken "Valor2", (dtValor2 < dtValor1)
    ElseIf mobjConsultaCampo.IsNumeric Then
        mobjValid.RuleBroken "Valor2", (CDbl(Val(Valor2)) < CDbl(Val(Valor1))) Or (Len(Valor2) = 0)
    Else
       mobjValid.RuleBroken "Valor2", (Valor2 < Valor1) Or (Len(Valor2) = 0)
    End If
    
    mflgDirty = True

End Property

Public Property Get Valor2() As String

    Valor2 = Trim(mudtprops.Valor2)

End Property

Public Property Get ClausulaWhere() As String

    Dim strWhere As String
    Dim strValor As String
    Dim strValor2 As String
    Dim strOperador As String

    If mobjConsultaCampo Is Nothing Then CargarConsultaCampo
    
    strValor2 = vbNullString
    Select Case Trim(Operador)
        Case "Empieza"
            strValor = "'" & Valor1 & "%'"
            strOperador = "like"
        Case "Contiene"
            strValor = "'%" & Valor1 & "%'"
            strOperador = "like"
        Case "Entre"
            strOperador = "between"
            If mobjConsultaCampo.TipoCampo = "A" Or mobjConsultaCampo.IsDate Then
                strValor = "'" & Valor1 & "'"
                strValor2 = "'" & Valor2 & "'"
            Else
                strValor = Valor1
                strValor2 = Valor2
            End If
        Case Else
            If mobjConsultaCampo.TipoCampo = "A" Or mobjConsultaCampo.IsDate Then
                strValor = "'" & Valor1 & "'"
            Else
                strValor = Valor1
            End If
            strOperador = Operador
    End Select
  
    strWhere = Trim(mudtprops.Campo) & " " & Trim(strOperador) & " " & strValor
    If Len(strValor2) <> 0 Then
        strWhere = strWhere & " AND " & strValor2
    End If
  
    ClausulaWhere = strWhere
    
End Property

Friend Function GetState() As String

    Dim udtdata As ConsultaItemData
  
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

    Dim udtdata As ConsultaItemData

    udtdata.Buffer = Buffer
    LSet mudtprops = udtdata
    mflgNew = False
    mobjValid.RuleBroken "Campo", False
    mobjValid.RuleBroken "Operador", False
    mobjValid.RuleBroken "Valor1", False
    
    Set mobjCampos = Nothing
    Set mobjConsultaCampo = Nothing
  
End Sub
