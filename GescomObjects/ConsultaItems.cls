VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConsultaItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolConsultaItems As Collection
Private mflgEditing As Boolean

Private mConsultaCampos As String

Private Sub Class_Initialize()

    Set mcolConsultaItems = New Collection

End Sub

Private Sub Class_Terminate()

    Set mcolConsultaItems = Nothing

End Sub

Public Function Count() As Long

    Count = mcolConsultaItems.Count

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mcolConsultaItems.[_NewEnum]

End Function

Public Sub BeginEdit()
    Dim objConsultaItem As ConsultaItem
  
    If mflgEditing Then Err.Raise 445
  
    For Each objConsultaItem In mcolConsultaItems
        objConsultaItem.ChildBeginEdit
    Next

    mflgEditing = True

End Sub

Public Sub CancelEdit()
    Dim objConsultaItem As ConsultaItem
    Dim lngIndex As Long
  
    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False
    For lngIndex = mcolConsultaItems.Count To 1 Step -1
        Set objConsultaItem = mcolConsultaItems.Item(lngIndex)
        With objConsultaItem
            If .IsDeleted Then _
                mcolConsultaItems.Remove lngIndex
            '.ChildCancelEdit
        End With

    Next lngIndex

End Sub

Public Sub ApplyEdit() 'ConsultaID As Long)

    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False

End Sub

Public Sub Delete()
    Dim objConsultaItem As ConsultaItem
  
    If Not mflgEditing Then Err.Raise 445
    
    ' code to delete the object's data goes here

    For Each objConsultaItem In mcolConsultaItems
        With objConsultaItem
            .BeginEdit
            .Delete
            .ApplyEdit
        End With

    Next

End Sub

Public Function Add() As ConsultaItem
    Dim objConsultaItem As ConsultaItem
  
    If Not mflgEditing Then Err.Raise 445

    Set objConsultaItem = New ConsultaItem
    'objConsultaItem.ConsultaCampos mConsultaCampos
    objConsultaItem.Objeto mConsultaCampos
    objConsultaItem.SetAsChild Me
    objConsultaItem.ChildBeginEdit
    Set Add = objConsultaItem
    Set objConsultaItem = Nothing

End Function

Public Function Item(ByVal Index As Variant) As ConsultaItem
Attribute Item.VB_UserMemId = 0

    Set Item = mcolConsultaItems(Index)

End Function


Friend Sub AddConsultaItem(Child As ConsultaItem)

    mcolConsultaItems.Add Item:=Child

End Sub

Public Sub Remove(ByVal Index As Variant)

    If Not mflgEditing Then Err.Raise 445
 
    With mcolConsultaItems(Index)
        .BeginEdit
        .Delete
        .ApplyEdit
    End With

End Sub

Friend Sub Load(ConsultaID As Long)
    Dim objPersist As ConsultaItemsPersist
  
    Set objPersist = CreateObject("GCServerMTS.ConsultaItemsPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(ConsultaID)
    Set objPersist = Nothing
    
End Sub

Public Property Get ClausulaWhere() As String
    Dim objConsultaItem As ConsultaItem
    Dim strAND As String

    strAND = vbNullString
    ClausulaWhere = vbNullString
    For Each objConsultaItem In mcolConsultaItems
        If Not objConsultaItem.IsDeleted Then
            ClausulaWhere = ClausulaWhere & strAND & objConsultaItem.ClausulaWhere
            strAND = " AND "
        End If
    Next

End Property

'Public Sub ConsultaCampos(Value As String)
'
'    mConsultaCampos = Value
'
'End Sub
'
Public Sub ConsultaCampos(Value As String)
    Dim objConsultaItem As ConsultaItem
  
    mConsultaCampos = Trim(Value)
    For Each objConsultaItem In mcolConsultaItems
        'objConsultaItem.ConsultaCampos Value
        objConsultaItem.Objeto Value
    Next
    
End Sub

Friend Function GetState() As String
    Dim objBuffer As Buffer
    Dim objConsultaItem As ConsultaItem
  
    Set objBuffer = New Buffer
    Set objConsultaItem = New ConsultaItem
    objBuffer.Initialize Len(objConsultaItem.GetState), 20

    For Each objConsultaItem In mcolConsultaItems
        objBuffer.Add objConsultaItem.GetState
    Next

    GetState = objBuffer.GetState
    Set objBuffer = Nothing

End Function

Friend Sub SetState(Buffer As String)
    Dim lngIndex As Long
    Dim objConsultaItem As ConsultaItem
    Dim objBuffer As Buffer

    Set mcolConsultaItems = Nothing
    Set mcolConsultaItems = New Collection

    Set objBuffer = New Buffer
    objBuffer.SetState Buffer

    For lngIndex = 1 To objBuffer.Count
        Set objConsultaItem = New ConsultaItem
        With objConsultaItem
            .SetState objBuffer.Item(lngIndex)
            .SetAsChild Me
        End With
        mcolConsultaItems.Add Item:=objConsultaItem
    Next

End Sub
