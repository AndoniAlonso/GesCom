VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TraspasoItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolTraspasoItems As Collection
Private mflgEditing As Boolean

Private Sub Class_Initialize()

    Set mcolTraspasoItems = New Collection

End Sub

Private Sub Class_Terminate()

    Set mcolTraspasoItems = Nothing

End Sub

Public Function Item(ByVal Index As Variant) As TraspasoItem
Attribute Item.VB_UserMemId = 0

    Set Item = mcolTraspasoItems(Index)

End Function

Public Function Count() As Long

    Count = mcolTraspasoItems.Count

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mcolTraspasoItems.[_NewEnum]

End Function

Public Sub BeginEdit()
    Dim objTraspasoItem As TraspasoItem
    
    If mflgEditing Then Err.Raise 445
    
    For Each objTraspasoItem In mcolTraspasoItems
        objTraspasoItem.ChildBeginEdit
    Next
    
    mflgEditing = True
    
End Sub

Public Sub CancelEdit()
    Dim objTraspasoItem As TraspasoItem
    Dim lngIndex As Long
    
    If Not mflgEditing Then Err.Raise 445
    
    mflgEditing = False
    For lngIndex = mcolTraspasoItems.Count To 1 Step -1
        Set objTraspasoItem = mcolTraspasoItems.Item(lngIndex)
        With objTraspasoItem
            If .IsNew Then _
                mcolTraspasoItems.Remove lngIndex
            .ChildCancelEdit
        End With
    
    Next lngIndex
    
End Sub

Public Sub ApplyEdit() 'TraspasoID As Long)

    If Not mflgEditing Then Err.Raise 445
    
    mflgEditing = False
    
End Sub

Public Sub Delete()
    Dim objTraspasoItem As TraspasoItem
    
    If Not mflgEditing Then Err.Raise 445
    ' code to delete the object's data goes here
    
    For Each objTraspasoItem In mcolTraspasoItems
        With objTraspasoItem
            .BeginEdit
            .Delete
            .ApplyEdit
        End With
    
    Next

End Sub

Public Function Add() As TraspasoItem
    Dim objTraspasoItem As TraspasoItem
  
    If Not mflgEditing Then Err.Raise 445
    
    Set objTraspasoItem = New TraspasoItem
    objTraspasoItem.SetAsChild Me
    objTraspasoItem.ChildBeginEdit
    Set Add = objTraspasoItem
    Set objTraspasoItem = Nothing

End Function

Friend Sub AddTraspasoItem(Child As TraspasoItem)

    mcolTraspasoItems.Add Item:=Child

End Sub

Public Sub Remove(ByVal Index As Variant)

    If Not mflgEditing Then Err.Raise 445
    
    With mcolTraspasoItems(Index)
        .BeginEdit
        .Delete
        .ApplyEdit
    End With
    
End Sub

Friend Sub Load(TraspasoID As Long)
    Dim objPersist As TraspasoItemsPersist
    
    Set objPersist = CreateObject("GCServerMTS.TraspasoItemsPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(TraspasoID)
    Set objPersist = Nothing
    
End Sub

Public Property Get Cantidad() As Double
    Dim objTraspasoItem As TraspasoItem
    
    Cantidad = 0
    For Each objTraspasoItem In mcolTraspasoItems
        If Not objTraspasoItem.IsDeleted Then _
            Cantidad = Cantidad + objTraspasoItem.Cantidad
    Next
    
End Property

Public Sub Enviar()
    Dim objTraspasoItem As TraspasoItem
    
    For Each objTraspasoItem In mcolTraspasoItems
        objTraspasoItem.Enviar
    Next

End Sub

Public Sub Recepcionar()
    Dim objTraspasoItem As TraspasoItem
    
    For Each objTraspasoItem In mcolTraspasoItems
        objTraspasoItem.Recepcionar
    Next

End Sub

Friend Function GetState() As String
    Dim objBuffer As Buffer
    Dim objTraspasoItem As TraspasoItem
    
    Set objBuffer = New Buffer
    Set objTraspasoItem = New TraspasoItem
    objBuffer.Initialize Len(objTraspasoItem.GetState), 20
    
    For Each objTraspasoItem In mcolTraspasoItems
        objBuffer.Add objTraspasoItem.GetState
    Next
    
    GetState = objBuffer.GetState
    Set objBuffer = Nothing
    
End Function

Friend Sub SetState(Buffer As String)
    Dim lngIndex As Long
    Dim objTraspasoItem As TraspasoItem
    Dim objBuffer As Buffer
    
    Set mcolTraspasoItems = Nothing
    Set mcolTraspasoItems = New Collection
    
    Set objBuffer = New Buffer
    objBuffer.SetState Buffer
    
    For lngIndex = 1 To objBuffer.Count
        Set objTraspasoItem = New TraspasoItem
        With objTraspasoItem
            .SetState objBuffer.Item(lngIndex)
            .SetAsChild Me
        End With
        mcolTraspasoItems.Add Item:=objTraspasoItem
    Next
    
End Sub

Friend Property Get IsDirty() As Boolean
    Dim objItem As TraspasoItem
    
    IsDirty = False
    
    For Each objItem In mcolTraspasoItems
        If objItem.IsDirty Or objItem.IsNew Then
            IsDirty = True
            Exit For
        End If
    
    Next
    
End Property
