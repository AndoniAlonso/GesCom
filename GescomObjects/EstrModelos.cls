VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EstrModelos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolEstrModelos As Collection
Private mflgEditing As Boolean

Private Sub Class_Initialize()

    Set mcolEstrModelos = New Collection

End Sub

Private Sub Class_Terminate()

    Set mcolEstrModelos = Nothing

End Sub

Public Function Item(ByVal Index As Variant) As EstrModelo
Attribute Item.VB_UserMemId = 0

    Set Item = mcolEstrModelos(Index)

End Function

Public Function Count() As Long

    Count = mcolEstrModelos.Count

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mcolEstrModelos.[_NewEnum]

End Function

Public Sub BeginEdit()
    Dim objEstrModelo As EstrModelo
  
    If mflgEditing Then Err.Raise 445
  
    For Each objEstrModelo In mcolEstrModelos
        objEstrModelo.ChildBeginEdit
    Next

    mflgEditing = True

End Sub

Public Sub CancelEdit()
    Dim objEstrModelo As EstrModelo
    Dim lngIndex As Long
  
    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False
    For lngIndex = mcolEstrModelos.Count To 1 Step -1
        Set objEstrModelo = mcolEstrModelos.Item(lngIndex)
        With objEstrModelo
            If .IsNew Then _
                mcolEstrModelos.Remove lngIndex
            .ChildCancelEdit
        End With

    Next lngIndex

End Sub

Public Sub ApplyEdit() 'ModeloID As Long)
  
    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False

End Sub

Public Sub Delete()
    Dim objEstrModelo As EstrModelo
  
    If Not mflgEditing Then Err.Raise 445
    
    ' code to delete the object's data goes here
    For Each objEstrModelo In mcolEstrModelos
        With objEstrModelo
            .BeginEdit
            .Delete
            .ApplyEdit
        End With

    Next

End Sub

Public Function Add() As EstrModelo
    Dim objEstrModelo As EstrModelo
  
    If Not mflgEditing Then Err.Raise 445

    Set objEstrModelo = New EstrModelo
    objEstrModelo.SetAsChild Me
    objEstrModelo.ChildBeginEdit
    Set Add = objEstrModelo
    Set objEstrModelo = Nothing

End Function

Friend Sub AddEstrModelo(Child As EstrModelo)

    mcolEstrModelos.Add Item:=Child

End Sub

Public Sub Remove(ByVal Index As Variant)

    If Not mflgEditing Then Err.Raise 445

    With mcolEstrModelos(Index)
        .BeginEdit
        .Delete
        .ApplyEdit
    End With

End Sub

Friend Sub Load(ModeloID As Long)

    Dim objPersist As EstrModelosPersist
  
    Set objPersist = CreateObject("GCServerMTS.EstrModelosPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(ModeloID)
    Set objPersist = Nothing

End Sub

Public Property Get Precio() As Double
    Dim objEstrModelo As EstrModelo

    Precio = 0
    For Each objEstrModelo In mcolEstrModelos
        Precio = Precio + objEstrModelo.Precio
    Next

End Property

Public Property Get Cantidad() As Double
    Dim objEstrModelo As EstrModelo

    Cantidad = 0
    For Each objEstrModelo In mcolEstrModelos
        If Not objEstrModelo.IsDeleted Then _
            Cantidad = Cantidad + objEstrModelo.Cantidad
    Next

End Property

Friend Function GetState() As String
    Dim objBuffer As Buffer
    Dim objEstrModelo As EstrModelo
  
    Set objBuffer = New Buffer
    Set objEstrModelo = New EstrModelo
    objBuffer.Initialize Len(objEstrModelo.GetState), 20

    For Each objEstrModelo In mcolEstrModelos
        objBuffer.Add objEstrModelo.GetState
    Next

    GetState = objBuffer.GetState
    Set objBuffer = Nothing

End Function

Friend Sub SetState(Buffer As String)
    Dim lngIndex As Long
    Dim objEstrModelo As EstrModelo
    Dim objBuffer As Buffer

    Set mcolEstrModelos = Nothing
    Set mcolEstrModelos = New Collection

    Set objBuffer = New Buffer
    objBuffer.SetState Buffer

    For lngIndex = 1 To objBuffer.Count
        Set objEstrModelo = New EstrModelo
        With objEstrModelo
            .SetState objBuffer.Item(lngIndex)
            .SetAsChild Me
        End With
        mcolEstrModelos.Add Item:=objEstrModelo
    Next

End Sub

Friend Property Get IsDirty() As Boolean
    Dim objItem As EstrModelo
    
    IsDirty = False
    
    For Each objItem In mcolEstrModelos
      If objItem.IsDirty Or objItem.IsNew Then
        IsDirty = True
        Exit For
      End If
    
    Next
    
End Property

