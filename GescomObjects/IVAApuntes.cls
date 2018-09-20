VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVAApuntes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolIVAApuntes As Collection
Private mflgEditing As Boolean

Private Sub Class_Initialize()

    Set mcolIVAApuntes = New Collection

End Sub

Private Sub Class_Terminate()

    Set mcolIVAApuntes = Nothing

End Sub

Public Function Item(ByVal Index As Variant) As IVAApunte
Attribute Item.VB_UserMemId = 0

    Set Item = mcolIVAApuntes(Index)

End Function

Public Function Count() As Long

    Count = mcolIVAApuntes.Count

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mcolIVAApuntes.[_NewEnum]

End Function

Public Sub BeginEdit()
    Dim objIVAApunte As IVAApunte
  
    If mflgEditing Then Err.Raise 445
  
    For Each objIVAApunte In mcolIVAApuntes
        objIVAApunte.ChildBeginEdit
    Next

    mflgEditing = True

End Sub

Public Sub CancelEdit()
    Dim objIVAApunte As IVAApunte
    Dim lngIndex As Long
  
    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False
    For lngIndex = mcolIVAApuntes.Count To 1 Step -1
        Set objIVAApunte = mcolIVAApuntes.Item(lngIndex)
        With objIVAApunte
            If .IsNew Then _
                mcolIVAApuntes.Remove lngIndex
            .ChildCancelEdit
        End With

    Next lngIndex

End Sub

Public Sub ApplyEdit() 'ModeloID As Long)
  
    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False

End Sub

Public Sub Delete()
    Dim objIVAApunte As IVAApunte
  
    If Not mflgEditing Then Err.Raise 445
    
    ' code to delete the object's data goes here
    For Each objIVAApunte In mcolIVAApuntes
        With objIVAApunte
            .BeginEdit
            .Delete
            .ApplyEdit
        End With

    Next

End Sub

Public Function Add() As IVAApunte
    Dim objIVAApunte As IVAApunte
  
    If Not mflgEditing Then Err.Raise 445

    Set objIVAApunte = New IVAApunte
    objIVAApunte.SetAsChild Me
    objIVAApunte.ChildBeginEdit
    Set Add = objIVAApunte
    Set objIVAApunte = Nothing

End Function

Friend Sub AddIVAApunte(Child As IVAApunte)

    mcolIVAApuntes.Add Item:=Child

End Sub

Public Sub Remove(ByVal Index As Variant)

    If Not mflgEditing Then Err.Raise 445

    With mcolIVAApuntes(Index)
        .BeginEdit
        .Delete
        .ApplyEdit
    End With

End Sub

Public Property Get AsientoID() As Long
    Dim objIVAApunte As IVAApunte

    AsientoID = 0
    For Each objIVAApunte In mcolIVAApuntes
        If objIVAApunte.AsientoID <> 0 Then
            AsientoID = objIVAApunte.AsientoID
            Exit For
        End If
    Next

End Property


Friend Function GetSuperState() As String
    Dim objPB As PropertyBag
    Dim lngCount As Long
    Dim objIVAApunte As IVAApunte
    
    Set objPB = New PropertyBag
    With objPB
      .WriteProperty "Count", mcolIVAApuntes.Count
      For Each objIVAApunte In mcolIVAApuntes
        lngCount = lngCount + 1
        .WriteProperty "IVAItem" & CStr(lngCount), objIVAApunte.GetSuperState
      Next
      GetSuperState = .Contents
    End With
    Set objPB = Nothing
    
End Function

Friend Sub SetSuperState(Buffer As String)
    Dim objPB As PropertyBag
    Dim arBuffer() As Byte
    Dim lngIndex As Long
    Dim objIVAApunte As IVAApunte
    
    Set objPB = New PropertyBag
    
    Set mcolIVAApuntes = Nothing
    Set mcolIVAApuntes = New Collection
    
    With objPB
      arBuffer = Buffer
      .Contents = arBuffer
      For lngIndex = 1 To .ReadProperty("Count")
        Set objIVAApunte = New IVAApunte
        objIVAApunte.SetSuperState .ReadProperty("IVAItem" & CStr(lngIndex))
        'objIVAApunte.BeginEdit
        mcolIVAApuntes.Add objIVAApunte
        Set objIVAApunte = Nothing
      Next
    End With
    Set objPB = Nothing
End Sub

Friend Property Get IsDirty() As Boolean
    Dim objItem As IVAApunte
    
    IsDirty = False
    
    For Each objItem In mcolIVAApuntes
      If objItem.IsDirty Or objItem.IsNew Then
        IsDirty = True
        Exit For
      End If
    
    Next
    
End Property

