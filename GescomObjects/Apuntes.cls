VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Apuntes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mMoneda As String
    
Private mcolApuntes As Collection
Private mflgEditing As Boolean

Private Sub Class_Initialize()

    Set mcolApuntes = New Collection

End Sub

Private Sub Class_Terminate()

    Set mcolApuntes = Nothing

End Sub

Public Function Item(ByVal Index As Variant) As Apunte
Attribute Item.VB_UserMemId = 0

    Set Item = mcolApuntes(Index)

End Function

Public Function Count() As Long

    Count = mcolApuntes.Count

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mcolApuntes.[_NewEnum]

End Function

Public Sub BeginEdit(Moneda As String)
    Dim objApunte As Apunte
  
    mMoneda = Moneda
    
    If mflgEditing Then Err.Raise 445
  
    For Each objApunte In mcolApuntes
        objApunte.ChildBeginEdit
    Next

    mflgEditing = True

End Sub

Public Sub CancelEdit()
    Dim objApunte As Apunte
    Dim lngIndex As Long
  
    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False
    For lngIndex = mcolApuntes.Count To 1 Step -1
        Set objApunte = mcolApuntes.Item(lngIndex)
        With objApunte
            If .IsNew Then _
                mcolApuntes.Remove lngIndex
            .ChildCancelEdit
        End With

    Next lngIndex

End Sub

Public Sub ApplyEdit() 'ModeloID As Long)
  
    If Not mflgEditing Then Err.Raise 445
  
    mflgEditing = False

End Sub

Public Sub Delete()
    Dim objApunte As Apunte
  
    If Not mflgEditing Then Err.Raise 445
    
    ' code to delete the object's data goes here
    For Each objApunte In mcolApuntes
        With objApunte
            .BeginEdit
            .Delete
            .ApplyEdit
        End With

    Next

End Sub

Public Function Add() As Apunte
    Dim objApunte As Apunte
  
    If Not mflgEditing Then Err.Raise 445

    Set objApunte = New Apunte
    objApunte.SetAsChild Me
    objApunte.ChildBeginEdit
    Set Add = objApunte
    Set objApunte = Nothing

End Function

Friend Sub AddApunte(Child As Apunte)

    mcolApuntes.Add Item:=Child

End Sub

Public Sub Remove(ByVal Index As Variant)

    If Not mflgEditing Then Err.Raise 445

    With mcolApuntes(Index)
        .BeginEdit mMoneda
        .Delete
        .ApplyEdit
    End With

End Sub

Public Property Get TotalDebe() As Double
    Dim objApunte As Apunte

    TotalDebe = 0
    For Each objApunte In mcolApuntes
        If Not objApunte.IsDeleted Then
            TotalDebe = TotalDebe + IIf(objApunte.CodigoTipoImporte = TipoImporteDebe, objApunte.Importe, 0)
        End If
    Next

End Property

Public Property Get TotalHaber() As Double
    Dim objApunte As Apunte

    TotalHaber = 0
    For Each objApunte In mcolApuntes
        If Not objApunte.IsDeleted Then
            TotalHaber = TotalHaber + IIf(objApunte.CodigoTipoImporte = TipoImporteHaber, objApunte.Importe, 0)
        End If
    Next

End Property

Public Property Get Saldo() As Double

    Saldo = Round(Me.TotalHaber - Me.TotalDebe, 2)
    
End Property

Public Property Get AsientoEquilibrado() As Boolean

    AsientoEquilibrado = IIf(Saldo = 0, True, False)

End Property

Public Property Get AsientoID() As Long
    Dim objApunte As Apunte

    AsientoID = 0
    For Each objApunte In mcolApuntes
        If objApunte.AsientoID <> 0 Then
            AsientoID = objApunte.AsientoID
            Exit For
        End If
    Next

End Property


'Friend Function GetState() As String
'    Dim objBuffer As Buffer
'    Dim objApunte As Apunte
'
'    Set objBuffer = New Buffer
'    Set objApunte = New Apunte
'    objBuffer.Initialize Len(objApunte.GetState), 20
'
'    For Each objApunte In mcolApuntes
'        objBuffer.Add objApunte.GetState
'    Next
'
'    GetState = objBuffer.GetState
'    Set objBuffer = Nothing
'
'End Function
'
'Friend Sub SetState(Buffer As String)
'    Dim lngindex As Long
'    Dim objApunte As Apunte
'    Dim objBuffer As Buffer
'
'    Set mcolApuntes = Nothing
'    Set mcolApuntes = New Collection
'
'    Set objBuffer = New Buffer
'    objBuffer.SetState Buffer
'
'    For lngindex = 1 To objBuffer.Count
'        Set objApunte = New Apunte
'        With objApunte
'            .SetState objBuffer.Item(lngindex)
'            .SetAsChild Me, mMoneda
'        End With
'        mcolApuntes.Add Item:=objApunte
'    Next
'
'End Sub
'
Friend Function GetSuperState() As String
    Dim objPB As PropertyBag
    Dim lngCount As Long
    Dim objApunte As Apunte
    
    Set objPB = New PropertyBag
    With objPB
      .WriteProperty "Count", mcolApuntes.Count
      For Each objApunte In mcolApuntes
        lngCount = lngCount + 1
        .WriteProperty "Item" & CStr(lngCount), objApunte.GetSuperState
      Next
      GetSuperState = .Contents
    End With
    Set objPB = Nothing
    
End Function

Friend Sub SetSuperState(Buffer As String)
    Dim objPB As PropertyBag
    Dim arBuffer() As Byte
    Dim lngIndex As Long
    Dim objApunte As Apunte
    
    Set objPB = New PropertyBag
    
    Set mcolApuntes = Nothing
    Set mcolApuntes = New Collection
    
    With objPB
      arBuffer = Buffer
      .Contents = arBuffer
      For lngIndex = 1 To .ReadProperty("Count")
        Set objApunte = New Apunte
        objApunte.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        'objApunte.BeginEdit "EUR"
        mcolApuntes.Add objApunte
        Set objApunte = Nothing
      Next
    End With
    Set objPB = Nothing
End Sub

Friend Property Get IsDirty() As Boolean
    Dim objItem As Apunte
    
    IsDirty = False
    
    For Each objItem In mcolApuntes
      If objItem.IsDirty Or objItem.IsNew Then
        IsDirty = True
        Exit For
      End If
    
    Next
    
End Property

