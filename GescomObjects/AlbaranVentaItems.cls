VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AlbaranVentaItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolAlbaranVentaItems As Collection
Private mflgEditing As Boolean

Private Sub Class_Initialize()

    Set mcolAlbaranVentaItems = New Collection

End Sub

Private Sub Class_Terminate()

    Set mcolAlbaranVentaItems = Nothing

End Sub

Public Function Item(ByVal Index As Variant) As AlbaranVentaItem
Attribute Item.VB_UserMemId = 0

    Set Item = mcolAlbaranVentaItems(Index)

End Function

Public Function Count() As Long

    Count = mcolAlbaranVentaItems.Count

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mcolAlbaranVentaItems.[_NewEnum]

End Function

Public Sub BeginEdit()
    Dim objAlbaranVentaItem As AlbaranVentaItem
    
    If mflgEditing Then Err.Raise 445
    
    For Each objAlbaranVentaItem In mcolAlbaranVentaItems
      objAlbaranVentaItem.ChildBeginEdit
    Next
    
    mflgEditing = True
    
End Sub

Public Sub CancelEdit()
    Dim objAlbaranVentaItem As AlbaranVentaItem
    Dim lngIndex As Long
    
    If Not mflgEditing Then Err.Raise 445
    
    mflgEditing = False
    For lngIndex = mcolAlbaranVentaItems.Count To 1 Step -1
      Set objAlbaranVentaItem = mcolAlbaranVentaItems.Item(lngIndex)
      With objAlbaranVentaItem
        If .IsNew Then _
           mcolAlbaranVentaItems.Remove lngIndex
        .ChildCancelEdit
      End With
    
    Next lngIndex
    
End Sub

Public Sub ApplyEdit(Comision As Double)

    If Not mflgEditing Then Err.Raise 445
    
    AsignarComision (Comision)
    
    mflgEditing = False
    
End Sub

Public Sub Delete()

    Dim objAlbaranVentaItem As AlbaranVentaItem
    
    If Not mflgEditing Then Err.Raise 445
    ' code to delete the object's data goes here
    
    For Each objAlbaranVentaItem In mcolAlbaranVentaItems
        With objAlbaranVentaItem
            .BeginEdit
            .Delete
            .ApplyEdit
        End With
    Next
    
End Sub

Public Function Add() As AlbaranVentaItem
    Dim objAlbaranVentaItem As AlbaranVentaItem
    
    If Not mflgEditing Then Err.Raise 445
    
    Set objAlbaranVentaItem = New AlbaranVentaItem
    objAlbaranVentaItem.SetAsChild Me
    objAlbaranVentaItem.ChildBeginEdit
    Set Add = objAlbaranVentaItem
    Set objAlbaranVentaItem = Nothing
    
End Function

Friend Sub AddAlbaranVentaItem(Child As AlbaranVentaItem)

    mcolAlbaranVentaItems.Add Item:=Child

End Sub

Public Sub Remove(ByVal Index As Variant)

    If Not mflgEditing Then Err.Raise 445
    
    With mcolAlbaranVentaItems(Index)
        If .IsNew Then      'OJO: quitamos el comportamiento habitual de Business Objects
            mcolAlbaranVentaItems.Remove (Index)
        Else
            .BeginEdit
            .Delete
            .ApplyEdit
        End If
    End With

End Sub

Friend Sub Load(AlbaranVentaID As Long)
    Dim objPersist As AlbaranVentaItemsPersist
    
    Set objPersist = CreateObject("GCServerMTS.AlbaranVentaItemsPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(AlbaranVentaID)
    Set objPersist = Nothing
    
End Sub
Public Property Get Bruto() As Double
    Dim objAlbaranVentaItem As AlbaranVentaItem
    
    Bruto = 0
    For Each objAlbaranVentaItem In mcolAlbaranVentaItems
        If Not objAlbaranVentaItem.IsDeleted Then _
            Bruto = Bruto + objAlbaranVentaItem.Bruto
    Next
    
End Property

Public Property Get Cantidad() As Double
    Dim objAlbaranVentaItem As AlbaranVentaItem
    
    Cantidad = 0
    For Each objAlbaranVentaItem In mcolAlbaranVentaItems
        If Not objAlbaranVentaItem.IsDeleted Then _
            Cantidad = Cantidad + objAlbaranVentaItem.Cantidad
    Next
    
End Property

' Esta facturado cuando lo esta alguna de los items
Public Property Get Facturado() As Boolean
    Dim objAlbaranVentaItem As AlbaranVentaItem
    
    Facturado = False
    For Each objAlbaranVentaItem In mcolAlbaranVentaItems
        Facturado = Facturado Or objAlbaranVentaItem.Facturado
    Next
    
End Property

' El importe pendiente de facturar es la suma de los importes de cada una de las lineas pendientes de facturar
Public Property Get ImportePendienteFacturar() As Double
    Dim objAlbaranVentaItem As AlbaranVentaItem
    
    ImportePendienteFacturar = 0
    For Each objAlbaranVentaItem In mcolAlbaranVentaItems
        If Not objAlbaranVentaItem.Facturado Then
            ImportePendienteFacturar = ImportePendienteFacturar + objAlbaranVentaItem.Bruto
        End If
    Next
    
End Property

' Existen pedidos facturados cuando hay algún albaran que tiene pedidos.
Public Property Get HayPedido() As Boolean
    Dim objAlbaranVentaItem As AlbaranVentaItem
    
    HayPedido = False
    For Each objAlbaranVentaItem In mcolAlbaranVentaItems
        HayPedido = HayPedido Or objAlbaranVentaItem.HayPedido
    Next
    
End Property

Friend Function GetState() As String
    Dim objBuffer As Buffer
    Dim objAlbaranVentaItem As AlbaranVentaItem
    
    Set objBuffer = New Buffer
    Set objAlbaranVentaItem = New AlbaranVentaItem
    objBuffer.Initialize Len(objAlbaranVentaItem.GetState), 20
    
    For Each objAlbaranVentaItem In mcolAlbaranVentaItems
        If Not (objAlbaranVentaItem.IsNew And objAlbaranVentaItem.IsDeleted) Then
            objBuffer.Add objAlbaranVentaItem.GetState
        End If
    Next
    
    GetState = objBuffer.GetState
    Set objBuffer = Nothing
    
End Function

Friend Sub SetState(Buffer As String)
    Dim lngIndex As Long
    Dim objAlbaranVentaItem As AlbaranVentaItem
    Dim objBuffer As Buffer
    
    Set mcolAlbaranVentaItems = Nothing
    Set mcolAlbaranVentaItems = New Collection
    
    Set objBuffer = New Buffer
    objBuffer.SetState Buffer
    
    For lngIndex = 1 To objBuffer.Count
        Set objAlbaranVentaItem = New AlbaranVentaItem
        With objAlbaranVentaItem
            .SetState objBuffer.Item(lngIndex)
            .SetAsChild Me
        End With
        mcolAlbaranVentaItems.Add Item:=objAlbaranVentaItem
    Next
    
End Sub

Friend Function GetSuperState() As Byte()
    Dim objPB As PropertyBag
    Dim objAlbaranVentaItem As AlbaranVentaItem
    Dim lngCount As Long
    
    Set objPB = New PropertyBag
    With objPB
        ' Store total number of child objects
        .WriteProperty "Count", mcolAlbaranVentaItems.Count
        
        For Each objAlbaranVentaItem In mcolAlbaranVentaItems
            lngCount = lngCount + 1
            .WriteProperty "Item" & CStr(lngCount), objAlbaranVentaItem.GetSuperState
        Next
        'Store object's state
        .WriteProperty "Editing", mflgEditing
        GetSuperState = .Contents
    End With
    
    Set objPB = Nothing
End Function

Friend Function SetSuperState(Buffer() As Byte)
    Dim objPB As PropertyBag
    Dim lngIndex As Long
    Dim objAlbaranVentaItem As AlbaranVentaItem
    
    Set objPB = New PropertyBag
    
    Set mcolAlbaranVentaItems = Nothing
    Set mcolAlbaranVentaItems = New Collection
    
    With objPB
        .Contents = Buffer
        
        For lngIndex = 1 To .ReadProperty("Count")
            Set objAlbaranVentaItem = New AlbaranVentaItem
            objAlbaranVentaItem.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
            
            mcolAlbaranVentaItems.Add objAlbaranVentaItem
            Set objAlbaranVentaItem = Nothing
        Next
        
        mflgEditing = .ReadProperty("Editing")
    End With
    
    Set objPB = Nothing
End Function

Public Sub Actualizar()
    Dim objAlbaranVentaItem As AlbaranVentaItem
    
    For Each objAlbaranVentaItem In mcolAlbaranVentaItems
        objAlbaranVentaItem.Actualizar
    Next

End Sub

Public Sub CambiarComision(Comision As Double)
    Dim objAlbaranVentaItem As AlbaranVentaItem
    
    For Each objAlbaranVentaItem In mcolAlbaranVentaItems
        objAlbaranVentaItem.BeginEdit
        objAlbaranVentaItem.Comision = Comision
        objAlbaranVentaItem.ApplyEdit
    Next
End Sub

' Asigna la comision a los items de albaran que son nuevos (que no lo tienen asignado).
Private Sub AsignarComision(Comision As Double)
    Dim objAlbaranVentaItem As AlbaranVentaItem
    
    For Each objAlbaranVentaItem In mcolAlbaranVentaItems
        If objAlbaranVentaItem.IsNew Then
            objAlbaranVentaItem.BeginEdit
            objAlbaranVentaItem.Comision = Comision
            objAlbaranVentaItem.ApplyEdit
        End If
    Next
End Sub

Friend Property Get IsDirty() As Boolean
    Dim objItem As AlbaranVentaItem
    
    IsDirty = False
    
    For Each objItem In mcolAlbaranVentaItems
        If objItem.IsDirty Or objItem.IsNew Then
            IsDirty = True
            Exit For
        End If
    Next
    
End Property

