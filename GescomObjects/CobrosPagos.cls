VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CobrosPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mMoneda As String

Private mcolCobrosPagos As Collection
Private mflgEditing As Boolean

Private Sub Class_Initialize()

    Set mcolCobrosPagos = New Collection

End Sub

Private Sub Class_Terminate()

    Set mcolCobrosPagos = Nothing

End Sub

Public Function Item(ByVal Index As Variant) As CobroPago
Attribute Item.VB_UserMemId = 0

    Set Item = mcolCobrosPagos(Index)

End Function

Public Function Count() As Long

    Count = mcolCobrosPagos.Count

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mcolCobrosPagos.[_NewEnum]

End Function

Public Sub BeginEdit(Moneda As String)
    Dim objCobroPago As CobroPago
    
    mMoneda = Moneda
    If mflgEditing Then Err.Raise 445
    
    For Each objCobroPago In mcolCobrosPagos
        objCobroPago.ChildBeginEdit Moneda
    Next
    
    mflgEditing = True
    
End Sub

Public Sub CancelEdit()
    Dim objCobroPago As CobroPago
    Dim lngIndex As Long
    
    If Not mflgEditing Then Err.Raise 445
    
    mflgEditing = False
    For lngIndex = mcolCobrosPagos.Count To 1 Step -1
        Set objCobroPago = mcolCobrosPagos.Item(lngIndex)
        With objCobroPago
            If .IsNew Then _
               mcolCobrosPagos.Remove lngIndex
'            .ChildCancelEdit
        End With
        
    Next lngIndex
    
End Sub

Public Sub ApplyEdit()
    
    If Not mflgEditing Then Err.Raise 445
    
    mflgEditing = False
    
End Sub

Public Sub Delete()
    Dim objCobroPago As CobroPago
    
    If Not mflgEditing Then Err.Raise 445
    ' code to delete the object's data goes here
    
    For Each objCobroPago In mcolCobrosPagos
        With objCobroPago
            .BeginEdit mMoneda
            .Delete
            .ApplyEdit
        End With
        
    Next
    
End Sub

Public Sub DeleteRemesa()
    Dim objCobroPago As CobroPago
    
    If Not mflgEditing Then Err.Raise 445
    ' code to delete the object's data goes here
    
    For Each objCobroPago In mcolCobrosPagos
        With objCobroPago
            .BeginEdit mMoneda
            .DeleteRemesa
            .ApplyEdit
        End With
        
    Next
    
End Sub

Public Function Add() As CobroPago
    Dim objCobroPago As CobroPago
    
    If Not mflgEditing Then Err.Raise 445
    
    Set objCobroPago = New CobroPago
    objCobroPago.SetAsChild Me, mMoneda
    objCobroPago.ChildBeginEdit mMoneda
    Set Add = objCobroPago
    Set objCobroPago = Nothing
    
End Function

Public Sub AddCobroPago(Child As CobroPago)

    'Child.SetAsChild Me, mMoneda
    'Child.ChildBeginEdit mMoneda
    mcolCobrosPagos.Add Item:=Child

End Sub

Public Sub Remove(ByVal Index As Variant)

    If Not mflgEditing Then Err.Raise 445
    
    With mcolCobrosPagos(Index)
        .BeginEdit mMoneda
        .Delete
        .ApplyEdit
    End With
    
End Sub

Friend Sub Load(FacturaVentaID As Long, Tipo As String, Moneda As String)
    Dim objPersist As CobrosPagosPersist
    
    mMoneda = Moneda
    Set objPersist = CreateObject("GCServerMTS.CobrosPagosPersist", PERSIST_SERVER)
    SetState objPersist.Fetch(FacturaVentaID, Tipo)
    Set objPersist = Nothing
    
End Sub

'Carga un conjunto de registros de cobros de una remesa.
Friend Sub LoadRemesa(BancoID As Long, FechaDomiciliacion As Date, Moneda As String)
    Dim objPersist As CobrosPagosPersist
    
    mMoneda = Moneda
    Set objPersist = CreateObject("GCServerMTS.CobrosPagosPersist", PERSIST_SERVER)
    SetState objPersist.FetchRemesa(BancoID, FechaDomiciliacion)
    Set objPersist = Nothing
    
End Sub

Public Sub RemoveRemesa(ByVal Index As Variant)

    If Not mflgEditing Then Err.Raise 445
    
    With mcolCobrosPagos(Index)
        .BeginEdit mMoneda
        .DeleteRemesa
        .ApplyEdit
    End With
    
End Sub

Public Property Get Importe() As Double
    Dim objCobroPago As CobroPago
    
    Importe = 0
    For Each objCobroPago In mcolCobrosPagos
        If Not objCobroPago.IsDeleted Then
            Importe = Importe + objCobroPago.Importe
        End If
    Next
    
End Property

' Esta Contabilizado cuando lo esta alguno de los items
Public Property Get Contabilizado() As Boolean
    Dim objCobroPago As CobroPago
    
    Contabilizado = False
    For Each objCobroPago In mcolCobrosPagos
        Contabilizado = Contabilizado Or objCobroPago.Contabilizado
    Next
    
End Property

' Esta Remesado si lo están todos los items
Public Property Get Remesado() As Boolean
    Dim objCobroPago As CobroPago
    
    ' Si existe algún elemento, se mira la coleccion,
    ' si no se indica que no está remesado.
    Remesado = IIf(mcolCobrosPagos.Count > 0, True, False)
    For Each objCobroPago In mcolCobrosPagos
        Remesado = Remesado And objCobroPago.Remesado
    Next
    
End Property

' Si lo está algún item
Public Property Get AlgunoRemesado() As Boolean
    Dim objCobroPago As CobroPago
    
    ' Si existe algún elemento, se mira la coleccion,
    ' si no se indica que no está remesado.
    AlgunoRemesado = False
    For Each objCobroPago In mcolCobrosPagos
        AlgunoRemesado = AlgunoRemesado Or objCobroPago.Remesado
        If AlgunoRemesado Then Exit Property
    Next
    
End Property

Friend Function GetState() As String
    Dim objBuffer As Buffer
    Dim objCobroPago As CobroPago
    
    Set objBuffer = New Buffer
    Set objCobroPago = New CobroPago
    objBuffer.Initialize Len(objCobroPago.GetState), 20
    
    For Each objCobroPago In mcolCobrosPagos
        objBuffer.Add objCobroPago.GetState
    Next
    
    GetState = objBuffer.GetState
    Set objBuffer = Nothing
    
End Function

Friend Sub SetState(Buffer As String)
    Dim lngIndex As Long
    Dim objCobroPago As CobroPago
    Dim objBuffer As Buffer
    
    Set mcolCobrosPagos = Nothing
    Set mcolCobrosPagos = New Collection
    
    Set objBuffer = New Buffer
    objBuffer.SetState Buffer
    
    For lngIndex = 1 To objBuffer.Count
        Set objCobroPago = New CobroPago
        With objCobroPago
            .SetState objBuffer.Item(lngIndex)
            .SetAsChild Me, mMoneda
        End With
        mcolCobrosPagos.Add Item:=objCobroPago
    Next
    
End Sub

Friend Property Get IsDirty() As Boolean
    Dim objItem As CobroPago
    
    IsDirty = False
    
    For Each objItem In mcolCobrosPagos
        If objItem.IsDirty Or objItem.IsNew Then
            IsDirty = True
            Exit For
        End If
    Next
    
End Property

'Graba los datos de una remesa en cada uno de los cobros:
' - Banco
' - Fecha de domiciliacion
Friend Sub SetDatosRemesa(BancoID As Long, FechaDomiciliacion As Date)
    Dim objItem As CobroPago
    
' Marcar de la remesa todos los que no han sido borrados de la misma.
    For Each objItem In mcolCobrosPagos
        If Not objItem.IsDeletedRemesa Then
            objItem.BeginEdit mMoneda
            objItem.SetDatosRemesa BancoID, FechaDomiciliacion
            objItem.ApplyEdit
        End If
    Next
    
End Sub

'Cambia la situacion de una remesa de seleccionada a remesada.
Friend Sub MarcarRemesado()
    Dim objItem As CobroPago
    
    For Each objItem In mcolCobrosPagos
        objItem.BeginEdit mMoneda
        objItem.MarcarRemesado
        objItem.ApplyEdit
    Next
    
End Sub


