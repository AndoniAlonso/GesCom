VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PedidoCompraItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum PedidoCompraItemTipos
  PEDIDOCOMPRAITEM_MATERIAL = 0
  PEDIDOCOMPRAITEM_ARTICULO = 1
End Enum

Public Sub BeginEdit() 'Moneda As String)

End Sub

Public Sub CancelEdit()

End Sub

Public Sub ApplyEdit()

End Sub

Public Sub Delete()

End Sub

Public Property Get IsValid() As Boolean

End Property

Public Property Get IsNew() As Boolean

End Property

Public Property Get IsDirty() As Boolean

End Property

Public Property Get IsDeleted() As Boolean

End Property

Public Sub Load(PedidoCompraItemID As Long, Tipo As PedidoCompraItemTipos)

End Sub

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Public Property Get Tipo() As PedidoCompraItemTipos

End Property

Public Property Get Cantidad() As Double

End Property

Public Sub Actualizar()
    
End Sub

Public Property Get PedidoCompraItemID() As Long

End Property

Public Property Get Situacion() As String
  
End Property

Public Property Let Servido(Value As Double)

End Property

Public Property Get Servido() As Double

End Property

Public Property Get Pendiente() As Double

End Property

Public Property Let Descuento(Value As Double)

End Property

Public Property Get Descuento() As Double

End Property

Public Property Let Comision(Value As Double)

End Property

Public Property Get Comision() As Double

End Property

Public Property Let PrecioCoste(Value As Double)
  
End Property

Public Property Get PrecioCoste() As Double
  
End Property

Public Property Let Bruto(Value As Double)
  
End Property

Public Property Get Bruto() As Double

End Property

