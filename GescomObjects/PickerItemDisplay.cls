VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PickerItemDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As PickerItemProps

Public Property Get DocumentoID() As Long

    DocumentoID = mudtprops.DocumentoID

End Property

Public Property Let DocumentoID(Value As Long)

    mudtprops.DocumentoID = Value

End Property

Public Property Get Nombre() As String

    Nombre = Trim(mudtprops.Nombre)

End Property

Public Property Let Nombre(Value As String)

    mudtprops.Nombre = Value

End Property

Public Property Get Numero() As Long

    Numero = mudtprops.Numero

End Property

Public Property Let Numero(Value As Long)

    mudtprops.Numero = Value

End Property

Public Property Get Descripcion() As String

    Descripcion = Trim(mudtprops.Descripcion)

End Property

Public Property Let Descripcion(Value As String)

    mudtprops.Descripcion = Value

End Property

Public Property Get Cantidad() As String

    Cantidad = Format$(mudtprops.Cantidad, "#,###,##0.00")
    
End Property

Public Property Let Cantidad(Value As String)

    mudtprops.Cantidad = Value

End Property

Public Property Get Fecha() As Date

    Fecha = mudtprops.Fecha

End Property

Public Property Let Fecha(Value As Date)

    mudtprops.Fecha = Value

End Property


