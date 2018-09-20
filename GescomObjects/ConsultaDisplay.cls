VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConsultaDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtprops As ConsultaProps

Public Property Get ConsultaID() As Long

    ConsultaID = mudtprops.ConsultaID

End Property

Friend Property Let ConsultaID(Value As Long)

    mudtprops.ConsultaID = Value

End Property

Public Property Get Nombre() As String

    Nombre = Trim(mudtprops.Nombre)

End Property

Friend Property Let Nombre(Value As String)

    mudtprops.Nombre = Value

End Property

Public Property Get Objeto() As String

    Objeto = Trim(mudtprops.Objeto)

End Property

Friend Property Let Objeto(Value As String)

    mudtprops.Objeto = Value

End Property

