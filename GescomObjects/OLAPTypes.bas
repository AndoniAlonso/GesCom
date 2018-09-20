Attribute VB_Name = "OLAPTypes"
Option Explicit

Public Const QRY_PrevEtiqueta As Integer = 1
Public Const QRY_TallajePedido As Integer = 2

' Modulo para la declaracion de tipos de consultas OLAP

' Consulta de prevision de etiquetas sobre las ordenes de corte.
Public Type PrevEtiquetaProps
  CodigoSerie As String * 2
  NombreSerie As String * 50
  CodigoPrenda As String * 2
  NombrePrenda As String * 50
  Cantidad As Double
End Type

Public Type PrevEtiquetaData
  Buffer As String * 108
End Type

Public Type TallajePedidoProps
  CodigoModelo As String * 2
  NombreModelo As String * 50
  TallaMinima As String * 10
  TallaMaxima As String * 10
End Type

Public Type TallajePedidoData
  Buffer As String * 72
End Type

