Attribute VB_Name = "CoreTypes"
Option Explicit

Public Type CoreConsultaDWProps
    ConsultaID As String * 10
    Descripcion As String * 100
    TiempoRefresco As Long
    DatePartRefresco As String * 4
    VistaOrigen As String * 50
    TablaDestino As String * 50
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type CoreConsultaDWData
   Buffer As String * 220
End Type

Public Type PrendaProps
  PrendaID As Long
  Nombre As String * 50
  Codigo As String * 1
  PlanchaPTA As Double
  PlanchaEUR As Double
  TransportePTA As Double
  TransporteEUR As Double
  PerchaPTA As Double
  PerchaEUR As Double
  CartonPTA As Double
  CartonEUR As Double
  EtiquetaPTA As Double
  EtiquetaEUR As Double
  Administracion As Double
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type PrendaData
  Buffer As String * 102
End Type


