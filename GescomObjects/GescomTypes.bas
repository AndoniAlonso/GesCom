Attribute VB_Name = "GescomTypes"
Option Explicit

Public Const dblAnchuraTelaEstandar = 1.5 ' Anchura de referencia de las telas españolas

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

Public Type DireccionProps
  DireccionID As Long
  Calle  As String * 50
  Poblacion As String * 50
  CodigoPostal As String * 5
  Provincia As String * 50
  Pais As String * 50
  Telefono1 As String * 30
  Telefono2 As String * 30
  Telefono3 As String * 30
  Fax As String * 30
  EMAIL As String * 50
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type DireccionData
  Buffer As String * 380
End Type

Public Type ParametroProps
  ParametroID As Long
  Alfanumero As String * 15
  Clave As String * 25
  Propietario As String * 50
  Usuario As String * 25
  EmpresaID As Long
  TemporadaID As Long
  Moneda As String * 3
  DireccionID As Long
  Direccion As DireccionData
  ServidorContawin As String * 100
  Proyecto As String * 50
  ServidorPersist As String * 50
  Sufijo As String * 10
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type ParametroData
  Buffer As String * 722
End Type

Public Type TextListProps
  Key As String * 30
  Item As String * 255
End Type

Public Type TextListData
  Buffer As String * 285
End Type

Public Type TemporadaProps
  TemporadaID As Long
  Nombre As String * 50
  Codigo As String * 2
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type TemporadaData
  Buffer As String * 58
End Type

Public Type EmpresaProps
  EmpresaID As Long
  Nombre As String * 50
  Codigo As String * 2
  Titular As String * 50
  DNINIF As String * 20
  Actividad As String * 50
  DireccionID As Long
  PedidoCompras As Long
  PedidoVentas As Long
  AlbaranCompras As Long
  AlbaranVentas As Long
  FacturaCompras As Long
  FacturaVentas As Long
  OrdenCorte As Long
  Direccion As DireccionData
  TratamientoIVA As String * 1
  EmpresaContawin As String * 50
  CodigoContawin As Long
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type EmpresaData
  Buffer As String * 628
End Type

Public Type MaterialProps
  MaterialID As Long
  Nombre As String * 50
  Codigo As String * 12
  UnidadMedida As String * 1
  StockActual As Double
  StockPendiente As Double
  StockMinimo As Double
  StockMaximo As Double
  PrecioCostePTA As Double
  PrecioCosteEUR As Double
  PrecioPonderadoPTA As Double
  PrecioPonderadoEUR As Double
  TipoMaterial As String * 1
  AnchuraTela As Double
  FechaAlta As Date
  Composicion1 As String * 20
  PorcComposicion1 As Double
  Composicion2 As String * 20
  PorcComposicion2 As Double
  Composicion3 As String * 20
  PorcComposicion3 As Double
  Composicion4 As String * 20
  PorcComposicion4 As Double
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type MaterialData
  Buffer As String * 208
End Type

' Tipo enumerado para unidades de medida
Public Const UMUnidades = "U"
Public Const UMMetros = "M"
Public Const UMCajas = "C"
Public Const UMKilos = "K"
Public Const UMGramos = "G"
Public Const UMUnidadesTexto = "Unidades"
Public Const UMMetrosTexto = "Metros"
Public Const UMCajasTexto = "Cajas"
Public Const UMKilosTexto = "Kilogramos"
Public Const UMGramosTexto = "Gramos"

Public Type MoviMaterialProps
  MoviMaterialID As Long
  Fecha As Date
  MaterialID As Long
  Tipo     As String * 1
  Concepto As String * 50
  Cantidad As Double
  StockFinal As Double
  PrecioEUR   As Double
  PrecioPTA   As Double
  PrecioCosteEUR As Double
  PrecioCostePTA As Double
  DocumentoID As Long
  TipoDocumento As String * 1
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type MoviMaterialData
  Buffer As String * 90
End Type

Public Type MoviArticuloProps
  MoviArticuloID As Long
  Fecha As Date
  ArticuloColorID As Long
  Tipo     As String * 1
  Concepto As String * 50
  CantidadT36 As Double
  CantidadT38 As Double
  CantidadT40 As Double
  CantidadT42 As Double
  CantidadT44 As Double
  CantidadT46 As Double
  CantidadT48 As Double
  CantidadT50 As Double
  CantidadT52 As Double
  CantidadT54 As Double
  CantidadT56 As Double
  StockFinal As Double
  PrecioEUR   As Double
  PrecioVentaEUR As Double
  PrecioCosteEUR As Double
  AlmacenID As Long
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type MoviArticuloData
  Buffer As String * 126
End Type


' Tipo enumerado para tipos de movimiento de material
Public Const TMMInventario = "I"
Public Const TMMEntrada = "E"
Public Const TMMSalida = "S"
Public Const TMMEntrega = "T"
Public Const TMMReserva = "R"
Public Const TMMInventarioTexto = "Inventario"
Public Const TMMEntradaTexto = "Entrada"
Public Const TMMSalidaTexto = "Salida"
Public Const TMMEntregaTexto = "Entrega"
Public Const TMMReservaTexto = "Reserva"


Public Type SerieProps
  SerieID As Long
  TemporadaID As Long
  Nombre As String * 50
  Codigo As String * 2
  MaterialID As Long
  NombreMaterial As String * 50
  AnchuraTela As Double
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type SerieData
  Buffer As String * 116
End Type

Public Type ModeloProps
  ModeloID As Long
  TemporadaID As Long
  Nombre As String * 50
  Codigo As String * 3
  Beneficio As Double
  CantidadTela As Double
  CorteEUR As Double
  TallerEUR As Double
  BeneficioPVP As Double
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type ModeloData
  Buffer As String * 82
End Type

Public Type EstrModeloProps
  EstrModeloID As Long
  ModeloID As Long
  MaterialID As Long
  Cantidad As Double
  Observaciones As String * 30
  'NombreMaterial As String * 50
  PrecioCostePTA As Double
  PrecioCosteEUR As Double
  PrecioPTA As Double
  PrecioEUR As Double
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type EstrModeloData
  Buffer As String * 60
End Type

Public Type ArticuloProps
  ArticuloID As Long
  Nombre As String * 6
  StockActual As Double
  StockPendiente As Double
  StockMinimo As Double
  StockMaximo As Double
  LoteEconomico As Double
  PrecioCosteEUR As Double
  PrecioVentaEUR As Double
  PrecioVentaPublico As Double
  PrendaID As Long
  ModeloID As Long
  SerieID As Long
  NombrePrenda As String * 20
  NombreModelo As String * 20
  NombreSerie As String * 20
  TemporadaID As Long
  SuReferencia As String * 30
  ProveedorID As Long
  PrecioCompraEUR As Double
  TallajeID As Long
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type ArticuloData
  Buffer As String * 150
End Type

Public Type CuentaBancariaProps
    CuentaBancariaID As Long
    Entidad As String * 4
    Sucursal As String * 4
    Control As String * 2
    Cuenta As String * 10
    NombreEntidad As String * 50
    NombreSucursal As String * 50
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type CuentaBancariaData
   Buffer As String * 126
End Type

Public Type BancoProps
    BancoID As Long
    EmpresaID As Long
    NombreEmpresa As String * 50
    CuentaBancariaID As Long
    NombreEntidad As String * 50
    Cuenta As String * 10
    DireccionID As Long
    Contacto As String * 50
    CuentaContable As String * 10
    SufijoNIF As String * 3
    Direccion As DireccionData
    CuentaBancaria As CuentaBancariaData
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type BancoData
    Buffer As String * 690
End Type

Public Type ArticuloColorProps
  ArticuloColorID As Long
  Nombre As String * 8
  NombreColor As String * 30
  ArticuloID As Long
  TemporadaID As Long
  StockActualT36 As Double
  StockActualT38 As Double
  StockActualT40 As Double
  StockActualT42 As Double
  StockActualT44 As Double
  StockActualT46 As Double
  StockActualT48 As Double
  StockActualT50 As Double
  StockActualT52 As Double
  StockActualT54 As Double
  StockActualT56 As Double
  StockPendienteT36 As Double
  StockPendienteT38 As Double
  StockPendienteT40 As Double
  StockPendienteT42 As Double
  StockPendienteT44 As Double
  StockPendienteT46 As Double
  StockPendienteT48 As Double
  StockPendienteT50 As Double
  StockPendienteT52 As Double
  StockPendienteT54 As Double
  StockPendienteT56 As Double
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type ArticuloColorData
  Buffer As String * 136
End Type

Public Type ArticuloColorAlmacenProps
  ArticuloColorID As Long
  AlmacenID As Long
  StockActualT36 As Double
  StockActualT38 As Double
  StockActualT40 As Double
  StockActualT42 As Double
  StockActualT44 As Double
  StockActualT46 As Double
  StockActualT48 As Double
  StockActualT50 As Double
  StockActualT52 As Double
  StockActualT54 As Double
  StockActualT56 As Double
  StockPendienteT36 As Double
  StockPendienteT38 As Double
  StockPendienteT40 As Double
  StockPendienteT42 As Double
  StockPendienteT44 As Double
  StockPendienteT46 As Double
  StockPendienteT48 As Double
  StockPendienteT50 As Double
  StockPendienteT52 As Double
  StockPendienteT54 As Double
  StockPendienteT56 As Double
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type ArticuloColorAlmacenData
  Buffer As String * 96
End Type

Public Type FormaDePagoProps
  FormaPagoID As Long
  Nombre As String * 50
  Giros As Long
  MesesPrimerGiro As Long
  MesesEntreGiros As Long
  Contado As Boolean
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type FormaDePagoData
  Buffer As String * 62
End Type

' tipo de datos para crear colecciones de vencimientos.
Public Type VencimientoProps
  Importe As Double
  Fecha As Date
  Giro As Long
End Type

Public Type RepresentanteProps
    RepresentanteID As Long
    Nombre As String * 50
    DNINIF As String * 20
    Contacto As String * 50
    Zona As String * 10
    Comision As Double
    IRPF As Double
    IVA As Double
    DireccionID As Long
    CuentaContable As String * 10
    Direccion As DireccionData
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type RepresentanteData
   Buffer As String * 540
End Type

Public Type DatoComercialProps
    DatoComercialID As Long
    Descuento As Double
    RecargoEquivalencia As Double
    IVA As Double
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type DatoComercialData
   Buffer As String * 18
End Type

Public Type ClienteProps
    ClienteID As Long
    Nombre As String * 50
    Titular As String * 50
    Contacto As String * 50
    DNINIF As String * 20
    DireccionFiscalID As Long
    DireccionEntregaID As Long
    TransportistaID As Long
    RepresentanteID As Long
    CuentaBancariaID As Long
    FormaPagoID As Long
    CuentaContable As String * 10
    DatoComercialID As Long
    DatoComercialBID As Long
    DiaPago1 As Integer
    DiaPago2 As Integer
    DiaPago3 As Integer
    PorcFacturacionAB As Integer
    DireccionFiscal As DireccionData
    DireccionEntrega As DireccionData
    CuentaBancaria As CuentaBancariaData
    DatoComercial As DatoComercialData
    DatoComercialB As DatoComercialData
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type ClienteData
   Buffer As String * 1128
End Type

Public Type PedidoVentaProps
    PedidoVentaID As Long
    ClienteID As Long
    NombreCliente As String * 50
    Fecha As Date
    FechaEntrega As Date
    FechaTopeServicio As Date
    Numero As Long
    Observaciones As String * 150
    RepresentanteID As Long
    NombreRepresentante As String * 50
    TransportistaID As Long
    NombreTransportista As String * 50
    FormaPagoID As Long
    DatoComercialID As Long
    DatoComercial As DatoComercialData
    TemporadaID As Long
    EmpresaID As Long
    TotalBrutoEUR As Double
    TotalBrutoPTA As Double
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type PedidoVentaData
   Buffer As String * 360
End Type

Public Type PedidoVentaItemProps
    PedidoVentaItemID As Long
    PedidoVentaID As Long
    ArticuloColorID As Long
    NombreArticuloColor As String * 50
    Situacion As String * 1
    SituacionCorte As String * 1
    CantidadT36 As Double
    CantidadT38 As Double
    CantidadT40 As Double
    CantidadT42 As Double
    CantidadT44 As Double
    CantidadT46 As Double
    CantidadT48 As Double
    CantidadT50 As Double
    CantidadT52 As Double
    CantidadT54 As Double
    CantidadT56 As Double
    ServidoT36 As Double
    ServidoT38 As Double
    ServidoT40 As Double
    ServidoT42 As Double
    ServidoT44 As Double
    ServidoT46 As Double
    ServidoT48 As Double
    ServidoT50 As Double
    ServidoT52 As Double
    ServidoT54 As Double
    ServidoT56 As Double
    PrecioVentaPTA As Double
    PrecioVentaEUR As Double
    Descuento As Double
    BrutoPTA As Double
    BrutoEUR As Double
    Comision As Double
    TemporadaID As Long
    ActualizarAlta As Boolean
    DesactualizarAlta As Boolean
    ActualizarAlbaran As Boolean
    DesactualizarAlbaran As Boolean
    Observaciones As String * 50
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type PedidoVentaItemData
   Buffer As String * 230
End Type

Public Type ProveedorProps
    ProveedorID As Long
    Nombre As String * 50
    Titular As String * 50
    Contacto As String * 50
    DNINIF As String * 20
    DireccionID As Long
    BancoID As Long
    TransportistaID As Long
    CuentaBancariaID As Long
    FormaPagoID As Long
    CuentaContable As String * 10
    CuentaContrapartida As String * 10
    DatoComercialID As Long
    MedioPagoID As Long
    Direccion As DireccionData
    CuentaBancaria As CuentaBancariaData
    DatoComercial As DatoComercialData
    Codigo As String * 3
    TipoProveedor As String * 1
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type ProveedorData
    Buffer As String * 738
End Type

Public Type TransportistaProps
    TransportistaID As Long
    Nombre As String * 50
    Titular As String * 50
    DNINIF As String * 20
    Contacto As String * 50
    Zona As String * 10
    DireccionID As Long
    Direccion As DireccionData
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type TransportistaData
    Buffer As String * 568
End Type

Public Type ConsultaProps
  ConsultaID As Long
  Nombre As String * 50
  Objeto As String * 50
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type ConsultaData
  Buffer As String * 106
End Type

Public Type ConsultaItemProps
  ConsultaItemID As Long
  ConsultaID As Long
  Alias As String * 30
  Campo As String * 30
  OperadorID As Long
  Valor1 As String * 50
  Valor2 As String * 50
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type ConsultaItemData
  Buffer As String * 170
End Type

' Tipo publico para la gestion de consultas.
' los tipos de campo posibles son:
' (N)umerico
' (A)lfanumerico
' (F)echa
' (B)ooleano
Public Type ConsultaCampoProps
    ConsultaCampoID As Long
    NombreCampo As String * 30
    Consulta As String * 50
    TipoCampo As String * 1
    Alias As String * 50
End Type

Public Type ConsultaCampoData
   Buffer As String * 134
End Type

Public Type PedidoCompraProps
    PedidoCompraID As Long
    ProveedorID As Long
    NombreProveedor As String * 50
    Fecha As Date
    FechaEntrega As Date
    Numero As Long
    NuestraReferencia As String * 20
    SuReferencia As String * 20
    Observaciones As String * 50
    BancoID As Long
    NombreBanco As String * 50
    TransportistaID As Long
    NombreTransportista As String * 50
    FormaPagoID As Long
    DatoComercialID As Long
    DatoComercial As DatoComercialData
    TemporadaID As Long
    EmpresaID As Long
    TotalBrutoPTA As Double
    TotalBrutoEUR As Double
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type PedidoCompraData
    Buffer As String * 296
End Type

Public Type PedidoCompraItemProps
    PedidoCompraItemID As Long
    PedidoCompraID As Long
    MaterialID As Long
    Situacion As String * 1
    Cantidad As Double
    Servido As Double
    PrecioCostePTA As Double
    PrecioCosteEUR As Double
    Descuento As Double
    BrutoPTA As Double
    BrutoEUR As Double
    Comision As Double
    ActualizarAlta As Boolean
    DesactualizarAlta As Boolean
    ActualizarAlbaran As Boolean
    DesactualizarAlbaran As Boolean
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type PedidoCompraItemData
    Buffer As String * 48
End Type

Public Type PedidoCompraArticuloProps
    PedidoCompraArticuloID As Long
    PedidoCompraID As Long
    ArticuloColorID As Long
    NombreArticuloColor As String * 50
    Situacion As String * 1
    CantidadT36 As Double
    CantidadT38 As Double
    CantidadT40 As Double
    CantidadT42 As Double
    CantidadT44 As Double
    CantidadT46 As Double
    CantidadT48 As Double
    CantidadT50 As Double
    CantidadT52 As Double
    CantidadT54 As Double
    CantidadT56 As Double
    ServidoT36 As Double
    ServidoT38 As Double
    ServidoT40 As Double
    ServidoT42 As Double
    ServidoT44 As Double
    ServidoT46 As Double
    ServidoT48 As Double
    ServidoT50 As Double
    ServidoT52 As Double
    ServidoT54 As Double
    ServidoT56 As Double
    PrecioCompraEUR As Double
    Descuento As Double
    BrutoEUR As Double
    Comision As Double
    TemporadaID As Long
    ActualizarAlta As Boolean
    DesactualizarAlta As Boolean
    ActualizarAlbaran As Boolean
    DesactualizarAlbaran As Boolean
    Observaciones As String * 50
    AlmacenID As Long
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type PedidoCompraArticuloData
   Buffer As String * 224
End Type

Public Type AlbaranVentaProps
    AlbaranVentaID As Long
    ClienteID As Long
    NombreCliente As String * 50
    Fecha As Date
    Numero As Long
    NuestraReferencia As String * 20
    SuReferencia As String * 20
    Observaciones As String * 50
    Bultos As Long
    PesoNeto As Long
    PesoBruto As Long
    PortesPTA As Double
    PortesEUR As Double
    EmbalajesPTA As Double
    EmbalajesEUR As Double
    TotalBrutoPTA As Double
    TotalBrutoEUR As Double
    RepresentanteID As Long
    NombreRepresentante As String * 50
    TransportistaID As Long
    NombreTransportista As String * 50
    FormaPagoID As Long
    DatoComercialID As Long
    TemporadaID As Long
    EmpresaID As Long
    DatoComercial As DatoComercialData
    FacturadoAB As Boolean
    FacturaVentaIDA As Long
    FacturaVentaIDB As Long
    AlmacenID As Long
    CentroGestionID As Long
    TerminalID As Long
    MedioPagoID As Long
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type AlbaranVentaData
    Buffer As String * 328
End Type

Public Type AlbaranVentaItemProps
    AlbaranVentaItemID As Long
    AlbaranVentaID As Long
    ArticuloColorID As Long
    Descripcion As String * 50
    PedidoVentaItemID As Long
    Situacion As String * 1
    CantidadT36 As Double
    CantidadT38 As Double
    CantidadT40 As Double
    CantidadT42 As Double
    CantidadT44 As Double
    CantidadT46 As Double
    CantidadT48 As Double
    CantidadT50 As Double
    CantidadT52 As Double
    CantidadT54 As Double
    CantidadT56 As Double
    PrecioVentaPTA As Double
    PrecioVentaEUR As Double
    Descuento As Double
    BrutoPTA As Double
    BrutoEUR As Double
    Comision As Double
    TemporadaID As Long
    ActualizarAlta As Boolean
    DesactualizarAlta As Boolean
    ActualizarFactura As Boolean
    DesactualizarFactura As Boolean
    FacturadoAB As Boolean
    FacturaVentaItemIDA As Long
    FacturaVentaItemIDB As Long
    AlmacenID As Long
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type AlbaranVentaItemData
   Buffer As String * 146
End Type

Public Type AlbaranCompraArticuloProps
    AlbaranCompraArticuloID As Long
    AlbaranCompraID As Long
    ArticuloColorID As Long
    Descripcion As String * 50
    PedidoCompraArticuloID As Long
    Situacion As String * 1
    CantidadT36 As Double
    CantidadT38 As Double
    CantidadT40 As Double
    CantidadT42 As Double
    CantidadT44 As Double
    CantidadT46 As Double
    CantidadT48 As Double
    CantidadT50 As Double
    CantidadT52 As Double
    CantidadT54 As Double
    CantidadT56 As Double
    PrecioCompraEUR As Double
    Descuento As Double
    BrutoEUR As Double
    Comision As Double
    TemporadaID As Long
    ActualizarAlta As Boolean
    DesactualizarAlta As Boolean
    ActualizarFactura As Boolean
    DesactualizarFactura As Boolean
    FacturadoAB As Boolean
    FacturaCompraArticuloID As Long
    AlmacenID As Long
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type AlbaranCompraArticuloData
   Buffer As String * 136
End Type

Public Type PickerItemProps
    DocumentoID As Long
    Numero As Long
    Nombre As String * 12
    Descripcion As String * 50
    Cantidad As Double
    Fecha As Date
End Type

Public Type PickerItemData
   Buffer As String * 74
End Type

Public Type FacturaVentaProps
    FacturaVentaID As Long
    ClienteID As Long
    NombreCliente As String * 50
    Fecha As Date
    Numero As Long
    NuestraReferencia As String * 20
    SuReferencia As String * 20
    Observaciones As String * 50
    SituacionContable As String * 1
    Bultos As Long
    PesoNeto As Long
    PesoBruto As Long
    BrutoPTA As Double
    BrutoEUR As Double
    DescuentoPTA As Double
    DescuentoEUR As Double
    PortesPTA As Double
    PortesEUR As Double
    EmbalajesPTA As Double
    EmbalajesEUR As Double
    BaseImponiblePTA As Double
    BaseImponibleEUR As Double
    IVAPTA As Double
    IVAEUR As Double
    RecargoPTA As Double
    RecargoEUR As Double
    NetoPTA As Double
    NetoEUR As Double
    RepresentanteID As Long
    NombreRepresentante As String * 50
    TransportistaID As Long
    NombreTransportista As String * 50
    FormaPagoID As Long
    DatoComercialID As Long
    TemporadaID As Long
    EmpresaID As Long
    FacturaVentaIDAB As Long
    AlmacenID As Long
    CentroGestionID As Long
    TerminalID As Long
    DatoComercial As DatoComercialData
    MedioPagoID As Long
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type FacturaVentaData
   Buffer As String * 366
End Type

Public Type FacturaVentaItemProps
    FacturaVentaItemID As Long
    FacturaVentaID As Long
    ArticuloColorID As Long
    Descripcion As String * 50
    AlbaranVentaItemID As Long
    SituacionImpresa As String * 1
    Cantidad As Double
    PrecioVentaPTA As Double
    PrecioVentaEUR As Double
    Descuento As Double
    BrutoPTA As Double
    BrutoEUR As Double
    Comision As Double
    ComisionPTA As Double
    ComisionEUR As Double
    TemporadaID As Long
    ActualizarAlta As Boolean
    DesactualizarAlta As Boolean
    AlmacenID As Long
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type FacturaVentaItemData
    Buffer As String * 106
End Type

Public Type CobroPagoProps
    CobroPagoID As Long
    Tipo As String * 1
    Vencimiento As Date
    PersonaID As Long
    NombrePersona As String * 50
    FormaPagoID As Long
    FormaPago As String * 50
    FacturaID As Long
    NumeroFactura As Long
    NumeroGiro As Long
    SituacionComercial As String * 1
    SituacionContable As String * 1
    ImportePTA As Double
    ImporteEUR As Double
    FechaEmision As Date
    FechaDomiciliacion As Date
    FechaContable As Date
    BancoID As Long
    NombreBanco As String * 50
    EmpresaID As Long
    TemporadaID As Long
    MedioPagoID As Long
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type CobroPagoData
   Buffer As String * 202
End Type

Public Type AlbaranCompraProps
    AlbaranCompraID As Long
    ProveedorID As Long
    NombreProveedor As String * 50
    Fecha As Date
    Numero As Long
    NuestraReferencia As String * 20
    SuReferencia As String * 20
    Observaciones As String * 50
    PortesPTA As Double
    PortesEUR As Double
    EmbalajesPTA As Double
    EmbalajesEUR As Double
    TotalBrutoPTA As Double
    TotalBrutoEUR As Double
    TransportistaID As Long
    NombreTransportista As String * 50
    DatoComercialID As Long
    TemporadaID As Long
    EmpresaID As Long
    DatoComercial As DatoComercialData
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type AlbaranCompraData
    Buffer As String * 254
End Type

Public Type AlbaranCompraItemProps
    AlbaranCompraItemID As Long
    AlbaranCompraID As Long
    MaterialID As Long
    PedidoCompraItemID As Long
    Situacion As String * 1
    Cantidad As Double
    PrecioCostePTA As Double
    PrecioCosteEUR As Double
    Descuento As Double
    BrutoPTA As Double
    BrutoEUR As Double
    Comision As Double
    ActualizarAlta As Boolean
    DesactualizarAlta As Boolean
    ActualizarFactura As Boolean
    DesactualizarFactura As Boolean
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type AlbaranCompraItemData
    Buffer As String * 46
End Type

Public Type FacturaCompraProps
    FacturaCompraID As Long
    ProveedorID As Long
    NombreProveedor As String * 50
    Fecha As Date
    FechaContable As Date
    Numero As Long
    Sufijo As String * 10
    NuestraReferencia As String * 20
    SuReferencia As String * 20
    Observaciones As String * 50
    SituacionContable As String * 1
    BrutoPTA As Double
    BrutoEUR As Double
    DescuentoPTA As Double
    DescuentoEUR As Double
    PortesPTA As Double
    PortesEUR As Double
    EmbalajesPTA As Double
    EmbalajesEUR As Double
    BaseImponiblePTA As Double
    BaseImponibleEUR As Double
    IVAPTA As Double
    IVAEUR As Double
    RecargoPTA As Double
    RecargoEUR As Double
    NetoPTA As Double
    NetoEUR As Double
    BancoID As Long
    NombreBanco As String * 50
    TransportistaID As Long
    NombreTransportista As String * 50
    FormaPagoID As Long
    DatoComercialID As Long
    TemporadaID As Long
    EmpresaID As Long
    MedioPagoID As Long
    DatoComercial As DatoComercialData
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type FacturaCompraData
   Buffer As String * 366
End Type

Public Type FacturaCompraItemProps
    FacturaCompraItemID As Long
    FacturaCompraID As Long
    MaterialID As Long
    NombreMaterial As String * 50
    AlbaranCompraItemID As Long
    SituacionImpresa As String * 1
    Cantidad As Double
    PrecioCostePTA As Double
    PrecioCosteEUR As Double
    Descuento As Double
    BrutoPTA As Double
    BrutoEUR As Double
    ActualizarAlta As Boolean
    DesactualizarAlta As Boolean
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type FacturaCompraItemData
    Buffer As String * 90
End Type

Public Type EtiquetaProps
    ArticuloColorID As Long
    CantidadT36 As Double
    CantidadT38 As Double
    CantidadT40 As Double
    CantidadT42 As Double
    CantidadT44 As Double
    CantidadT46 As Double
    CantidadT48 As Double
    CantidadT50 As Double
    CantidadT52 As Double
    CantidadT54 As Double
    CantidadT56 As Double
    TemporadaID As Long
    Composicion1 As String * 20
    PorcComposicion1 As Double
    Composicion2 As String * 20
    PorcComposicion2 As Double
    Composicion3 As String * 20
    PorcComposicion3 As Double
    Composicion4 As String * 20
    PorcComposicion4 As Double
    NombrePrenda As String * 50
    NombreModelo As String * 50
    NombreSerie As String * 50
    NombreColor As String * 50
    PrecioVentaPublico As Double
    CodigoProveedor As String * 3
    TallajeID As Long
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type OrdenCorteProps
    OrdenCorteID As Long
    Fecha As Date
    FechaCorte As Date
    Numero As Long
    ArticuloID As Long
    Observaciones As String * 50
    Nombre As String * 50
    TemporadaID As Long
    EmpresaID As Long
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type OrdenCorteData
    Buffer As String * 122
End Type

Public Type OrdenCorteItemProps
    OrdenCorteItemID As Long
    OrdenCorteID As Long
    ArticuloColorID As Long
    ArticuloID As Long
    Descripcion As String * 50
    PedidoVentaItemID As Long
    Situacion As String * 1
    CantidadT36 As Double
    CantidadT38 As Double
    CantidadT40 As Double
    CantidadT42 As Double
    CantidadT44 As Double
    CantidadT46 As Double
    CantidadT48 As Double
    CantidadT50 As Double
    CantidadT52 As Double
    CantidadT54 As Double
    CantidadT56 As Double
    Numero As Long
    Cliente As String * 50
    TemporadaID As Long
    Actualizar As Boolean
    Desactualizar As Boolean
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type OrdenCorteItemData
   Buffer As String * 166
End Type
    
Public Type RemesaProps
    'RemesaID As Long
    FechaDomiciliacion As Date
    BancoID As Long
    NombreEntidad As String * 50
    SituacionComercial As String * 1
    NumeroEfectos As Long
    ImportePTA As Double
    ImporteEUR As Double
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type RemesaData
   Buffer As String * 72
End Type
    
' Tipos de importes contables Debe/Haber
Public Enum enTipoImporte
    TipoImporteDebe = 1
    TipoImporteHaber = 2
End Enum

' Tipos de apuntes contables Pesetas/Euros
Public Enum enTipoApunte
    TipoApuntePesetas = 1
    TipoApunteEuros = 2
End Enum

Public Type AsientoProps
    AsientoID As Long
    Numero As Long
    Ejercicio As String * 4
    Concepto As String * 50
    FechaAlta As Date
    TemporadaID As Long
    EmpresaID As Long
    Situacion As String * 1
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type AsientoData
    Buffer As String * 70
End Type

Public Type ApunteProps
    ApunteID As Long
    AsientoID As Long
    Cuenta As String * 10
'    TipoImporte As String * 1
    TipoImporte As enTipoImporte
    ImportePTA As Double
    ImporteEUR As Double
    Descripcion As String * 50
    Fecha As Date
    Documento As String * 20
'    TipoApunte As String * 1
    TipoApunte As enTipoApunte
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type ApunteData
   Buffer As String * 104
End Type

Public Type IVAApunteProps
    IVAApunteID As Long
    AsientoID As Long
    TipoApunte As String * 1
    NumeroFactura As String * 20
    CuentaIVA As String * 10
    CuentaTotal As String * 10
    CuentaBase As String * 10
    Titular As String * 50
    DNINIF As String * 20
    BaseImponible As Double
    Total As Double
    IVA As Double
    CuotaIVA As Double
    RecargoEquivalencia As Double
    CuotaRecargo As Double
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type IVAApunteData
   Buffer As String * 154
End Type

Public Type FichaPedidoProps
    Numero As Long
    Fecha As Date
    FechaTopeServicio As Date
    CodigoColor As String * 2
    NombreModelo As String * 50
    Observaciones As String * 50
    NombreCliente As String * 50
    Cantidad As Long
    FechaOrden As Date
    NumeroOrden As Long
End Type

Public Type FichaPedidoData
   Buffer As String * 170
End Type

Public Type OperarioProps
  OperarioID As Long
  Nombre As String * 50
  PrecioHora As Double
  Activo As Boolean
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type OperarioData
  Buffer As String * 60
End Type

Public Type ParametroAplicacionProps
    ParametroAplicacionID As String * 10
    Nombre As String * 100
    Valor As String * 100
    Sistema As Boolean
    TipoParametro As Long
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type ParametroAplicacionData
   Buffer As String * 218
End Type

Public Type TerminalProps
    TerminalID As Long
    Nombre As String * 100
    CentroGestionID As Long
    AlmacenID As Long
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type TerminalData
   Buffer As String * 110
End Type

Public Type CentroGestionProps
    CentroGestionID As Long
    Nombre As String * 100
    DireccionID As Long
    Direccion As DireccionData
    ContadorTicketID As Long
    ContadorPedidoVentaID As Long
    ContadorAlbaranVentaID As Long
    ContadorFacturaVentaID As Long
    ContadorPedidoCompraID As Long
    ContadorAlbaranCompraID As Long
    ContadorFacturaCompraID As Long
    SedeCentral As Boolean
    EmpresaID As Long
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type CentroGestionData
   Buffer As String * 506
End Type

Public Type TraspasoProps
    TraspasoID As Long
    AlmacenOrigenID As Long
    AlmacenDestinoID As Long
    Situacion As String * 2
    FechaAlta As Date
    FechaTransito As Date
    FechaRecepcion As Date
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type TraspasoData
   Buffer As String * 24
End Type

Public Type TraspasoItemProps
    TraspasoItemID As Long
    TraspasoID As Long
    ArticuloColorID As Long
    Situacion As String * 2
    CantidadT36 As Double
    CantidadT38 As Double
    CantidadT40 As Double
    CantidadT42 As Double
    CantidadT44 As Double
    CantidadT46 As Double
    CantidadT48 As Double
    CantidadT50 As Double
    CantidadT52 As Double
    CantidadT54 As Double
    CantidadT56 As Double
    Observaciones As String * 50
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type TraspasoItemData
   Buffer As String * 106
End Type

Public Type TallajeProps
    TallajeID As Long
    Nombre As String * 30
    Estandar As Boolean
    DescripcionT36 As String * 5
    DescripcionT38 As String * 5
    DescripcionT40 As String * 5
    DescripcionT42 As String * 5
    DescripcionT44 As String * 5
    DescripcionT46 As String * 5
    DescripcionT48 As String * 5
    DescripcionT50 As String * 5
    DescripcionT52 As String * 5
    DescripcionT54 As String * 5
    DescripcionT56 As String * 5
    PermitidoT36 As Boolean
    PermitidoT38 As Boolean
    PermitidoT40 As Boolean
    PermitidoT42 As Boolean
    PermitidoT44 As Boolean
    PermitidoT46 As Boolean
    PermitidoT48 As Boolean
    PermitidoT50 As Boolean
    PermitidoT52 As Boolean
    PermitidoT54 As Boolean
    PermitidoT56 As Boolean
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type TallajeData
   Buffer As String * 102
End Type

