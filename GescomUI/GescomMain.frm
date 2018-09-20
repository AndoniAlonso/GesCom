VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm GescomMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Sistema de Gestión Comercial"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   630
   ClientWidth     =   10725
   Icon            =   "GescomMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComCtl3.CoolBar cbrHerramientas 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   1270
      BandCount       =   4
      BandBorders     =   0   'False
      VariantHeight   =   0   'False
      _CBWidth        =   10725
      _CBHeight       =   720
      _Version        =   "6.7.8988"
      Child1          =   "tlbPrincipal"
      MinWidth1       =   2505
      MinHeight1      =   270
      Width1          =   1005
      UseCoolbarPicture1=   0   'False
      NewRow1         =   0   'False
      Caption2        =   "Compras"
      Child2          =   "tlbCompras"
      MinWidth2       =   4005
      MinHeight2      =   330
      Width2          =   1500
      UseCoolbarPicture2=   0   'False
      NewRow2         =   0   'False
      Caption3        =   "Ventas"
      Child3          =   "tlbVentas"
      MinWidth3       =   2805
      MinHeight3      =   330
      Width3          =   2805
      UseCoolbarPicture3=   0   'False
      NewRow3         =   -1  'True
      Caption4        =   "Fabricación"
      Child4          =   "tlbFabricacion"
      MinWidth4       =   1995
      MinHeight4      =   270
      Width4          =   1995
      UseCoolbarPicture4=   0   'False
      NewRow4         =   0   'False
      Begin MSComctlLib.Toolbar tlbFabricacion 
         Height          =   270
         Left            =   8640
         TabIndex        =   4
         Top             =   390
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   476
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "mglIconosPequeños"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "OrdenCorte"
               Description     =   "Órdenes de corte"
               Object.ToolTipText     =   "Órdenes de corte"
               ImageKey        =   "Corte"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Etiquetas"
               Description     =   "Etiquetas de artículos"
               Object.ToolTipText     =   "Etiquetas de artículos"
               ImageKey        =   "Etiqueta"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbCompras 
         Height          =   330
         Left            =   3555
         TabIndex        =   2
         Top             =   30
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "mglIconosPequeños"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Material"
               Object.ToolTipText     =   "Materiales"
               ImageKey        =   "Material"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Proveedor"
               Object.ToolTipText     =   "Proveedores"
               ImageKey        =   "Proveedor"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "PedidoCompra"
               Object.ToolTipText     =   "Pedidos de compra"
               ImageKey        =   "PedidoCompra"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "AlbaranCompra"
               Object.ToolTipText     =   "Albaranes de compra"
               ImageKey        =   "AlbaranCompra"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FacturaCompra"
               Object.ToolTipText     =   "Facturas de compra"
               ImageKey        =   "FacturaCompra"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "MoviMateriales"
               Object.ToolTipText     =   "Movimientos de materiales"
               ImageKey        =   "Movimientos"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ListaPagos"
               Object.ToolTipText     =   "Lista de pagos"
               ImageKey        =   "CobroPago"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbVentas 
         Height          =   330
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "mglIconosPequeños"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   25
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Banco"
               Object.ToolTipText     =   "Bancos"
               ImageKey        =   "Banco"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Transportista"
               Object.ToolTipText     =   "Transportistas"
               ImageKey        =   "Transportista"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Representante"
               Object.ToolTipText     =   "Representantes"
               ImageKey        =   "Representante"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cliente"
               Object.ToolTipText     =   "Clientes"
               ImageKey        =   "Cliente"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Prenda"
               Object.ToolTipText     =   "Prendas"
               ImageKey        =   "Prenda"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Serie"
               Object.ToolTipText     =   "Series"
               ImageKey        =   "Serie"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Modelo"
               Object.ToolTipText     =   "Modelos"
               ImageKey        =   "Modelo"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Articulo"
               Object.ToolTipText     =   "Artículos"
               ImageKey        =   "Articulo"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ArticuloColor"
               Object.ToolTipText     =   "Artículos - colores"
               ImageKey        =   "ArticuloColor"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "PedidoVenta"
               Object.ToolTipText     =   "Pedidos de venta"
               ImageKey        =   "PedidoVenta"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "AlbaranVenta"
               Object.ToolTipText     =   "Albaranes de venta"
               ImageKey        =   "AlbaranVenta"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FacturaVenta"
               Object.ToolTipText     =   "Facturas de Venta"
               ImageKey        =   "FacturaVenta"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Remesa"
               Object.ToolTipText     =   "Gestión de remesas"
               ImageKey        =   "Remesa"
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ListaCobros"
               Object.ToolTipText     =   "ListaCobros"
               ImageKey        =   "CobroPago"
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "AlbaranAutomatico"
               Object.ToolTipText     =   "Albaranes automaticos"
               ImageKey        =   "BarCode"
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EstadisticaVentas"
               Object.ToolTipText     =   "Estadísticas de venta"
               ImageKey        =   "OLAPQuery"
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EstadisticaVentaTiendas"
               Object.ToolTipText     =   "Ventas por tipo de proveedor"
               ImageKey        =   "OLAPQuery"
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "PedidosPendientes"
               Object.ToolTipText     =   "Consulta de pedidos pendientes"
               ImageKey        =   "PedidoVenta"
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Traspasos"
               Object.ToolTipText     =   "Traspasos de artículos"
               ImageKey        =   "Traspaso"
            EndProperty
         EndProperty
         MouseIcon       =   "GescomMain.frx":08CA
      End
      Begin MSComctlLib.Toolbar tlbPrincipal 
         Height          =   270
         Left            =   165
         TabIndex        =   1
         Top             =   60
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   476
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "mglIconosPequeños"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Temporada"
               Object.ToolTipText     =   "Temporadas"
               ImageKey        =   "Temporada"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Empresa"
               Object.ToolTipText     =   "Empresas"
               ImageKey        =   "Empresa"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Parametro"
               Object.ToolTipText     =   "Parámetros de la aplicación"
               ImageKey        =   "Parametro"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Context"
               Object.ToolTipText     =   "Cambiar empresa y temporada"
               ImageKey        =   "Context"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cerrar"
               Object.ToolTipText     =   "Salir"
               ImageKey        =   "Cerrar"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList mglIconosGrandes 
      Left            =   360
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   27
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":191C
            Key             =   "Temporada"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":1D70
            Key             =   "FacturaVenta"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":2DC2
            Key             =   "FacturaCompra"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":3E14
            Key             =   "ArticuloColor"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":4C68
            Key             =   "AlbaranCompra"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":5CBC
            Key             =   "Proveedor"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":6998
            Key             =   "Material"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":77EC
            Key             =   "Articulo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":8640
            Key             =   "PedidoCompra"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":9694
            Key             =   "PedidoVenta"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":A6E8
            Key             =   "AlbaranVenta"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":B73C
            Key             =   "Modelo"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":C590
            Key             =   "Serie"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":C9EC
            Key             =   "Prenda"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":DA40
            Key             =   "Cliente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":E71C
            Key             =   "Representante"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":EA38
            Key             =   "Transportista"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":F314
            Key             =   "Banco"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":FBFC
            Key             =   "Empresa"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":10050
            Key             =   "Corte"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":1036A
            Key             =   "Etiqueta"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":10684
            Key             =   "Remesa"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":1099E
            Key             =   "CobroPago"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":10CB8
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":1110A
            Key             =   "Contabilidad"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":119E4
            Key             =   "Contawin"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":11CFE
            Key             =   "Cobrados"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList mglIconosPequeños 
      Left            =   360
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   59
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":12018
            Key             =   "IconosPequeños"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":1212C
            Key             =   "FacturaVenta"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":1317E
            Key             =   "FacturaCompra"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":141D0
            Key             =   "ArticuloColor"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":15024
            Key             =   "Context"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":1547C
            Key             =   "Parametro"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":15798
            Key             =   "Documento"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":15BF4
            Key             =   "AlbaranCompra"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":16C48
            Key             =   "EliminarItem"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":1709C
            Key             =   "ModificarItem"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":174F0
            Key             =   "NuevoItem"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":17944
            Key             =   "Proveedor"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":18620
            Key             =   "Material"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":19474
            Key             =   "PedidoCompra"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":1A4C8
            Key             =   "PedidoVenta"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":1B51C
            Key             =   "AlbaranVenta"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":1C570
            Key             =   "Articulo"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":1D3C4
            Key             =   "Modelo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":1E218
            Key             =   "Serie"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":1E66C
            Key             =   "Prenda"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":1F6C0
            Key             =   "Cliente"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":2039C
            Key             =   "Representante"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":206B8
            Key             =   "Transportista"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":20F94
            Key             =   "Banco"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":2187C
            Key             =   "Empresa"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":21CD0
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":21FEC
            Key             =   "Nuevo"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":22100
            Key             =   "Abrir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":22214
            Key             =   "Detalle"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":22328
            Key             =   "Lista"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":2243C
            Key             =   "IconosGrandes"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":22550
            Key             =   "Eliminar"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":22664
            Key             =   "Actualizar"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":22BA8
            Key             =   "Temporada"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":22FFC
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":23318
            Key             =   "Corte"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":23632
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":25DE4
            Key             =   "Etiqueta"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":260FE
            Key             =   "Cobrar"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":26550
            Key             =   "PrevEtiqueta"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":2686A
            Key             =   "OLAPQuery"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":26B84
            Key             =   "Remesa"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":26E9E
            Key             =   "PrintDocument"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":271B8
            Key             =   "CobroPago"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":274D2
            Key             =   "Recalcular"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":277EC
            Key             =   "Contabilidad"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":280C6
            Key             =   "Contawin"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":283E0
            Key             =   "Cobrados"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":286FA
            Key             =   "Movimientos"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":28FDC
            Key             =   "BarCode"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":2942E
            Key             =   "GroupBy"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":29588
            Key             =   "Propiedades"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":296E2
            Key             =   "Ordenar"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":2983C
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":2AB46
            Key             =   "PVP"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":2AF98
            Key             =   "Traspaso"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":2B2B2
            Key             =   "EnviarTraspaso"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":2BB8C
            Key             =   "RecepcionarTraspaso"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GescomMain.frx":2C466
            Key             =   "CentroGestion"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgFileSave 
      Left            =   360
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuFileTemporada 
         Caption         =   "&Temporadas"
         Begin VB.Menu mnuFileTemporadaList 
            Caption         =   "&Temporadas"
         End
         Begin VB.Menu mnuFileTemporadaLine1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileTemporadaNew 
            Caption         =   "&Nueva Temporada"
         End
      End
      Begin VB.Menu mnuFileEmpresa 
         Caption         =   "&Empresas"
         Begin VB.Menu mnuFileEmpresaList 
            Caption         =   "&Empresas"
         End
         Begin VB.Menu mnuFileEmpresaLine1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileEmpresaNew 
            Caption         =   "&Nueva Empresa"
         End
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuVentas 
      Caption         =   "&Ventas"
      Begin VB.Menu mnuVentasBancos 
         Caption         =   "&Bancos"
         Begin VB.Menu mnuVentasBancosList 
            Caption         =   "&Bancos"
         End
         Begin VB.Menu mnuVentasBancosSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVentasBancosNew 
            Caption         =   "&Nuevo Banco"
         End
      End
      Begin VB.Menu mnuVentasTransportistas 
         Caption         =   "&Transportistas"
         Begin VB.Menu mnuVentasTransportistasList 
            Caption         =   "&Transportistas"
         End
         Begin VB.Menu mnuVentasTransportistasSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVentasTransportistasNew 
            Caption         =   "&Nuevo Transportista"
         End
      End
      Begin VB.Menu mnuVentasRepresentantes 
         Caption         =   "&Representantes"
         Begin VB.Menu mnuVentasRepresentantesList 
            Caption         =   "&Representantes"
         End
         Begin VB.Menu mnuVentasRepresentantesSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVentasRepresentantesNew 
            Caption         =   "&Nuevo Representante"
         End
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVentasClientes 
         Caption         =   "&Clientes"
         Begin VB.Menu mnuVentasClienteList 
            Caption         =   "&Clientes"
         End
         Begin VB.Menu mnuVentasClienteLine1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVentasClienteNew 
            Caption         =   "&Nuevo Cliente"
         End
      End
      Begin VB.Menu mnuVentasPrendas 
         Caption         =   "&Prendas"
         Begin VB.Menu mnuVentasPrendaList 
            Caption         =   "&Prendas"
         End
         Begin VB.Menu mnuVentasPrendaLine1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVentasPrendaNew 
            Caption         =   "&Nueva Prenda"
         End
      End
      Begin VB.Menu nuVentasseries 
         Caption         =   "&Series"
         Begin VB.Menu mnuVentasSerieList 
            Caption         =   "&Series"
         End
         Begin VB.Menu nuVentasSerieLine1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVentasSerieNew 
            Caption         =   "&Nueva Serie"
         End
      End
      Begin VB.Menu nuVentasModelos 
         Caption         =   "&Modelos"
         Begin VB.Menu mnuVentasModeloList 
            Caption         =   "&Modelos"
         End
         Begin VB.Menu nuVentasModeloLine1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVentasModeloNew 
            Caption         =   "&Nuevo Modelo"
         End
      End
      Begin VB.Menu mnuVentasArticulos 
         Caption         =   "Artíc&ulos"
         Begin VB.Menu mnuVentasArticulosArticulos 
            Caption         =   "&Artículos"
            Begin VB.Menu mnuVentasArticulosArticuloList 
               Caption         =   "&Artículos"
            End
            Begin VB.Menu nuVentasArticuloLine1 
               Caption         =   "-"
            End
            Begin VB.Menu mnuVentasArticulosArticuloNew 
               Caption         =   "&Nuevo Artículo"
            End
         End
         Begin VB.Menu nuVentasArticuloLine2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVentasArticulosArticuloColores 
            Caption         =   "Artículo - &Colores"
            Begin VB.Menu mnuVentasArticulosArticuloColorList 
               Caption         =   "&Artículo - Colores"
            End
            Begin VB.Menu mnuSep6 
               Caption         =   "-"
            End
            Begin VB.Menu mnuVentasArticulosArticuloColorNew 
               Caption         =   "Nuevo Artículo - &Color"
            End
         End
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVentasPedidosVenta 
         Caption         =   "&Pedidos de Venta"
         Begin VB.Menu mnuVentasPedidoVentaList 
            Caption         =   "&Pedidos de Venta"
         End
         Begin VB.Menu mnuVentasPedidoVentaLine1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVentasPedidoVentaNew 
            Caption         =   "&Nuevo Pedido de Venta"
         End
         Begin VB.Menu mnuVentasPedidoVentaLine2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVentasPedidoSinCorte 
            Caption         =   "Pedidos pendientes de &corte"
         End
      End
      Begin VB.Menu mnuVentasSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVentasAlbaranes 
         Caption         =   "&Albaranes de Venta"
         Begin VB.Menu mnuVentasAlbaranesList 
            Caption         =   "&Albaranes de Venta"
         End
         Begin VB.Menu mnuVentasAlbaranesSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVentasAlbaranesNew 
            Caption         =   "&Nuevo Albarán de Ventas"
         End
      End
      Begin VB.Menu mnuVentasSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVentasFacturas 
         Caption         =   "&Facturas de Venta"
         Begin VB.Menu mnuVentasFacturasList 
            Caption         =   "&Facturas de Venta"
         End
         Begin VB.Menu mnuVentasFacturasSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVentasFacturasNew 
            Caption         =   "&Nueva Factura de Ventas"
         End
      End
      Begin VB.Menu mnuVentasSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVentasCobros 
         Caption         =   "&Lista de Cobros"
      End
      Begin VB.Menu mnuVentasRemesas 
         Caption         =   "&Gestión de remesas"
      End
      Begin VB.Menu mnuVentasSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVentasTraspasos 
         Caption         =   "&Traspasos de artículos"
      End
      Begin VB.Menu mnuVentasSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVentasEstadisticas 
         Caption         =   "&Estadísticas"
      End
   End
   Begin VB.Menu mnuCompras 
      Caption         =   "&Compras"
      Begin VB.Menu mnuComprasMaterial 
         Caption         =   "&Materiales"
         Begin VB.Menu mnuComprasMaterialList 
            Caption         =   "&Materiales"
         End
         Begin VB.Menu mnuComprasMaterialLine1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuComprasMaterialNew 
            Caption         =   "&Nuevo Material"
         End
         Begin VB.Menu mnuComprasMaterialLine2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuComprasMoviMateriales 
            Caption         =   "Mo&vimientos"
         End
      End
      Begin VB.Menu mnuComprasProveedores 
         Caption         =   "&Proveedores"
         Begin VB.Menu mnuComprasProveedoresList 
            Caption         =   "&Proveedores"
         End
         Begin VB.Menu mnuComprasProveedoresSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuComprasProveedoresNew 
            Caption         =   "&Nuevo Proveedor"
         End
      End
      Begin VB.Menu mnuComprasSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComprasPedidoCompra 
         Caption         =   "Pedidos de &Compra"
         Begin VB.Menu mnuComprasPedidoCompraList 
            Caption         =   "&Pedidos de Compra"
         End
         Begin VB.Menu mnuComprasPedidoCompraSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuComprasPedidoCompraNew 
            Caption         =   "&Nuevo Pedido de Compra"
         End
      End
      Begin VB.Menu mnuComprasSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComprasAlbaranes 
         Caption         =   "&Albaranes de Compra"
         Begin VB.Menu mnuComprasAlbaranesList 
            Caption         =   "&Albaranes de Compra"
         End
         Begin VB.Menu mnuComprasAlbaranesSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuComprasAlbaranEdit 
            Caption         =   "&Nuevo Albarán de Compra"
         End
      End
      Begin VB.Menu mnuComprasSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComprasFacturas 
         Caption         =   "&Facturas de Compra"
         Begin VB.Menu mnuComprasFacturasList 
            Caption         =   "&Facturas de Compra"
         End
         Begin VB.Menu mnuComprasFacturasSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuComprasFacturasNew 
            Caption         =   "&Nueva Factura de Compra"
         End
      End
      Begin VB.Menu mnuComprasSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComprasPagos 
         Caption         =   "&Lista de Pagos"
      End
   End
   Begin VB.Menu mnuFabricacion 
      Caption         =   "&Fabricación"
      Begin VB.Menu mnuFabrOrdenCorte 
         Caption         =   "&Ordenes de corte"
         Begin VB.Menu mnuFabrOrdenCorteList 
            Caption         =   "&Ordenes de corte"
         End
         Begin VB.Menu mnuFabrOrdenCorteSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFabrOrdenCorteNew 
            Caption         =   "&Nueva órden de corte"
         End
      End
      Begin VB.Menu mnuFabrOrdenCorteSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFabrEtiquetas 
         Caption         =   "&Etiquetas de artículos"
      End
   End
   Begin VB.Menu mnuContabilidad 
      Caption         =   "Conta&bilidad"
      Begin VB.Menu mnuFileAsiento 
         Caption         =   "&Asientos"
         Begin VB.Menu mnuFileAsientoList 
            Caption         =   "&Asientos"
         End
         Begin VB.Menu mnuFileAsientoLine1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileAsientoNew 
            Caption         =   "&Nuevo Asiento"
         End
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "&Utilidades"
      Begin VB.Menu mnuUtilityParam 
         Caption         =   "&Parámetros de la Aplicación"
      End
      Begin VB.Menu mnuUtilityContext 
         Caption         =   "&Seleccionar Empresa y Temporada"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Ve&ntana"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuListView 
      Caption         =   "Opciones del objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuListviewEdit 
         Caption         =   "&Abrir"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListViewNew 
         Caption         =   "&Nuevo"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListviewDel 
         Caption         =   "&Eliminar"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListViewSearch 
         Caption         =   "&Buscar"
      End
      Begin VB.Menu mnuListViewQuickSearch 
         Caption         =   "B&úsqueda rápida"
      End
   End
End
Attribute VB_Name = "GescomMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GesComSectionName = "GesCom"

Public mobjParametro As Parametro
Private mobjTerminal As Terminal

Private mobjGescomApp As clsMDIParentForm

Public Property Get Terminal()
    Set Terminal = mobjTerminal
End Property

Public Property Get objParametro()
    Set objParametro = mobjParametro
End Property

Private Sub MDIForm_Load()

    On Error GoTo ErrorManager
    
    Set mobjParametro = New Parametro
    mobjParametro.Load
  
    Set mobjTerminal = New Terminal
  
    GescomTerminal
    GescomTitulo
    GescomApp
    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
    TerminateProgram
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    Set mobjGescomApp = Nothing
    Set mobjParametro = Nothing
    Set mobjTerminal = Nothing

End Sub

Private Sub mnuComprasAlbaranEdit_Click()

    Dim objAlbaranCompra As AlbaranCompra
    Dim frmAlbaranCompra As AlbaranCompraEdit
  
    Set objAlbaranCompra = New AlbaranCompra
    Set frmAlbaranCompra = New AlbaranCompraEdit
  
    frmAlbaranCompra.Component objAlbaranCompra
    frmAlbaranCompra.Show
    
End Sub

Private Sub mnuComprasAlbaranesList_Click()
    Dim frmList As AlbaranCompraList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New AlbaranCompraList
    Set objRecordList = New RecordList

    With frmList
        .ComponentStatus objRecordList.Load("Select * from vAlbaranesCompra", _
                        "TemporadaID = " & GescomMain.mobjParametro.TemporadaActualID & " AND " & _
                        "EmpresaID = " & GescomMain.mobjParametro.EmpresaActualID)
        .Show

    End With

    Set objRecordList = Nothing

    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuComprasFacturasList_Click()
    Dim frmList As FacturaCompraList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New FacturaCompraList
    Set objRecordList = New RecordList

    With frmList
        .ComponentStatus objRecordList.Load("Select * from vFacturasCompra", "1=2")
' _
'                        "TemporadaID = " & GescomMain.mobjParametro.TemporadaActualID & " AND " & _
'                        "EmpresaID = " & GescomMain.mobjParametro.EmpresaActualID)
        .Show

    End With

    Set objRecordList = Nothing
    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuComprasFacturasNew_Click()

    Dim objFacturaCompra As FacturaCompra
    Dim frmFacturaCompra As FacturaCompraEdit
  
    Set objFacturaCompra = New FacturaCompra
    Set frmFacturaCompra = New FacturaCompraEdit
  
    frmFacturaCompra.Component objFacturaCompra
    frmFacturaCompra.Show
    
End Sub

Private Sub mnuComprasMoviMateriales_Click()
    Dim frmList As MoviMaterialList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New MoviMaterialList
    Set objRecordList = New RecordList

    With frmList
        .ComponentStatus objRecordList.Load("SELECT * FROM vMoviMateriales ", _
                        "Fecha BETWEEN '" & CStr(Date) & "'" & " AND '" & CStr(Date + 1) & "'")
        .Show

    End With

    Set objRecordList = Nothing
    
    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuComprasPagos_Click()
    Dim frmList As CobroPagoList

    On Error GoTo ErrorManager

    Set frmList = New CobroPagoList

    With frmList
        .ComponentQuery "vPagos", _
                        GescomMain.mobjParametro.EmpresaActualID, _
                        "Lista de Pagos", GCEuro.gcTipoPago
        .Show

    End With

    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuComprasPedidoCompraList_Click()
Dim frmList As PedidoCompraList
Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New PedidoCompraList
    Set objRecordList = New RecordList
    
    With frmList
        .ComponentStatus objRecordList.Load("SELECT * FROM vPedidosCompra ", _
                        "TemporadaID = " & GescomMain.mobjParametro.TemporadaActualID & _
                        " AND EmpresaID = " & GescomMain.mobjParametro.EmpresaActualID)
        .Show
  
    End With

    Set objRecordList = Nothing

    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuComprasPedidoCompraNew_Click()

    Dim objPedidoCompra As PedidoCompra
    Dim frmPedidoCompra As PedidoCompraEdit
  
    Set objPedidoCompra = New PedidoCompra
    Set frmPedidoCompra = New PedidoCompraEdit
  
    frmPedidoCompra.Component objPedidoCompra
    frmPedidoCompra.Show
    
End Sub

Private Sub mnuFabrEtiquetas_Click()
    Dim objEtiquetas As Etiquetas
    Dim frmEtiquetas As EtiquetasEdit
  
    Set objEtiquetas = New Etiquetas
    Set frmEtiquetas = New EtiquetasEdit
  
    frmEtiquetas.Component objEtiquetas
    frmEtiquetas.Show

End Sub

Private Sub mnuFabrOrdenCorteList_Click()
    Dim frmList As OrdenCorteList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New OrdenCorteList
    Set objRecordList = New RecordList
      
    With frmList
        .ComponentStatus objRecordList.Load("Select * from vOrdenesCorte", _
                    "TemporadaID = " & GescomMain.mobjParametro.TemporadaActualID & " AND " & _
                    "EmpresaID = " & GescomMain.mobjParametro.EmpresaActualID)
        .Show
    End With
    Set objRecordList = Nothing

    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuFabrOrdenCorteNew_Click()
    Dim objOrdenCorte As OrdenCorte
    Dim frmOrdenCorte As OrdenCorteEdit
  
    Set objOrdenCorte = New OrdenCorte
    Set frmOrdenCorte = New OrdenCorteEdit
  
    frmOrdenCorte.Component objOrdenCorte
    frmOrdenCorte.Show
    
End Sub

Private Sub mnuListViewSearch_Click()

    If GescomMain.ActiveForm Is Nothing Then Exit Sub
    If Not GescomMain.ActiveForm.IsList Then Exit Sub
    GescomMain.ActiveForm.ResultSearch
    
End Sub

Private Sub mnuListViewQuickSearch_Click()

    If GescomMain.ActiveForm Is Nothing Then Exit Sub
    If Not GescomMain.ActiveForm.IsList Then Exit Sub
    GescomMain.ActiveForm.QuickSearch

End Sub

Private Sub mnuVentasAlbaranesList_Click()
    Dim frmList As AlbaranVentaList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New AlbaranVentaList
    Set objRecordList = New RecordList

    With frmList
        .ComponentStatus objRecordList.Load("Select * from vAlbaranesVenta", _
                        "TemporadaID = " & GescomMain.mobjParametro.TemporadaActualID & " AND " & _
                        "EmpresaID = " & GescomMain.mobjParametro.EmpresaActualID)

        .Show
    End With

    Set objRecordList = Nothing


    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuVentasAlbaranesNew_Click()
    Dim objAlbaranVenta As AlbaranVenta
    Dim frmAlbaranVenta As AlbaranVentaEdit
  
    Set objAlbaranVenta = New AlbaranVenta
    Set frmAlbaranVenta = New AlbaranVentaEdit
  
    frmAlbaranVenta.Component objAlbaranVenta
    frmAlbaranVenta.Show
  
End Sub

Private Sub mnuAlbaranAutomaticoNew_Click()
    Dim objAlbaranVenta As AlbaranVenta
    Dim frmAlbaranVenta As AlbaranAutomaticoEdit
  
    Set objAlbaranVenta = New AlbaranVenta
    Set frmAlbaranVenta = New AlbaranAutomaticoEdit
  
    frmAlbaranVenta.Component objAlbaranVenta
    frmAlbaranVenta.Show
  
End Sub

Private Sub mnuVentasEstadisticas_Click()
    Dim frmList As EstadisticaVentaList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New EstadisticaVentaList
    Set objRecordList = New RecordList
    With frmList
'        .ComponentStatus objRecordList.Load("Select * from vEstadisticaVenta", "1=2")
        .Show

    End With
    Set objRecordList = Nothing
    
    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuVentasEstadisticaTiendas_Click()
    Dim frmList As VentasTiendaProveedor
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New VentasTiendaProveedor
    Set objRecordList = New RecordList
    With frmList
'        .ComponentStatus objRecordList.Load("Select * from vEstadisticaVenta", "1=2")
        .Show

    End With
    Set objRecordList = Nothing
    
    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuPedidosPendientes_Click()
    Dim frmList As PedidoVentaItemList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New PedidoVentaItemList
    Set objRecordList = New RecordList
    With frmList
        .Show

    End With
    Set objRecordList = Nothing
    
    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuVentasFacturasList_Click()
    Dim frmList As FacturaVentaList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New FacturaVentaList
    Set objRecordList = New RecordList
    With frmList
        .ComponentStatus objRecordList.Load("Select * from vFacturasVenta", _
                        "Anio = " & Year(Date) & " AND " & _
                        "EmpresaID = " & GescomMain.mobjParametro.EmpresaActualID)
'                        "TemporadaID = " & GescomMain.mobjParametro.TemporadaActualID & " AND " &
        .Show

    End With
    Set objRecordList = Nothing
    
    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuVentasFacturasNew_Click()
    Dim objFacturaVenta As FacturaVenta
    Dim frmFacturaVenta As FacturaVentaEdit
  
    Set objFacturaVenta = New FacturaVenta
    Set frmFacturaVenta = New FacturaVentaEdit
  
    frmFacturaVenta.Component objFacturaVenta
    frmFacturaVenta.Show
  
End Sub

Private Sub mnuVentasArticulosArticuloColorList_Click()
    Dim frmList As ArticuloColorList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New ArticuloColorList
    Set objRecordList = New RecordList

    With frmList
        .ComponentStatus objRecordList.Load("SELECT * FROM ArticuloColores", _
                                            "TemporadaID = " & GescomMain.mobjParametro.TemporadaActualID)
        .Show

    End With

    Set objRecordList = Nothing

    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuVentasClientesList_Click()
    Dim frmList As ClienteList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New ClienteList
    Set objRecordList = New RecordList

    With frmList
        .ComponentStatus objRecordList.Load("Select * from vClientes", vbNullString)
        .Show

    End With

    Set objRecordList = Nothing

    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuVentasCobros_Click()
    Dim frmList As CobroPagoList

    On Error GoTo ErrorManager

    Set frmList = New CobroPagoList

    With frmList
        .ComponentQuery "vCobros", _
                        GescomMain.mobjParametro.EmpresaActualID, _
                        "Lista de Cobros", GCEuro.gcTipoCobro
        .Show

    End With

    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub


Private Sub mnuVentasPedidoSinCorte_Click()
    Dim frmList As PedidoCorteList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New PedidoCorteList
    Set objRecordList = New RecordList

    With frmList
        .ComponentStatus objRecordList.Load("SELECT * FROM vPedidoVentaSinCorte ", _
                        "TemporadaID = " & GescomMain.mobjParametro.TemporadaActualID & _
                        " AND EmpresaID = " & GescomMain.mobjParametro.EmpresaActualID)
        .Show

    End With

    Set objRecordList = Nothing
    
    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)

End Sub

'Private Sub mnuVentasCobros_Click()
'    Dim frmList As CobroPagoList
'    Dim frmCobroPago As CobroPagoEdit
'    Dim objRecordList As RecordList
'    Dim objCobroPago As CobroPago
'
'    On Error GoTo ErrorManager
'
'    Set frmList = New CobroPagoList
'    Set objRecordList = New RecordList
'
'    With frmList
'        .ComponentStatus objRecordList.Load("Select * from vCobros", _
'                           "EmpresaID=" & GescomMain.mobjParametro.EmpresaActualID)
'        .Caption = "Lista de Cobros"
'        .Show
'
'    End With
'
'    Set objRecordList = Nothing
'
'    Exit Sub
'
'ErrorManager:
'    ManageErrors (Me.Caption)
'
'End Sub
'
Private Sub mnuVentasPedidoVentaList_Click()
    Dim frmList As PedidoVentaList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New PedidoVentaList
    Set objRecordList = New RecordList

    With frmList
        .ComponentStatus objRecordList.Load("SELECT * FROM vPedidosVenta ", _
                        "TemporadaID = " & GescomMain.mobjParametro.TemporadaActualID & _
                        " AND EmpresaID = " & GescomMain.mobjParametro.EmpresaActualID)
        .Show

    End With

    Set objRecordList = Nothing
    
    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

'Private Sub mnuVentasPrendaList_Click()
'    Dim frmList As PrendaList
'    Dim objRecordList As RecordList
'
'    On Error GoTo ErrorManager
'
'    Set frmList = New PrendaList
'    Set objRecordList = New RecordList
'    With frmList
'        .ComponentStatus objRecordList.Load("Select * From Prendas", vbNullString)
'        .Show
'
'    End With
'
'    Set objRecordList = Nothing
'    Exit Sub
'
'ErrorManager:
'    ManageErrors (Me.Caption)
'End Sub
'
Private Sub mnuVentasPrendaList_Click()
    
    On Error GoTo ErrorManager
    
    mobjGescomApp.OpenForm "uscPrendas", "uscPrendas"

    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuFileTemporadaNew_Click()
  
    Dim objTemporada As Temporada
    Dim frmTemporada As TemporadaEdit
  
    Set objTemporada = New Temporada
    Set frmTemporada = New TemporadaEdit
      
    frmTemporada.Component objTemporada
    frmTemporada.Show
    
End Sub

Private Sub mnuFileTemporadaList_Click()
    Dim frmList As TemporadaList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New TemporadaList
    Set objRecordList = New RecordList
    With frmList
        .ComponentStatus objRecordList.Load("SELECT * FROM Temporadas", vbNullString)
        .Show
  
    End With

    Set objRecordList = Nothing

    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuFileEmpresaNew_Click()
    
    Dim objEmpresa As Empresa
    Dim frmEmpresa As EmpresaEdit
  
    Set objEmpresa = New Empresa
    Set frmEmpresa = New EmpresaEdit
  
    frmEmpresa.Component objEmpresa
    frmEmpresa.Show
    
End Sub

Private Sub mnuFileEmpresaList_Click()
    Dim frmList As EmpresaList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New EmpresaList
    Set objRecordList = New RecordList

    With frmList
        .ComponentStatus objRecordList.Load("Select * from Empresas", vbNullString)
        .Show

    End With

    Set objRecordList = Nothing

    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuFileAsientoNew_Click()
    
    Dim objAsiento As Asiento
    Dim frmAsiento As AsientoEdit
  
    Set objAsiento = New Asiento
    Set frmAsiento = New AsientoEdit
  
    frmAsiento.Component objAsiento
    frmAsiento.Show
    
End Sub

Private Sub mnuFileAsientoList_Click()
    Dim frmList As AsientoList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New AsientoList
    Set objRecordList = New RecordList

    With frmList
        .ComponentStatus objRecordList.Load("Select * from vAsientosPendientes", _
                        "EmpresaID = " & GescomMain.mobjParametro.EmpresaActualID)
        .Show

    End With

    Set objRecordList = Nothing

    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuComprasMaterialNew_Click()
    
    Dim objMaterial As Material
    Dim frmMaterial As MaterialEdit
 
    Set objMaterial = New Material
    Set frmMaterial = New MaterialEdit
 
    frmMaterial.Component objMaterial
    frmMaterial.Show
    
End Sub

Private Sub mnuComprasMaterialList_Click()
    Dim frmList As MaterialList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New MaterialList
    Set objRecordList = New RecordList

    With frmList
        .ComponentStatus objRecordList.Load("Select * from Materiales", vbNullString)
        .Show

    End With

    Set objRecordList = Nothing
    
    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuComprasProveedoresList_Click()
    Dim frmList As ProveedorList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New ProveedorList
    Set objRecordList = New RecordList

    With frmList
        .ComponentStatus objRecordList.Load("Select * from vProveedores", vbNullString)
        .Show

    End With

    Set objRecordList = Nothing

    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuComprasProveedoresNew_Click()

    Dim objProveedor As Proveedor
    Dim frmProveedor As ProveedorEdit
  
    Set objProveedor = New Proveedor
    Set frmProveedor = New ProveedorEdit
  
    frmProveedor.Component objProveedor
    frmProveedor.Show
    
End Sub

Private Sub mnuVentasPrendaNew_Click()
  
    Dim objprenda As Prenda
    Dim frmprenda As PrendaEdit
  
    Set objprenda = New Prenda
    Set frmprenda = New PrendaEdit
  
    frmprenda.Component objprenda
    frmprenda.Show
    
End Sub

Private Sub mnuVentasPedidoVentaNew_Click()
  
    Dim objPedidoVenta As PedidoVenta
    Dim frmPedidoVenta As PedidoVentaEdit
  
    Set objPedidoVenta = New PedidoVenta
    Set frmPedidoVenta = New PedidoVentaEdit
      
    frmPedidoVenta.Component objPedidoVenta
    frmPedidoVenta.Show
    
End Sub

Private Sub mnuVentasClienteNew_Click()
    
    Dim objCliente As Cliente
    Dim frmCliente As ClienteEdit
  
    Set objCliente = New Cliente
    Set frmCliente = New ClienteEdit
      
    frmCliente.Component objCliente
    frmCliente.Show
    
End Sub

Private Sub mnuVentasArticulosArticuloList_Click()
    Dim frmList As ArticuloList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New ArticuloList
    Set objRecordList = New RecordList

    With frmList
        .ComponentStatus objRecordList.Load("Select * from vArticulos", _
                            "TemporadaID = " & GescomMain.mobjParametro.TemporadaActualID)
        .Show

    End With

    Set objRecordList = Nothing

    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuVentasBancosList_Click()
    Dim frmList As BancoList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New BancoList
    Set objRecordList = New RecordList

    With frmList
        .ComponentStatus objRecordList.Load("Select * from vBancos", vbNullString)
        .Show

    End With

    Set objRecordList = Nothing

    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuVentasBancosNew_Click()

    Dim objBanco As Banco
    Dim frmBanco As BancoEdit
  
    Set objBanco = New Banco
    Set frmBanco = New BancoEdit
  
    frmBanco.Component objBanco
    frmBanco.Show
    
End Sub

'Private Sub mnuVentasClientesNew_Click()
'
'    Dim objCliente As Cliente
'    Dim frmCliente As ClienteEdit
'
'    Set objCliente = New Cliente
'    Set frmCliente = New ClienteEdit
'
'    frmCliente.Component objCliente
'    frmCliente.Show
'
'End Sub
'
Private Sub mnuVentasRemesas_Click()
    Dim frmList As RemesaList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New RemesaList
    
    Set objRecordList = New RecordList

    With frmList
        .ComponentStatus objRecordList.Load("Select * from vRemesas", vbNullString)
        .Show

    End With

    Set objRecordList = Nothing

    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuVentasRepresentantesList_Click()
    Dim frmList As RepresentanteList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New RepresentanteList

    Set objRecordList = New RecordList

    With frmList
        .ComponentStatus objRecordList.Load("Select * from Representantes", vbNullString)
        .Show

    End With

    Set objRecordList = Nothing

    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuVentasRepresentantesNew_Click()

    Dim objRepresentante As Representante
    Dim frmRepresentante As RepresentanteEdit
  
    Set objRepresentante = New Representante
    Set frmRepresentante = New RepresentanteEdit
  
    frmRepresentante.Component objRepresentante
    frmRepresentante.Show
    
End Sub

Private Sub mnuVentasTransportistasList_Click()
    Dim frmList As TransportistaList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New TransportistaList
    Set objRecordList = New RecordList

    With frmList
        .ComponentStatus objRecordList.Load("Select * from Transportistas", vbNullString)
        .Show

    End With

    Set objRecordList = Nothing

    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuVentasTransportistasNew_Click()

    Dim objTransportista As Transportista
    Dim frmTransportista As TransportistaEdit
  
    Set objTransportista = New Transportista
    Set frmTransportista = New TransportistaEdit
  
    frmTransportista.Component objTransportista
    frmTransportista.Show
    
End Sub

Private Sub mnuFileSalir_Click()
  
    Unload Me
    
End Sub

Private Sub mnuUtilityContext_Click()
    
    Dim Result As VbMsgBoxResult
    Dim frmContext As ContextEdit
  
    If Not (GescomMain.ActiveForm Is Nothing) Then
        Result = MsgBox("Cierre todas las ventanas para cambiar de empresa y temporada", vbInformation + vbOKOnly)
        Exit Sub
    End If
    Set frmContext = New ContextEdit
  
    frmContext.Component mobjParametro
    frmContext.Show vbModal
    ' pongo el titulo de la aplicacion porque puede haber cambiado
    GescomTitulo

End Sub

Private Sub mnuUtilityParam_Click()
    
    Dim frmParametro As ParametroEdit
  
    Set frmParametro = New ParametroEdit
  
    frmParametro.Component mobjParametro
    frmParametro.Show vbModal

End Sub

Private Sub mnuVentasSerieNew_Click()
    
    Dim objSerie As Serie
    Dim frmSerie As SerieEdit
 
    Set objSerie = New Serie
    Set frmSerie = New SerieEdit
 
    objSerie.TemporadaID = GescomMain.mobjParametro.TemporadaActualID
    frmSerie.Component objSerie
    frmSerie.Show
    
End Sub

Private Sub mnuVentasSerieList_Click()
    Dim frmList As SerieList
    Dim objRecordList As RecordList
    
    On Error GoTo ErrorManager
    
    Set frmList = New SerieList
    Set objRecordList = New RecordList

    With frmList
        .ComponentStatus objRecordList.Load("SELECT * FROM vSeriesMateriales", _
                        "TemporadaID = " & GescomMain.mobjParametro.TemporadaActualID)
            
        .Show

    End With

    Set objRecordList = Nothing

    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuVentasModeloNew_Click()
    
    Dim objModelo As Modelo
    Dim frmModelo As ModeloEdit
 
    Set objModelo = New Modelo
    Set frmModelo = New ModeloEdit
 
    objModelo.TemporadaID = GescomMain.mobjParametro.TemporadaActualID
    frmModelo.Component objModelo
    frmModelo.Show
    
End Sub

Private Sub mnuVentasModeloList_Click()
    Dim frmList As ModeloList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New ModeloList
    Set objRecordList = New RecordList

    With frmList
        .ComponentStatus objRecordList.Load("Select * from Modelos", _
                        "TemporadaID = " & GescomMain.mobjParametro.TemporadaActualID)
        .Show

    End With

    Set objRecordList = Nothing

    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuVentasArticulosArticuloNew_Click()
    Dim objArticulo As Articulo
    Dim frmArticulo As ArticuloEdit
  
    On Error GoTo ErrorManager

    Set objArticulo = New Articulo
    Set frmArticulo = New ArticuloEdit
    frmArticulo.Component objArticulo
    frmArticulo.Show
    
    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub mnuVentasArticulosArticuloColorNew_Click()
    Dim objArticuloColor As ArticuloColor
    Dim frmArticuloColor As ArticuloColorEdit
  
    Set objArticuloColor = New ArticuloColor
    Set frmArticuloColor = New ArticuloColorEdit
    
    objArticuloColor.TemporadaID = GescomMain.mobjParametro.TemporadaActualID
    
    frmArticuloColor.Component objArticuloColor
    frmArticuloColor.Show

End Sub

Private Sub mnuListviewDel_Click()
   
    If GescomMain.ActiveForm Is Nothing Then Exit Sub
    If Not GescomMain.ActiveForm.IsList Then Exit Sub
    GescomMain.ActiveForm.DeleteSelected

End Sub

Private Sub mnuListViewEdit_Click()
   
    If GescomMain.ActiveForm Is Nothing Then Exit Sub
    If Not GescomMain.ActiveForm.IsList Then Exit Sub
    GescomMain.ActiveForm.EditSelected
   
End Sub

Private Sub mnuListViewNew_Click()
  
    If GescomMain.ActiveForm Is Nothing Then Exit Sub
    If Not GescomMain.ActiveForm.IsList Then Exit Sub
    GescomMain.ActiveForm.NewObject
  
End Sub

Public Sub GescomTerminal()
    Dim lngTerminalID As Long
    Dim strTerminalID As String
        
    WinIniRegister GesComSectionName
    strTerminalID = WinGetString("TerminalID", vbNullString)
    
    If strTerminalID = vbNullString Then
        ' ojoojo: Abrir el formulario de selección de terminal.
        Exit Sub
    End If
    
    ' ojoojo: si el terminal no existe, debería dejarse en un log...
    lngTerminalID = CLng(strTerminalID)
    mobjTerminal.Load lngTerminalID
    
End Sub

Public Sub GescomTitulo()
    
    GescomMain.Caption = "Sistema de Gestión Comercial   " & _
        "[" & Trim(mobjParametro.EmpresaActual) & "] - " & _
        "[" & Trim(mobjParametro.TemporadaActual) & "] - " & _
        "[" & mobjParametro.Moneda & "]"

End Sub

Private Sub GescomApp()
    
    Set mobjGescomApp = New clsMDIParentForm
    mobjGescomApp.GetUserSettings

End Sub

Private Sub mnuVentasTraspasos_Click()
    Dim frmList As TraspasoList
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager

    Set frmList = New TraspasoList
    Set objRecordList = New RecordList

    'ojoojo: la clausula where puede ser que se tenga en cuenta el almacén origen y el almacen destino
    With frmList
        .ComponentStatus objRecordList.Load("Select * from vTraspasos", vbNullString)
        .Show

    End With

    Set objRecordList = Nothing

    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub tlbCompras_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case Is = "Material"
            Call mnuComprasMaterialList_Click
        Case Is = "Proveedor"
            Call mnuComprasProveedoresList_Click
        Case Is = "PedidoCompra"
            Call mnuComprasPedidoCompraList_Click
        Case Is = "AlbaranCompra"
            Call mnuComprasAlbaranesList_Click
        Case Is = "FacturaCompra"
            Call mnuComprasFacturasList_Click
        Case Is = "Pago"
            Call mnuComprasPagos_Click
        Case Is = "MoviMateriales"
            Call mnuComprasMoviMateriales_Click
        Case Is = "ListaPagos"
            Call mnuComprasPagos_Click
    End Select
    
End Sub

Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case Is = "Temporada"
            Call mnuFileTemporadaList_Click
        Case Is = "Empresa"
            Call mnuFileEmpresaList_Click
        Case Is = "Parametro"
            Call mnuUtilityParam_Click
        Case Is = "Context"
            Call mnuUtilityContext_Click
        Case Is = "Cerrar"
            Call mnuFileSalir_Click
        
    End Select
        
End Sub

Private Sub tlbVentas_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case Is = "Banco"
            Call mnuVentasBancosList_Click
        Case Is = "Transportista"
            Call mnuVentasTransportistasList_Click
        Case Is = "Representante"
            Call mnuVentasRepresentantesList_Click
        Case Is = "Cliente"
            Call mnuVentasClientesList_Click
        Case Is = "Prenda"
            Call mnuVentasPrendaList_Click
        Case Is = "Serie"
            Call mnuVentasSerieList_Click
        Case Is = "Modelo"
            Call mnuVentasModeloList_Click
        Case Is = "Articulo"
            Call mnuVentasArticulosArticuloList_Click
        Case Is = "ArticuloColor"
            Call mnuVentasArticulosArticuloColorList_Click
        Case Is = "PedidoVenta"
            Call mnuVentasPedidoVentaList_Click
        Case Is = "AlbaranVenta"
            Call mnuVentasAlbaranesList_Click
        Case Is = "FacturaVenta"
            Call mnuVentasFacturasList_Click
        Case Is = "ListaCobros"
            Call mnuVentasCobros_Click
        Case Is = "Remesa"
            Call mnuVentasRemesas_Click
        Case Is = "AlbaranAutomatico"
            Call mnuAlbaranAutomaticoNew_Click
        Case Is = "EstadisticaVentas"
            Call mnuVentasEstadisticas_Click
        Case Is = "EstadisticaVentaTiendas"
            Call mnuVentasEstadisticaTiendas_Click
        Case Is = "PedidosPendientes"
            Call mnuPedidosPendientes_Click
        Case Is = "Traspasos"
            Call mnuVentasTraspasos_Click
    End Select
    
End Sub
Private Sub tlbFabricacion_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case Is = "OrdenCorte"
            Call mnuFabrOrdenCorteList_Click
        Case Is = "Etiquetas"
            Call mnuFabrEtiquetas_Click
    End Select
    
End Sub

