VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm TPVMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Sistema de Gestión Comercial"
   ClientHeight    =   7560
   ClientLeft      =   165
   ClientTop       =   630
   ClientWidth     =   10185
   Icon            =   "TPVMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList mglIconosGrandes 
      Left            =   600
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":08CA
            Key             =   "GesComTPV"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrHerramientas 
      Align           =   1  'Align Top
      Height          =   1080
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   1905
      BandCount       =   1
      BandBorders     =   0   'False
      VariantHeight   =   0   'False
      _CBWidth        =   10185
      _CBHeight       =   1080
      _Version        =   "6.7.9782"
      Child1          =   "tlbPrincipal"
      MinWidth1       =   2505
      MinHeight1      =   1020
      Width1          =   1005
      UseCoolbarPicture1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tlbPrincipal 
         Height          =   1020
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   10065
         _ExtentX        =   17754
         _ExtentY        =   1799
         ButtonWidth     =   1455
         ButtonHeight    =   1799
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "mglIconosGrandes"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Temporada"
               Object.ToolTipText     =   "Temporadas"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Context"
               Object.ToolTipText     =   "Cambiar empresa y temporada"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cerrar"
               Object.ToolTipText     =   "Salir"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "TPV"
               Key             =   "TPV"
               Description     =   "TPV"
               Object.ToolTipText     =   "TPV"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
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
         NumListImages   =   55
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":1103
            Key             =   "IconosPequeños"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":1217
            Key             =   "FacturaVenta"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":2269
            Key             =   "FacturaCompra"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":32BB
            Key             =   "ArticuloColor"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":410F
            Key             =   "Context"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":4567
            Key             =   "Parametro"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":4883
            Key             =   "Documento"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":4CDF
            Key             =   "AlbaranCompra"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":5D33
            Key             =   "EliminarItem"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":6187
            Key             =   "ModificarItem"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":65DB
            Key             =   "NuevoItem"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":6A2F
            Key             =   "Proveedor"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":770B
            Key             =   "Material"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":855F
            Key             =   "PedidoCompra"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":95B3
            Key             =   "PedidoVenta"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":A607
            Key             =   "AlbaranVenta"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":B65B
            Key             =   "Articulo"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":C4AF
            Key             =   "Modelo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":D303
            Key             =   "Serie"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":D757
            Key             =   "Prenda"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":E7AB
            Key             =   "Cliente"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":F487
            Key             =   "Representante"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":F7A3
            Key             =   "Transportista"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":1007F
            Key             =   "Banco"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":10967
            Key             =   "Empresa"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":10DBB
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":110D7
            Key             =   "Nuevo"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":111EB
            Key             =   "Abrir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":112FF
            Key             =   "Detalle"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":11413
            Key             =   "Lista"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":11527
            Key             =   "IconosGrandes"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":1163B
            Key             =   "Eliminar"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":1174F
            Key             =   "Actualizar"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":11C93
            Key             =   "Temporada"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":120E7
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":12403
            Key             =   "Corte"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":1271D
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":14ECF
            Key             =   "Etiqueta"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":151E9
            Key             =   "Cobrar"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":1563B
            Key             =   "PrevEtiqueta"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":15955
            Key             =   "OLAPQuery"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":15C6F
            Key             =   "Remesa"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":15F89
            Key             =   "PrintDocument"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":162A3
            Key             =   "CobroPago"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":165BD
            Key             =   "Recalcular"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":168D7
            Key             =   "Contabilidad"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":171B1
            Key             =   "Contawin"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":174CB
            Key             =   "Cobrados"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":177E5
            Key             =   "Movimientos"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":180C7
            Key             =   "BarCode"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":18519
            Key             =   "GroupBy"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":18673
            Key             =   "Propiedades"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":187CD
            Key             =   "Ordenar"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":18927
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TPVMain.frx":19C31
            Key             =   "PVP"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuFileTemporada 
         Caption         =   "&Temporadas"
         Begin VB.Menu mnuFileTemporadaList 
            Caption         =   "&Temporadas"
         End
         Begin VB.Menu mnuFileEmpresaLine1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileTemporadaNew 
            Caption         =   "&Nueva Temporada"
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
      Begin VB.Menu mnuVentasAlbaranes 
         Caption         =   "&TPV"
         Begin VB.Menu mnuTPVList 
            Caption         =   "&Lista de ventas"
         End
         Begin VB.Menu mnuVentasAlbaranesSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTPVNew 
            Caption         =   "&TPV"
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
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "&Utilidades"
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
Attribute VB_Name = "TPVMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GesComSectionName = "GesCom"

Private mobjParametro As Parametro
Private mobjTerminal As Terminal

Public Property Get Parametro()
    Set Parametro = mobjParametro
End Property

Public Property Get Terminal()
    Set Terminal = mobjTerminal
End Property

Private Sub MDIForm_Load()

    On Error GoTo ErrorManager
    
    Set mobjParametro = New Parametro
    mobjParametro.Load
    
    Set mobjTerminal = New Terminal
  
    GescomTerminal
    GescomTitulo
    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
    TerminateProgram
    
End Sub


Private Sub mnuListViewSearch_Click()

    If TPVMain.ActiveForm Is Nothing Then Exit Sub
    If Not TPVMain.ActiveForm.IsList Then Exit Sub
    TPVMain.ActiveForm.ResultSearch
    
End Sub

Private Sub mnuListViewQuickSearch_Click()

    If TPVMain.ActiveForm Is Nothing Then Exit Sub
    If Not TPVMain.ActiveForm.IsList Then Exit Sub
    TPVMain.ActiveForm.QuickSearch

End Sub

'
Private Sub mnuFileTemporadaList_Click()
'    Dim frmList As TemporadaList
'    Dim objRecordList As RecordList
'
'    On Error GoTo ErrorManager
'
'    Set frmList = New TemporadaList
'    Set objRecordList = New RecordList
'    With frmList
'        .ComponentStatus objRecordList.Load("SELECT * FROM Temporadas", vbNullString)
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
End Sub

Private Sub mnuTPVList_Click()
'    Dim frmList As AlbaranVentaList
'    Dim objRecordList As RecordList
'
'    On Error GoTo ErrorManager
'
'    Set frmList = New AlbaranVentaList
'    Set objRecordList = New RecordList
'
'    With frmList
'        .ComponentStatus objRecordList.Load("Select * from vAlbaranesVenta", _
'                        "TemporadaID = " & TPVMain.objParametro.TemporadaActualID & " AND " & _
'                        "EmpresaID = " & TPVMain.objParametro.EmpresaActualID)
'
'        .Show
'    End With
'
'    Set objRecordList = Nothing
'
'
'    Exit Sub
'
'ErrorManager:
'    ManageErrors (Me.Caption)
'
End Sub

Private Sub mnuTPVNew_Click()
    Dim objAlbaranVenta As AlbaranVenta
    Dim frmAlbaranVenta As TPVEdit
  
    Set objAlbaranVenta = New AlbaranVenta
    Set frmAlbaranVenta = New TPVEdit
  
    frmAlbaranVenta.Component objAlbaranVenta
    frmAlbaranVenta.Show
  
End Sub

Private Sub mnuFileSalir_Click()
  
    Unload Me
    
End Sub

Private Sub mnuUtilityContext_Click()
    
    Dim Result As VbMsgBoxResult
    Dim frmContext As ContextEdit
  
    If Not (TPVMain.ActiveForm Is Nothing) Then
        Result = MsgBox("Cierre todas las ventanas para cambiar de empresa y temporada", vbInformation + vbOKOnly)
        Exit Sub
    End If
    Set frmContext = New ContextEdit
  
    frmContext.Component mobjParametro
    frmContext.Show vbModal
    ' pongo el titulo de la aplicacion porque puede haber cambiado
    GescomTitulo

End Sub

Private Sub mnuListviewDel_Click()
   
    If TPVMain.ActiveForm Is Nothing Then Exit Sub
    If Not TPVMain.ActiveForm.IsList Then Exit Sub
    TPVMain.ActiveForm.DeleteSelected

End Sub

Private Sub mnuListViewEdit_Click()
   
    If TPVMain.ActiveForm Is Nothing Then Exit Sub
    If Not TPVMain.ActiveForm.IsList Then Exit Sub
    TPVMain.ActiveForm.EditSelected
   
End Sub

Private Sub mnuListViewNew_Click()
  
    If TPVMain.ActiveForm Is Nothing Then Exit Sub
    If Not TPVMain.ActiveForm.IsList Then Exit Sub
    TPVMain.ActiveForm.NewObject
  
End Sub

Public Sub GescomTitulo()
    
    TPVMain.Caption = "Sistema de Gestión Comercial   " & _
        "[" & Trim(mobjParametro.EmpresaActual) & "] - " & _
        "[" & Trim(mobjParametro.TemporadaActual) & "] - " & _
        "[" & mobjTerminal.Nombre & "]"
'        "[" & mobjParametro.Moneda & "]"

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
    
    lngTerminalID = CLng(strTerminalID)
    mobjTerminal.Load lngTerminalID
    
    
End Sub

Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case Is = "Temporada"
            Call mnuFileTemporadaList_Click
        Case Is = "Context"
            Call mnuUtilityContext_Click
        Case Is = "Cerrar"
            Call mnuFileSalir_Click
        Case Is = "TPV"
            Call mnuTPVNew_Click
        
        
    End Select
        
End Sub

