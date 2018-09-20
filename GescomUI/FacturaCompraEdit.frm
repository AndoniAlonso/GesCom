VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FacturaCompraEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas de Compra"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10335
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FacturaCompraEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAlbaranes 
      Caption         =   "Albara&nes..."
      Height          =   375
      Left            =   240
      TabIndex        =   37
      ToolTipText     =   "Incorporar pedidos pendientes"
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la Factura"
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin VB.TextBox txtSufijo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8880
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox cboMedioPago 
         Height          =   315
         Left            =   1440
         TabIndex        =   29
         Text            =   "cboFormaPago"
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7920
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cboProveedor 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Text            =   "cboProveedor"
         Top             =   340
         Width           =   3735
      End
      Begin VB.TextBox txtNuestraReferencia 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   1060
         Width           =   1935
      End
      Begin VB.TextBox txtSuReferencia 
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Top             =   1420
         Width           =   1935
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   285
         Left            =   1440
         TabIndex        =   26
         Top             =   2145
         Width           =   3735
      End
      Begin VB.ComboBox cboBanco 
         Height          =   315
         Left            =   6840
         TabIndex        =   11
         Text            =   "cboBanco"
         Top             =   720
         Width           =   2775
      End
      Begin VB.ComboBox cboTransportista 
         Height          =   315
         Left            =   6840
         TabIndex        =   15
         Text            =   "cboTransportista"
         Top             =   1060
         Width           =   2775
      End
      Begin VB.ComboBox cboFormaPago 
         Height          =   315
         Left            =   6840
         TabIndex        =   19
         Text            =   "cboFormaPago"
         Top             =   1420
         Width           =   2775
      End
      Begin VB.CommandButton btnDatoComercial 
         Caption         =   "Datos C&omerciales"
         Height          =   615
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtDatoComercial 
         Height          =   990
         Left            =   6840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   1780
         Width           =   2775
      End
      Begin VB.TextBox txtPortes 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         TabIndex        =   23
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtEmbalajes 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   21
         Top             =   1785
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   705
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58064897
         CurrentDate     =   36938
      End
      Begin MSComCtl2.DTPicker dtpFechaContable 
         Height          =   315
         Left            =   4080
         TabIndex        =   10
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58064897
         CurrentDate     =   36938
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Medio de Pago"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   2535
         Width           =   1050
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Contable"
         Height          =   195
         Left            =   2880
         TabIndex        =   9
         Top             =   735
         Width           =   1125
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Left            =   5520
         TabIndex        =   8
         Top             =   720
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   7080
         TabIndex        =   3
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "N/Referencia"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Su Referencia"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   2160
         Width           =   1065
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Transportista"
         Height          =   195
         Left            =   5520
         TabIndex        =   14
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pago"
         Height          =   195
         Left            =   5520
         TabIndex        =   18
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Portes"
         Height          =   195
         Left            =   2640
         TabIndex        =   22
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Embalajes"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   1800
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Líneas de la Factura de Compra"
      Height          =   3615
      Left            =   240
      TabIndex        =   30
      Top             =   3120
      Width           =   9855
      Begin VB.TextBox txtTotalBruto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   3225
         Width           =   1455
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "El&iminar"
         Height          =   375
         Left            =   8640
         TabIndex        =   35
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   7560
         TabIndex        =   34
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Aña&dir"
         Height          =   375
         Left            =   6480
         TabIndex        =   33
         Top             =   3120
         Width           =   975
      End
      Begin MSComctlLib.ListView lvwFacturaCompraItems 
         Height          =   2655
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Total Bruto de la Factura"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   3240
         Width           =   1785
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   9000
      TabIndex        =   40
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7800
      TabIndex        =   39
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6600
      TabIndex        =   38
      Top             =   6840
      Width           =   1095
   End
End
Attribute VB_Name = "FacturaCompraEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private mintProveedorSelStart As Integer
Private mintTransportistaSelStart As Integer
Private mintBancoSelStart As Integer
Private mintFormaPagoSelStart As Integer
Private mintMedioPagoSelStart As Integer

Private WithEvents mobjFacturaCompra As FacturaCompra
Attribute mobjFacturaCompra.VB_VarHelpID = -1

Public Sub Component(FacturaCompraObject As FacturaCompra)

    Set mobjFacturaCompra = FacturaCompraObject

End Sub

Private Sub btnDatoComercial_Click()

    Dim frmDatoComercial As DatoComercialEdit
  
    Set frmDatoComercial = New DatoComercialEdit
    frmDatoComercial.Component mobjFacturaCompra.DatoComercial
    frmDatoComercial.Show vbModal
    txtDatoComercial.Text = mobjFacturaCompra.DatoComercial.DatoComercialText
    
End Sub

Private Sub cboProveedor_Click()
    
    On Error GoTo ErrorManager
  
    If mflgLoading Then Exit Sub
    mobjFacturaCompra.Proveedor = cboProveedor.Text
  
    ' Al modificar el Proveedor se refrescan en el interface los datos relacionados.
    cboBanco.Text = mobjFacturaCompra.Banco
    cboBanco_Click
    cboTransportista.Text = mobjFacturaCompra.Transportista
    cboTransportista_Click
    cboFormaPago.Text = mobjFacturaCompra.FormaPago
    cboFormaPago_Click
    cboMedioPago.Text = mobjFacturaCompra.MedioPago
    cboMedioPago_Click
    txtDatoComercial.Text = mobjFacturaCompra.DatoComercial.DatoComercialText
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdApply_Click()
    
    On Error GoTo ErrorManager

    mobjFacturaCompra.ApplyEdit
    mobjFacturaCompra.BeginEdit GescomMain.objParametro.Moneda
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    If mobjFacturaCompra.Numero = GescomMain.objParametro.ObjEmpresaActual.FacturaCompras Then
        GescomMain.objParametro.ObjEmpresaActual.DecrementaFacturaCompras
    End If
  
    mobjFacturaCompra.CancelEdit
  
    Unload Me

End Sub

Private Sub cmdOK_Click()
    Dim frmResumenDatosFactura As FacturaCompraResEdit
    Dim intFacturaCompraID As Long

    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    If mobjFacturaCompra.IsNew Then
        mobjFacturaCompra.CalcularBruto
        mobjFacturaCompra.CrearPagos
    End If
    
    mobjFacturaCompra.ApplyEdit
    GescomMain.objParametro.ObjEmpresaActual.EstableceFacturaCompras (mobjFacturaCompra.Numero)
    
    intFacturaCompraID = mobjFacturaCompra.FacturaCompraID
    Set mobjFacturaCompra = Nothing
    Set mobjFacturaCompra = New FacturaCompra
    mobjFacturaCompra.Load intFacturaCompraID, GescomMain.objParametro.Moneda
  
    ' Aqui lanzo el formulario de edicion de los detalles de la factura.
    Set frmResumenDatosFactura = New FacturaCompraResEdit
  
    frmResumenDatosFactura.Component mobjFacturaCompra
    frmResumenDatosFactura.Show
  
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub dtpFecha_Change()
    
    mobjFacturaCompra.Fecha = dtpFecha.Value
    
End Sub

Private Sub dtpFechaContable_Change()
    
    mobjFacturaCompra.FechaContable = dtpFechaContable.Value
    
End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
    With mobjFacturaCompra
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Factura de Compra [(nueva)]"

        Else
            Caption = "Factura de Compra [" & .Proveedor & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtNumero.Text = .Numero
        txtSufijo.Text = .Sufijo
        dtpFecha.Value = .Fecha
        dtpFechaContable.Value = .FechaContable
        txtNuestraReferencia.Text = .NuestraReferencia
        txtSuReferencia.Text = .SuReferencia
        txtObservaciones.Text = .Observaciones
        txtEmbalajes.Text = .Embalajes
        txtPortes.Text = .Portes
        txtDatoComercial.Text = .DatoComercial.DatoComercialText
        txtTotalBruto.Text = FormatoMoneda(.TotalBruto, GescomMain.objParametro.Moneda)
        
        LoadCombo cboProveedor, .Proveedores
        cboProveedor.Text = .Proveedor
        
        LoadCombo cboBanco, .Bancos
        cboBanco.Text = .Banco
    
        LoadCombo cboTransportista, .Transportistas
        cboTransportista.Text = .Transportista
        
        LoadCombo cboFormaPago, .FormasPago
        cboFormaPago.Text = .FormaPago
    
        LoadCombo cboMedioPago, .MediosPago
        cboMedioPago.Text = .MedioPago
    
        .BeginEdit GescomMain.objParametro.Moneda
        
        If .IsNew Then
            .Numero = GescomMain.objParametro.ObjEmpresaActual.IncrementaFacturaCompras
            txtNumero = .Numero
           
            .TemporadaID = GescomMain.objParametro.TemporadaActualID
            .EmpresaID = GescomMain.objParametro.EmpresaActualID
        End If
    
    End With
  
    lvwFacturaCompraItems.SmallIcons = GescomMain.mglIconosPequeños
    
    lvwFacturaCompraItems.ColumnHeaders.Add , , vbNullString, ColumnSize(4)
    lvwFacturaCompraItems.ColumnHeaders.Add , , "Material", ColumnSize(25)
    lvwFacturaCompraItems.ColumnHeaders.Add , , "Cantidad", ColumnSize(6), vbRightJustify
    lvwFacturaCompraItems.ColumnHeaders.Add , , "Bruto", ColumnSize(10), vbRightJustify
    LoadFacturaCompraItems
      
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub lvwFacturaCompraItems_DblClick()
    
    Call cmdEdit_Click
    
End Sub

Private Sub mobjFacturaCompra_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub txtPortes_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtPortes

End Sub

Private Sub txtPortes_Change()

    If Not mflgLoading Then _
        TextChange txtPortes, mobjFacturaCompra, "Portes"

End Sub

Private Sub txtPortes_LostFocus()

    txtPortes = TextLostFocus(txtPortes, mobjFacturaCompra, "Portes")

End Sub

Private Sub txtEmbalajes_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtEmbalajes

End Sub

Private Sub txtEmbalajes_Change()

    If Not mflgLoading Then _
        TextChange txtEmbalajes, mobjFacturaCompra, "Embalajes"

End Sub

Private Sub txtEmbalajes_LostFocus()

    txtEmbalajes = TextLostFocus(txtEmbalajes, mobjFacturaCompra, "Embalajes")

End Sub

Private Sub txtObservaciones_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtObservaciones

End Sub

Private Sub txtObservaciones_Change()

    If Not mflgLoading Then _
        TextChange txtObservaciones, mobjFacturaCompra, "Observaciones"

End Sub

Private Sub txtObservaciones_LostFocus()

    txtObservaciones = TextLostFocus(txtObservaciones, mobjFacturaCompra, "Observaciones")

End Sub

Private Sub txtNumero_Change()

    If Not mflgLoading Then _
        TextChange txtNumero, mobjFacturaCompra, "Numero"

End Sub

Private Sub txtNumero_LostFocus()

    txtNumero = TextLostFocus(txtNumero, mobjFacturaCompra, "Numero")

End Sub

Private Sub txtNumero_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtNumero
        
End Sub

Private Sub txtSufijo_Change()

    If Not mflgLoading Then _
        TextChange txtSufijo, mobjFacturaCompra, "Sufijo"

End Sub

Private Sub txtSufijo_LostFocus()

    txtSufijo = TextLostFocus(txtSufijo, mobjFacturaCompra, "Sufijo")

End Sub

Private Sub txtSufijo_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtSufijo
        
End Sub

Private Sub cboTransportista_Click()

    If mflgLoading Then Exit Sub
    mobjFacturaCompra.Transportista = cboTransportista.Text

End Sub

Private Sub cboBanco_Click()

    If mflgLoading Then Exit Sub
    mobjFacturaCompra.Banco = cboBanco.Text

End Sub

Private Sub cboFormaPago_Click()

    If mflgLoading Then Exit Sub
    mobjFacturaCompra.FormaPago = cboFormaPago.Text

End Sub

Private Sub cboMedioPago_Click()

    If mflgLoading Then Exit Sub
    mobjFacturaCompra.MedioPago = cboMedioPago.Text

End Sub

Private Sub txtNuestraReferencia_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtNuestraReferencia

End Sub

Private Sub txtNuestraReferencia_Change()

    If Not mflgLoading Then _
        TextChange txtNuestraReferencia, mobjFacturaCompra, "NuestraReferencia"

End Sub

Private Sub txtNuestraReferencia_LostFocus()

    txtNuestraReferencia = TextLostFocus(txtNuestraReferencia, mobjFacturaCompra, "NuestraReferencia")

End Sub

Private Sub txtSuReferencia_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtSuReferencia

End Sub

Private Sub txtSuReferencia_Change()

    If Not mflgLoading Then _
        TextChange txtSuReferencia, mobjFacturaCompra, "SuReferencia"

End Sub

Private Sub txtSuReferencia_LostFocus()

    txtSuReferencia = TextLostFocus(txtSuReferencia, mobjFacturaCompra, "SuReferencia")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
    
    IsList = False
    
End Function

' a partir de aqui -----> child

Private Sub cmdAdd_Click()
  
    Dim frmFacturaCompraItem As FacturaCompraItemEdit
  
    On Error GoTo ErrorManager
    Set frmFacturaCompraItem = New FacturaCompraItemEdit
    frmFacturaCompraItem.Component mobjFacturaCompra.FacturaCompraItems.Add
    frmFacturaCompraItem.Show vbModal
    LoadFacturaCompraItems
    mobjFacturaCompra.CalcularBruto
    txtTotalBruto = FormatoMoneda(mobjFacturaCompra.TotalBruto, GescomMain.objParametro.Moneda)
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdEdit_Click()

    Dim frmFacturaCompraItem As FacturaCompraItemEdit
  
    On Error GoTo ErrorManager
    Set frmFacturaCompraItem = New FacturaCompraItemEdit
    frmFacturaCompraItem.Component _
        mobjFacturaCompra.FacturaCompraItems(Val(lvwFacturaCompraItems.SelectedItem.Key))
    frmFacturaCompraItem.Show vbModal
    LoadFacturaCompraItems
    mobjFacturaCompra.CalcularBruto
    txtTotalBruto = FormatoMoneda(mobjFacturaCompra.TotalBruto, GescomMain.objParametro.Moneda)
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdRemove_Click()
    
    mobjFacturaCompra.FacturaCompraItems.Remove Val(lvwFacturaCompraItems.SelectedItem.Key)
    LoadFacturaCompraItems
    mobjFacturaCompra.CalcularBruto
    txtTotalBruto = FormatoMoneda(mobjFacturaCompra.TotalBruto, GescomMain.objParametro.Moneda)
    
End Sub

Private Sub LoadFacturaCompraItems()
    Dim objFacturaCompraItem As FacturaCompraItem
    Dim itmList As ListItem
    Dim lngIndex As Long
  
    On Error GoTo ErrorManager
    lvwFacturaCompraItems.ListItems.Clear
    For lngIndex = 1 To mobjFacturaCompra.FacturaCompraItems.Count
        Set itmList = lvwFacturaCompraItems.ListItems.Add _
            (Key:=Format$(lngIndex) & "K")
        Set objFacturaCompraItem = mobjFacturaCompra.FacturaCompraItems(lngIndex)

        With itmList

            .SmallIcon = GescomMain.mglIconosPequeños.ListImages("NuevoItem").Key
    
            If objFacturaCompraItem.IsDeleted Then .SmallIcon = GescomMain.mglIconosPequeños.ListImages("EliminarItem").Key

            .SubItems(1) = Trim(objFacturaCompraItem.Material)
            .SubItems(2) = objFacturaCompraItem.Cantidad
            .SubItems(3) = FormatoMoneda(objFacturaCompraItem.Bruto, GescomMain.objParametro.Moneda)
        End With

    Next
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdAlbaranes_Click()
  
    Dim frmPicker As PickerList
    Dim objSelectedItems As PickerItems
    Dim objPickerItemDisplay As PickerItemDisplay
  
    On Error GoTo ErrorManager
      
    ' no mostrar nada si no se ha informado del proveedor
    If mobjFacturaCompra.ProveedorID = 0 Then Exit Sub
    
    Set frmPicker = New PickerList
  
    frmPicker.LoadData "vAlbaranCompraPendientes", mobjFacturaCompra.ProveedorID, _
        mobjFacturaCompra.EmpresaID, _
        mobjFacturaCompra.TemporadaID
    frmPicker.Show vbModal
    Set objSelectedItems = frmPicker.SelectedItems
    Unload frmPicker
  
    If objSelectedItems Is Nothing Then Exit Sub
  
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    For Each objPickerItemDisplay In objSelectedItems
        ' Primero hay que comprobar que no está ya seleccionado anteriormente
        If Not DocumentoSeleccionado(objPickerItemDisplay.DocumentoID) Then _
            FacturaDesdeAlbaran (objPickerItemDisplay.DocumentoID)
    Next
  
    Set frmPicker = Nothing
    Set objSelectedItems = Nothing
  
    LoadFacturaCompraItems
  
    ' Muestro el puntero normal
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
    Exit Sub

End Sub

Private Sub FacturaDesdeAlbaran(AlbaranItemID As Long)

    Dim objFacturaItem As FacturaCompraItem
    Dim objAlbaranItem As AlbaranCompraItem

    Set objAlbaranItem = New AlbaranCompraItem
    Set objFacturaItem = mobjFacturaCompra.FacturaCompraItems.Add
   
    objAlbaranItem.Load AlbaranItemID, GescomMain.objParametro.Moneda, ALBARANCOMPRAITEM_MATERIAL
     
    With objFacturaItem
        .BeginEdit GescomMain.objParametro.Moneda
        .FacturaDesdeAlbaran AlbaranItemID
        .ApplyEdit
    End With
   
    Set objFacturaItem = Nothing
    Set objAlbaranItem = Nothing

End Sub

Private Function DocumentoSeleccionado(DocumentoID As Long) As Boolean
    Dim objFacturaCompraItem As FacturaCompraItem
    
    ' Se trata de buscar si existe alguna referencia de ese documento en alguna linea de
    ' facturas y es nueva (no se ha actualizado).
    For Each objFacturaCompraItem In mobjFacturaCompra.FacturaCompraItems
        If objFacturaCompraItem.IsNew And _
           objFacturaCompraItem.AlbaranCompraItemID = DocumentoID Then
           DocumentoSeleccionado = True
           Exit Function
        End If
    Next
    
    DocumentoSeleccionado = False
    
    Set objFacturaCompraItem = Nothing
    
End Function

Private Sub cboTransportista_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintTransportistaSelStart = cboTransportista.SelStart
End Sub

Private Sub cboTransportista_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintTransportistaSelStart, cboTransportista
    
End Sub

Private Sub cboProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintProveedorSelStart = cboProveedor.SelStart
End Sub

Private Sub cboProveedor_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintProveedorSelStart, cboProveedor
    
End Sub

Private Sub cboFormaPago_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintFormaPagoSelStart = cboFormaPago.SelStart
End Sub

Private Sub cboFormaPago_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintFormaPagoSelStart, cboFormaPago
    
End Sub

Private Sub cboBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintBancoSelStart = cboBanco.SelStart
End Sub

Private Sub cboBanco_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintBancoSelStart, cboBanco
    
End Sub

Private Sub cboMedioPago_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintMedioPagoSelStart = cboMedioPago.SelStart
End Sub

Private Sub cboMedioPago_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintMedioPagoSelStart, cboMedioPago
    
End Sub


