VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PedidoCompraEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos de Compra"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10605
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PedidoCompraEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Pedido"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin VB.TextBox txtDatoComercial 
         Height          =   885
         Left            =   7080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   1680
         Width           =   3135
      End
      Begin VB.CommandButton btnDatoComercial 
         Caption         =   "Datos C&omerciales"
         Height          =   615
         Left            =   5760
         TabIndex        =   21
         Top             =   1680
         Width           =   1215
      End
      Begin VB.ComboBox cboFormaPago 
         Height          =   315
         Left            =   7080
         TabIndex        =   18
         Text            =   "cboFormaDePago"
         Top             =   1305
         Width           =   3135
      End
      Begin VB.ComboBox cboTransportista 
         Height          =   315
         Left            =   7080
         TabIndex        =   12
         Text            =   "cboTransportista"
         Top             =   945
         Width           =   3135
      End
      Begin VB.ComboBox cboBanco 
         Height          =   315
         Left            =   7080
         TabIndex        =   8
         Text            =   "cboBanco"
         Top             =   585
         Width           =   3135
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   285
         Left            =   1320
         TabIndex        =   20
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox txtSuReferencia 
         Height          =   285
         Left            =   4200
         TabIndex        =   16
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtNuestraReferencia 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   1305
         Width           =   1575
      End
      Begin VB.ComboBox cboProveedor 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Text            =   "cboProveedor"
         Top             =   225
         Width           =   4215
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9240
         TabIndex        =   4
         Top             =   225
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   585
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   23986177
         CurrentDate     =   36938
      End
      Begin MSComCtl2.DTPicker dtpFechaEntrega 
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   23986177
         CurrentDate     =   36938
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Forma de pago"
         Height          =   195
         Left            =   5760
         TabIndex        =   17
         Top             =   1320
         Width           =   1080
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Transportista"
         Height          =   195
         Left            =   5760
         TabIndex        =   11
         Top             =   960
         Width           =   960
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Left            =   5760
         TabIndex        =   7
         Top             =   600
         Width           =   435
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   1065
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Su Referencia"
         Height          =   195
         Left            =   3120
         TabIndex        =   15
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "N/Referencia"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha entrega"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   8400
         TabIndex        =   3
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   9360
      TabIndex        =   39
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8160
      TabIndex        =   38
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6960
      TabIndex        =   37
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Líneas del Pedido de Compra"
      Height          =   3855
      Left            =   120
      TabIndex        =   23
      Top             =   2760
      Width           =   10335
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3315
         Width           =   1455
      End
      Begin VB.TextBox txtTotalBruto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3315
         Width           =   1455
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "El&iminar"
         Height          =   375
         Left            =   8760
         TabIndex        =   35
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   7680
         TabIndex        =   34
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Aña&dir"
         Height          =   375
         Left            =   6600
         TabIndex        =   33
         Top             =   3360
         Width           =   975
      End
      Begin VB.PictureBox picPedidoCompraItems 
         Height          =   2535
         Index           =   1
         Left            =   240
         ScaleHeight     =   2475
         ScaleWidth      =   9315
         TabIndex        =   25
         Top             =   600
         Width           =   9375
         Begin MSComctlLib.ListView lvwPedidoCompraItems 
            Height          =   2500
            Left            =   -25
            TabIndex        =   26
            Top             =   -25
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   4419
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
      End
      Begin VB.PictureBox picPedidoCompraItems 
         Height          =   2535
         Index           =   2
         Left            =   240
         ScaleHeight     =   2475
         ScaleWidth      =   9315
         TabIndex        =   27
         Top             =   600
         Width           =   9375
         Begin MSComctlLib.ListView lvwPedidoCompraArticulos 
            Height          =   2500
            Left            =   -25
            TabIndex        =   28
            Top             =   -25
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   4419
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
      End
      Begin MSComctlLib.TabStrip tsPedidoCompraItems 
         Height          =   3015
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   5318
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Materiales"
               Key             =   "Materiales"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Artículos"
               Key             =   "Articulos"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total cantidad"
         Height          =   195
         Left            =   3480
         TabIndex        =   32
         Top             =   3330
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Bruto del Pedido"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   3330
         Width           =   1575
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Total Bruto del Pedido"
      Height          =   195
      Left            =   3600
      TabIndex        =   36
      Top             =   6240
      Width           =   1575
   End
End
Attribute VB_Name = "PedidoCompraEdit"
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
Private mintCurFrame As Integer ' Marco activo visible

Private WithEvents mobjPedidoCompra As PedidoCompra
Attribute mobjPedidoCompra.VB_VarHelpID = -1

Public Sub Component(PedidoCompraObject As PedidoCompra)

    Set mobjPedidoCompra = PedidoCompraObject

End Sub

Private Sub btnDatoComercial_Click()

    Dim frmDatoComercial As DatoComercialEdit
  
    Set frmDatoComercial = New DatoComercialEdit
    frmDatoComercial.Component mobjPedidoCompra.DatoComercial
    frmDatoComercial.Show vbModal
    txtDatoComercial.Text = mobjPedidoCompra.DatoComercial.DatoComercialText
    
End Sub

Private Sub cboProveedor_Click()

    On Error GoTo ErrorManager
    
    If mflgLoading Then Exit Sub
    mobjPedidoCompra.Proveedor = cboProveedor.Text
  
    ' Al modificar el proveedor se refrescan en el interface los datos relacionados.
    cboBanco.Text = mobjPedidoCompra.Banco
    cboBanco_Click
    cboTransportista.Text = mobjPedidoCompra.Transportista
    cboTransportista_Click
    cboFormaPago.Text = mobjPedidoCompra.FormaPago
    cboFormaPago_Click
    txtDatoComercial.Text = mobjPedidoCompra.DatoComercial.DatoComercialText
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
    Exit Sub
End Sub

Private Sub cmdApply_Click()

    On Error GoTo ErrorManager

    mobjPedidoCompra.ApplyEdit
    mobjPedidoCompra.BeginEdit GescomMain.objParametro.Moneda
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()
    
    If mobjPedidoCompra.Numero = GescomMain.objParametro.ObjEmpresaActual.PedidoCompras Then
        GescomMain.objParametro.ObjEmpresaActual.DecrementaPedidocompras
    End If
  
    mobjPedidoCompra.CancelEdit
  
    Unload Me

End Sub

Private Sub cmdOK_Click()

    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    mobjPedidoCompra.ApplyEdit
    GescomMain.objParametro.ObjEmpresaActual.EstablecePedidoCompras (mobjPedidoCompra.Numero)
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Private Sub dtpFecha_Change()
  
    mobjPedidoCompra.Fecha = dtpFecha.Value
    
End Sub

Private Sub dtpFechaEntrega_Change()
  
    mobjPedidoCompra.FechaEntrega = dtpFechaEntrega.Value
    
End Sub

Private Sub Form_Load()
    Dim objParametroAplicacion As ParametroAplicacion

    On Error GoTo ErrorManager
  
    DisableX Me
    
    mflgLoading = True
    
    With mobjPedidoCompra
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Pedido de Compra [(nuevo)]"

        Else
            Caption = "Pedido de Compra [" & Trim(.Proveedor) & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtNumero = .Numero
        dtpFecha.Value = .Fecha
        dtpFechaEntrega.Value = .FechaEntrega
        txtNuestraReferencia = .NuestraReferencia
        txtSuReferencia = .SuReferencia
        txtObservaciones = .Observaciones
        txtDatoComercial.Text = .DatoComercial.DatoComercialText
        txtTotalBruto = FormatoMoneda(.TotalBruto, GescomMain.objParametro.Moneda)
        txtCantidad = .Cantidad
        
        LoadCombo cboProveedor, .Proveedores
        cboProveedor.Text = .Proveedor
    
        LoadCombo cboBanco, .Bancos
        cboBanco.Text = .Banco
    
        LoadCombo cboTransportista, .Transportistas
        cboTransportista.Text = .Transportista
    
        LoadCombo cboFormaPago, .FormasPago
        cboFormaPago.Text = .FormaPago
    
        .BeginEdit GescomMain.objParametro.Moneda
        
        If .IsNew Then
            .Numero = GescomMain.objParametro.ObjEmpresaActual.IncrementaPedidoCompras
            txtNumero = .Numero
       
            .TemporadaID = GescomMain.objParametro.TemporadaActualID
            .EmpresaID = GescomMain.objParametro.EmpresaActualID
        End If
    
    End With
    
    lvwPedidoCompraItems.SmallIcons = GescomMain.mglIconosPequeños
    
    lvwPedidoCompraItems.ColumnHeaders.Add , , vbNullString, ColumnSize(2)
    lvwPedidoCompraItems.ColumnHeaders.Add , , "Material", ColumnSize(26)
    lvwPedidoCompraItems.ColumnHeaders.Add , , "Cantidad", ColumnSize(10), vbRightJustify
    lvwPedidoCompraItems.ColumnHeaders.Add , , "Bruto", ColumnSize(8), vbRightJustify
    
    lvwPedidoCompraArticulos.SmallIcons = GescomMain.mglIconosPequeños
    lvwPedidoCompraArticulos.ColumnHeaders.Add , , vbNullString, ColumnSize(2)
    lvwPedidoCompraArticulos.ColumnHeaders.Add , , "Lín", ColumnSize(4)
    lvwPedidoCompraArticulos.ColumnHeaders.Add , , "Artículo - Color", ColumnSize(15)
    lvwPedidoCompraArticulos.ColumnHeaders.Add , , "36", ColumnSize(3), vbRightJustify
    lvwPedidoCompraArticulos.ColumnHeaders.Add , , "38", ColumnSize(3), vbRightJustify
    lvwPedidoCompraArticulos.ColumnHeaders.Add , , "40", ColumnSize(3), vbRightJustify
    lvwPedidoCompraArticulos.ColumnHeaders.Add , , "42", ColumnSize(3), vbRightJustify
    lvwPedidoCompraArticulos.ColumnHeaders.Add , , "44", ColumnSize(3), vbRightJustify
    lvwPedidoCompraArticulos.ColumnHeaders.Add , , "46", ColumnSize(3), vbRightJustify
    lvwPedidoCompraArticulos.ColumnHeaders.Add , , "48", ColumnSize(3), vbRightJustify
    lvwPedidoCompraArticulos.ColumnHeaders.Add , , "50", ColumnSize(3), vbRightJustify
    lvwPedidoCompraArticulos.ColumnHeaders.Add , , "52", ColumnSize(3), vbRightJustify
    lvwPedidoCompraArticulos.ColumnHeaders.Add , , "54", ColumnSize(3), vbRightJustify
    lvwPedidoCompraArticulos.ColumnHeaders.Add , , "56", ColumnSize(3), vbRightJustify
    lvwPedidoCompraArticulos.ColumnHeaders.Add , , "Suma", ColumnSize(4), vbRightJustify
    lvwPedidoCompraArticulos.ColumnHeaders.Add , , "Pr.Venta", ColumnSize(10), vbRightJustify
    lvwPedidoCompraArticulos.ColumnHeaders.Add , , "Bruto", ColumnSize(10), vbRightJustify
    
    LoadPedidoCompraItems
      
    Set objParametroAplicacion = New ParametroAplicacion
    Select Case objParametroAplicacion.TipoInstalacion
    Case TIPOINSTALACION_FABRICA
        mintCurFrame = 2
        tsPedidoCompraItems.SelectedItem = tsPedidoCompraItems.Tabs(1)
        tsPedidoCompraItems_Click
    Case TIPOINSTALACION_PUNTOVENTA
        mintCurFrame = 1
        tsPedidoCompraItems.SelectedItem = tsPedidoCompraItems.Tabs(2)
        tsPedidoCompraItems_Click
    Case Else
        Err.Raise vbObjectError + 1001, "PedidoCompraEdit LOAD", "No existe el tipo de instalación:" & objParametroAplicacion.TipoInstalacion & ". Es necesario corregir el valor del parámetro."
    End Select
    
    Set objParametroAplicacion = Nothing
    
    mflgLoading = False
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub lvwPedidoCompraItems_DblClick()
  
    Call cmdEdit_Click
    
End Sub

Private Sub lvwPedidoCompraArticulos_DblClick()
  
    Call cmdEdit_Click
    
End Sub

Private Sub mobjPedidoCompra_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub tsPedidoCompraItems_Click()
   
   If tsPedidoCompraItems.SelectedItem.Index = mintCurFrame _
      Then Exit Sub ' No necesita cambiar el marco.
   ' Oculte el marco antiguo y muestre el nuevo.
   picPedidoCompraItems(tsPedidoCompraItems.SelectedItem.Index).Visible = True
   picPedidoCompraItems(mintCurFrame).Visible = False
   ' Establece mintCurFrame al nuevo valor.
   mintCurFrame = tsPedidoCompraItems.SelectedItem.Index
End Sub

Private Sub txtDatoComercial_DblClick()

    Call btnDatoComercial_Click
    
End Sub

Private Sub txtObservaciones_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtObservaciones

End Sub

Private Sub txtObservaciones_Change()

    If Not mflgLoading Then _
        TextChange txtObservaciones, mobjPedidoCompra, "Observaciones"

End Sub

Private Sub txtObservaciones_LostFocus()

    txtObservaciones = TextLostFocus(txtObservaciones, mobjPedidoCompra, "Observaciones")

End Sub

Private Sub txtNumero_Change()

    If Not mflgLoading Then _
        TextChange txtNumero, mobjPedidoCompra, "Numero"

End Sub

Private Sub txtNumero_LostFocus()

    txtNumero = TextLostFocus(txtNumero, mobjPedidoCompra, "Numero")

End Sub

Private Sub cboTransportista_Click()

    If mflgLoading Then Exit Sub
    mobjPedidoCompra.Transportista = cboTransportista.Text

End Sub

Private Sub cboBanco_Click()

    If mflgLoading Then Exit Sub
    mobjPedidoCompra.Banco = cboBanco.Text

End Sub

Private Sub cboFormaPago_Click()

    If mflgLoading Then Exit Sub
    mobjPedidoCompra.FormaPago = cboFormaPago.Text

End Sub

Private Sub txtNuestraReferencia_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtNuestraReferencia

End Sub

Private Sub txtNuestraReferencia_Change()

    If Not mflgLoading Then _
        TextChange txtNuestraReferencia, mobjPedidoCompra, "NuestraReferencia"

End Sub

Private Sub txtNuestraReferencia_LostFocus()

    txtNuestraReferencia = TextLostFocus(txtNuestraReferencia, mobjPedidoCompra, "NuestraReferencia")

End Sub

Private Sub txtSuReferencia_GotFocus()
      
    If Not mflgLoading Then _
        SelTextBox txtSuReferencia

End Sub

Private Sub txtSuReferencia_Change()

    If Not mflgLoading Then _
        TextChange txtSuReferencia, mobjPedidoCompra, "SuReferencia"

End Sub

Private Sub txtSuReferencia_LostFocus()

    txtSuReferencia = TextLostFocus(txtSuReferencia, mobjPedidoCompra, "SuReferencia")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   
   IsList = False
   
End Function

' a partir de aqui -----> child

Private Sub cmdAdd_Click()
  
    If mintCurFrame = 1 Then
        AddMaterial
    Else
        AddArticulo
    End If
  
End Sub

Private Sub AddMaterial()
  
    Dim frmPedidoCompraItem As PedidoCompraItemEdit
  
    On Error GoTo ErrorManager
    Set frmPedidoCompraItem = New PedidoCompraItemEdit
    frmPedidoCompraItem.Component mobjPedidoCompra.PedidoCompraItems.Add(PEDIDOCOMPRAITEM_MATERIAL)
    frmPedidoCompraItem.Show vbModal
    LoadPedidoCompraItems
'    txtTotalBruto = FormatoMoneda(mobjPedidoCompra.TotalBruto, GescomMain.objParametro.Moneda)
'    txtCantidad = mobjPedidoCompra.Cantidad
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub AddArticulo()
  
    Dim frmPedidoCompraItem As PedidoCompraArticuloEdit
  
    On Error GoTo ErrorManager
    Set frmPedidoCompraItem = New PedidoCompraArticuloEdit
    frmPedidoCompraItem.Component mobjPedidoCompra.PedidoCompraItems.Add(PEDIDOCOMPRAITEM_ARTICULO), mobjPedidoCompra.ProveedorID
    frmPedidoCompraItem.Show vbModal
    LoadPedidoCompraItems
'    txtTotalBruto = FormatoMoneda(mobjPedidoCompra.TotalBruto, GescomMain.objParametro.Moneda)
'    txtCantidad = mobjPedidoCompra.Cantidad
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdEdit_Click()
    Dim frmPedidoCompraItem As PedidoCompraItemEdit
    Dim frmPedidoCompraArticulo As PedidoCompraArticuloEdit
    
    On Error GoTo ErrorManager
    
    If mintCurFrame = 1 Then
        If lvwPedidoCompraItems.SelectedItem Is Nothing Then Exit Sub
    
        Set frmPedidoCompraItem = New PedidoCompraItemEdit
        frmPedidoCompraItem.Component _
            mobjPedidoCompra.PedidoCompraItems.Item(Val(lvwPedidoCompraItems.SelectedItem.Key))
        frmPedidoCompraItem.Show vbModal
    End If
    If mintCurFrame = 2 Then
        If lvwPedidoCompraArticulos.SelectedItem Is Nothing Then Exit Sub
    
        Set frmPedidoCompraArticulo = New PedidoCompraArticuloEdit
        frmPedidoCompraArticulo.Component _
            mobjPedidoCompra.PedidoCompraItems.Item(Val(lvwPedidoCompraArticulos.SelectedItem.Key)), _
            mobjPedidoCompra.ProveedorID
        frmPedidoCompraArticulo.Show vbModal
    End If
    LoadPedidoCompraItems
    
    txtTotalBruto = FormatoMoneda(mobjPedidoCompra.TotalBruto, GescomMain.objParametro.Moneda)
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

 Private Sub cmdRemove_Click()

    On Error GoTo ErrorManager
            
    If mintCurFrame = 1 Then
        If lvwPedidoCompraItems.SelectedItem Is Nothing Then Exit Sub
        mobjPedidoCompra.PedidoCompraItems.Remove Val(lvwPedidoCompraItems.SelectedItem.Key)
    End If
    
    If mintCurFrame = 2 Then
        If lvwPedidoCompraArticulos.SelectedItem Is Nothing Then Exit Sub
        mobjPedidoCompra.PedidoCompraItems.Remove Val(lvwPedidoCompraArticulos.SelectedItem.Key)
    End If
    
    LoadPedidoCompraItems
'    txtTotalBruto = FormatoMoneda(mobjPedidoCompra.TotalBruto, GescomMain.objParametro.Moneda)
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub LoadPedidoCompraItems()

    Dim objPedidoCompraItem As PedidoCompraItem
    Dim objPedidoCompraItemMaterial As PedidoCompraItemMaterial
    Dim objPedidoCompraItemArticulo As PedidoCompraItemArticulo
    Dim itmList As ListItem
    Dim lngIndex As Long
  
    On Error GoTo ErrorManager
    lvwPedidoCompraItems.ListItems.Clear
    lvwPedidoCompraArticulos.ListItems.Clear
    For lngIndex = 1 To mobjPedidoCompra.PedidoCompraItems.Count
        Set objPedidoCompraItem = mobjPedidoCompra.PedidoCompraItems.Item(lngIndex)
        
        Select Case objPedidoCompraItem.Tipo
        Case PEDIDOCOMPRAITEM_MATERIAL
            Set itmList = lvwPedidoCompraItems.ListItems.Add _
                (Key:=Format$(lngIndex) & "K")
    
            With itmList
                .SmallIcon = GescomMain.mglIconosPequeños.ListImages("NuevoItem").Key
                
                If objPedidoCompraItem.IsDeleted Then .SmallIcon = GescomMain.mglIconosPequeños.ListImages("EliminarItem").Key
                Set objPedidoCompraItemMaterial = objPedidoCompraItem
                .SubItems(1) = Trim(objPedidoCompraItemMaterial.Material)
                Set objPedidoCompraItemMaterial = Nothing
                .SubItems(2) = objPedidoCompraItem.Cantidad
                .SubItems(3) = FormatoMoneda(objPedidoCompraItem.Bruto, GescomMain.objParametro.Moneda)
            End With
        Case PEDIDOCOMPRAITEM_ARTICULO
            Set itmList = lvwPedidoCompraArticulos.ListItems.Add _
                (Key:=Format$(lngIndex) & "K")
    
            With itmList
                .SmallIcon = GescomMain.mglIconosPequeños.ListImages("NuevoItem").Key
                
                If objPedidoCompraItem.IsDeleted Then .SmallIcon = GescomMain.mglIconosPequeños.ListImages("EliminarItem").Key
                Set objPedidoCompraItemArticulo = objPedidoCompraItem
                .SubItems(1) = Space(4 - Len(Format$(lngIndex))) & Format$(lngIndex)
                .SubItems(2) = Trim(objPedidoCompraItemArticulo.ArticuloColor)
                .SubItems(3) = objPedidoCompraItemArticulo.CantidadT36
                .SubItems(4) = objPedidoCompraItemArticulo.CantidadT38
                .SubItems(5) = objPedidoCompraItemArticulo.CantidadT40
                .SubItems(6) = objPedidoCompraItemArticulo.CantidadT42
                .SubItems(7) = objPedidoCompraItemArticulo.CantidadT44
                .SubItems(8) = objPedidoCompraItemArticulo.CantidadT46
                .SubItems(9) = objPedidoCompraItemArticulo.CantidadT48
                .SubItems(10) = objPedidoCompraItemArticulo.CantidadT50
                .SubItems(11) = objPedidoCompraItemArticulo.CantidadT52
                .SubItems(12) = objPedidoCompraItemArticulo.CantidadT54
                .SubItems(13) = objPedidoCompraItemArticulo.CantidadT56
                Set objPedidoCompraItemArticulo = Nothing
                .SubItems(14) = objPedidoCompraItem.Cantidad
                .SubItems(15) = FormatoMoneda(objPedidoCompraItem.PrecioCoste, "EUR", False)
                .SubItems(16) = FormatoMoneda(objPedidoCompraItem.Bruto, "EUR")
            End With
            
        Case Else
            Err.Raise vbObjectError + 1001, "PedidoCompraEdit LoadPedidoCompraItems", "No existe el tipo de item de albaran de compra:" & objPedidoCompraItem.Tipo & ". Avisar al personal técnico."
        End Select
        
    Next
    txtTotalBruto = FormatoMoneda(mobjPedidoCompra.TotalBruto, GescomMain.objParametro.Moneda)
    txtCantidad = mobjPedidoCompra.Cantidad
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
    Exit Sub

End Sub

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

Private Sub cboBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintBancoSelStart = cboBanco.SelStart
End Sub

Private Sub cboBanco_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintBancoSelStart, cboBanco
    
End Sub

