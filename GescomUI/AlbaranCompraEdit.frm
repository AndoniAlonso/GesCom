VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form AlbaranCompraEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Albaranes de Compra"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AlbaranCompraEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6600
      TabIndex        =   35
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7800
      TabIndex        =   36
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   9000
      TabIndex        =   37
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Líneas del Albarán de Compra"
      Height          =   4335
      Left            =   240
      TabIndex        =   21
      Top             =   2640
      Width           =   9855
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1455
      End
      Begin VB.PictureBox picAlbaranCompraItems 
         Height          =   3015
         Index           =   2
         Left            =   240
         ScaleHeight     =   2955
         ScaleWidth      =   9315
         TabIndex        =   25
         Top             =   600
         Width           =   9375
         Begin MSComctlLib.ListView lvwAlbaranCompraArticulos 
            Height          =   3015
            Left            =   -30
            TabIndex        =   26
            Top             =   -25
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   5318
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
      Begin VB.PictureBox picAlbaranCompraItems 
         Height          =   3015
         Index           =   1
         Left            =   240
         ScaleHeight     =   2955
         ScaleWidth      =   9315
         TabIndex        =   23
         Top             =   600
         Width           =   9375
         Begin MSComctlLib.ListView lvwAlbaranCompraItems 
            Height          =   3015
            Left            =   -30
            TabIndex        =   24
            Top             =   -25
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   5318
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
      Begin MSComctlLib.TabStrip tsAlbaranCompraItems 
         Height          =   3495
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   6165
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
               Key             =   "Artículos"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Aña&dir"
         Height          =   375
         Left            =   6600
         TabIndex        =   29
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   7680
         TabIndex        =   30
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "El&iminar"
         Height          =   375
         Left            =   8760
         TabIndex        =   31
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txtTotalBruto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   3945
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total cantidad"
         Height          =   195
         Left            =   3720
         TabIndex        =   28
         Top             =   3975
         Width           =   1020
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Total bruto del albarán"
         Height          =   195
         Left            =   360
         TabIndex        =   27
         Top             =   3960
         Width           =   1635
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Albarán"
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin VB.TextBox txtEmbalajes 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6480
         TabIndex        =   18
         Top             =   2145
         Width           =   1095
      End
      Begin VB.TextBox txtPortes 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8520
         TabIndex        =   20
         Top             =   2145
         Width           =   1095
      End
      Begin VB.TextBox txtDatoComercial 
         Height          =   990
         Left            =   6840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   1065
         Width           =   2775
      End
      Begin VB.CommandButton btnDatoComercial 
         Caption         =   "Datos C&omerciales"
         Height          =   615
         Left            =   5520
         TabIndex        =   12
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox cboTransportista 
         Height          =   315
         Left            =   6840
         TabIndex        =   8
         Text            =   "cboTransportista"
         Top             =   705
         Width           =   2775
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   285
         Left            =   1560
         TabIndex        =   16
         Top             =   1785
         Width           =   3735
      End
      Begin VB.TextBox txtSuReferencia 
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Top             =   1420
         Width           =   1935
      End
      Begin VB.TextBox txtNuestraReferencia 
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   1060
         Width           =   1935
      End
      Begin VB.ComboBox cboProveedor 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Text            =   "cboProveedor"
         Top             =   340
         Width           =   3735
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8640
         TabIndex        =   4
         Top             =   340
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   700
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   71565313
         CurrentDate     =   36938
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Embalajes"
         Height          =   195
         Left            =   5520
         TabIndex        =   17
         Top             =   2160
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Portes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7860
         TabIndex        =   19
         Top             =   2160
         Width           =   465
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Transportista"
         Height          =   195
         Left            =   5520
         TabIndex        =   7
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   1800
         Width           =   1065
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Su Referencia"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "N/Referencia"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   7800
         TabIndex        =   3
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdPedidos 
      Caption         =   "&Pedidos..."
      Height          =   375
      Left            =   240
      TabIndex        =   34
      ToolTipText     =   "Incorporar pedidos pendientes"
      Top             =   7080
      Width           =   1335
   End
End
Attribute VB_Name = "AlbaranCompraEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private mintProveedorSelStart As Integer
Private mintTransportistaSelStart As Integer
Private mintCurFrame As Integer ' Marco activo visible

Private WithEvents mobjAlbaranCompra As AlbaranCompra
Attribute mobjAlbaranCompra.VB_VarHelpID = -1

Public Sub Component(AlbaranCompraObject As AlbaranCompra)

    Set mobjAlbaranCompra = AlbaranCompraObject

End Sub

Private Sub btnDatoComercial_Click()
    Dim frmDatoComercial As DatoComercialEdit
  
    Set frmDatoComercial = New DatoComercialEdit
    frmDatoComercial.Component mobjAlbaranCompra.DatoComercial
    frmDatoComercial.Show vbModal
    txtDatoComercial.Text = mobjAlbaranCompra.DatoComercial.DatoComercialText
    
End Sub

Private Sub cboProveedor_Click()
    
    On Error GoTo ErrorManager
  
    If mflgLoading Then Exit Sub
    mobjAlbaranCompra.Proveedor = cboProveedor.Text
  
    ' Al modificar el proveedor se refrescan en el interface los datos relacionados.
    cboTransportista.Text = mobjAlbaranCompra.Transportista
    cboTransportista_Click
    txtDatoComercial.Text = mobjAlbaranCompra.DatoComercial.DatoComercialText
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdApply_Click()

    On Error GoTo ErrorManager

    mobjAlbaranCompra.ApplyEdit
    mobjAlbaranCompra.BeginEdit GescomMain.objParametro.Moneda
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    If mobjAlbaranCompra.Numero = GescomMain.objParametro.ObjEmpresaActual.AlbaranCompras Then
        GescomMain.objParametro.ObjEmpresaActual.DecrementaAlbaranCompras
    End If
  
    mobjAlbaranCompra.CancelEdit
  
    Unload Me

End Sub

Private Sub cmdOK_Click()

    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    mobjAlbaranCompra.ApplyEdit
    GescomMain.objParametro.ObjEmpresaActual.EstableceAlbaranCompras (mobjAlbaranCompra.Numero)
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Private Sub dtpFecha_Change()
    
    mobjAlbaranCompra.Fecha = dtpFecha.Value
    
End Sub

Private Sub Form_Load()
    Dim objParametroAplicacion As ParametroAplicacion

    On Error GoTo ErrorManager
  
    DisableX Me
    
    mflgLoading = True
    With mobjAlbaranCompra
        EnableOK .IsValid
        
        If .IsNew Then
            Caption = "Albarán de Compra [(nuevo)]"

        Else
            Caption = "Albarán de Compra [" & Trim(.Proveedor) & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtNumero = .Numero
        dtpFecha.Value = .Fecha
        txtNuestraReferencia = .NuestraReferencia
        txtSuReferencia = .SuReferencia
        txtObservaciones = .Observaciones
        txtPortes = .Portes
        txtEmbalajes = .Embalajes
        txtDatoComercial.Text = .DatoComercial.DatoComercialText
        txtTotalBruto = FormatoMoneda(.TotalBruto, GescomMain.objParametro.Moneda)
        txtCantidad = .Cantidad
        
        LoadCombo cboProveedor, .Proveedores
        cboProveedor.Text = .Proveedor
        
        LoadCombo cboTransportista, .Transportistas
        cboTransportista.Text = .Transportista
        
        .BeginEdit GescomMain.objParametro.Moneda
        
        If .IsNew Then
            .Numero = GescomMain.objParametro.ObjEmpresaActual.IncrementaAlbaranCompras
            txtNumero = .Numero
       
            .TemporadaID = GescomMain.objParametro.TemporadaActualID
            .EmpresaID = GescomMain.objParametro.EmpresaActualID
        End If
    
    End With
    
    lvwAlbaranCompraItems.SmallIcons = GescomMain.mglIconosPequeños
    
    lvwAlbaranCompraItems.ColumnHeaders.Add , , vbNullString, ColumnSize(2)
    lvwAlbaranCompraItems.ColumnHeaders.Add , , "Material", ColumnSize(26)
    lvwAlbaranCompraItems.ColumnHeaders.Add , , "Cantidad", ColumnSize(10), vbRightJustify
    lvwAlbaranCompraItems.ColumnHeaders.Add , , "Bruto", ColumnSize(10), vbRightJustify
    
    lvwAlbaranCompraArticulos.SmallIcons = GescomMain.mglIconosPequeños
    lvwAlbaranCompraArticulos.ColumnHeaders.Add , , vbNullString, ColumnSize(2)
    lvwAlbaranCompraArticulos.ColumnHeaders.Add , , "Lín", ColumnSize(4)
    lvwAlbaranCompraArticulos.ColumnHeaders.Add , , "Artículo - Color", ColumnSize(15)
    lvwAlbaranCompraArticulos.ColumnHeaders.Add , , "36", ColumnSize(3), vbRightJustify
    lvwAlbaranCompraArticulos.ColumnHeaders.Add , , "38", ColumnSize(3), vbRightJustify
    lvwAlbaranCompraArticulos.ColumnHeaders.Add , , "40", ColumnSize(3), vbRightJustify
    lvwAlbaranCompraArticulos.ColumnHeaders.Add , , "42", ColumnSize(3), vbRightJustify
    lvwAlbaranCompraArticulos.ColumnHeaders.Add , , "44", ColumnSize(3), vbRightJustify
    lvwAlbaranCompraArticulos.ColumnHeaders.Add , , "46", ColumnSize(3), vbRightJustify
    lvwAlbaranCompraArticulos.ColumnHeaders.Add , , "48", ColumnSize(3), vbRightJustify
    lvwAlbaranCompraArticulos.ColumnHeaders.Add , , "50", ColumnSize(3), vbRightJustify
    lvwAlbaranCompraArticulos.ColumnHeaders.Add , , "52", ColumnSize(3), vbRightJustify
    lvwAlbaranCompraArticulos.ColumnHeaders.Add , , "54", ColumnSize(3), vbRightJustify
    lvwAlbaranCompraArticulos.ColumnHeaders.Add , , "56", ColumnSize(3), vbRightJustify
    lvwAlbaranCompraArticulos.ColumnHeaders.Add , , "Suma", ColumnSize(4), vbRightJustify
    lvwAlbaranCompraArticulos.ColumnHeaders.Add , , "Pr.Venta", ColumnSize(10), vbRightJustify
    lvwAlbaranCompraArticulos.ColumnHeaders.Add , , "Bruto", ColumnSize(10), vbRightJustify
    
    LoadAlbaranCompraItems
  
    Set objParametroAplicacion = New ParametroAplicacion
    Select Case objParametroAplicacion.TipoInstalacion
    Case TIPOINSTALACION_FABRICA
        mintCurFrame = 2
        tsAlbaranCompraItems.SelectedItem = tsAlbaranCompraItems.Tabs(1)
        tsAlbaranCompraItems_Click
    Case TIPOINSTALACION_PUNTOVENTA
        mintCurFrame = 1
        tsAlbaranCompraItems.SelectedItem = tsAlbaranCompraItems.Tabs(2)
        tsAlbaranCompraItems_Click
    Case Else
        Err.Raise vbObjectError + 1001, "AlbaranCompraEdit LOAD", "No existe el tipo de instalación:" & objParametroAplicacion.TipoInstalacion & ". Es necesario corregir el valor del parámetro."
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

Private Sub lvwAlbaranCompraItems_DblClick()
  
    Call cmdEdit_Click
    
End Sub

Private Sub lvwAlbaranCompraArticulos_DblClick()
  
    Call cmdEdit_Click
    
End Sub

Private Sub mobjAlbaranCompra_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub tsAlbaranCompraItems_Click()
   
   If tsAlbaranCompraItems.SelectedItem.Index = mintCurFrame _
      Then Exit Sub ' No necesita cambiar el marco.
   ' Oculte el marco antiguo y muestre el nuevo.
   picAlbaranCompraItems(tsAlbaranCompraItems.SelectedItem.Index).Visible = True
   picAlbaranCompraItems(mintCurFrame).Visible = False
   ' Establece mintCurFrame al nuevo valor.
   mintCurFrame = tsAlbaranCompraItems.SelectedItem.Index
End Sub

Private Sub txtDatoComercial_Click()

    Call btnDatoComercial_Click
    
End Sub

Private Sub txtObservaciones_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtObservaciones

End Sub

Private Sub txtObservaciones_Change()

    If Not mflgLoading Then _
        TextChange txtObservaciones, mobjAlbaranCompra, "Observaciones"

End Sub

Private Sub txtObservaciones_LostFocus()

    txtObservaciones = TextLostFocus(txtObservaciones, mobjAlbaranCompra, "Observaciones")

End Sub

Private Sub txtPortes_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPortes

End Sub

Private Sub txtPortes_Change()

    If Not mflgLoading Then _
        TextChange txtPortes, mobjAlbaranCompra, "Portes"

End Sub

Private Sub txtPortes_LostFocus()

    txtPortes = TextLostFocus(txtPortes, mobjAlbaranCompra, "Portes")

End Sub

Private Sub txtEmbalajes_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtEmbalajes

End Sub

Private Sub txtEmbalajes_Change()

    If Not mflgLoading Then _
        TextChange txtEmbalajes, mobjAlbaranCompra, "Embalajes"

End Sub

Private Sub txtEmbalajes_LostFocus()

    txtEmbalajes = TextLostFocus(txtEmbalajes, mobjAlbaranCompra, "Embalajes")

End Sub

Private Sub txtNumero_Change()

    If Not mflgLoading Then _
        TextChange txtNumero, mobjAlbaranCompra, "Numero"

End Sub

Private Sub txtNumero_LostFocus()

    txtNumero = TextLostFocus(txtNumero, mobjAlbaranCompra, "Numero")

End Sub

Private Sub cboTransportista_Click()

    If mflgLoading Then Exit Sub
    mobjAlbaranCompra.Transportista = cboTransportista.Text

End Sub

Private Sub txtNuestraReferencia_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtNuestraReferencia

End Sub

Private Sub txtNuestraReferencia_Change()

    If Not mflgLoading Then _
        TextChange txtNuestraReferencia, mobjAlbaranCompra, "NuestraReferencia"

End Sub

Private Sub txtNuestraReferencia_LostFocus()

    txtNuestraReferencia = TextLostFocus(txtNuestraReferencia, mobjAlbaranCompra, "NuestraReferencia")

End Sub

Private Sub txtSuReferencia_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtSuReferencia

End Sub

Private Sub txtSuReferencia_Change()

    If Not mflgLoading Then _
        TextChange txtSuReferencia, mobjAlbaranCompra, "SuReferencia"

End Sub

Private Sub txtSuReferencia_LostFocus()

    txtSuReferencia = TextLostFocus(txtSuReferencia, mobjAlbaranCompra, "SuReferencia")

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
    Dim frmAlbaranCompraItem As AlbaranCompraItemEdit
  
    On Error GoTo ErrorManager
    Set frmAlbaranCompraItem = New AlbaranCompraItemEdit
    frmAlbaranCompraItem.Component mobjAlbaranCompra.AlbaranCompraItems.Add(ALBARANCOMPRAITEM_MATERIAL)
    frmAlbaranCompraItem.Show vbModal
    LoadAlbaranCompraItems
'    txtTotalBruto = FormatoMoneda(mobjAlbaranCompra.TotalBruto, GescomMain.objParametro.Moneda)
'    txtCantidad = mobjAlbaranCompra.Cantidad
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub AddArticulo()
    Dim frmAlbaranCompraItem As AlbaranCompraArticuloEdit
  
    On Error GoTo ErrorManager
    Set frmAlbaranCompraItem = New AlbaranCompraArticuloEdit
    frmAlbaranCompraItem.Component mobjAlbaranCompra.AlbaranCompraItems.Add(ALBARANCOMPRAITEM_ARTICULO), mobjAlbaranCompra.ProveedorID
    frmAlbaranCompraItem.Show vbModal
    LoadAlbaranCompraItems
'    txtTotalBruto = FormatoMoneda(mobjAlbaranCompra.TotalBruto, GescomMain.objParametro.Moneda)
'    txtCantidad = mobjAlbaranCompra.Cantidad
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdEdit_Click()
    Dim frmAlbaranCompraItem As AlbaranCompraItemEdit
    Dim frmAlbaranCompraArticulo As AlbaranCompraArticuloEdit
  
    On Error GoTo ErrorManager
    
    If mintCurFrame = 1 Then
        If lvwAlbaranCompraItems.SelectedItem Is Nothing Then Exit Sub
        
        Set frmAlbaranCompraItem = New AlbaranCompraItemEdit
        frmAlbaranCompraItem.Component _
            mobjAlbaranCompra.AlbaranCompraItems(Val(lvwAlbaranCompraItems.SelectedItem.Key))
        frmAlbaranCompraItem.Show vbModal
    End If
    If mintCurFrame = 2 Then
        If lvwAlbaranCompraArticulos.SelectedItem Is Nothing Then Exit Sub
    
        Set frmAlbaranCompraArticulo = New AlbaranCompraArticuloEdit
        frmAlbaranCompraArticulo.Component _
            mobjAlbaranCompra.AlbaranCompraItems.Item(Val(lvwAlbaranCompraArticulos.SelectedItem.Key)), _
            mobjAlbaranCompra.ProveedorID
        frmAlbaranCompraArticulo.Show vbModal
    End If
        
    LoadAlbaranCompraItems
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

 Private Sub cmdRemove_Click()

    On Error GoTo ErrorManager
        
    If mintCurFrame = 1 Then
        If lvwAlbaranCompraItems.SelectedItem Is Nothing Then Exit Sub
        mobjAlbaranCompra.AlbaranCompraItems.Remove Val(lvwAlbaranCompraItems.SelectedItem.Key)
    End If
    If mintCurFrame = 2 Then
        If lvwAlbaranCompraArticulos.SelectedItem Is Nothing Then Exit Sub
        mobjAlbaranCompra.AlbaranCompraItems.Remove Val(lvwAlbaranCompraArticulos.SelectedItem.Key)
    End If
    
    LoadAlbaranCompraItems
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub LoadAlbaranCompraItems()
    Dim objAlbaranCompraItem As AlbaranCompraItem
    Dim objAlbaranCompraItemMaterial As AlbaranCompraItemMaterial
    Dim objAlbaranCompraItemArticulo As AlbaranCompraItemArticulo
    Dim itmList As ListItem
    Dim lngIndex As Long
  
    On Error GoTo ErrorManager
    lvwAlbaranCompraItems.ListItems.Clear
    lvwAlbaranCompraArticulos.ListItems.Clear
    For lngIndex = 1 To mobjAlbaranCompra.AlbaranCompraItems.Count
        Set objAlbaranCompraItem = mobjAlbaranCompra.AlbaranCompraItems(lngIndex)
        Select Case objAlbaranCompraItem.Tipo
        Case ALBARANCOMPRAITEM_MATERIAL
        
            Set itmList = lvwAlbaranCompraItems.ListItems.Add _
                (Key:=Format$(lngIndex) & "K")
            Set objAlbaranCompraItemMaterial = objAlbaranCompraItem
    
            With itmList
                'If objAlbaranCompraItem.IsNew Then
                '    .Text = "(new)"
    
                'Else
                '    .Text = objAlbaranCompraItem.AlbaranCompraItemID
    
                'End If
                .SmallIcon = GescomMain.mglIconosPequeños.ListImages("NuevoItem").Key
                
                If objAlbaranCompraItem.IsDeleted Then .SmallIcon = GescomMain.mglIconosPequeños.ListImages("EliminarItem").Key
                .SubItems(1) = Trim(objAlbaranCompraItemMaterial.Material)
                .SubItems(2) = objAlbaranCompraItem.Cantidad
                .SubItems(3) = FormatoMoneda(objAlbaranCompraItem.Bruto, GescomMain.objParametro.Moneda)
            End With
        Case ALBARANCOMPRAITEM_ARTICULO
            Set itmList = lvwAlbaranCompraArticulos.ListItems.Add _
                (Key:=Format$(lngIndex) & "K")
    
            With itmList
                .SmallIcon = GescomMain.mglIconosPequeños.ListImages("NuevoItem").Key
                
                If objAlbaranCompraItem.IsDeleted Then .SmallIcon = GescomMain.mglIconosPequeños.ListImages("EliminarItem").Key
                Set objAlbaranCompraItemArticulo = objAlbaranCompraItem
                .SubItems(1) = Space(4 - Len(Format$(lngIndex))) & Format$(lngIndex)
                .SubItems(2) = Trim(objAlbaranCompraItemArticulo.ArticuloColor)
                .SubItems(3) = objAlbaranCompraItemArticulo.CantidadT36
                .SubItems(4) = objAlbaranCompraItemArticulo.CantidadT38
                .SubItems(5) = objAlbaranCompraItemArticulo.CantidadT40
                .SubItems(6) = objAlbaranCompraItemArticulo.CantidadT42
                .SubItems(7) = objAlbaranCompraItemArticulo.CantidadT44
                .SubItems(8) = objAlbaranCompraItemArticulo.CantidadT46
                .SubItems(9) = objAlbaranCompraItemArticulo.CantidadT48
                .SubItems(10) = objAlbaranCompraItemArticulo.CantidadT50
                .SubItems(11) = objAlbaranCompraItemArticulo.CantidadT52
                .SubItems(12) = objAlbaranCompraItemArticulo.CantidadT54
                .SubItems(13) = objAlbaranCompraItemArticulo.CantidadT56
                Set objAlbaranCompraItemArticulo = Nothing
                .SubItems(14) = objAlbaranCompraItem.Cantidad
                .SubItems(15) = FormatoMoneda(objAlbaranCompraItem.PrecioCoste, "EUR", False)
                .SubItems(16) = FormatoMoneda(objAlbaranCompraItem.Bruto, "EUR")
            End With
            
        Case Else
            Err.Raise vbObjectError + 1001, "AlbaranCompraEdit LoadAlbaranCompraItems", "No existe el tipo de item de albaran de compra:" & objAlbaranCompraItem.Tipo & ". Avisar al personal técnico."
        End Select

    Next
    
    txtTotalBruto = FormatoMoneda(mobjAlbaranCompra.TotalBruto, GescomMain.objParametro.Moneda)
    txtCantidad = mobjAlbaranCompra.Cantidad
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdPedidos_Click()
    Dim frmPicker As PickerList
    Dim objSelectedItems As PickerItems
    Dim objPickerItemDisplay As PickerItemDisplay
    Dim strVistaPedidosPendientes As String
    Dim actAlbaranCompraItemTipo As AlbaranCompraItemTipos
  
    On Error GoTo ErrorManager
  
    ' No hacer nada si no se ha seleccionado un proveedor
    If mobjAlbaranCompra.ProveedorID = 0 Then Exit Sub
    
    Set frmPicker = New PickerList
  
    If mintCurFrame = 1 Then
        strVistaPedidosPendientes = "vPedidoCompraPendientes"
        actAlbaranCompraItemTipo = ALBARANCOMPRAITEM_MATERIAL
    Else
        strVistaPedidosPendientes = "vPedidoCompraArticuloPendientes"
        actAlbaranCompraItemTipo = ALBARANCOMPRAITEM_ARTICULO
    End If
    
    frmPicker.LoadData strVistaPedidosPendientes, mobjAlbaranCompra.ProveedorID, _
                                                  mobjAlbaranCompra.EmpresaID, _
                                                  mobjAlbaranCompra.TemporadaID
    frmPicker.Show vbModal
    Set objSelectedItems = frmPicker.SelectedItems
    Unload frmPicker
  
    If objSelectedItems Is Nothing Then Exit Sub
  
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    For Each objPickerItemDisplay In objSelectedItems
        ' Primero hay que comprobar que no está ya seleccionado anteriormente
        If Not DocumentoSeleccionado(objPickerItemDisplay.DocumentoID) Then _
            AlbaranDesdePedido objPickerItemDisplay.DocumentoID, actAlbaranCompraItemTipo
    Next
  
    Set frmPicker = Nothing
    Set objSelectedItems = Nothing
      
    LoadAlbaranCompraItems
  
    ' Muestro el puntero normal
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
    Exit Sub

End Sub

Private Sub AlbaranDesdePedido(PedidoItemID As Long, Tipo As AlbaranCompraItemTipos)
    Dim objAlbaranItem As AlbaranCompraItem
    Dim objPedidoItem As PedidoCompraItem

    Set objPedidoItem = New PedidoCompraItem
    Set objAlbaranItem = mobjAlbaranCompra.AlbaranCompraItems.Add(Tipo)
   
    objPedidoItem.Load PedidoItemID, Tipo
    With objAlbaranItem
        .BeginEdit GescomMain.objParametro.Moneda
        .AlbaranDesdePedido PedidoItemID
        .ApplyEdit
    End With
   
    Set objAlbaranItem = Nothing
    Set objPedidoItem = Nothing

End Sub

Private Function DocumentoSeleccionado(DocumentoID As Long) As Boolean
    Dim objAlbaranCompraItem As AlbaranCompraItem
    
    ' Se trata de buscar si existe alguna referencia de ese documento en alguna linea de
    ' albaranes y es nueva (no se ha actualizado).
    For Each objAlbaranCompraItem In mobjAlbaranCompra.AlbaranCompraItems
        If objAlbaranCompraItem.IsNew And _
           objAlbaranCompraItem.PedidoCompraItemID = DocumentoID Then
           DocumentoSeleccionado = True
           Exit Function
        End If
    Next
    
    DocumentoSeleccionado = False
    
    Set objAlbaranCompraItem = Nothing
    
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

