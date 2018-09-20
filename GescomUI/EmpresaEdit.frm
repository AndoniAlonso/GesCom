VERSION 5.00
Begin VB.Form EmpresaEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Empresas"
   ClientHeight    =   4170
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   8820
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EmpresaEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   7560
      TabIndex        =   35
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6360
      TabIndex        =   34
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   33
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la Empresa"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      Begin VB.TextBox txtCodigoContawin 
         Height          =   285
         Left            =   4680
         TabIndex        =   26
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtEmpresaContawin 
         Height          =   285
         Left            =   1680
         TabIndex        =   24
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtOrdenCorte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6960
         TabIndex        =   28
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tratamiento del IVA"
         Height          =   1095
         Left            =   4440
         TabIndex        =   31
         Top             =   2160
         Width           =   3975
         Begin VB.CheckBox chkAnularIVA 
            Caption         =   "Suprimir tratamiento del &IVA"
            Height          =   375
            Left            =   240
            TabIndex        =   32
            Top             =   480
            Width           =   3375
         End
      End
      Begin VB.TextBox txtFacturaVentas 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6960
         TabIndex        =   19
         Top             =   1420
         Width           =   1215
      End
      Begin VB.TextBox txtAlbaranVentas 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6960
         TabIndex        =   16
         Top             =   1060
         Width           =   1215
      End
      Begin VB.TextBox txtPedidoVentas 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6960
         TabIndex        =   10
         Top             =   700
         Width           =   1215
      End
      Begin VB.TextBox txtFacturaCompras 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4080
         TabIndex        =   18
         Top             =   1420
         Width           =   1215
      End
      Begin VB.TextBox txtAlbaranCompras 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4080
         TabIndex        =   14
         Top             =   1060
         Width           =   1215
      End
      Begin VB.TextBox txtPedidoCompras 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4080
         TabIndex        =   8
         Top             =   700
         Width           =   1215
      End
      Begin VB.TextBox txtDireccion 
         Height          =   1005
         Left            =   1080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   2145
         Width           =   3135
      End
      Begin VB.CommandButton btnDireccion 
         Caption         =   "Di&rección"
         Height          =   495
         Left            =   160
         TabIndex        =   29
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtActividad 
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   1060
         Width           =   1575
      End
      Begin VB.TextBox txtDNINIF 
         Height          =   285
         Left            =   1080
         TabIndex        =   20
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtTitular 
         Height          =   285
         Left            =   4080
         TabIndex        =   4
         Top             =   340
         Width           =   4095
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   700
         Width           =   495
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   340
         Width           =   1575
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Código ContaWin"
         Height          =   195
         Left            =   3240
         TabIndex        =   25
         Top             =   1815
         Width           =   1245
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Empresa ContaWin"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   1815
         Width           =   1365
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Orden de corte"
         Height          =   195
         Left            =   5520
         TabIndex        =   27
         Top             =   1815
         Width           =   1095
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Factura Ventas"
         Height          =   195
         Left            =   5520
         TabIndex        =   22
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Pedido Ventas"
         Height          =   195
         Left            =   5520
         TabIndex        =   9
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Albarán Ventas"
         Height          =   195
         Left            =   5520
         TabIndex        =   15
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Factura Compras"
         Height          =   195
         Left            =   2760
         TabIndex        =   21
         Top             =   1440
         Width           =   1230
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Albarán Compras"
         Height          =   195
         Left            =   2760
         TabIndex        =   13
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Actividad"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Pedido Compras"
         Height          =   195
         Left            =   2760
         TabIndex        =   7
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DNI/CIF"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Titular"
         Height          =   195
         Left            =   2760
         TabIndex        =   3
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   555
      End
   End
End
Attribute VB_Name = "EmpresaEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjEmpresa As Empresa
Attribute mobjEmpresa.VB_VarHelpID = -1

Public Sub Component(EmpresaObject As Empresa)

    Set mobjEmpresa = EmpresaObject

End Sub

Private Sub btnDireccion_Click()
    Dim frmDireccion As DireccionEdit
  
    Set frmDireccion = New DireccionEdit
    frmDireccion.Component mobjEmpresa.Direccion
    frmDireccion.Show vbModal
    txtDireccion.Text = mobjEmpresa.Direccion.DireccionText
  
End Sub

Private Sub cmdApply_Click()

    On Error GoTo ErrorManager

    mobjEmpresa.ApplyEdit
    mobjEmpresa.BeginEdit
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()
    Dim Respuesta As VbMsgBoxResult
    
    On Error GoTo ErrorManager
    
    If mobjEmpresa.IsDirty And Not mobjEmpresa.IsNew Then
        Respuesta = MostrarMensaje(MSG_MODIFY)
        If Respuesta = vbYes Then
            mobjEmpresa.CancelEdit
            Unload Me
        End If
    Else
        mobjEmpresa.CancelEdit
        Unload Me
    End If
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
    Exit Sub

End Sub

Private Sub cmdOK_Click()

    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    mobjEmpresa.ApplyEdit
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Private Sub chkAnularIVA_Click()
    
    If mflgLoading Then Exit Sub
    
    mobjEmpresa.AnularIVA = IIf(chkAnularIVA.Value = 1, True, False)

End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrorManager

    DisableX Me
    
    mflgLoading = True
    With mobjEmpresa
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Empresa [(nueva)]"

        Else
            Caption = "Empresa [" & .Nombre & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        'txtEmpresaID = .EmpresaID
        txtNombre = .Nombre
        txtCodigo = .Codigo
        txtTitular = .Titular
        txtDNINIF = .DNINIF
        txtActividad = .Actividad
        txtPedidoCompras = .PedidoCompras
        txtPedidoVentas = .PedidoVentas
        txtAlbaranCompras = .AlbaranCompras
        txtAlbaranVentas = .AlbaranVentas
        txtFacturaCompras = .FacturaCompras
        txtFacturaVentas = .FacturaVentas
        txtOrdenCorte = .OrdenCorte
        txtDireccion = .Direccion.DireccionText
        txtEmpresaContawin = .EmpresaContawin
        txtCodigoContawin = .CodigoContawin
        chkAnularIVA.Value = IIf(.AnularIVA, 1, 0)
  
        .BeginEdit
    
    End With
  
    mflgLoading = False
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub mobjEmpresa_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub txtActividad_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtActividad
        
End Sub

Private Sub txtAlbaranCompras_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtAlbaranCompras
        
End Sub

Private Sub txtAlbaranVentas_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtAlbaranVentas
        
End Sub

Private Sub txtCodigo_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCodigo
        
End Sub

Private Sub txtCodigoContawin_Change()

    If Not mflgLoading Then _
        TextChange txtCodigoContawin, mobjEmpresa, "CodigoContawin"

End Sub

Private Sub txtCodigoContawin_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCodigoContawin
        
End Sub

Private Sub txtCodigoContawin_LostFocus()

    txtCodigoContawin = TextLostFocus(txtCodigoContawin, mobjEmpresa, "CodigoContawin")

End Sub

Private Sub txtDireccion_DblClick()

    Call btnDireccion_Click
    
End Sub

Private Sub txtDNINIF_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtDNINIF
        
End Sub

Private Sub txtEmpresaContawin_Change()

    If Not mflgLoading Then _
        TextChange txtEmpresaContawin, mobjEmpresa, "EmpresaContawin"

End Sub

Private Sub txtEmpresaContawin_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtEmpresaContawin
        
End Sub

Private Sub txtEmpresaContawin_LostFocus()

    txtEmpresaContawin = TextLostFocus(txtEmpresaContawin, mobjEmpresa, "EmpresaContawin")

End Sub

Private Sub txtFacturaCompras_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtFacturaCompras
        
End Sub

Private Sub txtNombre_Change()

    If Not mflgLoading Then _
        TextChange txtNombre, mobjEmpresa, "Nombre"

End Sub

Private Sub txtNombre_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtNombre
        
End Sub

Private Sub txtNombre_LostFocus()

    txtNombre = TextLostFocus(txtNombre, mobjEmpresa, "Nombre")

End Sub

Private Sub txtCodigo_Change()

    If Not mflgLoading Then _
        TextChange txtCodigo, mobjEmpresa, "Codigo"

End Sub

Private Sub txtCodigo_LostFocus()

    txtCodigo = TextLostFocus(txtCodigo, mobjEmpresa, "Codigo")

End Sub

Private Sub txtPedidoCompras_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPedidoCompras
        
End Sub

Private Sub txtPedidoVentas_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPedidoVentas
        
End Sub

Private Sub txtTitular_Change()

    If Not mflgLoading Then _
        TextChange txtTitular, mobjEmpresa, "Titular"

End Sub

Private Sub txtTitular_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtTitular
        
End Sub

Private Sub txtTitular_LostFocus()

    txtTitular = TextLostFocus(txtTitular, mobjEmpresa, "Titular")

End Sub

Private Sub txtDNINIF_Change()

    If Not mflgLoading Then _
        TextChange txtDNINIF, mobjEmpresa, "DNINIF"

End Sub

Private Sub txtDNINIF_LostFocus()

    txtDNINIF = TextLostFocus(txtDNINIF, mobjEmpresa, "DNINIF")

End Sub

Private Sub txtActividad_Change()

    If Not mflgLoading Then _
        TextChange txtActividad, mobjEmpresa, "Actividad"

End Sub

Private Sub txtActividad_LostFocus()

    txtActividad = TextLostFocus(txtActividad, mobjEmpresa, "Actividad")

End Sub

Private Sub txtPedidoCompras_Change()

    If Not mflgLoading Then _
        TextChange txtPedidoCompras, mobjEmpresa, "PedidoCompras"

End Sub

Private Sub txtPedidoCompras_LostFocus()

    txtPedidoCompras = TextLostFocus(txtPedidoCompras, mobjEmpresa, "PedidoCompras")

End Sub

Private Sub txtAlbaranCompras_Change()

    If Not mflgLoading Then _
        TextChange txtAlbaranCompras, mobjEmpresa, "AlbaranCompras"

End Sub

Private Sub txtAlbaranCompras_LostFocus()

    txtAlbaranCompras = TextLostFocus(txtAlbaranCompras, mobjEmpresa, "AlbaranCompras")

End Sub

Private Sub txtFacturaCompras_Change()

    If Not mflgLoading Then _
        TextChange txtFacturaCompras, mobjEmpresa, "FacturaCompras"

End Sub

Private Sub txtFacturaCompras_LostFocus()

    txtFacturaCompras = TextLostFocus(txtFacturaCompras, mobjEmpresa, "FacturaCompras")

End Sub

Private Sub txtPedidoVentas_Change()

    If Not mflgLoading Then _
        TextChange txtPedidoVentas, mobjEmpresa, "PedidoVentas"

End Sub

Private Sub txtPedidoVentas_LostFocus()

    txtPedidoVentas = TextLostFocus(txtPedidoVentas, mobjEmpresa, "PedidoVentas")

End Sub

Private Sub txtAlbaranVentas_Change()

    If Not mflgLoading Then _
        TextChange txtAlbaranVentas, mobjEmpresa, "AlbaranVentas"

End Sub

Private Sub txtAlbaranVentas_LostFocus()

    txtAlbaranVentas = TextLostFocus(txtAlbaranVentas, mobjEmpresa, "AlbaranVentas")

End Sub

Private Sub txtFacturaVentas_Change()

    If Not mflgLoading Then _
        TextChange txtFacturaVentas, mobjEmpresa, "FacturaVentas"

End Sub

Private Sub txtFacturaVentas_LostFocus()

    txtFacturaVentas = TextLostFocus(txtFacturaVentas, mobjEmpresa, "FacturaVentas")

End Sub

Private Sub txtFacturaVentas_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtFacturaVentas
        
End Sub

Private Sub txtOrdenCorte_Change()

    If Not mflgLoading Then _
        TextChange txtOrdenCorte, mobjEmpresa, "OrdenCorte"

End Sub

Private Sub txtOrdenCorte_LostFocus()

    txtOrdenCorte = TextLostFocus(txtOrdenCorte, mobjEmpresa, "OrdenCorte")

End Sub

Private Sub txtOrdenCorte_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtOrdenCorte
        
End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
    
    IsList = False
    
End Function
