VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form ProveedorEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proveedores"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9525
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ProveedorEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEmail 
      Height          =   615
      Left            =   240
      Picture         =   "ProveedorEdit.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   8280
      TabIndex        =   34
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7080
      TabIndex        =   33
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   5880
      TabIndex        =   32
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proveedores"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin VB.ComboBox cboTipoProveedor 
         Height          =   315
         Left            =   6960
         TabIndex        =   22
         Text            =   "cboTipoProveedor"
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   5640
         TabIndex        =   24
         Top             =   2160
         Width           =   615
      End
      Begin VB.ComboBox cboMedioPago 
         Height          =   315
         Left            =   6960
         TabIndex        =   18
         Text            =   "cboMedioPago"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtCuentaContrapartida 
         Height          =   285
         Left            =   7680
         TabIndex        =   26
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton btnDatoComercial 
         Caption         =   "Datos c&omerciales"
         Height          =   615
         Left            =   4560
         TabIndex        =   29
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtDatoComercial 
         Height          =   1005
         Left            =   5760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   2640
         Width           =   3255
      End
      Begin VB.CommandButton btnCuentaBancaria 
         Caption         =   "Cuenta &bancaria"
         Height          =   615
         Left            =   160
         TabIndex        =   27
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtCuentaBancaria 
         Height          =   1005
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   2625
         Width           =   3255
      End
      Begin VB.ComboBox cboFormaPago 
         Height          =   315
         Left            =   6960
         TabIndex        =   12
         Text            =   "cboFormaPago"
         Top             =   1060
         Width           =   2055
      End
      Begin VB.ComboBox cboBanco 
         Height          =   315
         Left            =   6960
         TabIndex        =   8
         Text            =   "cboBanco"
         Top             =   700
         Width           =   2055
      End
      Begin VB.ComboBox cboTransportista 
         Height          =   315
         Left            =   6960
         TabIndex        =   4
         Text            =   "cboTransportista"
         Top             =   340
         Width           =   2055
      End
      Begin VB.TextBox txtCuentaContable 
         Height          =   285
         Left            =   3960
         TabIndex        =   16
         Top             =   1420
         Width           =   1335
      End
      Begin VB.TextBox txtDNINIF 
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   1420
         Width           =   1215
      End
      Begin VB.TextBox txtDireccion 
         Height          =   750
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   1800
         Width           =   3255
      End
      Begin VB.CommandButton btnDireccion 
         Caption         =   "Di&rección"
         Height          =   615
         Left            =   160
         TabIndex        =   19
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtTitular 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   700
         Width           =   4095
      End
      Begin VB.TextBox txtContacto 
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   1060
         Width           =   4095
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   340
         Width           =   4095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de proveedor"
         Height          =   195
         Left            =   5640
         TabIndex        =   21
         Top             =   1815
         Width           =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Código corto"
         Height          =   195
         Left            =   4560
         TabIndex        =   23
         Top             =   2160
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Medio de pago"
         Height          =   195
         Left            =   5640
         TabIndex        =   17
         Top             =   1455
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contrapartida"
         Height          =   195
         Left            =   6600
         TabIndex        =   25
         Top             =   2175
         Width           =   1005
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Forma de pago"
         Height          =   195
         Left            =   5640
         TabIndex        =   11
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Left            =   5640
         TabIndex        =   7
         Top             =   720
         Width           =   435
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Transportista"
         Height          =   195
         Left            =   5640
         TabIndex        =   3
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Contable"
         Height          =   195
         Left            =   2640
         TabIndex        =   15
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Titular"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "DNI/NIF"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Contacto"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   660
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
   Begin MSMAPI.MAPIMessages mpmErrorMail 
      Left            =   2160
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession mpsErrorMail 
      Left            =   1440
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "ProveedorEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private mintBancoSelStart As Integer
Private mintTransportistaSelStart As Integer
Private mintFormaPagoSelStart As Integer
Private mintMedioPagoSelStart As Integer
Private mintTipoProveedorSelStart As Integer

Private WithEvents mobjProveedor As Proveedor
Attribute mobjProveedor.VB_VarHelpID = -1

Public Sub Component(ProveedorObject As Proveedor)

    Set mobjProveedor = ProveedorObject

End Sub

Private Sub btnDireccion_Click()
    
    Dim frmDireccion As DireccionEdit
  
    Set frmDireccion = New DireccionEdit
    frmDireccion.Component mobjProveedor.Direccion
    frmDireccion.Show vbModal
    txtDireccion.Text = mobjProveedor.Direccion.DireccionText
  
End Sub

Private Sub btnCuentaBancaria_Click()
    
    Dim frmCuentaBancaria As CuentaBancariaEdit
  
    Set frmCuentaBancaria = New CuentaBancariaEdit
    frmCuentaBancaria.Component mobjProveedor.CuentaBancaria
    frmCuentaBancaria.Show vbModal
    txtCuentaBancaria.Text = mobjProveedor.CuentaBancaria.CuentaBancariaText
  
End Sub

Private Sub btnDatoComercial_Click()
    
    Dim frmDatoComercial As DatoComercialEdit
  
    Set frmDatoComercial = New DatoComercialEdit
    frmDatoComercial.Component mobjProveedor.DatoComercial
    frmDatoComercial.Show vbModal
    txtDatoComercial.Text = mobjProveedor.DatoComercial.DatoComercialText
  
End Sub

Private Sub cmdApply_Click()
    
    On Error GoTo ErrorManager

    mobjProveedor.ApplyEdit
    mobjProveedor.BeginEdit
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    Dim Respuesta As VbMsgBoxResult
    
    If mobjProveedor.IsDirty And Not mobjProveedor.IsNew Then
        Respuesta = MostrarMensaje(MSG_MODIFY)
        If Respuesta = vbYes Then
            mobjProveedor.CancelEdit
            Unload Me
        End If
    Else
        mobjProveedor.CancelEdit
        Unload Me
    End If

End Sub

Private Sub cmdEmail_Click()
    On Error GoTo SendErrorMailError

    If mobjProveedor.Direccion.EMAIL = vbNullString Then
        Err.Raise vbObjectError + 1001, "El proveedor no tiene asignada dirección de eMail"
        Exit Sub
    End If
    ' Sign on to the mail system.
    mpsErrorMail.SignOn

    ' Send the message.
    mpmErrorMail.SessionID = mpsErrorMail.SessionID
    mpmErrorMail.Compose
    mpmErrorMail.RecipDisplayName = mobjProveedor.Nombre
    mpmErrorMail.RecipAddress = mobjProveedor.Direccion.EMAIL
    mpmErrorMail.AddressResolveUI = False
    'mpmErrorMail.MsgSubject = "Subject"
    'mpmErrorMail.MsgNoteText = "Mensaje..."
    mpmErrorMail.Send True

    ' Sign off of the mail system.
    mpsErrorMail.SignOff
    Exit Sub

SendErrorMailError:
    ManageErrors (Me.Caption)
End Sub

 Private Sub cmdOK_Click()
    
    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    mobjProveedor.ApplyEdit
    Unload Me

    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrorManager

    DisableX Me
    
    mflgLoading = True
    With mobjProveedor
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Proveedor [(nuevo)]"

        Else
            Caption = "Proveedor [" & .Nombre & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        'txtProveedorID = .ProveedorID
        txtNombre = .Nombre
        txtCodigo = .Codigo
        txtTitular = .Titular
        txtContacto = .Contacto
        txtDNINIF = .DNINIF
        txtDireccion = .Direccion.DireccionText
        txtCuentaBancaria = .CuentaBancaria.CuentaBancariaText
        txtDatoComercial = .DatoComercial.DatoComercialText
        txtCuentaContable = .CuentaContable
        txtCuentaContrapartida = .CuentaContrapartida
   
        LoadCombo cboTransportista, .Transportistas
        cboTransportista.Text = .Transportista

        LoadCombo cboBanco, .Bancos
        cboBanco.Text = .Banco

        LoadCombo cboFormaPago, .FormasDePago
        cboFormaPago.Text = .FormaDePago

        LoadCombo cboMedioPago, .MediosPago
        cboMedioPago.Text = .MedioPago
    
        LoadCombo cboTipoProveedor, .TiposProveedor
        cboTipoProveedor.Text = .TipoProveedor
    
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

Private Sub mobjProveedor_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub txtContacto_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtContacto
        
End Sub

Private Sub txtCuentaBancaria_DblClick()

    Call btnCuentaBancaria_Click
    
End Sub

Private Sub txtCuentaContable_Change()

    If Not mflgLoading Then _
        TextChange txtCuentaContable, mobjProveedor, "CuentaContable"
        
End Sub

Private Sub txtCuentaContable_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCuentaContable
        
End Sub

Private Sub txtCuentaContable_LostFocus()

    txtCuentaContable = TextLostFocus(txtCuentaContable, mobjProveedor, "CuentaContable")
    
End Sub

Private Sub txtCuentaContrapartida_Change()

    If Not mflgLoading Then _
        TextChange txtCuentaContrapartida, mobjProveedor, "CuentaContrapartida"
        
End Sub

Private Sub txtCuentaContrapartida_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCuentaContrapartida
        
End Sub

Private Sub txtCuentaContrapartida_LostFocus()

    txtCuentaContrapartida = TextLostFocus(txtCuentaContrapartida, mobjProveedor, "CuentaContrapartida")
    
End Sub

Private Sub txtDatoComercial_DblClick()

    Call btnDatoComercial_Click
    
End Sub

Private Sub txtDireccion_DblClick()

    Call btnDireccion_Click
    
End Sub

Private Sub txtDNINIF_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtDNINIF
        
End Sub

Private Sub txtNombre_Change()

    If Not mflgLoading Then _
        TextChange txtNombre, mobjProveedor, "Nombre"

End Sub

Private Sub txtNombre_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtNombre
        
End Sub

Private Sub txtNombre_LostFocus()

    txtNombre = TextLostFocus(txtNombre, mobjProveedor, "Nombre")

End Sub

Private Sub txtCodigo_Change()

    If Not mflgLoading Then _
        TextChange txtCodigo, mobjProveedor, "Codigo"

End Sub

Private Sub txtCodigo_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCodigo
        
End Sub

Private Sub txtCodigo_LostFocus()

    txtCodigo = TextLostFocus(txtCodigo, mobjProveedor, "Codigo")

End Sub

Private Sub txtTitular_Change()

    If Not mflgLoading Then _
        TextChange txtTitular, mobjProveedor, "Titular"

End Sub

Private Sub txtTitular_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtTitular
        
End Sub

Private Sub txtTitular_LostFocus()

    txtTitular = TextLostFocus(txtTitular, mobjProveedor, "Titular")

End Sub

Private Sub txtContacto_Change()

    If Not mflgLoading Then _
        TextChange txtContacto, mobjProveedor, "Contacto"

End Sub

Private Sub txtContacto_LostFocus()

    txtContacto = TextLostFocus(txtContacto, mobjProveedor, "Contacto")

End Sub

Private Sub txtDNINIF_Change()

    If Not mflgLoading Then _
        TextChange txtDNINIF, mobjProveedor, "DNINIF"

End Sub

Private Sub txtDNINIF_LostFocus()

    txtDNINIF = TextLostFocus(txtDNINIF, mobjProveedor, "DNINIF")

End Sub

Private Sub cboTransportista_Click()

    If mflgLoading Then Exit Sub
    mobjProveedor.Transportista = cboTransportista.Text

End Sub

Private Sub cboBanco_Click()

    If mflgLoading Then Exit Sub
    mobjProveedor.Banco = cboBanco.Text

End Sub

Private Sub cboFormaPago_Click()

    If mflgLoading Then Exit Sub
    mobjProveedor.FormaDePago = cboFormaPago.Text

End Sub

Private Sub cboMedioPago_Click()

    If mflgLoading Then Exit Sub
    mobjProveedor.MedioPago = cboMedioPago.Text

End Sub

Private Sub cboTipoProveedor_Click()

    If mflgLoading Then Exit Sub
    mobjProveedor.TipoProveedor = cboTipoProveedor.Text

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   
   IsList = False
   
End Function

Private Sub cboBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintBancoSelStart = cboBanco.SelStart
End Sub

Private Sub cboBanco_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintBancoSelStart, cboBanco
    
End Sub

Private Sub cboTransportista_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintTransportistaSelStart = cboTransportista.SelStart
End Sub

Private Sub cboTransportista_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintTransportistaSelStart, cboTransportista
    
End Sub

Private Sub cboFormaPago_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintFormaPagoSelStart = cboFormaPago.SelStart
End Sub

Private Sub cboFormaPago_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintFormaPagoSelStart, cboFormaPago
    
End Sub

Private Sub cboMedioPago_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintMedioPagoSelStart = cboMedioPago.SelStart
End Sub

Private Sub cboMedioPago_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintMedioPagoSelStart, cboMedioPago
    
End Sub

Private Sub cboTipoProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintTipoProveedorSelStart = cboTipoProveedor.SelStart
End Sub

Private Sub cboTipoProveedor_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintTipoProveedorSelStart, cboTipoProveedor
    
End Sub



