VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Begin VB.Form ClienteEdit 
   Caption         =   "Clientes"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ClienteEdit.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   9495
   Begin VB.CommandButton cmdEmail 
      Height          =   615
      Left            =   240
      Picture         =   "ClienteEdit.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5040
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Cliente"
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.TextBox txtPorcFacturacionAB 
         Height          =   285
         Left            =   1440
         TabIndex        =   22
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtDiaPago1 
         Height          =   285
         Left            =   6960
         TabIndex        =   18
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtDiaPago2 
         Height          =   285
         Left            =   7440
         TabIndex        =   19
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtDiaPago3 
         Height          =   285
         Left            =   7920
         TabIndex        =   20
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton btnDatoComercialB 
         Caption         =   "Datos Co&merciales B"
         Height          =   615
         Left            =   4600
         TabIndex        =   29
         Top             =   3135
         Width           =   1095
      End
      Begin VB.TextBox txtDatoComercialB 
         Height          =   1005
         Left            =   5760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   3120
         Width           =   3255
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   4095
      End
      Begin VB.TextBox txtContacto 
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   945
         Width           =   4095
      End
      Begin VB.TextBox txtTitular 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   585
         Width           =   4095
      End
      Begin VB.CommandButton btnDireccionFiscal 
         Caption         =   "Dirección &Fiscal"
         Height          =   615
         Left            =   160
         TabIndex        =   23
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtDireccionFiscal 
         Height          =   1005
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   2025
         Width           =   3255
      End
      Begin VB.TextBox txtDNINIF 
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   1305
         Width           =   1215
      End
      Begin VB.TextBox txtCuentaContable 
         Height          =   285
         Left            =   3960
         TabIndex        =   16
         Top             =   1305
         Width           =   1335
      End
      Begin VB.TextBox txtDireccionEntrega 
         Height          =   1005
         Left            =   5760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   2025
         Width           =   3255
      End
      Begin VB.CommandButton btnDireccionEntrega 
         Caption         =   "Di&rección Entrega"
         Height          =   615
         Left            =   4600
         TabIndex        =   25
         Top             =   2040
         Width           =   1095
      End
      Begin VB.ComboBox cboTransportista 
         Height          =   315
         Left            =   6960
         TabIndex        =   4
         Text            =   "cboTransportista"
         Top             =   225
         Width           =   2055
      End
      Begin VB.ComboBox cboRepresentante 
         Height          =   315
         Left            =   6960
         TabIndex        =   8
         Text            =   "cboRepresentante"
         Top             =   585
         Width           =   2055
      End
      Begin VB.ComboBox cboFormaPago 
         Height          =   315
         Left            =   6960
         TabIndex        =   12
         Text            =   "cboFormaPago"
         Top             =   945
         Width           =   2055
      End
      Begin VB.TextBox txtCuentaBancaria 
         Height          =   630
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   4185
         Width           =   3255
      End
      Begin VB.CommandButton btnCuentaBancaria 
         Caption         =   "Cuenta &Bancaria"
         Height          =   495
         Left            =   165
         TabIndex        =   31
         Top             =   4200
         Width           =   975
      End
      Begin VB.TextBox txtDatoComercial 
         Height          =   1005
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   3105
         Width           =   3255
      End
      Begin VB.CommandButton btnDatoComercial 
         Caption         =   "Datos C&omerciales"
         Height          =   615
         Left            =   165
         TabIndex        =   27
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Facturacion A/B"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   1695
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Dias de pago"
         Height          =   195
         Left            =   5640
         TabIndex        =   17
         Top             =   1335
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Contacto"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "DNI/NIF"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Titular"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Contable"
         Height          =   195
         Left            =   2640
         TabIndex        =   15
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Transportista"
         Height          =   195
         Left            =   5640
         TabIndex        =   3
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Representante"
         Height          =   195
         Left            =   5640
         TabIndex        =   7
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pago"
         Height          =   195
         Left            =   5640
         TabIndex        =   11
         Top             =   960
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   5880
      TabIndex        =   34
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7080
      TabIndex        =   35
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   8280
      TabIndex        =   36
      Top             =   5040
      Width           =   1095
   End
   Begin MSMAPI.MAPIMessages mpmErrorMail 
      Left            =   2040
      Top             =   4800
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
      Left            =   1320
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "ClienteEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private mintRepresentanteSelStart As Integer
Private mintTransportistaSelStart As Integer
Private mintFormaPagoSelStart As Integer

Private WithEvents mobjCliente As Cliente
Attribute mobjCliente.VB_VarHelpID = -1


Private Sub cmdEmail_Click()
    On Error GoTo SendErrorMailError

    If mobjCliente.DireccionFiscal.EMAIL = vbNullString Then
        Err.Raise vbObjectError + 1001, "El cliente no tiene asignada dirección de eMail"
        Exit Sub
    End If
    ' Sign on to the mail system.
    mpsErrorMail.SignOn

    ' Send the message.
    mpmErrorMail.SessionID = mpsErrorMail.SessionID
    mpmErrorMail.Compose
    mpmErrorMail.RecipDisplayName = mobjCliente.Nombre
    mpmErrorMail.RecipAddress = mobjCliente.DireccionFiscal.EMAIL
    mpmErrorMail.AddressResolveUI = False
    mpmErrorMail.Send True

    ' Sign off of the mail system.
    mpsErrorMail.SignOff
    Exit Sub

SendErrorMailError:
    ManageErrors (Me.Caption)
End Sub

Public Sub Component(ClienteObject As Cliente)

    Set mobjCliente = ClienteObject

End Sub

Private Sub btnDireccionFiscal_Click()
    
    Dim frmDireccionFiscal As DireccionEdit
  
    Set frmDireccionFiscal = New DireccionEdit
    frmDireccionFiscal.Component mobjCliente.DireccionFiscal
    frmDireccionFiscal.Show vbModal
    txtDireccionFiscal.Text = mobjCliente.DireccionFiscal.DireccionText
  
End Sub

Private Sub btnDireccionEntrega_Click()
    
    Dim frmDireccionEntrega As DireccionEdit
  
    Set frmDireccionEntrega = New DireccionEdit
    frmDireccionEntrega.Component mobjCliente.DireccionEntrega
    frmDireccionEntrega.Show vbModal
    txtDireccionEntrega.Text = mobjCliente.DireccionEntrega.DireccionText
  
End Sub

Private Sub btnCuentaBancaria_Click()
    
    Dim frmCuentaBancaria As CuentaBancariaEdit
  
    Set frmCuentaBancaria = New CuentaBancariaEdit
    frmCuentaBancaria.Component mobjCliente.CuentaBancaria
    frmCuentaBancaria.Show vbModal
    txtCuentaBancaria.Text = mobjCliente.CuentaBancaria.CuentaBancariaText
  
End Sub

Private Sub btnDatoComercial_Click()
    
    Dim frmDatoComercial As DatoComercialEdit
  
    Set frmDatoComercial = New DatoComercialEdit
    frmDatoComercial.Component mobjCliente.DatoComercial
    frmDatoComercial.Show vbModal
    txtDatoComercial.Text = mobjCliente.DatoComercial.DatoComercialText
  
End Sub

Private Sub btnDatoComercialB_Click()
    Dim frmDatoComercial As DatoComercialEdit
  
    Set frmDatoComercial = New DatoComercialEdit
    frmDatoComercial.Component mobjCliente.DatoComercialB
    frmDatoComercial.Show vbModal
    txtDatoComercialB.Text = mobjCliente.DatoComercialB.DatoComercialText
  
End Sub

Private Sub cmdApply_Click()
    
    On Error GoTo ErrorManager

    mobjCliente.ApplyEdit
    mobjCliente.BeginEdit
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    Dim Respuesta As VbMsgBoxResult
    
    If mobjCliente.IsDirty And Not mobjCliente.IsNew Then
        Respuesta = MostrarMensaje(MSG_MODIFY)
        If Respuesta = vbYes Then
            mobjCliente.CancelEdit
            Unload Me
        End If
    Else
        mobjCliente.CancelEdit
        Unload Me
    End If

End Sub

 Private Sub cmdOK_Click()
    
    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass

    mobjCliente.ApplyEdit
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
    With mobjCliente
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Cliente [(nuevo)]"

        Else
            Caption = "Cliente [" & .Nombre & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        'txtClienteID = .ClienteID
        txtNombre = .Nombre
        txtTitular = .Titular
        txtContacto = .Contacto
        txtDNINIF = .DNINIF
        txtDireccionFiscal = .DireccionFiscal.DireccionText
        txtDireccionEntrega = .DireccionEntrega.DireccionText
        txtCuentaBancaria = .CuentaBancaria.CuentaBancariaText
        txtDatoComercial = .DatoComercial.DatoComercialText
        txtDatoComercialB = .DatoComercialB.DatoComercialText
        txtCuentaContable = .CuentaContable
      
        LoadCombo cboTransportista, .Transportistas
        cboTransportista.Text = .Transportista

        LoadCombo cboRepresentante, .Representantes
        cboRepresentante.Text = .Representante

        LoadCombo cboFormaPago, .FormasPago
        cboFormaPago.Text = .FormaPago
        txtDiaPago1 = .DiaPago1
        txtDiaPago2 = .DiaPago2
        txtDiaPago3 = .DiaPago3
        txtPorcFacturacionAB = .PorcFacturacionAB

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

Private Sub mobjCliente_Valid(IsValid As Boolean)

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
        TextChange txtCuentaContable, mobjCliente, "CuentaContable"
        
End Sub

Private Sub txtCuentaContable_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCuentaContable
        
End Sub

Private Sub txtCuentaContable_LostFocus()

    txtCuentaContable = TextLostFocus(txtCuentaContable, mobjCliente, "CuentaContable")
    
End Sub

Private Sub txtDatoComercial_DblClick()

    Call btnDatoComercial_Click
    
End Sub

Private Sub txtDatoComercialB_DblClick()

    Call btnDatoComercialB_Click
    
End Sub

Private Sub txtDiaPago1_Change()

    If Not mflgLoading Then _
        TextChange txtDiaPago1, mobjCliente, "DiaPago1"

End Sub

Private Sub txtDiaPago1_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtDiaPago1
        
End Sub

Private Sub txtDiaPago1_LostFocus()

    txtDiaPago1 = TextLostFocus(txtDiaPago1, mobjCliente, "DiaPago1")

End Sub

Private Sub txtDiaPago2_Change()

    If Not mflgLoading Then _
        TextChange txtDiaPago2, mobjCliente, "DiaPago2"

End Sub

Private Sub txtDiaPago2_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtDiaPago2
        
End Sub

Private Sub txtDiaPago2_LostFocus()

    txtDiaPago2 = TextLostFocus(txtDiaPago2, mobjCliente, "DiaPago2")

End Sub

Private Sub txtDiaPago3_Change()

    If Not mflgLoading Then _
        TextChange txtDiaPago3, mobjCliente, "DiaPago3"

End Sub

Private Sub txtDiaPago3_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtDiaPago3
        
End Sub

Private Sub txtDiaPago3_LostFocus()

    txtDiaPago3 = TextLostFocus(txtDiaPago3, mobjCliente, "DiaPago3")

End Sub

Private Sub txtPorcFacturacionAB_Change()

    If Not mflgLoading Then _
        TextChange txtPorcFacturacionAB, mobjCliente, "PorcFacturacionAB"

End Sub

Private Sub txtPorcFacturacionAB_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPorcFacturacionAB
        
End Sub

Private Sub txtPorcFacturacionAB_LostFocus()

    txtPorcFacturacionAB = TextLostFocus(txtPorcFacturacionAB, mobjCliente, "PorcFacturacionAB")

End Sub

Private Sub txtDireccionEntrega_DblClick()

    Call btnDireccionEntrega_Click
    
End Sub

Private Sub txtDireccionFiscal_DblClick()

    Call btnDireccionFiscal_Click
    
End Sub

Private Sub txtDNINIF_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtDNINIF
        
End Sub

Private Sub txtNombre_Change()

    If Not mflgLoading Then _
        TextChange txtNombre, mobjCliente, "Nombre"

End Sub

Private Sub txtNombre_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtNombre
        
End Sub

Private Sub txtNombre_LostFocus()

    txtNombre = TextLostFocus(txtNombre, mobjCliente, "Nombre")

End Sub

Private Sub txtTitular_Change()

    If Not mflgLoading Then _
        TextChange txtTitular, mobjCliente, "Titular"

End Sub

Private Sub txtTitular_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtTitular
        
End Sub

Private Sub txtTitular_LostFocus()

    txtTitular = TextLostFocus(txtTitular, mobjCliente, "Titular")

End Sub

Private Sub txtContacto_Change()

    If Not mflgLoading Then _
        TextChange txtContacto, mobjCliente, "Contacto"

End Sub

Private Sub txtContacto_LostFocus()

    txtContacto = TextLostFocus(txtContacto, mobjCliente, "Contacto")

End Sub

Private Sub txtDNINIF_Change()

    If Not mflgLoading Then _
        TextChange txtDNINIF, mobjCliente, "DNINIF"

End Sub

Private Sub txtDNINIF_LostFocus()

    txtDNINIF = TextLostFocus(txtDNINIF, mobjCliente, "DNINIF")

End Sub

Private Sub cboTransportista_Click()

    If mflgLoading Then Exit Sub
    mobjCliente.Transportista = cboTransportista.Text

End Sub

Private Sub cboRepresentante_Click()

    If mflgLoading Then Exit Sub
    mobjCliente.Representante = cboRepresentante.Text

End Sub

Private Sub cboFormaPago_Click()

    If mflgLoading Then Exit Sub
    mobjCliente.FormaPago = cboFormaPago.Text

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   
   IsList = False
   
End Function

Private Sub cboRepresentante_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintRepresentanteSelStart = cboRepresentante.SelStart
End Sub

Private Sub cboRepresentante_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintRepresentanteSelStart, cboRepresentante
    
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

