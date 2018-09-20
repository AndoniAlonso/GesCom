VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CobroPagoEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CobrosPagos"
   ClientHeight    =   4665
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CobroPagoEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   24
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6480
      TabIndex        =   25
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   7680
      TabIndex        =   26
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CheckBox chkHayFactura 
      Caption         =   "Está relacionado con una Factura"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del cobro/pago"
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VB.ComboBox cboMedioPago 
         Height          =   315
         Left            =   6240
         TabIndex        =   6
         Text            =   "cboMedioPago"
         Top             =   720
         Width           =   2055
      End
      Begin VB.ComboBox cboPersona 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Text            =   "cboPersona"
         Top             =   340
         Width           =   3015
      End
      Begin VB.Frame Frame4 
         Caption         =   "Datos de gestión"
         Height          =   855
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   8175
         Begin VB.TextBox txtImporte 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3720
            TabIndex        =   11
            Top             =   340
            Width           =   1455
         End
         Begin VB.TextBox txtNumeroGiro 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6240
            TabIndex        =   13
            Top             =   340
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpVencimiento 
            Height          =   315
            Left            =   1320
            TabIndex        =   9
            Top             =   340
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   77725697
            CurrentDate     =   36938
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Importe"
            Height          =   195
            Left            =   3000
            TabIndex        =   10
            Top             =   360
            Width           =   570
         End
         Begin VB.Label Vencimiento 
            AutoSize        =   -1  'True
            Caption         =   "Vencimiento"
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Nº Giro"
            Height          =   195
            Left            =   5520
            TabIndex        =   12
            Top             =   360
            Width           =   510
         End
      End
      Begin VB.ComboBox cboFormaPago 
         Height          =   315
         Left            =   6240
         TabIndex        =   4
         Text            =   "cboFormaPago"
         Top             =   340
         Width           =   2055
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos contables"
         Height          =   855
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Width           =   8175
         Begin VB.CheckBox chkContabilizado 
            Caption         =   "Contabilizado"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker dtpFechaContable 
            Height          =   315
            Left            =   3840
            TabIndex        =   17
            Top             =   340
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   77725697
            CurrentDate     =   36938
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Contabilización"
            Height          =   195
            Left            =   1920
            TabIndex        =   16
            Top             =   360
            Width           =   1770
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos de remesa"
         Height          =   855
         Left            =   240
         TabIndex        =   18
         Top             =   3000
         Width           =   8175
         Begin VB.CheckBox chkRemesado 
            Caption         =   "Remesado"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox cboBanco 
            Height          =   315
            Left            =   5280
            TabIndex        =   22
            Text            =   "cboBanco"
            Top             =   340
            Width           =   2175
         End
         Begin MSComCtl2.DTPicker dtpFechaDomiciliacion 
            Height          =   315
            Left            =   3720
            TabIndex        =   21
            Top             =   340
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   77725697
            CurrentDate     =   36938
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Domiciliación"
            Height          =   195
            Left            =   1920
            TabIndex        =   20
            Top             =   360
            Width           =   1590
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Medio de Pago"
         Height          =   195
         Left            =   5040
         TabIndex        =   5
         Top             =   735
         Width           =   1050
      End
      Begin VB.Label lblPersona 
         AutoSize        =   -1  'True
         Caption         =   "Cliente/Proveedor"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pago"
         Height          =   195
         Left            =   5040
         TabIndex        =   3
         Top             =   360
         Width           =   1080
      End
   End
End
Attribute VB_Name = "CobroPagoEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private mintPersonaSelStart As Integer
Private mintFormaPagoSelStart As Integer
Private mintMedioPagoSelStart As Integer
Private mintBancoSelStart As Integer

Private WithEvents mobjCobroPago As CobroPago
Attribute mobjCobroPago.VB_VarHelpID = -1

'Public Tipo As String

Public Sub Component(CobroPagoObject As CobroPago)

    Set mobjCobroPago = CobroPagoObject
    'Tipo = mobjCobroPago.Tipo

End Sub

Private Sub cmdApply_Click()
    
    On Error GoTo ErrorManager

    mobjCobroPago.ApplyEdit
    mobjCobroPago.BeginEdit GescomMain.objParametro.Moneda
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    mobjCobroPago.CancelEdit
    Unload Me

End Sub

 Private Sub cmdOK_Click()
    
    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass

    mobjCobroPago.ApplyEdit
    Unload Me

    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
    With mobjCobroPago
        EnableOK .IsValid
    
        If .IsNew Then
            '.Tipo = Tipo
            Caption = .TipoText & " [(nuevo)]"

        Else
            Caption = .TipoText & " [" & .Vencimiento & " - " & Format(.Importe, "##,##0.00") & " - " & .Persona & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        LoadCombo cboPersona, .Personas
        cboPersona.Text = .Persona
    
        LoadCombo cboFormaPago, .FormasPago
        cboFormaPago.Text = .FormaPago
    
        LoadCombo cboMedioPago, .MediosPago
        cboMedioPago.Text = .MedioPago
    
        LoadCombo cboBanco, .Bancos
        cboBanco.Text = .Banco
    
        dtpVencimiento.Value = .Vencimiento
        txtImporte = .Importe
        txtNumeroGiro = .NumeroGiro
        
        chkContabilizado = IIf(.Contabilizado, vbChecked, vbUnchecked)
        dtpFechaContable.Value = .FechaContable
        dtpFechaContable.Enabled = .Contabilizado

        chkRemesado = IIf(.Remesado, vbChecked, vbUnchecked)
        dtpFechaDomiciliacion.Value = .FechaDomiciliacion
        dtpFechaDomiciliacion.Enabled = .Remesado
        If .Tipo = "C" Then
            cboBanco.Enabled = .Remesado
        End If
    
        chkHayFactura = IIf(.HayFactura, vbChecked, vbUnchecked)

        .BeginEdit GescomMain.objParametro.Moneda
    
    End With
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub mobjCobroPago_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub cboPersona_Click()

    On Error GoTo ErrorManager

    If mflgLoading Then Exit Sub
    mobjCobroPago.Persona = cboPersona.Text
  
    cboFormaPago.Text = mobjCobroPago.FormaPago
    cboFormaPago_Click

    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cboFormaPago_Click()
    
    On Error GoTo ErrorManager

    If mflgLoading Then Exit Sub
    mobjCobroPago.FormaPago = cboFormaPago.Text
  
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cboMedioPago_Click()
    
    On Error GoTo ErrorManager

    If mflgLoading Then Exit Sub
    mobjCobroPago.MedioPago = cboMedioPago.Text
  
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cboBanco_Click()
    
    On Error GoTo ErrorManager

    If mflgLoading Then Exit Sub
    mobjCobroPago.Banco = cboBanco.Text
  
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub
  
Private Sub dtpVencimiento_Change()
    
    mobjCobroPago.Vencimiento = dtpVencimiento.Value
    
End Sub
  
Private Sub txtImporte_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtImporte

End Sub

Private Sub txtImporte_Change()

    If Not mflgLoading Then
        TextChange txtImporte, mobjCobroPago, "Importe"
        txtImporte = mobjCobroPago.Importe
    End If

End Sub

Private Sub txtImporte_LostFocus()

    txtImporte = TextLostFocus(txtImporte, mobjCobroPago, "Importe")

End Sub

Private Sub txtNumeroGiro_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtNumeroGiro

End Sub

Private Sub txtNumeroGiro_Change()

    If Not mflgLoading Then
        TextChange txtNumeroGiro, mobjCobroPago, "NumeroGiro"
        txtImporte = mobjCobroPago.Importe
    End If

End Sub

Private Sub txtNumeroGiro_LostFocus()

    txtNumeroGiro = TextLostFocus(txtNumeroGiro, mobjCobroPago, "NumeroGiro")

End Sub

Private Sub dtpFechaDomiciliacion_Change()
    
    mobjCobroPago.FechaDomiciliacion = dtpFechaDomiciliacion.Value
    
End Sub

Private Sub dtpFechaContable_Change()

    mobjCobroPago.FechaContable = dtpFechaContable.Value
    
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

Private Sub cboPersona_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintPersonaSelStart = cboPersona.SelStart
End Sub

Private Sub cboPersona_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintPersonaSelStart, cboPersona
    
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


