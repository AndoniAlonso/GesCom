VERSION 5.00
Begin VB.Form TPVGeneralEdit 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboFormaPago 
      BackColor       =   &H00E7CDCD&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00591E1E&
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Text            =   "cboFormaPago"
      Top             =   1680
      Width           =   5055
   End
   Begin VB.ComboBox cboRepresentante 
      BackColor       =   &H00E7CDCD&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00591E1E&
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Text            =   "cboRepresentante"
      Top             =   1080
      Width           =   5055
   End
   Begin VB.ComboBox cboCliente 
      BackColor       =   &H00E7CDCD&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00591E1E&
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Text            =   "cboCliente"
      Top             =   480
      Width           =   5055
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C99497&
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8040
      MaskColor       =   &H00591E1E&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Forma de Pago"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00591E1E&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1695
      Width           =   2130
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Vendedor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00591E1E&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1350
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00591E1E&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   495
      Width           =   960
   End
End
Attribute VB_Name = "TPVGeneralEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintClienteSelStart As Integer
Private mintRepresentanteSelStart As Integer
Private mintFormaPagoSelStart As Integer
Private mflgLoading As Boolean

Private WithEvents mobjAlbaranVenta As AlbaranVenta
Attribute mobjAlbaranVenta.VB_VarHelpID = -1

Private Sub cboCliente_Click()
    
    On Error GoTo ErrorManager
  
    If mflgLoading Then Exit Sub
    
    mobjAlbaranVenta.Cliente = cboCliente.Text
'    ' Tener en cuenta que si la empresa "anula" el tratamiento del IVA, se debe poner a 0
'    If GescomMain.objParametro.ObjEmpresaActual.AnularIVA Then
'        mobjAlbaranVenta.DatoComercial.ChildBeginEdit
'        mobjAlbaranVenta.DatoComercial.IVA = 0
'        mobjAlbaranVenta.DatoComercial.ChildApplyEdit
'    End If
  
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cboCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintClienteSelStart = cboCliente.SelStart
End Sub

Private Sub cboCliente_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintClienteSelStart, cboCliente
    
End Sub


Public Sub Component(AlbaranVentaObject As AlbaranVenta)

    Set mobjAlbaranVenta = AlbaranVentaObject

End Sub

Private Sub cboRepresentante_Change()
    If mflgLoading Then Exit Sub
    
    mobjAlbaranVenta.Representante = cboRepresentante.Text

End Sub

Private Sub cboRepresentante_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintRepresentanteSelStart = cboRepresentante.SelStart
End Sub

Private Sub cboRepresentante_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintRepresentanteSelStart, cboRepresentante
    
End Sub

Private Sub cmdOK_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()
    mflgLoading = True
    
    EnableOK mobjAlbaranVenta.IsValid
    
    LoadCombo cboCliente, mobjAlbaranVenta.Clientes
    cboCliente.Text = mobjAlbaranVenta.Cliente
 
    LoadCombo cboRepresentante, mobjAlbaranVenta.Representantes
    cboRepresentante.Text = mobjAlbaranVenta.Representante
        
    LoadCombo cboFormaPago, mobjAlbaranVenta.FormasPago
    cboFormaPago.Text = mobjAlbaranVenta.FormaPago
        
    mflgLoading = False
End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
'    cmdApply.Enabled = flgValid

End Sub

Private Sub cboFormaPago_Change()
    If mflgLoading Then Exit Sub
    
    mobjAlbaranVenta.FormaPago = cboFormaPago.Text

End Sub

Private Sub cboFormaPago_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintFormaPagoSelStart = cboFormaPago.SelStart
End Sub

Private Sub cboFormaPago_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintFormaPagoSelStart, cboFormaPago
    
End Sub


