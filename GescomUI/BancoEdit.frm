VERSION 5.00
Begin VB.Form BancoEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bancos"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BancoEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Banco"
      Height          =   3405
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.TextBox txtSufijoNIF 
         Height          =   285
         Left            =   4080
         TabIndex        =   10
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtCuentaContable 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   2860
         Width           =   1575
      End
      Begin VB.TextBox txtContacto 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   340
         Width           =   3855
      End
      Begin VB.CommandButton btnCuentaBancaria 
         Caption         =   "Cuenta &Bancaria"
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtCuentaBancaria 
         Height          =   1005
         Left            =   1560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1780
         Width           =   3855
      End
      Begin VB.TextBox txtDireccion 
         Height          =   1005
         Left            =   1560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   700
         Width           =   3855
      End
      Begin VB.CommandButton btnDireccion 
         Caption         =   "Di&rección"
         Height          =   480
         Left            =   240
         TabIndex        =   3
         Top             =   735
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sufijo NIF:"
         Height          =   195
         Left            =   3240
         TabIndex        =   9
         Top             =   2895
         Width           =   765
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Contable"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Contacto"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   660
      End
   End
End
Attribute VB_Name = "BancoEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjBanco As Banco
Attribute mobjBanco.VB_VarHelpID = -1

Public Sub Component(BancoObject As Banco)

    Set mobjBanco = BancoObject

End Sub

Private Sub btnDireccion_Click()
    
    Dim frmDireccion As DireccionEdit
  
    Set frmDireccion = New DireccionEdit
    frmDireccion.Component mobjBanco.Direccion
    frmDireccion.Show vbModal
    txtDireccion.Text = mobjBanco.Direccion.DireccionText
  
End Sub

Private Sub btnCuentaBancaria_Click()
    
    Dim frmCuentaBancaria As CuentaBancariaEdit
  
    Set frmCuentaBancaria = New CuentaBancariaEdit
    frmCuentaBancaria.Component mobjBanco.CuentaBancaria
    frmCuentaBancaria.Show vbModal
    txtCuentaBancaria.Text = mobjBanco.CuentaBancaria.CuentaBancariaText
  
End Sub

Private Sub cmdApply_Click()
    
    On Error GoTo ErrorManager
    
    If mobjBanco.EmpresaID = 0 Then _
        mobjBanco.EmpresaID = GescomMain.objParametro.EmpresaActualID
    mobjBanco.ApplyEdit
    mobjBanco.BeginEdit
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    Dim Respuesta As VbMsgBoxResult
    
    If mobjBanco.IsDirty And Not mobjBanco.IsNew Then
        Respuesta = MostrarMensaje(MSG_MODIFY)
        If Respuesta = vbYes Then
            mobjBanco.CancelEdit
            Unload Me
        End If
    Else
        mobjBanco.CancelEdit
        Unload Me
    End If

End Sub

Private Sub cmdOK_Click()
    
    On Error GoTo ErrorManager
    
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    If mobjBanco.EmpresaID = 0 Then _
        mobjBanco.EmpresaID = GescomMain.objParametro.EmpresaActualID
    mobjBanco.ApplyEdit
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
    With mobjBanco
        EnableOK .IsValid
        
        If .IsNew Then
            Caption = "Banco [(nuevo)]"

        Else
            Caption = "Banco [" & .CuentaBancaria.NombreEntidad & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtContacto = .Contacto
        txtDireccion = .Direccion.DireccionText
        txtCuentaBancaria = .CuentaBancaria.CuentaBancariaText
        txtCuentaContable = .CuentaContable
        txtSufijoNIF = .SufijoNIF
        
        .BeginEdit
        .EmpresaID = GescomMain.objParametro.EmpresaActualID
        .CancelEdit
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

Private Sub mobjBanco_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub txtContacto_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtContacto
        
End Sub

Private Sub txtContacto_Change()

    If Not mflgLoading Then _
        TextChange txtContacto, mobjBanco, "Contacto"

End Sub

Private Sub txtContacto_LostFocus()

    txtContacto = TextLostFocus(txtContacto, mobjBanco, "Contacto")

End Sub

Private Sub txtCuentaBancaria_DblClick()

    Call btnCuentaBancaria_Click
    
End Sub

Private Sub txtCuentaContable_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCuentaContable
        
End Sub

Private Sub txtCuentaContable_Change()

    If Not mflgLoading Then _
        TextChange txtCuentaContable, mobjBanco, "CuentaContable"

End Sub

Private Sub txtCuentaContable_LostFocus()

    TextLostFocus txtCuentaContable, mobjBanco, "CuentaContable"

End Sub

Private Sub txtSufijoNIF_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtSufijoNIF
        
End Sub

Private Sub txtSufijoNIF_Change()

    If Not mflgLoading Then _
        TextChange txtSufijoNIF, mobjBanco, "SufijoNIF"

End Sub

Private Sub txtSufijoNIF_LostFocus()

    TextLostFocus txtSufijoNIF, mobjBanco, "SufijoNIF"

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   
   IsList = False
   
End Function

Private Sub txtDireccion_DblClick()

    Call btnDireccion_Click
    
End Sub
