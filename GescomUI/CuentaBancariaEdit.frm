VERSION 5.00
Begin VB.Form CuentaBancariaEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuentas Bancarias"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CuentaBancariaEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   6240
      TabIndex        =   15
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5040
      TabIndex        =   14
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la Cuenta Bancaria"
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.TextBox txtEntidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   340
         Width           =   855
      End
      Begin VB.TextBox txtSucursal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   700
         Width           =   855
      End
      Begin VB.TextBox txtControl 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   1060
         Width           =   375
      End
      Begin VB.TextBox txtNombreEntidad 
         Height          =   285
         Left            =   4200
         TabIndex        =   4
         Top             =   340
         Width           =   2655
      End
      Begin VB.TextBox txtCuenta 
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Top             =   1420
         Width           =   1575
      End
      Begin VB.TextBox txtNombreSucursal 
         Height          =   285
         Left            =   4200
         TabIndex        =   8
         Top             =   700
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Entidad"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Control"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Entidad"
         Height          =   195
         Left            =   2760
         TabIndex        =   3
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nº cuenta"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "NombreSucursal"
         Height          =   195
         Left            =   2760
         TabIndex        =   7
         Top             =   720
         Width           =   1155
      End
   End
End
Attribute VB_Name = "CuentaBancariaEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjCuentaBancaria As CuentaBancaria
Attribute mobjCuentaBancaria.VB_VarHelpID = -1

Public Sub Component(CuentaBancariaObject As CuentaBancaria)

    Set mobjCuentaBancaria = CuentaBancariaObject

End Sub

Private Sub cmdApply_Click()

    mobjCuentaBancaria.ChildApplyEdit
    mobjCuentaBancaria.ChildBeginEdit

End Sub

Private Sub cmdCancel_Click()

    mobjCuentaBancaria.ChildCancelEdit
    Unload Me

End Sub

Private Sub cmdOK_Click()

    mobjCuentaBancaria.ChildApplyEdit
    Unload Me

End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
    With mobjCuentaBancaria
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Cuenta Bancaria [(nueva)]"

        Else
            Caption = "Cuenta Bancaria [" & .NombreEntidad & "]"
      
        End If
    
        txtEntidad = .Entidad
        txtSucursal = .Sucursal
        txtControl = .Control
        txtCuenta = .Cuenta
        txtNombreEntidad = .NombreEntidad
        txtNombreSucursal = .NombreSucursal
        .ChildBeginEdit
    End With
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub mobjCuentaBancaria_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub txtControl_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtControl
        
End Sub

Private Sub txtCuenta_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCuenta
        
End Sub


Private Sub txtEntidad_Change()

    If Not mflgLoading Then _
        TextChange txtEntidad, mobjCuentaBancaria, "Entidad"

End Sub

Private Sub txtEntidad_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtEntidad
        
End Sub

Private Sub txtEntidad_LostFocus()

    TextLostFocus txtEntidad, mobjCuentaBancaria, "Entidad"

End Sub

Private Sub txtNombreEntidad_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtNombreEntidad
        
End Sub

Private Sub txtNombreSucursal_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtNombreSucursal
        
End Sub

Private Sub txtSucursal_Change()

    If Not mflgLoading Then _
        TextChange txtSucursal, mobjCuentaBancaria, "Sucursal"

End Sub

Private Sub txtSucursal_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtSucursal
        
End Sub

Private Sub txtSucursal_LostFocus()

    TextLostFocus txtSucursal, mobjCuentaBancaria, "Sucursal"

End Sub

Private Sub txtControl_Change()

    If Not mflgLoading Then _
        TextChange txtControl, mobjCuentaBancaria, "Control"

End Sub

Private Sub txtControl_LostFocus()

    TextLostFocus txtControl, mobjCuentaBancaria, "Control"

End Sub

Private Sub txtCuenta_Change()

    If Not mflgLoading Then _
        TextChange txtCuenta, mobjCuentaBancaria, "Cuenta"

End Sub

Private Sub txtCuenta_LostFocus()

    TextLostFocus txtCuenta, mobjCuentaBancaria, "Cuenta"

End Sub

Private Sub txtNombreEntidad_Change()

    If Not mflgLoading Then _
        TextChange txtNombreEntidad, mobjCuentaBancaria, "NombreEntidad"

End Sub

Private Sub txtNombreEntidad_LostFocus()

    TextLostFocus txtNombreEntidad, mobjCuentaBancaria, "NombreEntidad"

End Sub

Private Sub txtNombreSucursal_Change()

    If Not mflgLoading Then _
        TextChange txtNombreSucursal, mobjCuentaBancaria, "NombreSucursal"

End Sub

Private Sub txtNombreSucursal_LostFocus()

    TextLostFocus txtNombreSucursal, mobjCuentaBancaria, "NombreSucursal"

End Sub

