VERSION 5.00
Begin VB.Form RepresentanteEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Representantes"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RepresentanteEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   7320
      TabIndex        =   24
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6120
      TabIndex        =   23
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   22
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Representante"
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin VB.TextBox txtDireccion 
         Height          =   1005
         Left            =   1080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   1800
         Width           =   3615
      End
      Begin VB.CommandButton btnDireccion 
         Caption         =   "Di&rección"
         Height          =   495
         Left            =   160
         TabIndex        =   20
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtCuentaContable 
         Height          =   285
         Left            =   6360
         TabIndex        =   19
         Top             =   1420
         Width           =   1575
      End
      Begin VB.TextBox txtIVA 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6360
         TabIndex        =   14
         Top             =   1060
         Width           =   975
      End
      Begin VB.TextBox txtIRPF 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6360
         TabIndex        =   9
         Top             =   700
         Width           =   975
      End
      Begin VB.TextBox txtComision 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6360
         TabIndex        =   4
         Top             =   340
         Width           =   975
      End
      Begin VB.TextBox txtZona 
         Height          =   285
         Left            =   1080
         TabIndex        =   17
         Top             =   1420
         Width           =   1575
      End
      Begin VB.TextBox txtContacto 
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   1060
         Width           =   3615
      End
      Begin VB.TextBox txtDNINIF 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   700
         Width           =   2175
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   340
         Width           =   3615
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   7440
         TabIndex        =   10
         Top             =   720
         Width           =   165
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   7440
         TabIndex        =   5
         Top             =   360
         Width           =   165
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   7440
         TabIndex        =   15
         Top             =   1080
         Width           =   165
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Contable"
         Height          =   195
         Left            =   4920
         TabIndex        =   18
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "I.V.A."
         Height          =   195
         Left            =   4920
         TabIndex        =   13
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "I.R.P.F."
         Height          =   195
         Left            =   4920
         TabIndex        =   8
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Comisión"
         Height          =   195
         Left            =   4920
         TabIndex        =   3
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Zona"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Contacto"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "DNI/NIF"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label1 
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
Attribute VB_Name = "RepresentanteEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjRepresentante As Representante
Attribute mobjRepresentante.VB_VarHelpID = -1

Public Sub Component(RepresentanteObject As Representante)

    Set mobjRepresentante = RepresentanteObject

End Sub

Private Sub btnDireccion_Click()
    
    Dim frmDireccion As DireccionEdit
  
    Set frmDireccion = New DireccionEdit
    frmDireccion.Component mobjRepresentante.Direccion
    frmDireccion.Show vbModal
    txtDireccion.Text = mobjRepresentante.Direccion.DireccionText
  
End Sub

Private Sub cmdApply_Click()
    
    On Error GoTo ErrorManager

    mobjRepresentante.ApplyEdit
    mobjRepresentante.BeginEdit
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    Dim Respuesta As VbMsgBoxResult
    
    If mobjRepresentante.IsDirty And Not mobjRepresentante.IsNew Then
        Respuesta = MostrarMensaje(MSG_MODIFY)
        If Respuesta = vbYes Then
            mobjRepresentante.CancelEdit
            Unload Me
        End If
    Else
        mobjRepresentante.CancelEdit
        Unload Me
    End If

End Sub

 Private Sub cmdOK_Click()
    
    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass

    mobjRepresentante.ApplyEdit
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
    With mobjRepresentante
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Representante [(nuevo)]"

        Else
            Caption = "Representante [" & .Nombre & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        'txtRepresentanteID = .RepresentanteID
        txtNombre = .Nombre
        txtContacto = .Contacto
        txtDNINIF = .DNINIF
        txtDireccion = .Direccion.DireccionText
        txtZona = .Zona
        txtComision = .Comision
        txtIRPF = .IRPF
        txtIVA = .IVA
        txtCuentaContable = .CuentaContable
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

Private Sub mobjRepresentante_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub txtContacto_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtContacto
        
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
        TextChange txtNombre, mobjRepresentante, "Nombre"

End Sub

Private Sub txtNombre_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtNombre
        
End Sub

Private Sub txtNombre_LostFocus()

    txtNombre = TextLostFocus(txtNombre, mobjRepresentante, "Nombre")

End Sub

Private Sub txtCuentaContable_Change()

    If Not mflgLoading Then _
        TextChange txtCuentaContable, mobjRepresentante, "CuentaContable"

End Sub

Private Sub txtCuentaContable_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCuentaContable
        
End Sub

Private Sub txtCuentaContable_LostFocus()

    txtCuentaContable = TextLostFocus(txtCuentaContable, mobjRepresentante, "CuentaContable")

End Sub

Private Sub txtContacto_Change()

    If Not mflgLoading Then _
        TextChange txtContacto, mobjRepresentante, "Contacto"

End Sub

Private Sub txtContacto_LostFocus()

    txtContacto = TextLostFocus(txtContacto, mobjRepresentante, "Contacto")

End Sub

Private Sub txtDNINIF_Change()

    If Not mflgLoading Then _
        TextChange txtDNINIF, mobjRepresentante, "DNINIF"

End Sub

Private Sub txtDNINIF_LostFocus()

    txtDNINIF = TextLostFocus(txtDNINIF, mobjRepresentante, "DNINIF")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   
   IsList = False
   
End Function

Private Sub txtZona_Change()

    If Not mflgLoading Then _
        TextChange txtZona, mobjRepresentante, "Zona"
        
End Sub

Private Sub txtZona_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtZona
        
End Sub

Private Sub txtZona_LostFocus()

    txtZona = TextLostFocus(txtZona, mobjRepresentante, "Zona")
    
End Sub

Private Sub txtComision_Change()

    If Not mflgLoading Then _
        TextChange txtComision, mobjRepresentante, "Comision"

End Sub

Private Sub txtComision_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtComision
        
End Sub

Private Sub txtComision_LostFocus()

    txtComision = TextLostFocus(txtComision, mobjRepresentante, "Comision")

End Sub

Private Sub txtIRPF_Change()

    If Not mflgLoading Then _
        TextChange txtIRPF, mobjRepresentante, "IRPF"

End Sub

Private Sub txtIRPF_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtIRPF
        
End Sub

Private Sub txtIRPF_LostFocus()

    txtIRPF = TextLostFocus(txtIRPF, mobjRepresentante, "IRPF")

End Sub

Private Sub txtIVA_Change()

    If Not mflgLoading Then _
        TextChange txtIVA, mobjRepresentante, "IVA"

End Sub

Private Sub txtIVA_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtIVA
        
End Sub

Private Sub txtIVA_LostFocus()

    txtIVA = TextLostFocus(txtIVA, mobjRepresentante, "IVA")

End Sub
