VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ApunteEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Apunte"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ApunteEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Apunte"
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.ComboBox cboTipoImporte 
         Height          =   315
         ItemData        =   "ApunteEdit.frx":08CA
         Left            =   3240
         List            =   "ApunteEdit.frx":08D4
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1420
         Width           =   1455
      End
      Begin VB.TextBox txtCuenta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtDocumento 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4200
         TabIndex        =   6
         Top             =   1060
         Width           =   960
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   1420
         Width           =   1695
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   705
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   1080
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
         Format          =   78184449
         CurrentDate     =   36938
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1095
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Documento"
         Height          =   195
         Left            =   3120
         TabIndex        =   8
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   570
      End
      Begin VB.Label label12 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   5040
      TabIndex        =   14
      Top             =   2160
      Width           =   1095
   End
End
Attribute VB_Name = "ApunteEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjApunte As Apunte
Attribute mobjApunte.VB_VarHelpID = -1

Public Sub Component(ApunteObject As Apunte)

    Set mobjApunte = ApunteObject

End Sub

Private Sub cboTipoImporte_Click()
    
    If mflgLoading Then Exit Sub
    mobjApunte.TipoImporte = cboTipoImporte

End Sub

Private Sub cmdApply_Click()
    
    On Error GoTo ErrorManager

    mobjApunte.ApplyEdit
    mobjApunte.BeginEdit
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    mobjApunte.CancelEdit
    Unload Me

End Sub

Private Sub cmdOK_Click()

    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    mobjApunte.ApplyEdit
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
    With mobjApunte
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Línea del Apunte [(nueva)]"

        Else
            Caption = "Línea del Apunte [" & .Descripcion & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtCuenta = .Cuenta
        txtDocumento = .Documento
        txtImporte = .Importe
        txtImporte = .Importe
        txtDescripcion = .Descripcion
        dtpFecha.Value = .Fecha
        cboTipoImporte = .TipoImporte
        
        .BeginEdit
    
    End With
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub mobjApunte_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub dtpFecha_Change()
    
    mobjApunte.Fecha = dtpFecha.Value
    
End Sub

Private Sub txtDocumento_Change()

    If Not mflgLoading Then _
        TextChange txtDocumento, mobjApunte, "Documento"

End Sub

Private Sub txtDocumento_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtDocumento
        
End Sub

Private Sub txtDocumento_LostFocus()

    txtDocumento = TextLostFocus(txtDocumento, mobjApunte, "Documento")

End Sub

Private Sub txtCuenta_Change()

    If Not mflgLoading Then _
        TextChange txtCuenta, mobjApunte, "Cuenta"

End Sub

Private Sub txtCuenta_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCuenta
        
End Sub

Private Sub txtCuenta_LostFocus()

    txtCuenta = TextLostFocus(txtCuenta, mobjApunte, "Cuenta")

End Sub

Private Sub txtDescripcion_Change()

    If Not mflgLoading Then _
        TextChange txtDescripcion, mobjApunte, "Descripcion"

End Sub

Private Sub txtDescripcion_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtDescripcion
        
End Sub

Private Sub txtDescripcion_LostFocus()

    txtDescripcion = TextLostFocus(txtDescripcion, mobjApunte, "Descripcion")

End Sub

Private Sub txtImporte_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtImporte

End Sub

Private Sub txtImporte_Change()

    If Not mflgLoading Then _
        TextChange txtImporte, mobjApunte, "Importe"

End Sub

Private Sub txtImporte_LostFocus()

    txtImporte = TextLostFocus(txtImporte, mobjApunte, "Importe")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   
   IsList = False
   
End Function

