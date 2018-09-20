VERSION 5.00
Begin VB.Form DatoComercialEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos C&omerciales"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DatoComercialEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   3000
      TabIndex        =   12
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Dato Comercial"
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.TextBox txtDescuento 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   340
         Width           =   1095
      End
      Begin VB.TextBox txtRecargoEquivalencia 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Top             =   705
         Width           =   1095
      End
      Begin VB.TextBox txtIVA 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   1065
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   3480
         TabIndex        =   9
         Top             =   1080
         Width           =   165
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   3480
         TabIndex        =   6
         Top             =   720
         Width           =   165
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   3480
         TabIndex        =   3
         Top             =   360
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descuento"
         Height          =   195
         Left            =   255
         TabIndex        =   1
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Recargo de equivalencia"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1755
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "I.V.A."
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   435
      End
   End
End
Attribute VB_Name = "DatoComercialEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjDatoComercial As DatoComercial
Attribute mobjDatoComercial.VB_VarHelpID = -1

Public Sub Component(DatoComercialObject As DatoComercial)

    Set mobjDatoComercial = DatoComercialObject

End Sub

Private Sub cmdApply_Click()

    mobjDatoComercial.ChildApplyEdit
    mobjDatoComercial.ChildBeginEdit

End Sub

Private Sub cmdCancel_Click()

    mobjDatoComercial.ChildCancelEdit
    Unload Me

End Sub

Private Sub cmdOK_Click()

    mobjDatoComercial.ChildApplyEdit
    Unload Me

End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
    With mobjDatoComercial
        EnableOK .IsValid
        
        If .IsNew Then
            Caption = "DatoComercial [(nuevo)]"
    
        Else
            Caption = "DatoComercial [Dto.:" & .Descuento & "%, I.V.A.:" & .IVA & "%]"
      
        End If
        
        txtDescuento = .Descuento
        txtRecargoEquivalencia = .RecargoEquivalencia
        txtIVA = .IVA
        .ChildBeginEdit
    End With
    
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub mobjDatoComercial_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub txtDescuento_Change()

    If Not mflgLoading Then _
        TextChange txtDescuento, mobjDatoComercial, "Descuento"

End Sub

Private Sub txtDescuento_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtDescuento
        
End Sub

Private Sub txtDescuento_LostFocus()

    TextLostFocus txtDescuento, mobjDatoComercial, "Descuento"

End Sub

Private Sub txtIVA_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtIVA
        
End Sub

Private Sub txtRecargoEquivalencia_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtRecargoEquivalencia
        
End Sub

Private Sub txtRecargoEquivalencia_Change()

    If Not mflgLoading Then _
        TextChange txtRecargoEquivalencia, mobjDatoComercial, "RecargoEquivalencia"

End Sub

Private Sub txtRecargoEquivalencia_LostFocus()

    TextLostFocus txtRecargoEquivalencia, mobjDatoComercial, "RecargoEquivalencia"

End Sub

Private Sub txtIVA_Change()

    If Not mflgLoading Then _
        TextChange txtIVA, mobjDatoComercial, "IVA"

End Sub

Private Sub txtIVA_LostFocus()

    TextLostFocus txtIVA, mobjDatoComercial, "IVA"

End Sub
