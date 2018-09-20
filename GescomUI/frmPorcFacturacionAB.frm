VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPorcFacturacionAB 
   Caption         =   "Seleccione el porcentaje de facturación AB para facturar los albaranes"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   7530
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   330
      Left            =   6240
      TabIndex        =   6
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   6240
      TabIndex        =   7
      Top             =   510
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Porcentaje de facturación"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.TextBox txtCliente 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox txtPorcFacturacionAB 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Porcentaje AB"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmPorcFacturacionAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintPorcFacturacionAB As Integer
Private mstrCliente As String
Private mblnOk As Boolean


Private Sub Form_Initialize()
    
    mintPorcFacturacionAB = 100
    mstrCliente = vbNullString

End Sub

Private Sub Form_Load()
    ' Por defecto todo al 100 en A
    mblnOk = False

End Sub

Public Property Get PorcFacturacionAB() As Integer

    PorcFacturacionAB = mintPorcFacturacionAB

End Property

Public Property Let PorcFacturacionAB(ByVal intPorcFacturacionAB As Integer)

    If intPorcFacturacionAB < 0 Or intPorcFacturacionAB > 100 Then Err.Raise vbObjectError + 1001, "frmPorcFacturacionAB.PorcFacturacionAB", "El valor del porcentaje de facturacion no se encuentra entre 0 y 100"
    
    mintPorcFacturacionAB = intPorcFacturacionAB
    Slider1.Value = intPorcFacturacionAB
    txtPorcFacturacionAB.Text = intPorcFacturacionAB

End Property

Private Sub cmdCancel_Click()

    Hide
    
End Sub

Private Sub cmdOK_Click()

    mblnOk = True
    Hide
    
End Sub


Private Sub Slider1_Click()
    mintPorcFacturacionAB = Slider1.Value
    txtPorcFacturacionAB.Text = mintPorcFacturacionAB
End Sub

Private Sub txtPorcFacturacionAB_Change()
    mintPorcFacturacionAB = Val(txtPorcFacturacionAB.Text)
    Slider1.Value = mintPorcFacturacionAB
End Sub

Public Property Get Ok() As Boolean

    Ok = mblnOk

End Property


Public Property Let Cliente(ByVal strCliente As String)

    mstrCliente = strCliente
    txtCliente.Text = mstrCliente

End Property
