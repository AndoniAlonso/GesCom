VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones de Impresora"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraPrinter 
      Caption         =   "Impresora"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.ComboBox cboPrinter 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblName 
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.Frame fraCopies 
      Caption         =   "Copias"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   3975
      Begin VB.CheckBox chkCollate 
         Appearance      =   0  'Flat
         Caption         =   "In&tercalar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   16
         Top             =   840
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox txtCopies 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         TabIndex        =   15
         Text            =   "1"
         Top             =   240
         Width           =   975
      End
      Begin VB.PictureBox picCopies 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   2055
         TabIndex        =   14
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      TabIndex        =   18
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Imprimir"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   17
      Top             =   3720
      Width           =   975
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Seleccion de páginas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   3975
      Begin VB.TextBox txtStart 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3000
         TabIndex        =   8
         Top             =   630
         Width           =   855
      End
      Begin VB.TextBox txtEnd 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3000
         TabIndex        =   11
         Top             =   990
         Width           =   855
      End
      Begin VB.OptionButton optPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Rango"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Todas las páginas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton optPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Página actual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label lblEnd 
         Caption         =   "Fin:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label lblStart 
         Caption         =   "Inicio:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label lblStatus 
         Caption         =   "Página:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   300
         Width           =   615
      End
      Begin VB.Label lblPages 
         Alignment       =   2  'Center
         Caption         =   "0 / 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         Top             =   300
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================
' Project:      Enhance Print object
' Author:       edward moth
' Copyright:    © 2000-2001 qbd software ltd
'
' ==============================================================
' Module:       frmPrint
' Purpose:      User options for printing
' ==============================================================
' Credits:      Thanks to Jarry of Jacsoft for his contributions
'               toward improving the functionality of this
'               module.
' ==============================================================

Option Explicit


Private mvarCurrent As Integer
Private mvarMax As Integer
Private mvarStart As Integer
Private mvarEnd As Integer
Private mvarPrint As Boolean
Private mvarCollate As Boolean
Private mvarFlags As qePrintOptionFlags
Private mvarPrinter As Integer
Private mvarCopies As Integer
Private bInternal As Boolean




Friend Property Let Flags(ByVal eFlag As qePrintOptionFlags)

mvarFlags = eFlag
End Property


Private Sub cboPrinter_Click()
If Not bInternal And cboPrinter.ListIndex > -1 Then
mvarPrinter = cboPrinter.ItemData(cboPrinter.ListIndex)
End If

End Sub

Private Sub chkCollate_Click()

mvarCollate = CBool(chkCollate.Value = vbChecked)
Copies_ShowImage

End Sub
Private Sub Copies_ShowImage()

Dim sX As Single
Dim sY As Single
Dim iPage As Integer
Dim iNum As Integer

picCopies.Cls
picCopies.FontSize = 8
If mvarCollate Then
sX = 1400: sY = 0
iNum = 2
Do While iNum > 0
iPage = 3
Do While iPage > 0
picCopies.Line (sX, sY)-Step(300, 420), vbWhite, BF
picCopies.Line (sX, sY)-Step(300, 420), vbBlack, B
picCopies.CurrentX = sX + 300 - picCopies.TextWidth(iPage) - 60
picCopies.CurrentY = sY + 420 - picCopies.TextHeight(iPage)
picCopies.Print iPage
sX = sX - 150
sY = sY + 210
iPage = iPage - 1
Loop
sX = sX - 400
sY = 0
iNum = iNum - 1
Loop
Else
sX = 1400: sY = 105
iNum = 3
Do While iNum > 0
iPage = 2
Do While iPage > 0
picCopies.Line (sX, sY)-Step(300, 420), vbWhite, BF
picCopies.Line (sX, sY)-Step(300, 420), vbBlack, B
picCopies.CurrentX = sX + 300 - picCopies.TextWidth(iNum) - 60
picCopies.CurrentY = sY + 420 - picCopies.TextHeight(iNum)
picCopies.Print iNum
sX = sX - 150
sY = sY + 210
iPage = iPage - 1
Loop
sX = sX - 200
sY = 105
iNum = iNum - 1
Loop

End If




End Sub


Private Sub cmdCancel_Click()

mvarPrint = False
Me.Hide

End Sub

Private Sub cmdPrint_Click()

Dim lStart As Integer, lEnd As Integer
Dim bEnable As Boolean

bEnable = True
lStart = Val(txtStart.Text)
lEnd = Val(txtEnd.Text)

If lStart = 0 Or lEnd = 0 Then
  bEnable = False
End If
If lStart > lEnd Then
  bEnable = False
End If
If lStart <> CInt(lStart) Then
  bEnable = False
End If
If lEnd <> CInt(lEnd) Then
  bEnable = False
End If

If optPrint(0).Value Then
  bEnable = True
  mvarStart = mvarCurrent
  mvarEnd = mvarCurrent
ElseIf optPrint(1).Value Then
  bEnable = True
  mvarStart = 1
  mvarEnd = mvarMax
ElseIf optPrint(2).Value Then
  mvarStart = lStart
  mvarEnd = lEnd
End If
mvarCopies = Val(txtCopies.Text)

If Not bEnable Then
  MsgBox "Please enter a valid page range.", vbOKOnly, "Range Error"
  mvarPrint = False
Else
  mvarPrint = True
  Me.Hide
End If

End Sub





Public Property Get PrintDoc() As Boolean
PrintDoc = mvarPrint
End Property


Public Property Let PageCurrent(ByVal vNewValue As Integer)
mvarCurrent = vNewValue
End Property

Public Property Get PageStart() As Integer
PageStart = mvarStart
End Property

Public Property Get PageEnd() As Integer
PageEnd = mvarEnd
End Property

Public Property Get Collate() As Boolean
Collate = mvarCollate
End Property

Public Property Get Copies() As Integer
    Copies = mvarCopies
End Property

Public Property Let Copies(NumCopies As Integer)

    mvarCopies = NumCopies

End Property

Public Property Get PrinterNumber() As Integer
PrinterNumber = mvarPrinter
End Property


Public Property Let PageMax(ByVal vNewValue As Integer)
mvarMax = vNewValue
End Property

Private Sub Form_Load()
Dim prtPrinter As Printer
Dim iPrinter As Integer
Dim sPrinter As String

Dim sY As Single
' Change options dependent on flags and default values
bInternal = True
' Display Printer List
If CBool(mvarFlags And ShowPrinter_po) Then
fraPrinter.Visible = True
sY = sY + fraPrinter.Height + 120
cboPrinter.Clear
iPrinter = 0
For Each prtPrinter In Printers
sPrinter = prtPrinter.DeviceName & " on "
If Right$(prtPrinter.Port, 1) = ":" Then
sPrinter = sPrinter & Left$(prtPrinter.Port, Len(prtPrinter.Port) - 1)
Else
sPrinter = sPrinter & prtPrinter.Port
End If
cboPrinter.AddItem sPrinter
cboPrinter.ItemData(cboPrinter.NewIndex) = iPrinter
If prtPrinter.DeviceName = Printer.DeviceName And Printer.Port = prtPrinter.Port Then
cboPrinter.ListIndex = cboPrinter.NewIndex
mvarPrinter = iPrinter
End If
iPrinter = iPrinter + 1
Next
Else
fraPrinter.Visible = False
End If

' Display Page range print options
optPrint(0).Value = True
optPrint(1).Enabled = CBool(mvarMax > 1)
optPrint(2).Enabled = CBool(mvarMax > 1)
lblPages.Caption = mvarCurrent & " / " & mvarMax
fraOptions.Top = sY
sY = sY + fraOptions.Height + 120
txtStart.Enabled = CBool(mvarMax > 1)
txtEnd.Enabled = CBool(mvarMax > 1)
txtStart.Text = "1"
txtEnd.Text = mvarMax
txtCopies.Text = mvarCopies


' Display copy and collation options
If CBool(mvarFlags And ShowCopies_po) Then
fraCopies.Visible = True
mvarCollate = True
Copies_ShowImage
fraCopies.Top = sY
sY = sY + fraCopies.Height + 120
Else
fraCopies.Visible = False
End If

' Adjust form size
cmdPrint.Top = sY
cmdCancel.Top = sY
sY = sY + cmdPrint.Height + 120
sY = (Me.Height - Me.ScaleHeight) + sY
Me.Height = sY
bInternal = False

End Sub

Private Sub optPrint_Click(Index As Integer)
txtStart.Locked = CBool(Index <> 2)
txtEnd.Locked = CBool(Index <> 2)
If Index = 0 Then
  txtStart.Text = mvarCurrent
  txtEnd.Text = mvarCurrent
Else
  txtStart.Text = 1
  txtEnd.Text = mvarMax
End If

End Sub



Private Sub txtEnd_GotFocus()
optPrint(2).Value = True

End Sub

Private Sub txtStart_GotFocus()
optPrint(2).Value = True

End Sub
