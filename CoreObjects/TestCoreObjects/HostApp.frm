VERSION 5.00
Begin VB.Form frmHostApp 
   Caption         =   "Tested Application (Standard EXE)"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   ScaleHeight     =   1170
   ScaleWidth      =   4035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Run 
      Caption         =   "Run Tests!"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmHostApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pTestRunner As TestRunner
Dim pSimpleTestCase As TestCase
Dim pTestFrameworkCase As TestCase

Private Sub Form_Initialize()
    'Tests initialization
    Set pTestRunner = New TestRunner
    
    'Add Simple Test
    Set pSimpleTestCase = New TestCase
    Set pSimpleTestCase.TestFixture = New SimpleTest
    pTestRunner.AddSuite pSimpleTestCase

    'Add Framework Test
    Set pTestFrameworkCase = New TestCase
    Set pTestFrameworkCase.TestFixture = New TestFramework
    pTestRunner.AddSuite pTestFrameworkCase
End Sub

Private Sub Run_Click()
    'Select a test
    Dim pTest As ITest
    
    Set pTest = pSimpleTestCase
    pTestRunner.SelectTest pTest.Name
    
    'Then using TestComplete as a runner, the parameters of this procedure are not used.
    pTestRunner.Run False, False
End Sub

