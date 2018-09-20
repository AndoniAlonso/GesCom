VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestTestError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITestFixture

Private m_pTestCase As TestCase

Private Property Get ITestFixture_Name() As String
    ITestFixture_Name = "TestTestError"
End Property

Private Sub ITestFixture_RunTest()
    Dim colTest As Collection
    Dim objTest As Object

    Set colTest = New Collection
    Set objTest = colTest.Item(0)
    
End Sub

Private Sub ITestFixture_Setup(TestCaseContainer As VBUnit.TestCase)
    Set m_pTestCase = TestCaseContainer
End Sub

'Test case reference must be released here.
Private Sub ITestFixture_TearDown()
    Set m_pTestCase = Nothing
End Sub
