VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimpleTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITestFixture

Private m_dblValue1 As Double
Private m_dblValue2 As Double
Private m_pTestCase As TestCase

'Name of the test
Private Property Get ITestFixture_Name() As String
    ITestFixture_Name = "Simple Test"
End Property

'Run all tests
Private Sub ITestFixture_RunTest()
    TestAdd
    TestEqualsLong
    TestEqualsDouble
'    TestDivideByZero
    TestCoreConsultaDW_Load
    TestCoreConsultaDW_UpdateTable
End Sub


'Initialize variables used by the test. Save a reference to the ContainingTestCase
'to use asserts.
Private Sub ITestFixture_Setup(TestCaseContainer As VBUnit.TestCase)
    Set m_pTestCase = TestCaseContainer
    m_dblValue1 = 2#
    m_dblValue2 = 3#
End Sub

'References to TestCaseContainer must be released here (not in Class_Terminate)
Private Sub ITestFixture_TearDown()
    Set m_pTestCase = Nothing
End Sub

Public Sub TestAdd()
    Dim dblResult As Double
    
    m_pTestCase.Trace "TestAdd"
    dblResult = m_dblValue1 + m_dblValue2
    m_pTestCase.Assert (dblResult = 6#), "TestAdd: Result is not 6"
End Sub

Public Sub TestEqualsLong()
    m_pTestCase.Trace "TestEqualsLong"
    m_pTestCase.AssertEqualsLong 12, 12
    m_pTestCase.AssertEqualsLong 12, 13, "size"
End Sub

Public Sub TestEqualsDouble()
    m_pTestCase.Trace "TestEqualsDouble"
    m_pTestCase.AssertEqualsDouble 12#, 11.99, 0#, "capacity"
End Sub

Public Sub TestCoreConsultaDW_Load()
    Dim objConsulta As CoreConsultaDW

    m_pTestCase.Trace "TestCoreConsultaDW_Load"
    
    Set objConsulta = New CoreConsultaDW
    
    objConsulta.Load "CVTempProv"
    
    Set objConsulta = Nothing
End Sub

Public Sub TestCoreConsultaDW_UpdateTable()
    Dim objConsulta As CoreConsultaDW

    Set objConsulta = New CoreConsultaDW
    
    m_pTestCase.Trace "TestCoreConsultaDW_UpdateTable"
    
    objConsulta.Load "CVTempProv"
    
    objConsulta.UpdateTable
    
    Set objConsulta = Nothing
End Sub

