VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestFramework"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Test case to test the framework

Implements ITestFixture

Private m_pTestSuccess As ITest
Private m_pTestFailure As ITest
Private m_pTestError As ITest
Private m_pTestCase As TestCase

Private Property Get ITestFixture_Name() As String
    ITestFixture_Name = "Test Framework"
End Property

Private Sub ITestFixture_RunTest()
    TestSuccess
    TestError
    TestFailure
End Sub

Private Sub ITestFixture_Setup(TestCaseContainer As VBUnit.TestCase)
    Dim pTempCase As TestCase
    
    Set m_pTestSuccess = New TestCase
    Set pTempCase = m_pTestSuccess
    Set pTempCase.TestFixture = New TestTestSuccess
    
    Set m_pTestError = New TestCase
    Set pTempCase = m_pTestError
    Set pTempCase.TestFixture = New TestTestError
    
    Set m_pTestFailure = New TestCase
    Set pTempCase = m_pTestFailure
    Set pTempCase.TestFixture = New TestTestFailure
    
    Set pTempCase = Nothing
    
    Set m_pTestCase = TestCaseContainer
End Sub

Private Sub ITestFixture_TearDown()
    Set m_pTestSuccess = Nothing
    Set m_pTestError = Nothing
    Set m_pTestFailure = Nothing
    
    'This reference must be released here
    Set m_pTestCase = Nothing
End Sub
Public Sub TestSuccess()
    Dim pResult As TestResult
    
    m_pTestCase.Trace "Test Success"
    Set pResult = m_pTestSuccess.Run()
    m_pTestCase.Assert pResult.RunTests = 1
    m_pTestCase.Assert pResult.TestFailures = 0
    m_pTestCase.Assert pResult.TestErrors = 0
    m_pTestCase.Assert pResult.WasSuccessful

End Sub
Public Sub TestError()
    Dim pResult As TestResult
    
    m_pTestCase.Trace "Test Error"
    Set pResult = m_pTestError.Run()
    m_pTestCase.Assert pResult.RunTests = 1
    m_pTestCase.Assert pResult.TestFailures = 0
    m_pTestCase.Assert pResult.TestErrors = 1
    m_pTestCase.Assert Not pResult.WasSuccessful
End Sub

Public Sub TestFailure()
    Dim pResult As TestResult
    
    m_pTestCase.Trace "Test Failure"
    Set pResult = m_pTestFailure.Run()
    m_pTestCase.Assert pResult.RunTests = 1
    m_pTestCase.Assert pResult.TestFailures = 1
    m_pTestCase.Assert pResult.TestErrors = 0
    m_pTestCase.Assert Not pResult.WasSuccessful
End Sub
