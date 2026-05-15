' Test Suite with SetUp and TearDown
' This example demonstrates SetUp/TearDown support in the test runner

Option Explicit

' Global state for testing
Dim testCounter As Integer
Dim testState As String

' Setup routine - called before each test
Sub SetUp()
    testCounter = 0
    testState = "initialized"
End Sub

' Teardown routine - called after each test
Sub TearDown()
    testCounter = 0
    testState = ""
End Sub

' Test 1: Verify SetUp initializes counter
Sub Test_SetupInitializesCounter(testResult)
    If testCounter <> 0 Then testResult = False
End Sub

' Test 2: Verify SetUp sets state
Sub Test_SetupInitializesState(testResult)
    If testState <> "initialized" Then testResult = False
End Sub

' Test 3: Modify counter
Sub Test_CounterIncrement(testResult)
    testCounter = testCounter + 1
    testCounter = testCounter + 1
    If testCounter <> 2 Then testResult = False
End Sub

' Test 4: String operations
Sub Test_StringConcat(testResult)
    testState = testState & " and working"
    If testState <> "initialized and working" Then testResult = False
End Sub

' Test 5: Boolean operations
Sub Test_LogicalAnd(testResult)
    Dim result As Boolean
    result = (testCounter = 0) And (testState = "initialized")
    If Not result Then testResult = False
End Sub
