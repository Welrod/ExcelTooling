Attribute VB_Name = "Test"
'runs through all tests, displaying err values or 0 if correct
Public Sub TestSuit()
    MsgBox "Test Results:" & vbNewLine & _
        "TestBooleanObj(): " & TestBooleanObj() & vbNewLine & _
        "TestDateObj(): " & TestDateObj() & vbNewLine & _
        "TestDoubleObj(): " & TestDoubleObj() & vbNewLine & _
        "TestIntegerObj(): " & TestIntegerObj() & vbNewLine & _
        "TestStringObj(): " & TestStringObj() & vbNewLine & _
        ""
End Sub

Public Function TestBooleanObj() As Integer
    On Error Resume Next
        Dim myBoolean As Boolean
        Dim myObj As BooleanObj
        myBoolean = True
        Set myObj = New BooleanObj
        myObj.Value = myBoolean
        
        TestBooleanObj = Err.Number
    On Error GoTo 0
End Function

Public Function TestDateObj() As Integer
    On Error Resume Next
        Dim myDate As Date
        Dim myObj As DateObj
        myDate = Now 'forget if this is right
        Set myObj = New DateObj
        myObj.Value = myDate
        
        TestDateObj = Err.Number
    On Error GoTo 0
End Function

Public Function TestDoubleObj() As Integer
    On Error Resume Next
        Dim myDouble As Double
        Dim myObj As DoubleObj
        myDouble = 2.22
        Set myObj = New DoubleObj
        myObj.Value = myDouble
        
        TestDoubleObj = Err.Number
    On Error GoTo 0
End Function

Public Function TestIntegerObj() As Integer
    On Error Resume Next
        Dim myInt As Integer
        Dim myObj As IntegerObj
        myInt = 2
        Set myObj = New IntegerObj
        myObj.Value = myInt
        
        TestIntegerObj = Err.Number
    On Error GoTo 0
End Function

Public Function TestStringObj() As Integer
    On Error Resume Next
        Dim myString As String
        Dim myObj As StringObj
        myString = "Hello World!"
        Set myObj = New StringObj
        myObj.Value = myString
        
        TestStringObj = Err.Number
    On Error GoTo 0
End Function