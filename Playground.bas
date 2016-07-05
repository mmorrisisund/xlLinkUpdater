Attribute VB_Name = "Playground"
Option Explicit

Sub TestStringConstants()
    Dim uninitializedString As String
    Dim emptyString As String
    Dim testString As String: testString = "test"
    
    'Debug.Print "vbEmpty:uninitializedString", uninitializedString = vbEmpty
    'Debug.Print "vbNull:uninitializedString", uninitializedString = vbNull
    Debug.Print "vbNullString:uninitializedString", uninitializedString = vbNullString
    
    Debug.Print "vbEmpty:emptyString", emptyString = vbEmpty
    Debug.Print "vbNull:emptyString", emptyString = vbNull
    Debug.Print "vbNullString:emptyString", emptyString = vbNullString
    
    Debug.Print "vbEmpty:testString", testString = vbEmpty
    Debug.Print "vbNull:testString", testString = vbNull
    Debug.Print "vbNullString:testString", testString = vbNullString
End Sub

Sub TestUnprotect()
    Debug.Print HasPasswordProtection(Sheet1)
End Sub
Function HasPasswordProtection(ws As Worksheet) As Boolean
    On Error Resume Next
        ' Attempt to Unprotect sheet with no password.
        Call ws.Unprotect("")
    On Error GoTo 0
    HasPasswordProtection = ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios
End Function

