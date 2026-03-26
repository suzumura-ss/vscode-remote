Option Explicit

Dim totalTests, passedTests, failedTests
totalTests = 0
passedTests = 0
failedTests = 0

Sub AssertEqual(actual, expected, msg)
    If actual <> expected Then
        Err.Raise vbObjectError + 1, , _
            msg & " | Expected: [" & expected & "] Got: [" & actual & "]"
    End If
End Sub

Sub AssertTrue(condition, msg)
    If Not condition Then
        Err.Raise vbObjectError + 1, , msg
    End If
End Sub

Sub RunTest(testName)
    totalTests = totalTests + 1
    Dim testProc
    On Error Resume Next
    Set testProc = GetRef(testName)
    If Err.Number <> 0 Then
        failedTests = failedTests + 1
        WScript.Echo "FAIL: " & testName & " - not found"
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear
    Call testProc
    If Err.Number <> 0 Then
        failedTests = failedTests + 1
        WScript.Echo "FAIL: " & testName & " - " & Err.Description
    Else
        passedTests = passedTests + 1
        WScript.Echo "PASS: " & testName
    End If
    On Error GoTo 0
End Sub

' --- Load production code ---
Dim fso, f, code, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
Set f = fso.OpenTextFile(scriptDir & "\vscode-remote.vbs", 1)
code = f.ReadAll
f.Close
ExecuteGlobal code

' ============================================================
' UriDecode tests
' ============================================================

Sub Test_UriDecode_BasicSpace
    AssertEqual UriDecode("hello%20world"), "hello world", "Basic space decoding"
End Sub

Sub Test_UriDecode_Slash
    AssertEqual UriDecode("%2Fusr%2Fbin"), "/usr/bin", "Slash decoding"
End Sub

Sub Test_UriDecode_NoEncoding
    AssertEqual UriDecode("plaintext"), "plaintext", "No encoding pass-through"
End Sub

Sub Test_UriDecode_EmptyString
    AssertEqual UriDecode(""), "", "Empty string"
End Sub

Sub Test_UriDecode_ConsecutiveEncoded
    AssertEqual UriDecode("%48%65%6C%6C%6F"), "Hello", "Consecutive encoded chars"
End Sub

Sub Test_UriDecode_TrailingPercent
    AssertEqual UriDecode("abc%"), "abc%", "Trailing percent treated as literal"
End Sub

' ============================================================
' WslUriToUncPath tests
' ============================================================

Sub Test_WslUriToUncPath_Basic
    Dim result, outError
    result = WslUriToUncPath("vscode-remote://wsl+Ubuntu/home/user", outError)
    AssertEqual outError, "", "Should have no error"
    AssertEqual result, "\\wsl.localhost\Ubuntu\home\user", "Basic UNC path conversion"
End Sub

Sub Test_WslUriToUncPath_RootPath
    Dim result, outError
    result = WslUriToUncPath("vscode-remote://wsl+Debian/", outError)
    AssertEqual outError, "", "Should have no error"
    AssertEqual result, "\\wsl.localhost\Debian\", "Root path conversion"
End Sub

Sub Test_WslUriToUncPath_EncodedSpaces
    Dim result, outError
    result = WslUriToUncPath("vscode-remote://wsl+Ubuntu/home/my%20folder", outError)
    AssertEqual outError, "", "Should have no error"
    AssertEqual result, "\\wsl.localhost\Ubuntu\home\my folder", "Encoded spaces in path"
End Sub

Sub Test_WslUriToUncPath_DeepPath
    Dim result, outError
    result = WslUriToUncPath("vscode-remote://wsl+Ubuntu/home/user/a/b/c", outError)
    AssertEqual outError, "", "Should have no error"
    AssertEqual result, "\\wsl.localhost\Ubuntu\home\user\a\b\c", "Deep path with slash conversion"
End Sub

Sub Test_WslUriToUncPath_ControlCharInPath
    Dim result, outError
    result = WslUriToUncPath("vscode-remote://wsl+Ubuntu/home/%01user", outError)
    AssertTrue outError <> "", "Control char in decoded path should return error"
    AssertEqual result, "", "Result should be empty on error"
End Sub

Sub Test_WslUriToUncPath_QuoteInPath
    Dim result, outError
    result = WslUriToUncPath("vscode-remote://wsl+Ubuntu/home/%22user", outError)
    AssertTrue outError <> "", "Quote in decoded path should return error"
    AssertEqual result, "", "Result should be empty on error"
End Sub

Sub Test_WslUriToUncPath_PathTraversal
    Dim result, outError
    result = WslUriToUncPath("vscode-remote://wsl+Ubuntu/home/user/../etc", outError)
    AssertTrue outError <> "", "Path traversal should return error"
    AssertTrue InStr(outError, "traversal") > 0, "Error should mention traversal"
    AssertEqual result, "", "Result should be empty on error"
End Sub

' ============================================================
' Run all tests
' ============================================================

' UriDecode
RunTest "Test_UriDecode_BasicSpace"
RunTest "Test_UriDecode_Slash"
RunTest "Test_UriDecode_NoEncoding"
RunTest "Test_UriDecode_EmptyString"
RunTest "Test_UriDecode_ConsecutiveEncoded"
RunTest "Test_UriDecode_TrailingPercent"

' WslUriToUncPath
RunTest "Test_WslUriToUncPath_Basic"
RunTest "Test_WslUriToUncPath_RootPath"
RunTest "Test_WslUriToUncPath_EncodedSpaces"
RunTest "Test_WslUriToUncPath_DeepPath"
RunTest "Test_WslUriToUncPath_ControlCharInPath"
RunTest "Test_WslUriToUncPath_QuoteInPath"
RunTest "Test_WslUriToUncPath_PathTraversal"

' Summary
WScript.Echo ""
WScript.Echo passedTests & " passed, " & failedTests & " failed, " & totalTests & " total"
WScript.Quit failedTests
