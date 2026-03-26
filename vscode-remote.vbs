If WScript.ScriptName = "vscode-remote.vbs" Then
    Set WshShell = CreateObject("WScript.Shell")
    scriptPath = WScript.ScriptFullName
    uri = ""
    If WScript.Arguments.Count > 0 Then
        uri = WScript.Arguments(0)
    End If

    If uri = "--install" Then
        Dim regBase
        regBase = "HKCU\Software\Classes\vscode-remote"

        WshShell.RegWrite regBase & "\", "URL:vscode-remote Protocol", "REG_SZ"
        WshShell.RegWrite regBase & "\URL Protocol", "", "REG_SZ"
        WshShell.RegWrite regBase & "\shell\open\command\", _
            "wscript.exe """ & scriptPath & """ ""%1""", "REG_SZ"

        WScript.Echo "Done. " & scriptPath
        WScript.Quit 0
    End If

    If uri = "--uninstall" Then
        DeleteRegistryTree "HKCU\Software\Classes\vscode-remote"
        WScript.Echo "Uninstalled."
        WScript.Quit 0
    End If

    Dim uncPath, conversionError
    uncPath = WslUriToUncPath(uri, conversionError)
    If conversionError <> "" Then
        WScript.Echo conversionError
        WScript.Quit 1
    End If

    WshShell.Run "explorer.exe """ & uncPath & """", 1, False
End If

Sub DeleteRegistryTree(regPath)
    Dim objReg, subPath, lReturn
    Const HKCU = &H80000001

    On Error Resume Next
    Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    If Err.Number <> 0 Then
        WScript.Echo "Failed to access registry provider: " & Err.Description
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    subPath = Mid(regPath, InStr(regPath, "\") + 1)
    DeleteSubKeys objReg, HKCU, subPath

    lReturn = objReg.DeleteKey(HKCU, subPath)
    If lReturn <> 0 Then
        WScript.Echo "Failed to delete registry key: " & subPath
    End If
End Sub

Sub DeleteSubKeys(objReg, hive, keyPath)
    Dim arrSubKeys, subKey, lReturn
    objReg.EnumKey hive, keyPath, arrSubKeys
    If Not IsNull(arrSubKeys) Then
        For Each subKey In arrSubKeys
            DeleteSubKeys objReg, hive, keyPath & "\" & subKey
            lReturn = objReg.DeleteKey(hive, keyPath & "\" & subKey)
            If lReturn <> 0 Then
                WScript.Echo "Failed to delete registry key: " & keyPath & "\" & subKey
            End If
        Next
    End If
End Sub

Function WslUriToUncPath(wslUri, ByRef outError)
    If Left(wslUri, 16) <> "vscode-remote://" Then
        outError = "Invalid URI: unsupported scheme."
        WslUriToUncPath = ""
        Exit Function
    End If

    Dim re, matches, distro, path
    Set re = New RegExp
    re.Pattern = "^vscode-remote://wsl\+([^/]+)(/.*)$"
    re.IgnoreCase = False
    Set matches = re.Execute(wslUri)
    distro = matches(0).SubMatches(0)
    path = UriDecode(matches(0).SubMatches(1))

    Dim i
    For i = 1 To Len(path)
        If Asc(Mid(path, i, 1)) < 32 Then
            outError = "Invalid path: contains control characters."
            WslUriToUncPath = ""
            Exit Function
        End If
    Next
    If InStr(path, """") > 0 Then
        outError = "Invalid path."
        WslUriToUncPath = ""
        Exit Function
    End If
    If InStr(path, "/..") > 0 Then
        outError = "Invalid path: path traversal detected."
        WslUriToUncPath = ""
        Exit Function
    End If

    outError = ""
    WslUriToUncPath = "\\wsl.localhost\" & distro & Replace(path, "/", "\")
End Function

Function UriDecode(s)
    Dim result, j, ch
    result = ""
    j = 1
    Do While j <= Len(s)
        ch = Mid(s, j, 1)
        If ch = "%" And j + 2 <= Len(s) Then
            result = result & Chr(CInt("&H" & Mid(s, j + 1, 2)))
            j = j + 3
        Else
            result = result & ch
            j = j + 1
        End If
    Loop
    UriDecode = result
End Function
