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

If InStr(uri, """") > 0 Then
    WScript.Echo "Invalid URI."
    WScript.Quit 1
End If

Set re = New RegExp
re.Pattern = "^vscode-remote://wsl\+([^/]+)(/.*)$"
re.IgnoreCase = False

If re.Test(uri) Then
    Set matches = re.Execute(uri)
    distro = matches(0).SubMatches(0)
    wslPath = Replace(matches(0).SubMatches(1), "/", "\")
    uncPath = "\\wsl.localhost\" & distro & wslPath
    WshShell.Run "explorer.exe """ & uncPath & """", 1, False
Else
    codePath = FindCode()
    WshShell.Run """" & codePath & """" & " " & """" & uri & """", 1, False
End If

Function FindCode()
    Dim exec, line
    On Error Resume Next
    Set exec = WshShell.Exec("where code")
    If Err.Number = 0 Then
        line = Trim(exec.StdOut.ReadLine())
        If Len(line) > 0 Then
            FindCode = line
            On Error GoTo 0
            Exit Function
        End If
    End If
    On Error GoTo 0
    FindCode = WshShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & _
        "\Programs\Microsoft VS Code\Code.exe"
End Function

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
