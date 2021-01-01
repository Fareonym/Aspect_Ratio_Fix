Const cHKLM = &H80000002
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
Set oShell = CreateObject("WScript.Shell")
sKeyPath = "SYSTEM\CurrentControlSet\Control\GraphicsDrivers\Configuration"
oReg.EnumKey cHKLM, sKeyPath, aSubKeys
If Not (isnull(aSubKeys)) Then
    For Each sSubkey In aSubKeys
        sTmpValueName = "HKLM\SYSTEM\CurrentControlSet\Control\GraphicsDrivers\Configuration\" & sSubKey & "\00\00\Scaling"
        Wscript.Echo "Attempting to write 4 value to: """ & sTmpValueName & """"
        oShell.RegWrite sTmpValueName, 4, "REG_DWORD"
        If Err.Number <> 0 Then
            WScript.Echo "Write failed with errors: " & Err.Number
        Else
            WScript.Echo "Write succedded."
        End If
    Next
End If
