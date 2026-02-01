Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

' --- Script dizini ---
scriptDir = FSO.GetParentFolderName(WScript.ScriptFullName)
adbPath = scriptDir & "\adb.exe"
scrcpyPath = scriptDir & "\scrcpy.exe"

If Not FSO.FileExists(adbPath) Or Not FSO.FileExists(scrcpyPath) Then
    MsgBox "adb.exe veya scrcpy.exe bulunamadi!", 16, "Hata"
    WScript.Quit
End If

' --- Pair adresi sor (iptal = sadece scrcpy) ---
addr = InputBox( _
    "PAIR gerekiyorsa IP:PORT gir" & vbCrLf & _
    "PAIR zaten yapiliysa BOS birak", _
    "ADB Pair (Opsiyonel)")

' --- Pair (sessiz) ---
If addr <> "" Then
    WshShell.Run """" & adbPath & """ pair " & addr, 0, True
End If

' --- scrcpy + mDNS auto-connect AYNI PROCESS ---
cmd = "cmd /c set ADB_MDNS_AUTO_CONNECT=adb-tls-connect && " & _
"""" & scrcpyPath & _
""" --video-bit-rate 12M --max-size 2400 --max-fps 120 --video-codec=h265"

WshShell.Run cmd, 0, False
