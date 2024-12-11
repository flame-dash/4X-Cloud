' Prompt user for username
username = InputBox("Enter your username:", "Login")

' Prompt user for password
password = InputBox("Enter your password:", "Login")

' Validate username and password
If username = "admin" And password = "x-cloud" Then
    MsgBox "Login successful!", vbInformation, "Success"
    
    ' Define the path to your custom PowerShell script
    powerShellScriptPath = "C:\path\to\your\script.ps1" ' Replace with the actual path to your PowerShell script
    
    ' Launch the PowerShell script
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "powershell.exe -ExecutionPolicy Bypass -File """ & powerShellScriptPath & """", 1, False
    
    ' Optionally, you can also display a message after launching the PowerShell script
    MsgBox "Your PowerShell script is now running.", vbInformation, "Script Started"
    
    Set objShell = Nothing
Else
    MsgBox "Invalid username or password!", vbCritical, "Login Failed"
End If
