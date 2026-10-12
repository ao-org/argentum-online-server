Attribute VB_Name = "modManagedConfig"
' Argentum 20 Game Server
'
'    Copyright (C) 2023-2026 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit

Private Declare Sub ExitProcess Lib "kernel32" (ByVal exitCode As Long)

Public Enum e_ManagedConfig
    mcServer = 0
    mcMain = 1
    mcFeatures = 2
    mcBalance = 3
    mcEvents = 4
    mcIntervals = 5
End Enum

Private mConfigDirectory As String
Private mStartupComplete As Boolean

' Quotes group an argument; backslashes in Windows paths are literal.
' Unknown options are left alone. Only duplicate occurrences of the requested option fail.
Public Function StartupOption(ByVal commandLine As String, ByVal optionName As String, ByRef present As Boolean) As String
    Dim tokens As New Collection
    Dim token As String, ch As String
    Dim quoted As Boolean, started As Boolean
    Dim i As Long, value As String, prefix As String
    present = False
    For i = 1 To Len(commandLine)
        ch = Mid$(commandLine, i, 1)
        If ch = Chr$(34) Then
            quoted = Not quoted
            started = True
        ElseIf (ch = " " Or ch = vbTab) And Not quoted Then
            If started Then Call tokens.Add(token)
            token = vbNullString
            started = False
        Else
            If AscW(ch) < 32 And ch <> vbTab Then Err.Raise 5, "StartupOption", "Invalid control character in command line"
            token = token & ch
            started = True
        End If
    Next i
    If quoted Then Err.Raise 5, "StartupOption", "Unclosed quote in command line"
    If started Then Call tokens.Add(token)
    prefix = LCase$(optionName) & "="
    For i = 1 To tokens.Count
        token = tokens(i)
        If LCase$(Left$(token, Len(prefix))) = prefix Or LCase$(token) = LCase$(optionName) Then
            If present Then Err.Raise 5, "StartupOption", "Duplicate " & optionName & " argument"
            present = True
            If LCase$(token) = LCase$(optionName) Then
                If i = tokens.Count Then Err.Raise 5, "StartupOption", "Missing value for " & optionName
                i = i + 1
                value = tokens(i)
                If Left$(value, 1) = "/" Or Left$(value, 2) = "--" Then Err.Raise 5, "StartupOption", "Missing value for " & optionName
            Else
                value = Mid$(token, Len(prefix) + 1)
            End If
            If LenB(Trim$(value)) = 0 Then Err.Raise 5, "StartupOption", "Empty value for " & optionName
            StartupOption = value
        End If
    Next i
End Function

Public Function ManagedConfigName(ByVal config As e_ManagedConfig) As String
    Select Case config
        Case mcServer: ManagedConfigName = "Server.ini"
        Case mcMain: ManagedConfigName = "Configuracion.ini"
        Case mcFeatures: ManagedConfigName = "feature_toggle.ini"
        Case mcBalance: ManagedConfigName = "Balance.dat"
        Case mcEvents: ManagedConfigName = "Eventos.dat"
        Case mcIntervals: ManagedConfigName = "intervalos.ini"
        Case Else: Err.Raise 5, "ManagedConfigName", "Unknown managed configuration"
    End Select
End Function

Public Function ExternalConfigEnabled() As Boolean
    ExternalConfigEnabled = LenB(mConfigDirectory) <> 0
End Function

Public Function ManagedConfigPath(ByVal config As e_ManagedConfig) As String
    Dim base As String
    If ExternalConfigEnabled() Then
        base = mConfigDirectory
    ElseIf config = mcBalance Or config = mcEvents Then
        base = DatPath
    Else
        base = IniPath
    End If
    ' Legacy bases already include their separator; preserve their exact semantics.
    ManagedConfigPath = base & ManagedConfigName(config)
End Function

Public Sub RequireReadableConfig(ByVal filePath As String)
    On Error GoTo Failed
    Dim handle As Integer, opened As Boolean, chunk As String, remaining As Long
    If (GetAttr(filePath) And vbDirectory) <> 0 Then Err.Raise 5
    handle = FreeFile()
    ' Access Read never creates a missing file. Deny writes while checking the contents.
    Open filePath For Binary Access Read Lock Write As #handle
    opened = True
    remaining = LOF(handle)
    Do While remaining > 0
        chunk = Space$(IIf(remaining > 4096, 4096, remaining))
        Get #handle, , chunk
        remaining = remaining - Len(chunk)
    Loop
    Close #handle
    Exit Sub
Failed:
    Dim failure As Long
    failure = Err.Number
    If opened Then Close #handle
    Err.Raise failure, "ManagedConfig", "Cannot read required configuration: " & filePath & " (error " & failure & ")"
End Sub

Public Sub InitializeManagedConfig(ByVal commandLine As String)
    Dim present As Boolean, directory As String, fs As New FileSystemObject
    Dim config As e_ManagedConfig
    directory = StartupOption(commandLine, "/config-dir", present)
    If present Then
        directory = fs.GetAbsolutePathName(directory)
        If Not fs.FolderExists(directory) Then Err.Raise 76, "ManagedConfig", "Configuration directory does not exist or is inaccessible: " & directory
        If Right$(directory, 1) <> "\" Then directory = directory & "\"
        ' Validate the entire required set before publishing the selection or reading Server.ini.
        For config = mcServer To mcIntervals
            Call RequireReadableConfig(directory & ManagedConfigName(config))
        Next config
    End If
    mConfigDirectory = directory
End Sub

Public Function ReadManagedConfig(ByVal config As e_ManagedConfig) As clsIniManager
    Dim reader As New clsIniManager
    Call RequireReadableConfig(ManagedConfigPath(config))
    Call reader.Initialize(ManagedConfigPath(config), True)
    Set ReadManagedConfig = reader
End Function

Public Sub WriteManagedConfigValue(ByVal config As e_ManagedConfig, ByVal section As String, ByVal key As String, ByVal value As String)
    Dim filePath As String
    filePath = ManagedConfigPath(config)
    Call RequireReadableConfig(filePath)
    If writeprivateprofilestring(section, key, value, filePath) = 0 Then
        Err.Raise 75, "ManagedConfig", "Cannot write configuration: " & filePath
    End If
End Sub

Public Sub LogManagedConfigPaths()
    Dim config As e_ManagedConfig, fs As New FileSystemObject, mode As String
    If ExternalConfigEnabled() Then
        mode = "External configuration directory: " & mConfigDirectory
    Else
        mode = "Legacy configuration locations (no /config-dir)"
    End If
    Call LogThis(0, mode, vbLogEventTypeInformation)
    Call AgregarAConsola(mode)
    For config = mcServer To mcIntervals
        mode = ManagedConfigName(config) & " -> " & fs.GetAbsolutePathName(ManagedConfigPath(config))
        Call LogThis(0, mode, vbLogEventTypeInformation)
        Call AgregarAConsola(mode)
    Next config
End Sub

Public Sub CompleteConfigStartup()
    mStartupComplete = True
End Sub

Public Function ConfigStartupComplete() As Boolean
    ConfigStartupComplete = mStartupComplete
End Function

Public Sub ReportManagedConfigFailure(ByVal operation As String, ByVal number As Long, ByVal description As String)
    Call TraceError(number, description, operation, 0)
    Call AgregarAConsola("Configuration load failed: " & operation & ": " & description)
    If Not mStartupComplete Then Err.Raise number, operation, description
End Sub

Public Sub FatalConfigStartup(ByVal number As Long, ByVal description As String)
    On Error GoTo Terminate
    Call TraceError(number, "Fatal startup: " & description, "General.Main", 0)
Terminate:
    ' An unattended service must fail, not wait indefinitely for a modal dialog.
    Call ExitProcess(1)
End Sub

Public Function ReadManagedProfileValue(ByVal config As e_ManagedConfig, ByVal section As String, ByVal key As String) As String
    On Error GoTo Failed
    Dim filePath As String, handle As Integer, opened As Boolean, failure As Long
    Dim checked As clsIniManager
    filePath = ManagedConfigPath(config)
    Set checked = ReadManagedConfig(config)
    handle = FreeFile()
    ' Keep the file readable and stable while using the original Win32 INI semantics
    ' (including quoted values). Input mode cannot create a missing file.
    Open filePath For Input Access Read Lock Write As #handle
    opened = True
    ReadManagedProfileValue = GetVar(filePath, section, key)
    Close #handle
    Exit Function
Failed:
    failure = Err.Number
    If opened Then Close #handle
    Err.Raise failure, "ManagedConfig", "Cannot read configuration: " & filePath & " (error " & failure & ")"
End Function
