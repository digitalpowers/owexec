Imports System
Imports System.Runtime.InteropServices
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Security
Imports System.Security.Principal
Imports System.Management

Public Class ExecuteProcess

    Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As IntPtr, _
    ByVal DesiredAccess As Integer, ByRef TokenHandle As IntPtr) As Boolean

    Declare Auto Function CloseHandle Lib "kernel32.dll" (ByVal hObject As IntPtr) As Boolean
    Declare Function CloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Integer) As Integer

    Declare Function DuplicateToken Lib "advapi32.dll" (ByVal ExistingTokenHandle As IntPtr, _
    ByVal SECURITY_IMPERSONATION_LEVEL As Int16, ByRef DuplicateTokenHandle As IntPtr) As Boolean

    Private Declare Auto Function CreateProcessAsUser Lib "advapi32" ( _
        ByVal hToken As IntPtr, _
        ByVal strApplicationName As String, _
        ByVal strCommandLine As String, _
        ByRef lpProcessAttributes As SECURITY_ATTRIBUTES, _
        ByRef lpThreadAttributes As SECURITY_ATTRIBUTES, _
        ByVal bInheritHandles As Boolean, _
        ByVal dwCreationFlags As Integer, _
        ByVal lpEnvironment As IntPtr, _
        ByVal lpCurrentDriectory As String, _
        ByRef lpStartupInfo As STARTUPINFO, _
        ByRef lpProcessInformation As PROCESS_INFORMATION) As Boolean

    '<DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = True)> _ 
    'Public extern Shared Property CreateProcessWithTokenW(() As Boolean
    'End Property
    '    IntPtr hToken,
    '    Integer LogonFlags,
    '    String lpApplicationName,
    '    <In) String lpCommandLine,
    '    Integer CreationFlags,
    '    IntPtr lpEnvironment,
    '    String lpCurrentDirectory,
    '     STARTUPINFO lpStartupInfo,
    '     PROCESS_INFORMATION lpProcessInformation)

    Private Shared Function OpenDesktop(ByVal lpszDesktop As String, ByVal dwFlags As Integer, _
    ByVal fInderit As Boolean, ByVal dwDesiredAccess As Integer) As IntPtr
    End Function

    Public Structure SECURITY_ATTRIBUTES
        Public nLength As Integer
        Public lpSecurityDescriptor As IntPtr
        Public bInheritHandle As Integer
    End Structure

    Public Structure PROCESS_INFORMATION
        Public hProcess As IntPtr
        Public hThread As IntPtr
        Public dwProcessId As Integer
        Public dwThreadId As Integer
    End Structure

    Public Structure STARTUPINFO
        Public cb As Int32
        Public lpReserved As String
        Public lpDesktop As IntPtr
        Public lpTitle As String
        Public dwX As Int32
        Public dwY As Int32
        Public dwXSize As Int32
        Public dwYSize As Int32
        Public dwXCountChars As Int32
        Public dwYCountChars As Int32
        Public dwFillAttribute As Int32
        Public dwFlags As Int32
        Public wShowWindow As Int16
        Public cbReserved2 As Int16
        Public lpReserved2 As IntPtr
        Public hStdInput As IntPtr
        Public hStdOutput As IntPtr
        Public hStdError As IntPtr
    End Structure


    Public Const TOKEN_DUPLICATE As Integer = 2
    Public Const TOKEN_QUERY As Integer = &H8
    Public Const TOKEN_IMPERSONATE As Integer = &H4

    Public Const LOGON_WITH_PROFILE As Integer = &H1

    'READ_CONTROL | WRITE_DAC |DESKTOP_WRITEOBJECTS | DESKTOP_READOBJECTS);
    Public Const READ_CONTROL As Long = &H20000L
    Public Const WRITE_DAC As Long = &H40000L
    Public Const DESKTOP_WRITEOBJECTS As Long = &H80L
    Public Const DESKTOP_READOBJECTS As Long = &H1L

    Public Shared Sub execute(ByVal processid As Integer, ByVal appName As String, ByVal arguments As String, Optional ByVal visible As Boolean = True)
        Dim hToken As IntPtr = IntPtr.Zero
        Dim dupeTokenHandle As IntPtr = IntPtr.Zero
        ' Must have passed in a process id, get a handle to the process
        Dim proc As Process = Process.GetProcessById(processid)
        'try to get a token from the process
        If OpenProcessToken(proc.Handle, TOKEN_QUERY Or TOKEN_IMPERSONATE Or TOKEN_DUPLICATE, hToken) <> 0 Then

            'create a new idenity from the token
            Dim NewId As WindowsIdentity = New WindowsIdentity(hToken)
            Try
                Const SecurityImpersonation As Integer = 2
                DuplicateToken(hToken, SecurityImpersonation, dupeTokenHandle)

                If IntPtr.Zero = dupeTokenHandle Then

                End If

                Dim impersonatedUser As WindowsImpersonationContext = NewId.Impersonate()

                'Execute the process with wmi, now that we are impersonateing the user
                Dim objConnectionOptions As ConnectionOptions = New ConnectionOptions()
                objConnectionOptions.Impersonation = ImpersonationLevel.Impersonate
                objConnectionOptions.EnablePrivileges = True
                Dim objManagementScope As ManagementScope = New ManagementScope("\\" + "." + "\root\cimv2")
                objManagementScope.Connect()
                Dim processClassFinal As ManagementClass = New ManagementClass(objManagementScope, New ManagementPath("Win32_Process"), Nothing)
                Dim inParamsFinal As ManagementBaseObject = proWcessClassFinal.GetMethodParameters("Create")
                Dim commandline As String = appName + " " + arguments
                report(commandline)
                inParamsFinal("CommandLine") = commandline
                Dim ProcessStartup As ManagementClass
                If (Not visible) Then
                    ''Then make it INVISIBLE!
                    ProcessStartup = New ManagementClass(objManagementScope, New ManagementPath("Win32_ProcessStartup"), Nothing)
                    ProcessStartup.Properties.Add("ShowWindow", 0) 'SW_HIDE
                    inParamsFinal("ProcessStartupInformation") = ProcessStartup
                End If

                Dim outParamsFinal As ManagementBaseObject = processClassFinal.InvokeMethod("Create", inParamsFinal, Nothing)

                'Diagnostics.Process.Start(appName, arguments)


                'Switch back to our old context
                impersonatedUser.Undo()

            Catch ex As Exception
                report("ExecuteProcess.execute " + ex.Message)
            Finally
                CloseHandle(hToken)

            End Try

        Else
            report("Could not open process for token")
        End If
        Return
    End Sub

    Public Shared Function getExplorerProcID(ByVal username As String) As Integer
        If (username = "") Then
            Return getExplorerProcID()
        End If
        ''Create a query object
        Dim query As ManagementObjectSearcher
        Dim oq As ObjectQuery
        Dim queryCollection As ManagementObjectCollection

        Dim ms As ManagementScope = New ManagementScope("\\" + "." + "\root\cimv2")
        oq = New System.Management.ObjectQuery("SELECT * FROM Win32_Process Where Name = '" + "explorer.exe" + "'")

        query = New ManagementObjectSearcher(ms, oq)

        queryCollection = query.Get()

        For Each mo As ManagementObject In queryCollection
            Dim param() As Object = {"", ""}
            mo.InvokeMethod("GetOwner", param)
            Dim exOwner As String = param(1).ToString() + "\" + param(0).ToString()
            report(exOwner)
            If (exOwner.ToLower() = username.ToLower()) Then
                Return Integer.Parse(mo("ProcessID").ToString())
            End If
        Next

        Return 0
    End Function

    Public Shared Function getExplorerProcID() As Integer

        ''Create a query object
        Dim query As ManagementObjectSearcher
        Dim oq As ObjectQuery
        Dim queryCollection As ManagementObjectCollection

        Dim ms As ManagementScope = New ManagementScope("\\" + "." + "\root\cimv2")
        oq = New System.Management.ObjectQuery("SELECT * FROM Win32_Process Where Name = '" + "explorer.exe" + "'")

        query = New ManagementObjectSearcher(ms, oq)

        queryCollection = query.Get()

        For Each mo As ManagementObject In queryCollection
            Dim param() As Object = {"", ""}
            mo.InvokeMethod("GetOwner", param)
            
            Return Integer.Parse(mo("ProcessID").ToString())

        Next

        Return 0
    End Function


    Shared Sub report(ByVal str As String)
        'Dim fstream As New IO.StreamWriter("owexec.log", True)
        'fstream.WriteLine(str)
        'fstream.Close()
    End Sub

End Class

