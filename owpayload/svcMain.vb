Public Class svcMain

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
        Dim cLineArgs = Environment.GetCommandLineArgs()
        report("Command Line: " + Environment.CommandLine)
        report("Num Args: " + cLineArgs.Length.ToString())

        'report("ArgsLen: " + cLineArgs.Length.ToString())
        'For Each Str As String In cLineArgs
        '    report("arg: " + Str)
        'Next
        Dim user As String = ""
        Dim compname As String = ""
        Dim app As String = ""
        Dim params As String = ""

        For i As Integer = 1 To cLineArgs.Length - 1
            If (cLineArgs(i) = "-u") Then
                report("-u found")
                If (i + 1 < cLineArgs.Length) Then
                    report("-u recorded")
                    user = cLineArgs(i + 1)
                End If
            ElseIf (cLineArgs(i) = "-c") Then
                report("-c found")
                If (i + 1 < cLineArgs.Length) Then
                    report("-c recorded")
                    compname = cLineArgs(i + 1)
                End If
            ElseIf (cLineArgs(i) = "-k") Then
                report("-k found")
                If (i + 1 < cLineArgs.Length) Then
                    report("-k recorded")
                    app = cLineArgs(i + 1)
                End If
            ElseIf (cLineArgs(i) = "-p") Then
                report("-p found")
                If (i + 1 < cLineArgs.Length) Then
                    report("-p recorded")
                    params = cLineArgs(i + 1)
                End If

            End If
        Next
        'If (cLineArgs.Length >= 3) Then
        '    ''Load the parameters
        '    Dim user As String = cLineArgs(1)
        '    Dim process As String = cLineArgs(2)
        '    Dim params As String = ""
        '    If (cLineArgs.Length >= 4) Then
        '        params = cLineArgs(3)
        '    End If
        report("app: " + app)
        report("user: " + user)
        report("params: " + params)

        If (app <> "") Then
            Dim procID As Integer = ExecuteProcess.getExplorerProcID(user)
            report("procid: " + procID.ToString())
            ExecuteProcess.execute(procID, app, params)

        End If

        report("closing")
        report("")
        Me.Stop()
    End Sub

    Sub report(ByVal str As String)
        'Dim fstream As New IO.StreamWriter("c:\owexec.log", True)
        'fstream.WriteLine(str)
        'fstream.Close()
    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
    End Sub

End Class
