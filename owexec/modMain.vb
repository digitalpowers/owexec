Imports System.Management

Module modMain
    Dim args() As String
    Dim user As String = ""
    Dim compname As String = ""
    Dim app As String = ""
    Dim params As String = ""
    Dim newargs As String = ""

    Dim fname As String
    Dim copy As Boolean = False
    Dim nowait As Boolean = False

    Sub Main()
        args = Environment.GetCommandLineArgs()
        If (args.Length = 1) Then
            displayUsage()
            Return
        End If

        For i As Integer = 1 To args.Length - 1
            If (args(i) = "-copy") Then
                copy = True
            ElseIf (args(i) = "-nowait") Then
                nowait = True
            Else
                If (args(i).Contains(" ")) Then
                    newargs += """" + args(i) + """ "
                Else
                    newargs += args(i) + " "
                End If

                If (args(i) = "-u") Then
                    If (i + 1 < args.Length) Then
                        user = args(i + 1)
                    End If
                ElseIf (args(i) = "-c") Then
                    If (i + 1 < args.Length) Then
                        compname = args(i + 1)
                    End If
                ElseIf (args(i) = "-k") Then
                    If (i + 1 < args.Length) Then
                        app = args(i + 1)
                    End If
                ElseIf (args(i) = "-p") Then
                    If (i + 1 < args.Length) Then
                        params = args(i + 1)
                    End If
                End If
            End If
        Next

        If (compname <> "" And app <> "") Then
            If (copy) Then
                If (app.Contains("/") Or app.Contains("\")) Then
                    Dim fnameindex As Integer = app.LastIndexOf("\")
                    If (fnameindex = -1) Then fnameindex = app.LastIndexOf("/")
                    fname = app.Substring(fnameindex)
                Else
                    fname = app
                End If
            End If
            If (installService()) Then
                waitForServiceToExit()
                deleteService()
            End If
        Else
            displayUsage()
            Return
        End If

        waitAtEnd()

    End Sub

    Private Sub waitAtEnd()
        If (Not nowait) Then
            Console.WriteLine("Press any key to close")
            Console.ReadKey()
        End If
    End Sub

    Private Sub displayUsage()
        Console.WriteLine("owexec v-1.1 USAGE")
        Console.WriteLine("owexec -c computername -k command [ -p parameters ] [ -u domain\user ] [ -copy ] [ -nowait ]")
        Console.WriteLine("")
        Console.WriteLine(vbTab + "-c the computer host name or ip of the target computer")
        Console.WriteLine("")
        Console.WriteLine(vbTab + "-k the command to be run, relative to the destination ")
        Console.WriteLine(vbTab + "   computer. ex: c:\windows\system32\notepad.exe")
        Console.WriteLine("")
        Console.WriteLine(vbTab + "-p the parameters to pass to the program, optional")
        Console.WriteLine("")
        Console.WriteLine(vbTab + "-u the user whose context the program should be run in")
        Console.WriteLine(vbTab + "   if ommitted the first user that is found will be used")
        Console.WriteLine("")
        Console.WriteLine(vbTab + "-copy finds the command referenced with -k on the local")
        Console.WriteLine(vbTab + "   machine and copies it to the comptuer referenced in ")
        Console.WriteLine(vbTab + "   -c on the admin$ share then runs it from there")
        Console.WriteLine("")
        Console.WriteLine(vbTab + "-nowait does not ask to press a key when the program finishes")
        Console.WriteLine("")
        Console.WriteLine("")
        Console.WriteLine("download the current version at officewarfare.net")
        Console.WriteLine("")
        waitAtEnd()
    End Sub

    Private Sub waitForServiceToExit()
        Try
            Dim ms As New ManagementScope("\\" + compname + "\root\CIMV2")

            Dim oq As New ObjectQuery("Select * from Win32_Service WHERE Name='OWPayload'")
            Dim query As New ManagementObjectSearcher(ms, oq)
            Console.WriteLine("waiting for the service to stop")
            Dim count As Integer = 0
            While (True)

                Dim queryCollection As ManagementObjectCollection = query.Get()
                For Each mo As ManagementObject In queryCollection
                    Console.Write(".")
                    count += 1
                    If (Boolean.Parse((mo.Properties("Started").Value.ToString())) = False) Then
                        Console.WriteLine("")
                        Return
                    End If
                    If (count > 50) Then
                        Console.WriteLine("")
                        Console.WriteLine("timeout waiting for service to stop")
                        Return
                    End If
                    System.Threading.Thread.Sleep(500)
                    Exit For
                Next
            End While
        Catch ex As Exception
            Console.WriteLine("Error waiting for service to close")
        End Try
    End Sub

    Public Function ping(ByVal ComputerName As String) As Boolean
        Try
            Dim ms As ManagementScope = New ManagementScope("root\cimv2")
            Dim oq As ObjectQuery = New ObjectQuery("Select * From Win32_PingStatus Where Address = '" + ComputerName + "'")
            Dim query As ManagementObjectSearcher = New ManagementObjectSearcher(ms, oq)
            Dim queryCollection As ManagementObjectCollection = query.Get
            For Each mo As ManagementObject In queryCollection
                Select Case Integer.Parse(mo("StatusCode").ToString)
                    Case 0
                        Return True

                    Case Else
                        Return False
                End Select
            Next
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function installService() As Boolean
        Try
            Console.WriteLine("installing service remotely")
            Dim sysRoot As String

            If (Not ping(compname)) Then
                Console.WriteLine("Ping Failed")
                'Return False
            End If


            If (System.IO.Directory.Exists("\\" + compname + "\admin$\")) Then
                sysRoot = "\\" + compname + "\admin$"
            Else
                Console.WriteLine("Admin$ Share Not Shared: Not Found")
                Return False
            End If

            If (Not System.IO.Directory.Exists(sysRoot + "\Microsoft.NET\Framework\v2.0.50727")) Then
                Console.WriteLine(".NET 2.0 Not Installed")
                Return False
            End If

            If Not System.IO.File.Exists(sysRoot + "\system32\owpayload.exe") Then
                Dim writer As New IO.BinaryWriter(New System.IO.FileStream(sysRoot + "\system32\owpayload.exe", IO.FileMode.Create))
                writer.Write(My.Resources.owpayload)
                writer.Close()
            End If

            If copy And Not System.IO.File.Exists(sysRoot + "\system32\" + fname) Then
                IO.File.Copy(app, sysRoot + "\system32\" + fname)
                app = fname
            End If

            Dim ms As New ManagementScope("\\" + compname + "\root\CIMV2")
            ''IPEnabled = True should give us the primary network card
            Dim mc As New Management.ManagementClass(ms, New ManagementPath("Win32_Service"), Nothing)
            Dim inArgs As ManagementBaseObject = mc.Methods("Create").InParameters

            ' Add the input parameters. 
            inArgs("Name") = "OWPayload" '< - Service Name 
            inArgs("DisplayName") = "OWPayload" '< - Display Name, what you see in the Services control panel 
            inArgs("PathName") = """owpayload.exe""" + " " + newargs '+ " " + user + " " + app + " " + params + "" '< - Path and Command Line of the executable 
            'Console.WriteLine("""owpayload.exe""" + " " + newargs)
            inArgs("ServiceType") = 16
            inArgs("ErrorControl") = 0
            inArgs("StartMode") = "Automatic"
            inArgs("DesktopInteract") = True
            inArgs("StartName") = "LocalSystem"
            inArgs("StartPassword") = ""

            mc.InvokeMethod("Create", inArgs, Nothing)
        Catch ex As Exception
            Console.WriteLine("Error installing the service")
        End Try

        Try
            Dim ms As New ManagementScope("\\" + compname + "\root\CIMV2")
            ''IPEnabled = True should give us the primary network card
            Dim oq As New ObjectQuery("Select * from Win32_Service WHERE Name='OWPayload'")
            Dim query As New ManagementObjectSearcher(ms, oq)

            Dim queryCollection As ManagementObjectCollection = query.Get()
            For Each mo As ManagementObject In queryCollection
                mo.InvokeMethod("StartService", Nothing)
                Return True
            Next
        Catch ex As Exception
            Console.WriteLine("Error starting the service")
            Return False
        End Try

    End Function

    Private Sub deleteService()
        Try
            Console.WriteLine("removing service")
            Dim sysRoot As String

            If (Not ping(compname)) Then
                Console.WriteLine("Ping Failed")
                Return
            End If


            If (System.IO.Directory.Exists("\\" + compname + "\admin$\")) Then
                sysRoot = "\\" + compname + "\admin$"
            Else
                Console.WriteLine("Admin$ Share Not Shared: Not Found")
                Return
            End If

            Dim ms As New ManagementScope("\\" + compname + "\root\CIMV2")
            ''IPEnabled = True should give us the primary network card
            Dim oq As New ObjectQuery("Select * from Win32_Service WHERE Name='OWPayload'")
            Dim query As New ManagementObjectSearcher(ms, oq)

            Dim queryCollection As ManagementObjectCollection = query.Get()
            For Each mo As ManagementObject In queryCollection
                mo.InvokeMethod("StopService", Nothing)
                mo.InvokeMethod("Delete", Nothing)
                Exit For
            Next

            If System.IO.File.Exists(sysRoot + "\system32\owpayload.exe") Then
                IO.File.Delete(sysRoot + "\system32\owpayload.exe")
            End If
        Catch ex As Exception
            Console.WriteLine("Error removing the service")
        End Try

    End Sub

End Module
