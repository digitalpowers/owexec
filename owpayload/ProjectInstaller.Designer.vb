<System.ComponentModel.RunInstaller(True)> Partial Class ProjectInstaller
    Inherits System.Configuration.Install.Installer

    'Installer overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.spInstaller = New System.ServiceProcess.ServiceProcessInstaller
        Me.svcInstaller = New System.ServiceProcess.ServiceInstaller
        '
        'spInstaller
        '
        Me.spInstaller.Password = Nothing
        Me.spInstaller.Username = Nothing
        '
        'svcInstaller
        '
        Me.svcInstaller.Description = "Payload for owexec, does the starting of the program"
        Me.svcInstaller.DisplayName = "OW Payload"
        Me.svcInstaller.ServiceName = "OW Payload"
        Me.svcInstaller.StartType = System.ServiceProcess.ServiceStartMode.Automatic
        '
        'ProjectInstaller
        '
        Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.spInstaller, Me.svcInstaller})

    End Sub
    Friend WithEvents spInstaller As System.ServiceProcess.ServiceProcessInstaller
    Friend WithEvents svcInstaller As System.ServiceProcess.ServiceInstaller

End Class
