Imports QuestPDF.Infrastructure

Class Application

    Protected Overrides Sub OnStartup(e As StartupEventArgs)
        ' Set QuestPDF license here
        QuestPDF.Settings.License = LicenseType.Community

        MyBase.OnStartup(e)
    End Sub
End Class
