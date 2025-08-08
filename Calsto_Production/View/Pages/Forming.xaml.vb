Imports System.Collections.ObjectModel

Public Class Forming
    Inherits Page

    Private FormingJobList As ObservableCollection(Of FormingModel)

    Public Sub New()
        InitializeComponent()
        LoadformingJobs()
    End Sub

    ' Load all forming jobs
    Private Sub LoadformingJobs()
        Try
            Dim data = FormingDBHelper.GetFormingJobs()
            FormingJobList = New ObservableCollection(Of FormingModel)(data)
            FormDG.ItemsSource = FormingJobList
        Catch ex As Exception
            MessageBox.Show("Error loading project list: " & ex.Message)
        End Try
    End Sub


End Class
