Imports System.Collections.ObjectModel

Public Class Forming
    Inherits Page

    Private FormingJobList As ObservableCollection(Of FormingModel)
    Private ProjectList As ObservableCollection(Of PlanningModel)

    Public Sub New()
        InitializeComponent()
        LoadProjects()
    End Sub


    Private Sub LoadProjects()
        Try
            Dim data = PlanningDBHelper.GetProjectNO()
            ProjectList = New ObservableCollection(Of PlanningModel)(data)
            ProjDG.ItemsSource = ProjectList
        Catch ex As Exception
            MessageBox.Show("Error loading project list: " & ex.Message)
        End Try
    End Sub

    ' Load all forming jobs
    Private Sub LoadformingJobs(proNo As String)
        Try
            Dim data = FormingDBHelper.GetFormingJobs(proNo)
            FormingJobList = New ObservableCollection(Of FormingModel)(data)
            FormDG.ItemsSource = FormingJobList
        Catch ex As Exception
            MessageBox.Show("Error loading project list: " & ex.Message)
        End Try
    End Sub


    Private Sub FormProj_Sel_Changed(sender As Object, e As SelectionChangedEventArgs)
        Dim selectedProj As PlanningModel = TryCast(ProjDG.SelectedItem, PlanningModel)
        If selectedProj IsNot Nothing Then
            LoadformingJobs(selectedProj.PROJECTNO)
            txt_Projectno.Text = selectedProj.PROJECTNO
            txt_Customer.Text = selectedProj.Customer
        End If
    End Sub





End Class
