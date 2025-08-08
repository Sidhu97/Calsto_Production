Imports System.Collections.ObjectModel


Public Class Dashboard

    Inherits Page

    Private ProjectList As ObservableCollection(Of DashboardModel)

    Public Sub New()
        InitializeComponent()
        LoadProjects()
    End Sub

    Private Sub LoadProjects()
        Try
            Dim data = dash_Plan_Dbhelper.GetProjectList()
            ProjectList = New ObservableCollection(Of DashboardModel)(data)
            DashDG.ItemsSource = ProjectList
        Catch ex As Exception
            MessageBox.Show("Error loading project list: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadProjectSideDetails(proNo As String)
        Try
            Dim data = dash_Plan_Dbhelper.GetProjectSideDetails(proNo)
            SideDG.ItemsSource = New ObservableCollection(Of SideDGModel)(data)
        Catch ex As Exception
            MessageBox.Show("Error loading project side details: " & ex.Message)
        End Try
    End Sub

    Private Sub dgProjects_Sel_Changed(sender As Object, e As SelectionChangedEventArgs)
        Dim selectedProject As DashboardModel = TryCast(DashDG.SelectedItem, DashboardModel)
        If selectedProject IsNot Nothing Then
            LoadProjectSideDetails(selectedProject.BOM_No) ' or .PRO_No
            dpDispatchDate.SelectedDate = selectedProject.Dispatch_Date
        End If
    End Sub


End Class
