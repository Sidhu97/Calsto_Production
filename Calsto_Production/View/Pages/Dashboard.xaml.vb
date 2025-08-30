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
            dpDispatchDate.SelectedDate = selectedProject.PO_Dispatch_Date
        End If
    End Sub

    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedProject As DashboardModel = TryCast(DashDG.SelectedItem, DashboardModel)
        If selectedProject IsNot Nothing AndAlso dpDispatchDate.SelectedDate.HasValue Then
            Try
                dash_Plan_Dbhelper.UpdateProjectDispatchDate(selectedProject.BOM_No, dpDispatchDate.SelectedDate.Value)
                MessageBox.Show("Dispatch date updated successfully.")

                ' Refresh the project list
                LoadProjects()

                ' Keep the same project selected after refresh
                Dim refreshedProject = ProjectList.FirstOrDefault(Function(p) p.BOM_No = selectedProject.BOM_No)
                If refreshedProject IsNot Nothing Then
                    DashDG.SelectedItem = refreshedProject
                    DashDG.ScrollIntoView(refreshedProject)
                End If

            Catch ex As Exception
                MessageBox.Show("Error updating date: " & ex.Message)
            End Try
        Else
            MessageBox.Show("Please select a project and a date.")
        End If
    End Sub

End Class
