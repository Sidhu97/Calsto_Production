Imports System.Collections.ObjectModel


Public Class Planning

    Inherits Page

    Private ProjectList As ObservableCollection(Of PlanningModel)

    Public Sub New()
        InitializeComponent()
        LoadProjects()
    End Sub

    Private Sub LoadProjects()
        Try
            Dim data = PlanningDBHelper.GetProjectNO()
            ProjectList = New ObservableCollection(Of PlanningModel)(data)
            ProjNoDG.ItemsSource = ProjectList
        Catch ex As Exception
            MessageBox.Show("Error loading project list: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadBOMitems(proNo As String)
        Try
            Dim data = PlanningDBHelper.GetBOMItems(proNo)
            WIDBOMDG.ItemsSource = New ObservableCollection(Of PlanningSideDgModel)(data)
        Catch ex As Exception
            MessageBox.Show("Error loading project side details: " & ex.Message)
        End Try
    End Sub


    Private Sub LoadJobcrads(proNo As String)
        Try
            Dim data = PlanningDBHelper.GetJobcards(proNo)
            JobcardDG.ItemsSource = New ObservableCollection(Of JobCardModel)(data)
        Catch ex As Exception
            MessageBox.Show("Error loading project side details: " & ex.Message)
        End Try
    End Sub

    Private Sub dgProj_Sel_Changed(sender As Object, e As SelectionChangedEventArgs)
        Dim selectedProj As PlanningModel = TryCast(ProjNoDG.SelectedItem, PlanningModel)
        If selectedProj IsNot Nothing Then
            LoadBOMitems(selectedProj.PROJECTNO)
            txt_Projectno.Text = selectedProj.PROJECTNO
            txt_Customer.Text = selectedProj.Customer
        End If
    End Sub


    Private Sub dgWId_Sel_Changed(sender As Object, e As SelectionChangedEventArgs)
        Dim selectedWid As PlanningSideDgModel = TryCast(WIDBOMDG.SelectedItem, PlanningSideDgModel)
        If selectedWid IsNot Nothing Then
            LoadJobcrads(selectedWid.WID)
        End If
    End Sub




End Class
