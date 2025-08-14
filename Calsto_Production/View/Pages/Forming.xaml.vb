Imports System.Collections.ObjectModel

Public Class Forming
    Inherits Page

    Private FormingJobList As ObservableCollection(Of FormingModel)
    Private LotList As ObservableCollection(Of FormingLotModel)
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


    Private Sub LoadLotIDs(jobc As String)
        Try
            Dim data = FormingDBHelper.GetFormingLots(jobc)
            LotList = New ObservableCollection(Of FormingLotModel)(data)
            JobcardDG.ItemsSource = LotList
        Catch ex As Exception
            MessageBox.Show("Error loading project side details: " & ex.Message)
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


    Private Sub dgFORM_Sel_Changed(sender As Object, e As SelectionChangedEventArgs)
        Dim selectedJC As FormingModel = TryCast(FormDG.SelectedItem, FormingModel)
        If selectedJC IsNot Nothing Then
            LoadLotIDs(selectedJC.JC_no)
            Txt_ReqQty.Text = selectedJC.BOMQty
            Txt_BalQty.Text = selectedJC.BalanceQty
            Txt_Colour.Text = selectedJC.Colour
            Txt_Description.Text = selectedJC.Description
        End If
    End Sub


    Private Sub LotEntry_Click(sender As Object, e As RoutedEventArgs)
        If FormDG.SelectedItem Is Nothing Then
            MessageBox.Show("Please select a job to apply the lot entry.")
            Exit Sub
        End If

        Dim selectedJC As FormingModel = CType(FormDG.SelectedItem, FormingModel)

        If MessageBox.Show($"Apply the Qty for this item: {selectedJC.JC_no}?",
                           "Confirm Lot Entry",
                           MessageBoxButton.YesNo,
                           MessageBoxImage.Question) = MessageBoxResult.No Then
            Exit Sub
        End If


        Dim JCQty As String = Lot_qty.Text
        Dim createdBy As String = Environment.UserName
        Dim Remarks As String = ""

        ' Call your method properly with actual values
        FormingDBHelper.CreateJobEntry(selectedJC.JC_no,
                                       JCQty,
                                       createdBy,
                                       Remarks) ' replace with the correct property

        ' Clear selection
        FormDG.SelectedItem = Nothing

        If selectedJC IsNot Nothing Then
            LoadLotIDs(selectedJC.JC_no)
            Txt_ReqQty.Text = selectedJC.BOMQty
            Txt_BalQty.Text = selectedJC.BalanceQty
            Txt_Colour.Text = selectedJC.Colour
            Txt_Description.Text = selectedJC.Description
        End If


        Dim selectedProj As PlanningModel = TryCast(ProjDG.SelectedItem, PlanningModel)
        If selectedProj IsNot Nothing Then
            LoadformingJobs(selectedProj.PROJECTNO)

        End If

        MessageBox.Show($"Lot entry applied for Job Card {selectedJC.JC_no} successfully.")
    End Sub









End Class
