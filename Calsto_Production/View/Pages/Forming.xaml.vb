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
            Txt_CompQty.Text = selectedJC.ProducedQty

            If selectedJC.Status = "CLOSED" Then
                DetailsTab.IsEnabled = False
            Else
                DetailsTab.IsEnabled = True
            End If
        End If






    End Sub


    Private Sub LotEntry_Click(sender As Object, e As RoutedEventArgs)
        If FormDG.SelectedItem Is Nothing Then
            MessageBox.Show("Please select a job to apply the lot entry.", "Lot Entry", MessageBoxButton.OK, MessageBoxImage.Warning)
            Exit Sub
        End If

        Dim selectedJC As FormingModel = CType(FormDG.SelectedItem, FormingModel)

        ' Confirm
        If MessageBox.Show($"Apply the Qty for this item: {selectedJC.JC_no}?",
                           "Confirm Lot Entry",
                           MessageBoxButton.YesNo,
                           MessageBoxImage.Question) = MessageBoxResult.No Then
            Exit Sub
        End If

        ' Validate Qty
        Dim JCQty As Integer
        If Not Integer.TryParse(Lot_qty.Text, JCQty) OrElse JCQty <= 0 Then
            MessageBox.Show("Please enter a valid quantity.", "Invalid Entry", MessageBoxButton.OK, MessageBoxImage.Error)
            Exit Sub
        End If

        ' 🔹 Validate against Balance Qty
        Dim balanceQty As Integer
        If Not Integer.TryParse(Txt_BalQty.Text, balanceQty) Then
            balanceQty = 0
        End If

        If JCQty > balanceQty Then
            MessageBox.Show($"Lot entry quantity ({JCQty}) cannot be greater than Balance Qty ({balanceQty}).",
                            "Lot Entry Error",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error)
            Exit Sub
        End If

        Dim createdBy As String = Environment.UserName
        Dim Remarks As String = ""

        ' Save to DB
        FormingDBHelper.CreateJobEntry(selectedJC.JC_no, JCQty, createdBy, Remarks)

        ' 🔹 Remember JC_no before reload
        Dim jcNoToSelect As String = selectedJC.JC_no

        ' Refresh data (new values loaded from DB)
        Dim selectedProj As PlanningModel = TryCast(ProjDG.SelectedItem, PlanningModel)
        If selectedProj IsNot Nothing Then
            LoadformingJobs(selectedProj.PROJECTNO)
        End If

        ' 🔹 Find the refreshed row and reselect
        Dim refreshedList = CType(FormDG.ItemsSource, IEnumerable(Of FormingModel))
        Dim refreshedItem = refreshedList.FirstOrDefault(Function(x) x.JC_no = jcNoToSelect)

        If refreshedItem IsNot Nothing Then
            FormDG.SelectedItem = refreshedItem
            FormDG.ScrollIntoView(refreshedItem)

            ' Update details with **refreshed values**
            Txt_ReqQty.Text = refreshedItem.BOMQty.ToString()
            Txt_BalQty.Text = refreshedItem.BalanceQty.ToString()
            Txt_Colour.Text = refreshedItem.Colour
            Txt_Description.Text = refreshedItem.Description
            Txt_CompQty.Text = refreshedItem.ProducedQty.ToString()

            If refreshedItem.Status = "CLOSED" Then
                DetailsTab.IsEnabled = False
            Else
                DetailsTab.IsEnabled = True
            End If


        End If

        Lot_qty.Text = ""
        MessageBox.Show($"Lot entry applied for Job Card {jcNoToSelect} successfully.", "Lot Entry", MessageBoxButton.OK, MessageBoxImage.Information)
    End Sub









End Class
