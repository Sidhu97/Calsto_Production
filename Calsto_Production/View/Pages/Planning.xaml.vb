Imports System.Collections.ObjectModel
Imports System.Threading.Tasks
Imports ClosedXML.Excel
Imports TACKERCALSTO.PlanningModel

Class Planning
    Private PlanningList As ObservableCollection(Of PlanningModel)
    Private BOMList As ObservableCollection(Of PlanningModel.PlanningSideDgModel)
    Private JobCards As ObservableCollection(Of JobCardModel)

    Private Sub Page_Loaded(sender As Object, e As RoutedEventArgs)
        LoadProjectsAsync()
    End Sub

    Private Async Sub LoadProjectsAsync()
        Dim data = Await Task.Run(Function() dash_Plan_Dbhelper.GetWids())
        PlanningList = New ObservableCollection(Of PlanningModel)(data)
        MasterDG.ItemsSource = PlanningList
    End Sub

    Private Async Sub MasterDG_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        Dim selected = TryCast(MasterDG.SelectedItem, PlanningModel)
        If selected IsNot Nothing Then
            Dim data = Await Task.Run(Function() dash_Plan_Dbhelper.GetDetailsByBOM(selected.PROJECTNO))
            BOMList = New ObservableCollection(Of PlanningModel.PlanningSideDgModel)(data)
            WIDBOMDG.ItemsSource = BOMList
        End If
    End Sub

    Private Async Sub WIDBOMDG_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        Dim selected = TryCast(WIDBOMDG.SelectedItem, PlanningModel.PlanningSideDgModel)
        If selected IsNot Nothing Then
            Dim jobCards = Await Task.Run(Function() dash_Plan_Dbhelper.GetJobCardsByWID(selected.WID))
            jobCards = New ObservableCollection(Of JobCardModel)(jobCards)
            JobcardDG.ItemsSource = jobCards
        End If
    End Sub

    Private Async Sub CreateJobCard_Click(sender As Object, e As RoutedEventArgs)
        Dim selected = TryCast(WIDBOMDG.SelectedItem, PlanningModel.PlanningSideDgModel)
        If selected Is Nothing Then
            MessageBox.Show("Select a BOM item")
            Return
        End If

        If jobCards IsNot Nothing AndAlso jobCards.Any(Function(x) x.WID = selected.WID) Then
            MessageBox.Show("Job card already created.")
            Return
        End If

        Await Task.Run(Sub() dash_Plan_Dbhelper.CreateJobCard(selected.WID, Environment.UserName))

        ' Refresh Job Card
        Dim jobCards = Await Task.Run(Function() dash_Plan_Dbhelper.GetJobCardsByWID(selected.WID))
        jobCards = New ObservableCollection(Of JobCardModel)(jobCards)
        JobcardDG.ItemsSource = jobCards
    End Sub

    Private Sub SelectAllFilteredCheckBox_Click(sender As Object, e As RoutedEventArgs)
        If BOMList Is Nothing Then Return
        For Each item In BOMList
            item.IsSelectedToCreate = SelectAllFilteredCheckBox.IsChecked = True
        Next
        WIDBOMDG.Items.Refresh()
    End Sub
End Class
