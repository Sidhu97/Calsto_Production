Imports System.Collections.ObjectModel
Imports Microsoft.Data.SqlClient
Imports System.Data


Public Class Planning

    Inherits Page

    Private ProjectList As ObservableCollection(Of PlanningModel)
    Private BOMList As ObservableCollection(Of PlanningSideDgModel)
    Private JobCardList As ObservableCollection(Of JobCardModel)

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
            BOMList = New ObservableCollection(Of PlanningSideDgModel)(data)
            WIDBOMDG.ItemsSource = BOMList
        Catch ex As Exception
            MessageBox.Show("Error loading project side details: " & ex.Message)
        End Try
    End Sub



    Private Sub LoadJobcrads(proNo As String)
        Try
            Dim data = PlanningDBHelper.GetJobCards(proNo)
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
            SelectAllFilteredCheckBox.IsChecked = False
        End If
    End Sub


    Private Sub dgWId_Sel_Changed(sender As Object, e As SelectionChangedEventArgs)
        Dim selectedWid As PlanningSideDgModel = TryCast(WIDBOMDG.SelectedItem, PlanningSideDgModel)
        If selectedWid IsNot Nothing Then
            LoadJobcrads(selectedWid.WID)
            Process_combo.Text = selectedWid.Process


        End If
    End Sub


    Private Sub BtnSelectAll_Click(sender As Object, e As RoutedEventArgs)
        For Each item In WIDBOMDG.Items
            Dim model = TryCast(item, PlanningSideDgModel)
            If model IsNot Nothing AndAlso model.Jobcard <> "Y" Then
                model.IsSelected = True
            End If
        Next
        WIDBOMDG.Items.Refresh()
    End Sub


    Private Sub BtnClearSelection_Click(sender As Object, e As RoutedEventArgs)
        For Each item In WIDBOMDG.Items
            Dim model = TryCast(item, PlanningSideDgModel)
            If model IsNot Nothing Then
                model.IsSelected = False
            End If
        Next
        WIDBOMDG.Items.Refresh()
    End Sub

    Private Async Sub CreateJobButton_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedItems = WIDBOMDG.Items.Cast(Of PlanningSideDgModel)().
                        Where(Function(x) x.IsSelected).ToList()

        If Not selectedItems.Any() Then
            MessageBox.Show("Please select at least one item to create a job card.")
            Exit Sub
        End If

        Dim createdBy As String = Environment.UserName

        Dim progressWindow As New JobProgress()
        progressWindow.Show()


        Dim errorMessage As String = Nothing
        Try
            For i As Integer = 0 To selectedItems.Count - 1
                Dim item = selectedItems(i)

                Await Task.Run(Sub()
                                   PlanningDBHelper.CreateJobCard(item.WID, createdBy)
                               End Sub)

                item.IsSelected = False

                Dim percent = CInt(((i + 1) / selectedItems.Count) * 100)
                progressWindow.Dispatcher.Invoke(Sub()
                                                     progressWindow.LoadingBar.Value = percent
                                                 End Sub)
            Next

            progressWindow.Dispatcher.Invoke(Sub()
                                                 progressWindow.StatusText.Text = $"{selectedItems.Count} job cards created successfully."
                                                 progressWindow.LoadingBar.Value = 100
                                             End Sub)

        Catch ex As Exception
            errorMessage = ex.Message
        Finally

        End Try

        If errorMessage IsNot Nothing Then
            progressWindow.Dispatcher.Invoke(Sub()
                                                 progressWindow.StatusText.Text = "Error: " & errorMessage
                                             End Sub)
            Await Task.Delay(2500)
        Else
            Await Task.Delay(1500)
        End If

        progressWindow.Close()
        SelectAllFilteredCheckBox.IsChecked = False
        RefreshBOM()
        Refreshproj()



    End Sub




    Private Sub DeleteJobButton_Click(sender As Object, e As RoutedEventArgs)
        If WIDBOMDG.SelectedItem Is Nothing Then
            MessageBox.Show("Please select a job to delete.")
            Exit Sub
        End If

        Dim selectedRow = CType(WIDBOMDG.SelectedItem, PlanningSideDgModel)

        If MessageBox.Show($"Delete job card for WID {selectedRow.WID}?",
                       "Confirm Delete",
                       MessageBoxButton.YesNo,
                       MessageBoxImage.Warning) = MessageBoxResult.No Then
            Exit Sub
        End If

        ' Call your shared delete method
        PlanningDBHelper.DeleteJobCard(selectedRow.WID)

        ' Clear selection
        'WIDBOMDG.SelectedItem = Nothing
        JobcardDG.ItemsSource = Nothing

        MessageBox.Show($"Job card for WID {selectedRow.WID} deleted successfully.")

        RefreshBOM()
        Refreshproj()

    End Sub



    Private Sub Refreshproj()
        Try
            ' ==== Refresh Projects ====
            ProjNoDG.SavePreset()

            Dim selectedProjNo As String = Nothing
            If ProjNoDG.SelectedItem IsNot Nothing Then
                selectedProjNo = CType(ProjNoDG.SelectedItem, PlanningModel).PROJECTNO
            End If

            Dim projData = PlanningDBHelper.GetProjectNO()
            ProjectList = New ObservableCollection(Of PlanningModel)(projData)
            ProjNoDG.ItemsSource = ProjectList

            ProjNoDG.LoadPreset()

            ' Restore project selection
            Dim currentProjNo As String = Nothing
            If selectedProjNo IsNot Nothing Then
                Dim projRow = ProjectList.FirstOrDefault(Function(x) x.PROJECTNO = selectedProjNo)
                If projRow IsNot Nothing Then
                    ProjNoDG.SelectedItem = projRow
                    ProjNoDG.UpdateLayout()
                    ProjNoDG.ScrollIntoView(projRow)
                    currentProjNo = projRow.PROJECTNO

                End If


            End If

        Catch ex As Exception
            MessageBox.Show("Error refreshing: " & ex.Message)
        End Try
    End Sub

    Private Sub RefreshBOM()
        Try

            WIDBOMDG.SavePreset()
            ' 1. Save selected item key (if any)
            Dim selectedWID As String = Nothing
            If WIDBOMDG.SelectedItem IsNot Nothing Then
                selectedWID = CType(WIDBOMDG.SelectedItem, PlanningSideDgModel).WID
            End If

            ' 2. Get current project no. from Project grid selection
            Dim selectedProj = CType(ProjNoDG.SelectedItem, PlanningModel)
            If selectedProj Is Nothing Then Exit Sub

            Dim proNo = selectedProj.PROJECTNO



            ' 3. Reload BOM items
            Dim data = PlanningDBHelper.GetBOMItems(proNo)
            BOMList = New ObservableCollection(Of PlanningSideDgModel)(data)
            WIDBOMDG.ItemsSource = BOMList
            WIDBOMDG.LoadPreset()
            ' 4. Restore selection
            If selectedWID IsNot Nothing Then
                Dim rowToSelect = BOMList.FirstOrDefault(Function(x) x.WID = selectedWID)
                If rowToSelect IsNot Nothing Then
                    WIDBOMDG.SelectedItem = rowToSelect
                    WIDBOMDG.UpdateLayout()
                    WIDBOMDG.ScrollIntoView(rowToSelect)
                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Error refreshing BOM items: " & ex.Message)
        End Try
    End Sub




End Class
