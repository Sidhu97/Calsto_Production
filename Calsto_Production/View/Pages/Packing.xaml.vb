Imports System.Collections.ObjectModel
Imports Calsto_Production.PackingModel
Imports DocumentFormat.OpenXml.Vml.Office

Class Packing

    Private PackingJobList As ObservableCollection(Of PackingModel)
    Private PackingBundleList As ObservableCollection(Of PackingHeaderModel)
    Private ProjectList As ObservableCollection(Of PackingProjModel)


    Public Sub New()
        InitializeComponent()
        LoadProjects()
    End Sub


    Private Sub LoadPacklist()
        Dim Packlist = PackingDBHelper.GetPacktypeList()
        CMB_Packtype.ItemsSource = Packlist
        CMB_Packtype.DisplayMemberPath = "Packtype"
        CMB_Packtype.SelectedValuePath = "Packtype"
    End Sub


    Private Sub LoadProjects()
        Dim projects = PackingDBHelper.GetProjectList()
        CMB_PackProj.ItemsSource = projects
        CMB_PackProj.DisplayMemberPath = "Proj_no"
        CMB_PackProj.SelectedValuePath = "Proj_no"
    End Sub


    Private Sub LoadpackingJobs(proNo As String)
        Try
            Dim data = PackingDBHelper.GetPackingJobs(proNo)
            PackingJobList = New ObservableCollection(Of PackingModel)(data)
            ProjDG.ItemsSource = PackingJobList
        Catch ex As Exception
            MessageBox.Show("Error loading project list: " & ex.Message)
        End Try
    End Sub


    Private Sub LoadBundles(proNo As String)
        Try
            Dim data = PackingDBHelper.GetBundleList(proNo)
            PackingBundleList = New ObservableCollection(Of PackingHeaderModel)(data)
            BundleDG.ItemsSource = PackingBundleList
        Catch ex As Exception
            MessageBox.Show("Error loading project list: " & ex.Message)
        End Try
    End Sub


    Private Sub Applybtn_Click(sender As Object, e As RoutedEventArgs)


        Dim selectedProj = TryCast(CMB_PackProj.SelectedItem, PackingProjModel)
        If selectedProj IsNot Nothing Then
            Txt_customer.Text = selectedProj.Customer
            LoadpackingJobs(selectedProj.Proj_no)
            LoadBundles(selectedProj.Proj_no)
        End If
        ProjDG.IsEnabled = True
        Btn_Projapply.IsEnabled = False
        PROJPANEL.IsEnabled = True
        CMB_PackProj.IsEnabled = False
        Btn_ProjClear.IsEnabled = True
        LoadPacklist()
    End Sub


    Private Sub Clearbtn_Click(sender As Object, e As RoutedEventArgs)

        Txt_customer.Text = ""
        CMB_Packtype.SelectedItem = ""
        ProjDG.ItemsSource = Nothing
        BundleDG.ItemsSource = Nothing
        ProjDG.IsEnabled = False
        Btn_Projapply.IsEnabled = True
        PROJPANEL.IsEnabled = False
        CMB_PackProj.IsEnabled = True
        Btn_ProjClear.IsEnabled = False
    End Sub

End Class
