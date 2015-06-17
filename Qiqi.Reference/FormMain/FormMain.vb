Public Class FormMain

    ' Main menu
    Dim MenuStrip As _FormMain.MenuStrip
    ' Tool bar
    Dim ToolStrip As _FormMain.PrimaryToolStrip
    ' Status bar
    Dim StatusStrip As _FormMain.StatusStrip
    ' Paper tab control
    Dim DataBaseTabControl As _FormMain.DataBaseTabControl
    ' Configuration
    Dim Configuration As _FormConfiguration.Configuration

    ' Invoke
    Private Delegate Sub DelegateProgressUpdate(ByVal Progress As Double)
    Private Delegate Sub DelegateAddTabPage(ByVal TabPage As _FormMain.DataBaseTabPage)
    Private Delegate Sub DelegateRemoveTabPage(ByVal TabPage As _FormMain.DataBaseTabPage)
    Private Delegate Sub DelegateDataBaseLoading()

    Public Sub New()
        ' Initialize component
        InitializeComponent()

        ' Initialize configuration
        InitializeConfiguration()

        ' Initialize paper tab control
        InitialPaperTabControl()

        ' Initialize tool bar
        InitializeToolStrip()

        ' Initialize main menu
        InitializeMenuStrip()

        ' Initialize status bar
        InitializeStatusStrip()

        ' Initialize current database
        InitializeCurrentDataBase()
    End Sub

    Private Sub InitializeCurrentDataBase()
        Dim DataTable As New DataTable
        If Configuration.GetConfig(TableName.CurrentOpenFileList, DataTable) Then
            For Each Row As DataRow In DataTable.Rows
                If Not My.Computer.FileSystem.FileExists(Row("FileName")) Then
                    Continue For
                End If

                If DataBaseTabControl.Exist(Row("FileName")) Then
                    Continue For
                End If

                Dim NewTabPage As New _FormMain.DataBaseTabPage(Row("FileName"), Configuration)
                AddHandler NewTabPage.ProgressUpdate, AddressOf TabPage_ProgressUpdate
                AddHandler NewTabPage.ColumnDisplayChanged, AddressOf TabPage_ColumnDisplayChanged
                AddHandler NewTabPage.SplitterMoved, AddressOf TabPage_SplitterMoved
                DataBaseTabControl.TabPages.Add(NewTabPage)

                Dim Selected As Boolean = Row("Selected")
                If Selected Then
                    DataBaseTabControl.SelectedTab = NewTabPage
                End If
            Next

            Dim OpenDataBase As New Threading.Thread(AddressOf ThreadOpenDataBase)
            OpenDataBase.Start()
        End If
    End Sub

    Private Sub InitializeConfiguration()
        Configuration = New _FormConfiguration.Configuration
        If Not Configuration.Load(Application.StartupPath & "\Config.xml") Then
            MsgBox("There is an error when load configuration.", MsgBoxStyle.OkOnly, "Configuration Error")
        End If
    End Sub

    Private Sub CloseDataBase() ' Handles MenuStrip.MenuFileCloseDataBaseClick
        If Not DataBaseTabControl.SelectedTab Is Nothing Then
            DataBaseTabControl.TabPages.Remove(DataBaseTabControl.SelectedTab)
        End If
    End Sub

    Private Sub OpenDataBase() ' Handles MenuStrip.MenuFileOpenDataBase_Click
        Dim OpenFileDialog As New OpenFileDialog
        With OpenFileDialog
            .Filter = "BibTeX File|*.bib"
            .FileName = ""
            .Multiselect = True
        End With

        If Not OpenFileDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            Exit Sub
        End If

        For Each FileName As String In OpenFileDialog.FileNames
            If DataBaseTabControl.Exist(FileName) Then
                Continue For
            End If

            Dim NewTabPage As New _FormMain.DataBaseTabPage(FileName, Configuration)
            AddHandler NewTabPage.ProgressUpdate, AddressOf TabPage_ProgressUpdate
            AddHandler NewTabPage.ColumnDisplayChanged, AddressOf TabPage_ColumnDisplayChanged
            AddHandler NewTabPage.SplitterMoved, AddressOf TabPage_SplitterMoved
            DataBaseTabControl.TabPages.Add(NewTabPage)
            DataBaseTabControl.SelectedTab = NewTabPage
        Next

        Dim OpenDataBase As New Threading.Thread(AddressOf ThreadOpenDataBase)
        OpenDataBase.Start()

        ' Add this file into the history list
        Dim DataTable As New DataTable

        If Not Configuration.GetConfig(TableName.OpenFileHistoryList, DataTable) Then
            DataTable.TableName = TableName.OpenFileHistoryList
            DataTable.Columns.Add("FileFullName")
        End If

        For Each FileName As String In OpenFileDialog.FileNames
            For Each Row As DataRow In DataTable.Rows
                If CType(Row.Item(0), String).Trim.ToLower = FileName.Trim.ToLower Then
                    Row.Delete()
                    Exit For
                End If
            Next

            If DataTable.Rows.Count > 10 Then
                DataTable.Rows(0).Delete()
            End If
            DataTable.Rows.Add(FileName)
        Next

        Configuration.SetConfig(DataTable)
    End Sub

    Private Sub OpenRecentDataBase(ByVal RecentDataBaseFullName As String)
        If Not My.Computer.FileSystem.FileExists(RecentDataBaseFullName) Then
            MsgBox("There is no database" & vbCr & vbCr & RecentDataBaseFullName, MsgBoxStyle.OkOnly, "Error")
            Exit Sub
        End If

        If DataBaseTabControl.Exist(RecentDataBaseFullName) Then
            Exit Sub
        End If

        Dim NewTabPage As New _FormMain.DataBaseTabPage(RecentDataBaseFullName, Configuration)
        AddHandler NewTabPage.ProgressUpdate, AddressOf TabPage_ProgressUpdate
        DataBaseTabControl.TabPages.Add(NewTabPage)


        Dim OpenDataBase As New Threading.Thread(AddressOf ThreadOpenDataBase)
        OpenDataBase.Start()

        ' Add this file into the history list
        Dim DataTable As New DataTable

        If Not Configuration.GetConfig(TableName.OpenFileHistoryList, DataTable) Then
            DataTable.TableName = TableName.OpenFileHistoryList
            DataTable.Columns.Add("FileFullName")
        End If


        For Each Row As DataRow In DataTable.Rows
            If CType(Row.Item(0), String).Trim.ToLower = RecentDataBaseFullName.Trim.ToLower Then
                Row.Delete()
                Exit For
            End If
        Next

        If DataTable.Rows.Count > 10 Then
            DataTable.Rows(0).Delete()
        End If
        DataTable.Rows.Add(RecentDataBaseFullName)

        Configuration.SetConfig(DataTable)
    End Sub

    Private Sub ExitProgram() ' Handles MenuStrip.MenuFile_Exit_Click
        Me.Dispose()
        End
    End Sub

    Public Sub ThreadOpenDataBase()
        For Each TabPage As _FormMain.DataBaseTabPage In DataBaseTabControl.TabPages
            If Not TabPage.Loaded Then
                Dim ErrorMessage As New _BibTeX.ErrorMessage
                If Not TabPage.DataBaseLoading(ErrorMessage) Then
                    Me.Invoke(New DelegateRemoveTabPage(AddressOf RemoveTabPage), TabPage)
                    StatusStrip.ShowErrorMessage(TabPage.Name & " : " & ErrorMessage.GetErrorMessage)

                    MsgBox("There is an error in file """ & TabPage.Name & """ line : " & ErrorMessage.LineNumber & vbCr & _
                           "Error message is """ & ErrorMessage.GetErrorMessage & """" & vbCr & vbCr & _
                           "Click OK to continue.", _
                           MsgBoxStyle.OkOnly, "Load Error")
                End If
            End If
        Next
    End Sub

    Private Sub ShowFormConfiguration()

        For Each TabPage As _FormMain.DataBaseTabPage In DataBaseTabControl.TabPages
            TabPage.SetDataBaseGridViewEventEnable(False)
        Next

        Configuration.Reload()
        Dim FormConfiguration As New FormConfiguration(Configuration)
        If FormConfiguration.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Configuration.Reload()
            DataBaseTabPagesRefresh()
        End If

        For Each TabPage As _FormMain.DataBaseTabPage In DataBaseTabControl.TabPages
            TabPage.SetDataBaseGridViewEventEnable(True)
        Next
    End Sub

    Private Sub DataBaseTabPagesRefresh()
        For Each TabPage As _FormMain.DataBaseTabPage In DataBaseTabControl.TabPages
            TabPage.DataBaseGridViewRefresh(Configuration)
            TabPage.LiteratureTabControlRefresh(Configuration)
        Next
    End Sub

    Private Sub TabPage_ColumnDisplayChanged(ByVal sender As Object)
        Configuration.Reload()
   
        For Each TabPage As _FormMain.DataBaseTabPage In DataBaseTabControl.TabPages
            TabPage.SetDataBaseGridViewEventEnable(False)
            TabPage.DataBaseGridViewRefresh(Configuration)
            TabPage.SetDataBaseGridViewEventEnable(True)
        Next

    End Sub

    Private Sub TabPage_ProgressUpdate(ByVal Progress As Double)
        Me.Invoke(New DelegateProgressUpdate(AddressOf ProgressUpdate), Progress)
    End Sub

    Private Sub ProgressUpdate(ByVal Progress As Double)
        If Progress < 1 Then
            StatusStrip.DataBaseLoading(Progress)
        ElseIf Progress = 1 Then
            StatusStrip.DataBaseLoaded()
        Else
        End If
    End Sub

    Private Sub AddTabPage(ByVal TabPage As _FormMain.DataBaseTabPage)
        DataBaseTabControl.TabPages.Add(TabPage)
    End Sub

    Private Sub RemoveTabPage(ByVal TabPage As _FormMain.DataBaseTabPage)
        DataBaseTabControl.TabPages.Remove(TabPage)
    End Sub

    ''' <summary>
    ''' Initialize main menu
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitializeMenuStrip()
        MenuStrip = New _FormMain.MenuStrip(Configuration)
        Me.Controls.Add(MenuStrip)

        AddHandler MenuStrip.MenuFileOpenDataBaseClick, AddressOf OpenDataBase
        AddHandler MenuStrip.MenuFileExitClick, AddressOf ExitProgram
        AddHandler MenuStrip.MenuFileRecentDataBaseItemClick, AddressOf OpenRecentDataBase
        AddHandler MenuStrip.MenuFileCloseDataBaseClick, AddressOf CloseDataBase
        AddHandler MenuStrip.MenuOptionPreferencesClick, AddressOf ShowFormConfiguration
    End Sub

    ''' <summary>
    ''' Initialize tool bar 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitializeToolStrip()
        ToolStrip = New _FormMain.PrimaryToolStrip
        Me.Controls.Add(ToolStrip)
    End Sub

    ''' <summary>
    ''' Initialize status bar
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitializeStatusStrip()
        StatusStrip = New _FormMain.StatusStrip
        Me.Controls.Add(StatusStrip)
    End Sub

    ''' <summary>
    ''' Initialize paper tab control
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitialPaperTabControl()
        DataBaseTabControl = New _FormMain.DataBaseTabControl
        Me.Controls.Add(DataBaseTabControl)
        AddHandler DataBaseTabControl.TabPageChanged, AddressOf DataBaseTabControl_TabPageChanged
    End Sub

    ''' <summary>
    ''' If the TabPage of DataBaseTabControl has changed, save the current open file into the configuration
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <remarks></remarks>
    Private Sub DataBaseTabControl_TabPageChanged(ByVal sender As Object)
        MenuStrip.MenuFileCloseDataBase.Enabled = Not (DataBaseTabControl.TabPages.Count = 0)

        Dim DataTable As New DataTable
        DataTable.TableName = TableName.CurrentOpenFileList
        DataTable.Columns.Add("FileName")
        DataTable.Columns.Add("Selected")
        For Each TabPage As _FormMain.DataBaseTabPage In DataBaseTabControl.TabPages
            Dim Row(1) As String
            Row(0) = TabPage.Name
            If DataBaseTabControl.SelectedTab Is Nothing Then
                Row(1) = False
            Else
                Row(1) = DataBaseTabControl.SelectedTab.Equals(TabPage)
            End If
            DataTable.Rows.Add(Row)
        Next
        Configuration.SetConfig(DataTable)
        Configuration.Save()
    End Sub

    Private Sub TabPage_SplitterMoved(ByVal sender As Object, ByVal e As System.Windows.Forms.SplitterEventArgs)
        'For Each TabPage As _FormMain.DataBaseTabPage In DataBaseTabControl.TabPages
        '    TabPage.SetSplitContainerEventEnable(False)

        '    With CType(sender, SplitContainer)
        '        Select Case .Name
        '            Case TabPage.SplitContainerPrimary.Name
        '                TabPage.SplitContainerPrimary.SplitterDistance = .SplitterDistance

        '            Case TabPage.SplitContainerSecondary.Name
        '                TabPage.SplitContainerSecondary.SplitterDistance = .SplitterDistance
        '                TabPage.SplitContainerSecondary.Panel2Collapsed = .Panel2Collapsed
        '        End Select
        '    End With

        '    TabPage.SetSplitContainerEventEnable(True)
        'Next
    End Sub



End Class
