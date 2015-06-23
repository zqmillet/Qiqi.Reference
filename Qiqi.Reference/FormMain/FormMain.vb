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
    Private Delegate Sub DelegateRecentDataBaseListRemove(ByVal DataBaseFullName As String)

    Public Sub New()
        ' Initialize component
        InitializeComponent()

        ' Initialize configuration
        InitializeConfiguration()

        ' Initialize paper tab control
        InitialPaperTabControl()

        ' Initialize tool bar
        InitializeToolStrip()

        ' Initialize status bar
        InitializeStatusStrip()

        ' Initialize main menu
        InitializeMenuStrip()

        ' Initialize current database
        InitializeCurrentDataBase()

        ' Initialize view of FormMain
        InitializeView()

    End Sub

    Private Sub InitializeView()
        Dim DataTable As New DataTable
        If Not Configuration.GetConfig(TableName.FormMainViewConfiguration, DataTable) Then
            Exit Sub
        End If

        For Each Row As DataRow In DataTable.Rows
            Select Case Row("Control")
                Case MenuStrip.MenuViewShowGroupTreeView.Name
                    For Each TabPage As _FormMain.DataBaseTabPage In DataBaseTabControl.TabPages
                        TabPage.SplitContainerPrimary.Panel1Collapsed = Not CType(Row("Parameter"), Boolean)
                    Next
                Case MenuStrip.MenuViewShowStatusStrip.Name
                    StatusStrip.Visible = Row("Parameter")
                Case MenuStrip.MenuViewShowToolStrip.Name
                    ToolStrip.Visible = Row("Parameter")
            End Select
        Next
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
                AddHandler NewTabPage.DataBaseGridViewDoubleClick, AddressOf NewTabPage_DataBaseGridViewDoubleClick
                DataBaseTabControl.TabPages.Add(NewTabPage)
                MenuStrip.RecentDataBaseOrderDelete(Row("FileName"))

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
            MenuStrip.RecentDataBaseOrderUpdate(DataBaseTabControl.SelectedTab.Name.ToString)
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
            AddHandler NewTabPage.DataBaseGridViewDoubleClick, AddressOf NewTabPage_DataBaseGridViewDoubleClick
            DataBaseTabControl.TabPages.Add(NewTabPage)
            DataBaseTabControl.SelectedTab = NewTabPage
            MenuStrip.RecentDataBaseOrderDelete(FileName)
        Next

        Dim OpenDataBase As New Threading.Thread(AddressOf ThreadOpenDataBase)
        OpenDataBase.Start()
    End Sub

    ''' <summary>
    ''' MenuFileRecentDataBaseItem is clicked
    ''' </summary>
    ''' <param name="RecentDataBaseFullName">The full name of recent database which is clicked</param>
    ''' <remarks></remarks>
    Private Sub OpenRecentDataBase(ByVal RecentDataBaseFullName As String)
        MenuStrip.RecentDataBaseOrderDelete(RecentDataBaseFullName)

        ' If the file do not exist, show the error message and exit sub
        If Not My.Computer.FileSystem.FileExists(RecentDataBaseFullName) Then
            MsgBox("There is no database" & vbCr & vbCr & RecentDataBaseFullName, MsgBoxStyle.OkOnly, "Error")
            MenuStrip.RecentDataBaseOrderDelete(RecentDataBaseFullName)
            Exit Sub
        End If

        ' If the file has been opened, exit sub
        If DataBaseTabControl.Exist(RecentDataBaseFullName) Then
            MenuStrip.RecentDataBaseOrderUpdate(RecentDataBaseFullName)
            For Each TabPage As _FormMain.DataBaseTabPage In DataBaseTabControl.TabPages
                If TabPage.Name.Trim.ToLower = RecentDataBaseFullName.Trim.ToLower Then
                    DataBaseTabControl.SelectedTab = TabPage
                    Exit For
                End If
            Next
            Exit Sub
        End If

        ' Create a new TabPage for the database
        Dim NewTabPage As New _FormMain.DataBaseTabPage(RecentDataBaseFullName, Configuration)
        AddHandler NewTabPage.ProgressUpdate, AddressOf TabPage_ProgressUpdate
        AddHandler NewTabPage.ColumnDisplayChanged, AddressOf TabPage_ColumnDisplayChanged
        AddHandler NewTabPage.SplitterMoved, AddressOf TabPage_SplitterMoved
        AddHandler NewTabPage.DataBaseGridViewDoubleClick, AddressOf NewTabPage_DataBaseGridViewDoubleClick

        DataBaseTabControl.TabPages.Add(NewTabPage)
        DataBaseTabControl.SelectedTab = NewTabPage

        ' Start a thread to load the database
        Dim OpenDataBase As New Threading.Thread(AddressOf ThreadOpenDataBase)
        OpenDataBase.Start()
    End Sub

    Private Sub NewTabPage_DataBaseGridViewDoubleClick(ByVal sender As Object)
        For Each TabPage As _FormMain.DataBaseTabPage In DataBaseTabControl.TabPages
            If CType(sender, _FormMain.DataBaseTabPage).Equals(TabPage) Then
                Continue For
            End If

            If Not TabPage.SplitContainerSecondary.Panel2Collapsed Then
                CType(sender, _FormMain.DataBaseTabPage).SplitContainerSecondary.SplitterDistance = _
                    TabPage.SplitContainerSecondary.SplitterDistance
                Exit Sub
            End If
        Next
    End Sub

    Private Sub ExitProgram() ' Handles MenuStrip.MenuFile_Exit_Click
        Me.Dispose()
        End
    End Sub

    Public Sub ThreadOpenDataBase()
        For Each TabPage As _FormMain.DataBaseTabPage In DataBaseTabControl.TabPages
            If TabPage.Loaded Then
                Continue For
            End If

            Dim ErrorMessage As New _BibTeX.ErrorMessage
            If Not TabPage.DataBaseLoading(ErrorMessage) Then
                Me.Invoke(New DelegateRemoveTabPage(AddressOf RemoveTabPage), TabPage)
                Me.Invoke(New DelegateRecentDataBaseListRemove(AddressOf RemoveDataBaseListItem), TabPage.Name)
                StatusStrip.ShowErrorMessage(TabPage.Name & " : " & ErrorMessage.GetErrorMessage)
                MsgBox("There is an error in file """ & TabPage.Name & """ line : " & ErrorMessage.LineNumber & vbCr & _
                       "Error message is """ & ErrorMessage.GetErrorMessage & """" & vbCr & vbCr & _
                       "Click OK to continue.", _
                       MsgBoxStyle.OkOnly, "Load Error")
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
        If Me.InvokeRequired Then
            Me.Invoke(New DelegateProgressUpdate(AddressOf ProgressUpdate), Progress)
        Else
            ProgressUpdate(Progress)
        End If


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

    Private Sub RemoveDataBaseListItem(ByVal DataBaseFullName As String)
        MenuStrip.RecentDataBaseOrderDelete(DataBaseFullName)
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
        AddHandler MenuStrip.MenuViewNextDataBaseClick, AddressOf ShowNextDataBase
        AddHandler MenuStrip.MenuViewPreviousDataBaseClick, AddressOf ShowPreviousDataBase
        AddHandler MenuStrip.MenuViewShowToolStripClick, AddressOf ShowToolStrip
        AddHandler MenuStrip.MenuViewShowStatusStripClick, AddressOf ShowStatusStrip
        AddHandler MenuStrip.MenuViewShowGroupTreeViewClick, AddressOf ShowGroupTreeView
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


        MenuStrip.SetNextPreviousDataBaseEnable(DataBaseTabControl.TabPages.Count > 1)

    End Sub

    Private Sub TabPage_SplitterMoved(ByVal sender As Object, ByVal e As System.Windows.Forms.SplitterEventArgs)
        If DataBaseTabControl.SelectedTab Is Nothing Then
            Exit Sub
        End If

        Dim Height As Integer = CType(DataBaseTabControl.SelectedTab, _FormMain.DataBaseTabPage).SplitContainerSecondary.SplitterDistance

        For Each TabPage As _FormMain.DataBaseTabPage In DataBaseTabControl.TabPages
            TabPage.SetSplitContainerEventEnable(False)
            If Not TabPage.SplitContainerSecondary.Panel2Collapsed Then
                TabPage.SplitContainerSecondary.SplitterDistance = Height
            End If
            TabPage.SetSplitContainerEventEnable(True)
        Next
    End Sub

    Private Sub ShowPreviousDataBase()
        If DataBaseTabControl.TabPages.Count = 0 Then
            Exit Sub
        End If

        Dim Index As Integer = DataBaseTabControl.SelectedIndex
        If Index > 0 Then
            Index -= 1
        Else
            Index = DataBaseTabControl.TabPages.Count - 1
        End If

        DataBaseTabControl.SelectedIndex = Index
    End Sub

    Private Sub ShowNextDataBase()
        If DataBaseTabControl.TabPages.Count = 0 Then
            Exit Sub
        End If

        Dim Index As Integer = DataBaseTabControl.SelectedIndex
        If Index < DataBaseTabControl.TabPages.Count - 1 Then
            Index += 1
        Else
            Index = 0
        End If

        DataBaseTabControl.SelectedIndex = Index
    End Sub

    Private Sub ShowToolStrip(ByVal Visible As Boolean)
        ToolStrip.Visible = Visible
        ViewConfigurationSave()
    End Sub

    Private Sub ShowStatusStrip(ByVal Visible As Boolean)
        StatusStrip.Visible = Visible
        ViewConfigurationSave()
    End Sub

    Private Sub ShowGroupTreeView(ByVal Visible As Boolean)
        For Each TabPage As _FormMain.DataBaseTabPage In DataBaseTabControl.TabPages
            TabPage.SplitContainerPrimary.Panel1Collapsed = Not Visible
        Next
        ViewConfigurationSave()
    End Sub

    Private Sub ViewConfigurationSave()
        Dim OldDataTable As New DataTable
        If Not Configuration.GetConfig(TableName.FormMainViewConfiguration, OldDataTable) Then
            Exit Sub
        End If

        Dim GroupTreeWidth As Integer = 0
        Dim DetailTabControlHeight As Integer = 0
        For Each Row As DataRow In OldDataTable.Rows
            Select Case Row("Control")
                Case "GroupTreeWidth"
                    GroupTreeWidth = Row("Parameter")
                Case "DetailTabControlHeight"
                    DetailTabControlHeight = Row("Parameter")
            End Select
        Next

        Dim NewDataTable As DataTable = MenuStrip.GetFormMainViewConfiguration

        NewDataTable.Rows.Add("GroupTreeWidth", GroupTreeWidth)
        NewDataTable.Rows.Add("DetailTabControlHeight", DetailTabControlHeight)
        NewDataTable.Rows.Add("FormMainHeight", Me.Size.Height)
        NewDataTable.Rows.Add("FormMainWidth", Me.Size.Width)
        NewDataTable.Rows.Add("FormMainWindowState", Me.WindowState)

        Configuration.SetConfig(NewDataTable)
        Configuration.Save()

        MsgBox(CType(DataBaseTabControl.SelectedTab, _FormMain.DataBaseTabPage).GroupTreeViewWidth)
    End Sub


End Class
