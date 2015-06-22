Namespace _FormMain
    Public Class MenuStrip
        Inherits System.Windows.Forms.MenuStrip

        Public MenuFile As ToolStripMenuItem
        Public MenuFileNewDataBase As ToolStripMenuItem
        Public MenuFileOpenDataBase As ToolStripMenuItem
        Public MenuFileAppendDataBase As ToolStripMenuItem
        Public MenuFileExit As ToolStripMenuItem
        Public MenuFileRecentDataBase As ToolStripMenuItem
        Public MenuFileCloseDataBase As ToolStripMenuItem

        Public MenuView As ToolStripMenuItem
        Public MenuViewNextDataBase As ToolStripMenuItem
        Public MenuViewPreviousDataBase As ToolStripMenuItem
        Public MenuViewShowGroupTreeView As ToolStripMenuItem
        Public MenuViewShowToolStrip As ToolStripMenuItem
        Public MenuViewShowStatusStrip As ToolStripMenuItem

        Public MenuOption As ToolStripMenuItem
        Public MenuOptionPreferences As ToolStripMenuItem

        Public Event MenuFileNewDataBaseClick()
        Public Event MenuFileOpenDataBaseClick()
        Public Event MenuFileCloseDataBaseClick()
        Public Event MenuFileExitClick()
        Public Event MenuFileRecentDataBaseItemClick(ByVal RecentDataBaseFullName As String)
        Public Event MenuOptionPreferencesClick()
        Public Event MenuViewNextDataBaseClick()
        Public Event MenuViewPreviousDataBaseClick()
        Public Event MenuViewShowGroupTreeViewClick(ByVal Visible As Boolean)
        Public Event MenuViewShowToolStripClick(ByVal Visible As Boolean)
        Public Event MenuViewShowStatusStripClick(ByVal Visible As Boolean)

        Dim Configuration As _FormConfiguration.Configuration

        Public Sub New(ByVal Configuration As _FormConfiguration.Configuration)
            With Me
                .Anchor = AnchorStyles.Left Or AnchorStyles.Top
                .Configuration = Configuration
            End With

            InitializeFileMenu()
        End Sub

        Private Sub InitializeFileMenu()
            MenuFileNewDataBase = New ToolStripMenuItem
            With MenuFileNewDataBase
                .Name = "MenuFileNewDataBase"
                .Text = "&New Database"
                .Image = Resource.Icon.MenuFileNewDataBase
                .ShortcutKeys = Keys.Control Or Keys.N
                AddHandler .Click, AddressOf ToolStripMenuItem_Click
            End With

            MenuFileOpenDataBase = New ToolStripMenuItem
            With MenuFileOpenDataBase
                .Name = "MenuFileOpenDataBase"
                .Text = "&Open Database"
                .Image = Resource.Icon.MenuFileOpenDataBase
                .ShortcutKeys = Keys.Control Or Keys.O
                AddHandler .Click, AddressOf ToolStripMenuItem_Click
            End With

            MenuFileCloseDataBase = New ToolStripMenuItem
            With MenuFileCloseDataBase
                .Name = "MenuFileCloseDataBase"
                .Text = "&Close Database"
                .Enabled = False
                .Image = Resource.Icon.MenuFileCloseDataBase
                .ShortcutKeys = Keys.Control Or Keys.W
                AddHandler .Click, AddressOf ToolStripMenuItem_Click
            End With

            MenuFileExit = New ToolStripMenuItem
            With MenuFileExit
                .Name = "MenuFileExit"
                .Text = "E&xit"
                .Image = Resource.Icon.MenuFileExit
                .ShortcutKeys = Keys.Alt Or Keys.X
                AddHandler .Click, AddressOf ToolStripMenuItem_Click
            End With

            MenuFileRecentDataBase = New ToolStripMenuItem
            With MenuFileRecentDataBase
                .Name = "MenuFileRecentDataBase"
                .Text = "&Recent Database"
                .Enabled = False

                Dim OpenFileHistoryListDataTable As New DataTable
                If Configuration.GetConfig(TableName.OpenFileHistoryList, OpenFileHistoryListDataTable) = True Then
                    .Enabled = True

                    For Index As Integer = OpenFileHistoryListDataTable.Rows.Count - 1 To 0 Step -1
                        Dim MenuFileRecentDataBaseItem As New ToolStripMenuItem

                        With MenuFileRecentDataBaseItem
                            .Name = "MenuFileRecentDataBaseItem"
                            .Tag = OpenFileHistoryListDataTable.Rows(Index).Item(0)
                            .Text = OpenFileHistoryListDataTable.Rows.Count - Index & ". " & OpenFileHistoryListDataTable.Rows(Index).Item(0)
                            AddHandler .Click, AddressOf ToolStripMenuItem_Click
                            MenuFileRecentDataBase.DropDownItems.Add(MenuFileRecentDataBaseItem)
                        End With
                    Next
                Else
                    .Enabled = False
                End If

            End With

            MenuFile = New ToolStripMenuItem
            With MenuFile
                .Text = "&File"
                .DropDownItems.Add(MenuFileNewDataBase)
                .DropDownItems.Add(MenuFileOpenDataBase)
                .DropDownItems.Add(New ToolStripSeparator)
                .DropDownItems.Add(MenuFileCloseDataBase)
                .DropDownItems.Add(New ToolStripSeparator)
                .DropDownItems.Add(MenuFileRecentDataBase)
                .DropDownItems.Add(New ToolStripSeparator)
                .DropDownItems.Add(MenuFileExit)
            End With


            MenuViewNextDataBase = New ToolStripMenuItem
            With MenuViewNextDataBase
                .Name = "MenuViewNextDataBase"
                .Text = "&Next Database"
                .Enabled = False
                .Image = Resource.Icon.MenuViewNextDataBase
                .ShortcutKeys = Keys.Control Or Keys.Right
                AddHandler .Click, AddressOf ToolStripMenuItem_Click
            End With

            MenuViewPreviousDataBase = New ToolStripMenuItem
            With MenuViewPreviousDataBase
                .Name = "MenuViewPreviousDataBase"
                .Text = "&Previous Database"
                .Enabled = False
                .Image = Resource.Icon.MenuViewPreviousDataBase
                .ShortcutKeys = Keys.Control Or Keys.Left
                AddHandler .Click, AddressOf ToolStripMenuItem_Click
            End With

            MenuViewShowGroupTreeView = New ToolStripMenuItem
            With MenuViewShowGroupTreeView
                .Name = "MenuViewShowGroupTreeView"
                .Text = "Show &Group"
                .ShortcutKeys = Keys.Control Or Keys.G
                AddHandler .Click, AddressOf ToolStripMenuItem_Click
            End With

            MenuViewShowToolStrip = New ToolStripMenuItem
            With MenuViewShowToolStrip
                .Name = "MenuViewShowToolStrip"
                .Text = "Show &Toolbar"
                .ShortcutKeys = Keys.Control Or Keys.T
                AddHandler .Click, AddressOf ToolStripMenuItem_Click
            End With

            MenuViewShowStatusStrip = New ToolStripMenuItem
            With MenuViewShowStatusStrip
                .Name = "MenuViewShowStatusStrip"
                .Text = "Show &Statusbar"
                .ShortcutKeys = Keys.Control Or Keys.Shift Or Keys.S
                AddHandler .Click, AddressOf ToolStripMenuItem_Click
            End With

            Dim FormMainViewConfiguration As New DataTable
            If Configuration.GetConfig(TableName.FormMainViewConfiguration, FormMainViewConfiguration) Then
                For Each Row As DataRow In FormMainViewConfiguration.Rows
                    Select Case Row("Control")
                        Case MenuViewShowGroupTreeView.Name
                            MenuViewShowGroupTreeView.Checked = Row("Parameter")
                        Case MenuViewShowStatusStrip.Name
                            MenuViewShowStatusStrip.Checked = Row("Parameter")
                        Case MenuViewShowToolStrip.Name
                            MenuViewShowToolStrip.Checked = Row("Parameter")
                    End Select
                Next
            Else
                MenuViewShowGroupTreeView.Checked = True
                MenuViewShowStatusStrip.Checked = True
                MenuViewShowToolStrip.Checked = True
            End If

            MenuView = New ToolStripMenuItem
            With MenuView
                .Text = "&View"
                .DropDownItems.Add(MenuViewNextDataBase)
                .DropDownItems.Add(MenuViewPreviousDataBase)
                .DropDownItems.Add(New ToolStripSeparator)
                .DropDownItems.Add(MenuViewShowGroupTreeView)
                .DropDownItems.Add(MenuViewShowToolStrip)
                .DropDownItems.Add(MenuViewShowStatusStrip)
            End With


            MenuOptionPreferences = New ToolStripMenuItem
            With MenuOptionPreferences
                .Name = "MenuOptionPreferences"
                .Text = "&Preferences"
                .Image = Resource.Icon.MenuOptionPreferences
                .ShortcutKeys = Keys.Control Or Keys.P
                AddHandler .Click, AddressOf ToolStripMenuItem_Click
            End With

            MenuOption = New ToolStripMenuItem
            With MenuOption
                .Text = "&Option"
                .DropDownItems.Add(MenuOptionPreferences)
            End With

            Me.Items.Clear()
            Me.Items.Add(MenuFile)
            Me.Items.Add(MenuView)
            Me.Items.Add(MenuOption)
        End Sub

        Private Sub ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Select Case CType(sender, ToolStripMenuItem).Name
                Case MenuFileNewDataBase.Name
                    RaiseEvent MenuFileNewDataBaseClick()
                Case MenuFileOpenDataBase.Name
                    RaiseEvent MenuFileOpenDataBaseClick()
                Case MenuFileCloseDataBase.Name
                    RaiseEvent MenuFileCloseDataBaseClick()
                Case MenuFileExit.Name
                    RaiseEvent MenuFileExitClick()
                Case "MenuFileRecentDataBaseItem"
                    ' RecentDataBaseOrderUpdate(CType(sender, ToolStripMenuItem))
                    RaiseEvent MenuFileRecentDataBaseItemClick(CType(sender, ToolStripMenuItem).Tag.ToString)
                Case MenuOptionPreferences.Name
                    RaiseEvent MenuOptionPreferencesClick()
                Case MenuViewNextDataBase.Name
                    RaiseEvent MenuViewNextDataBaseClick()
                Case MenuViewPreviousDataBase.Name
                    RaiseEvent MenuViewPreviousDataBaseClick()
                Case MenuViewShowGroupTreeView.Name
                    MenuViewShowGroupTreeView.Checked = Not MenuViewShowGroupTreeView.Checked
                    RaiseEvent MenuViewShowGroupTreeViewClick(MenuViewShowGroupTreeView.Checked)
                Case MenuViewShowToolStrip.Name
                    MenuViewShowToolStrip.Checked = Not MenuViewShowToolStrip.Checked
                    RaiseEvent MenuViewShowToolStripClick(MenuViewShowToolStrip.Checked)
                Case MenuViewShowStatusStrip.Name
                    MenuViewShowStatusStrip.Checked = Not MenuViewShowStatusStrip.Checked
                    RaiseEvent MenuViewShowStatusStripClick(MenuViewShowStatusStrip.Checked)
            End Select
        End Sub

        Public Sub RecentDataBaseOrderUpdate(ByVal RecentDataBaseFullName As String)
            Dim ToolStripMenuItemList As New ArrayList

            For Each ToolStripMenuItem As ToolStripMenuItem In MenuFileRecentDataBase.DropDownItems
                If ToolStripMenuItem.Tag.ToString.Trim.ToLower = RecentDataBaseFullName.Trim.ToLower Then
                    ToolStripMenuItemList.Add(ToolStripMenuItem)
                    Exit For
                End If
            Next

            If ToolStripMenuItemList.Count = 0 Then
                Dim ToolStripMenuItem As New ToolStripMenuItem
                With ToolStripMenuItem
                    .Name = "MenuFileRecentDataBaseItem"
                    .Tag = RecentDataBaseFullName
                    AddHandler .Click, AddressOf ToolStripMenuItem_Click
                End With
                ToolStripMenuItemList.Add(ToolStripMenuItem)
            End If

            For Each ToolStripMenuItem As ToolStripMenuItem In MenuFileRecentDataBase.DropDownItems
                If Not ToolStripMenuItem.Tag.ToString.Trim.ToLower = RecentDataBaseFullName.Trim.ToLower Then
                    ToolStripMenuItemList.Add(ToolStripMenuItem)
                End If
            Next

            MenuFileRecentDataBase.DropDownItems.Clear()
            Dim Index As Integer = 0
            For Each ToolStripMenuItem As ToolStripMenuItem In ToolStripMenuItemList
                Index += 1
                ToolStripMenuItem.Text = Index & ". " & ToolStripMenuItem.Tag.ToString
                MenuFileRecentDataBase.DropDownItems.Add(ToolStripMenuItem)
            Next

            SaveConfiguration()

            If MenuFileRecentDataBase.DropDownItems.Count > 0 Then
                MenuFileRecentDataBase.Enabled = True
            End If
        End Sub

        Public Sub RecentDataBaseOrderDelete(ByVal RecentDataBaseFullName As String)
            For Each ToolStripMenuItem As ToolStripMenuItem In MenuFileRecentDataBase.DropDownItems
                If ToolStripMenuItem.Tag.ToString.Trim.ToLower = RecentDataBaseFullName.Trim.ToLower Then
                    MenuFileRecentDataBase.DropDownItems.Remove(ToolStripMenuItem)
                    Exit For
                End If
            Next

            Dim Index As Integer = 0
            For Each ToolStripMenuItem As ToolStripMenuItem In MenuFileRecentDataBase.DropDownItems
                Index += 1
                ToolStripMenuItem.Text = Index & ". " & ToolStripMenuItem.Tag.ToString
            Next

            SaveConfiguration()

            If MenuFileRecentDataBase.DropDownItems.Count = 0 Then
                MenuFileRecentDataBase.Enabled = False
            End If
        End Sub

        Private Sub SaveConfiguration()
            Dim DataTable As New DataTable

            With DataTable
                .TableName = TableName.OpenFileHistoryList
                .Columns.Add("FileFullName")
            End With

            For Index As Integer = MenuFileRecentDataBase.DropDownItems.Count - 1 To 0 Step -1
                DataTable.Rows.Add(MenuFileRecentDataBase.DropDownItems(Index).Tag.ToString)
            Next

            Configuration.SetConfig(DataTable)
            Configuration.Save()
        End Sub

        Public Sub SetNextPreviousDataBaseEnable(ByVal Enable As Boolean)
            MenuViewNextDataBase.Enabled = Enable
            MenuViewPreviousDataBase.Enabled = Enable
        End Sub

        Public Function GetFormMainViewConfiguration() As DataTable
            Dim DataTable As New DataTable
            With DataTable
                .TableName = TableName.FormMainViewConfiguration
                .Columns.Add("Control")
                .Columns.Add("Parameter")

                .Rows.Add({MenuViewShowGroupTreeView.Name, MenuViewShowGroupTreeView.Checked})
                .Rows.Add({MenuViewShowStatusStrip.Name, MenuViewShowStatusStrip.Checked})
                .Rows.Add({MenuViewShowToolStrip.Name, MenuViewShowToolStrip.Checked})
            End With

            Return DataTable
        End Function
    End Class
End Namespace

