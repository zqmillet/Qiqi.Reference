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

        Public MenuOption As ToolStripMenuItem
        Public MenuOptionPreferences As ToolStripMenuItem

        Public Event MenuFileNewDataBaseClick()
        Public Event MenuFileOpenDataBaseClick()
        Public Event MenuFileCloseDataBaseClick()
        Public Event MenuFileExitClick()
        Public Event MenuFileRecentDataBaseItemClick(ByVal RecentDataBaseFullName As String)
        Public Event MenuOptionPreferencesClick()

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

                Dim DataTable As New DataTable
                If Configuration.GetConfig(TableName.OpenFileHistoryList, DataTable) = True Then
                    .Enabled = True

                    For Index As Integer = DataTable.Rows.Count - 1 To 0 Step -1
                        Dim MenuFileRecentDataBaseItem As New ToolStripMenuItem

                        With MenuFileRecentDataBaseItem
                            .Name = "MenuFileRecentDataBaseItem"
                            .Tag = DataTable.Rows(Index).Item(0)
                            .Text = DataTable.Rows.Count - Index & ". " & DataTable.Rows(Index).Item(0)
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
            Me.Items.Add(MenuOption)
        End Sub

        Private Sub ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Select Case CType(sender, ToolStripMenuItem).Name
                Case "MenuFileNewDataBase"
                    RaiseEvent MenuFileNewDataBaseClick()
                Case "MenuFileOpenDataBase"
                    RaiseEvent MenuFileOpenDataBaseClick()
                Case "MenuFileCloseDataBase"
                    RaiseEvent MenuFileCloseDataBaseClick()
                Case "MenuFileExit"
                    RaiseEvent MenuFileExitClick()
                Case "MenuFileRecentDataBaseItem"
                    RecentDataBaseOrderUpdate(CType(sender, ToolStripMenuItem))
                    RaiseEvent MenuFileRecentDataBaseItemClick(CType(sender, ToolStripMenuItem).Tag.ToString)
                Case "MenuOptionPreferences"
                    RaiseEvent MenuOptionPreferencesClick()
            End Select
        End Sub

        Public Sub RecentDataBaseOrderUpdate(ByVal LastDataBaseItem As ToolStripMenuItem)
            Dim ToolStripMenuItemList As New ArrayList

            ToolStripMenuItemList.Add(LastDataBaseItem)
            For Each ToolStripMenuItem As ToolStripMenuItem In MenuFileRecentDataBase.DropDownItems
                If Not ToolStripMenuItem.Equals(LastDataBaseItem) Then
                    ToolStripMenuItemList.Add(ToolStripMenuItem)
                End If
            Next

            MenuFileRecentDataBase.DropDownItems.Clear()
            For Each ToolStripMenuItem As ToolStripMenuItem In ToolStripMenuItemList
                MenuFileRecentDataBase.DropDownItems.Add(ToolStripMenuItem)
            Next
        End Sub
    End Class
End Namespace

