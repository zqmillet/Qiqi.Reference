Public Class FormConfiguration

    Dim Configuration As _FormConfiguration.Configuration
    Dim TreeView As TreeView
    Dim ButtonSize = New Size(75, 25)
    Dim ControlDistance As Integer = 4
    Dim TreeViewWidth As Integer = 170

    Dim ButtonOK As Button
    Dim ButtonCancel As Button

    Public BackupConfigurationPath As String

    Public Sub New(ByVal Configuration As _FormConfiguration.Configuration)

        ' Initialize component
        InitializeComponent()

        With Me
            ' .ControlBox = False
            .MinimumSize = New Size(900, 550)
            .MaximizeBox = False
            .Configuration = Configuration
            .FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
            .Icon = Resource.Icon.FormConfiguration
            .Text = "Preferences"
            .BackupConfigurationPath = Application.StartupPath & "\Config.bkp"
        End With

        ' Initialize tree view
        InitializeTreeView()

        ' Initialize buttons
        InitializeButton()

        ' Backup the configuration
        BackupConfiguration()
    End Sub

    Private Sub BackupConfiguration()
        Dim Configuration As _FormConfiguration.Configuration = Me.Configuration
        Configuration.SaveAs(BackupConfigurationPath)
    End Sub

    Private Sub InitializeButton()
        ButtonOK = New Button
        With ButtonOK
            .Size = ButtonSize
            .Location = New Point(Me.ClientSize.Width - 2 * ControlDistance - 2 * ButtonSize.width, _
                                  Me.ClientSize.Height - ControlDistance - ButtonSize.height)
            .FlatStyle = FlatStyle.Flat
            .FlatAppearance.BorderColor = Color.FromArgb(100, 100, 100)
            .Text = "OK"
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        End With

        ButtonCancel = New Button
        With ButtonCancel
            .Size = ButtonSize
            .Location = New Point(Me.ClientSize.Width - ControlDistance - ButtonSize.width, _
                                  Me.ClientSize.Height - ControlDistance - ButtonSize.height)
            .FlatStyle = FlatStyle.Flat
            .FlatAppearance.BorderColor = Color.FromArgb(100, 100, 100)
            .Text = "Cancel"
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        End With

        AddHandler ButtonOK.Click, AddressOf ButtonOK_Click
        AddHandler ButtonCancel.Click, AddressOf ButtonCancel_Click

        Me.Controls.Add(ButtonOK)
        Me.Controls.Add(ButtonCancel)

        Me.AcceptButton = ButtonOK
        Me.CancelButton = ButtonCancel
    End Sub

    Private Sub ButtonOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        For Each TreeNode As TreeNode In TreeView.Nodes
            Configuration.SetConfig(TreeNode.Tag.GetConfiguration)
        Next

        Configuration.Save()

        Me.DialogResult = Windows.Forms.DialogResult.OK
    End Sub

    Private Sub ButtonCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
    End Sub

    Private Sub InitializeTreeView()
        TreeView = New TreeView
        With TreeView
            .Size = New Size(TreeViewWidth, Me.ClientSize.Height - 3 * ControlDistance - ButtonSize.Height)
            .Location = New Point(ControlDistance, ControlDistance)
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Bottom
            .ItemHeight = 18
            .ShowRootLines = False
        End With

        Dim DataBaseColumnDisplay As New TreeNode
        With DataBaseColumnDisplay
            .Text = "Database Column Display"
            .Name = "DataBaseColumnDisplay"
            .Tag = New _FormConfiguration.DataBaseColumnDisplay.MainPanel(Configuration)
        End With
        TreeView.Nodes.Add(DataBaseColumnDisplay)

        Dim LiteratureDetailDisplay As New TreeNode
        With LiteratureDetailDisplay
            .Text = "Literature Detail Display"
            .Name = "LiteratureDetailDisplay"
            .Tag = New _FormConfiguration.LiteratureDetailDisplay.MainPanel(Configuration)
        End With
        TreeView.Nodes.Add(LiteratureDetailDisplay)

        Dim FontConfiguration As New TreeNode
        With FontConfiguration
            .Text = "Interface Font"
            .Name = "InterfaceFont"
            .Tag = New _FormConfiguration.InterfaceFont.MainPanel(Configuration)
        End With
        TreeView.Nodes.Add(FontConfiguration)

        Me.Controls.Add(TreeView)

        For Each TreeNode As TreeNode In TreeView.Nodes
            With TreeNode.Tag
                .Location = New Point(2 * ControlDistance + TreeViewWidth, ControlDistance)
                .Size = New Size(Me.ClientSize.Width - 3 * ControlDistance - TreeViewWidth, _
                                  Me.ClientSize.Height - 3 * ControlDistance - ButtonSize.Height)
                .Visible = False
                .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
            End With

            Me.Controls.Add(TreeNode.Tag)
        Next

        AddHandler TreeView.AfterSelect, AddressOf TreeView_AfterSelect
    End Sub

    Private Sub TreeView_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs)
        For Each TreeNode As TreeNode In TreeView.Nodes
            TreeNode.Tag.Visible = False
        Next

        TreeView.SelectedNode.Tag.Visible = True
    End Sub
End Class