﻿Namespace _FormMain
    Public Class DataBaseTabPage
        Inherits System.Windows.Forms.TabPage

        Public Loaded As Boolean

        Public BibTeXFullName As String

        Public SplitContainerPrimary As SplitContainer
        Public SplitContainerSecondary As SplitContainer

        Dim DataBaseGridView As _FormMain.DataGridView
        Dim StructureTreeView As TreeView
        Dim LiteratureTabControl As _FormMain.LiteratureTabControl
        Dim LiteratureToolStrip As _FormMain.SecondaryToolStrip

        Dim Configuration As _FormConfiguration.Configuration

        Public Event ProgressUpdate(ByVal Progress As Double)
        Public Event ColumnDisplayChanged(ByVal sender As Object)
        Public Event SplitterMoved(ByVal sender As Object, ByVal e As System.Windows.Forms.SplitterEventArgs)
        Public Event DataBaseGridViewDoubleClick(ByVal sender As Object)

        ''' <summary>
        ''' Constructor with parameter
        ''' </summary>
        ''' <param name="BibTeXFullName">The full name of BibTeX file</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal BibTeXFullName As String, ByVal Configuration As _FormConfiguration.Configuration)
            Loaded = False

            ' If the full name of file is wrong, exit sub
            If Not My.Computer.FileSystem.FileExists(BibTeXFullName) Then
                Exit Sub
            End If

            With Me
                .Configuration = Configuration
                .BibTeXFullName = BibTeXFullName
                .Name = BibTeXFullName
                .Text = BibTeXFullName.Remove(0, BibTeXFullName.LastIndexOf("\") + 1)
                .ToolTipText = BibTeXFullName
            End With

            ' Initialize primary split container 
            InitializeSplitContainerPrimary()

            ' Initialize secondary split container 
            InitializeSplitContainerSecondary()

            ' Initialize structure tree view
            InitializeStuctureTreeView()

            ' Initialize database grid view
            InitializeDataBaseGridView()

            ' Initialize literature tab control
            InitializeLiteratureTabControl()
        End Sub

        Public Sub DataBaseGridView_ProgressUpdate(ByVal Progress As Double)
            RaiseEvent ProgressUpdate(Progress)
        End Sub


        Private Sub InitializeLiteratureToolStrip()

        End Sub

        Private Sub InitializeSplitContainerPrimary()
            ' Create a new split container
            SplitContainerPrimary = New SplitContainer

            ' Configuration
            With SplitContainerPrimary
                .Name = "SplitContainerPrimary"
                .Dock = DockStyle.Fill
                .Orientation = Orientation.Vertical
                .Panel1Collapsed = True
                AddHandler .SplitterMoved, AddressOf SplitContainer_SplitterMoved

            End With

            ' Add this split container in the tab page
            Me.Controls.Add(SplitContainerPrimary)
        End Sub

        Private Sub InitializeSplitContainerSecondary()
            ' Create a new split container
            SplitContainerSecondary = New SplitContainer

            ' Configuration
            With SplitContainerSecondary
                .Name = "SplitContainerSecondary"
                .Dock = DockStyle.Fill
                .Orientation = Orientation.Horizontal
                .Panel2Collapsed = True
                AddHandler .SplitterMoved, AddressOf SplitContainer_SplitterMoved
            End With

            ' Add this split container in the tab page
            SplitContainerPrimary.Panel2.Controls.Add(SplitContainerSecondary)
        End Sub

        Private Sub SplitContainer_SplitterMoved(ByVal sender As Object, ByVal e As System.Windows.Forms.SplitterEventArgs)
            RaiseEvent SplitterMoved(sender, e)
        End Sub

        Private Sub InitializeStuctureTreeView()
            ' Create a new tree view
            StructureTreeView = New TreeView

            With StructureTreeView
                .Dock = DockStyle.Fill
            End With

            SplitContainerPrimary.Panel1.Controls.Add(StructureTreeView)
        End Sub

        Public Function DataBaseLoading(ByRef ErrorMessage As _BibTeX.ErrorMessage) As Boolean
            Loaded = True
            Return DataBaseGridView.DataBaseLoading(ErrorMessage)
        End Function

        Private Sub InitializeDataBaseGridView()
            Dim DataTableColumnConfiguration As New DataTable
            If Not Configuration.GetConfig(TableName.DataBaseGridViewColunmConfiguration, DataTableColumnConfiguration) Then
                MsgBox("There is an error when load configuration.", MsgBoxStyle.OkOnly, "Configuration Error")
            End If

            ' Create a new data grid view
            DataBaseGridView = New _FormMain.DataGridView(BibTeXFullName, DataTableColumnConfiguration)

            SplitContainerSecondary.Panel1.Controls.Add(DataBaseGridView)

            AddHandler DataBaseGridView.ProgressUpdate, AddressOf DataBaseGridView_ProgressUpdate
            AddHandler DataBaseGridView.CellDoubleClick, AddressOf DataBaseGridView_CellDoubleClick
            AddHandler DataBaseGridView.SelectionChanged, AddressOf DataBaseGridView_SelectionChanged
            AddHandler DataBaseGridView.ColumnDisplayChanged, AddressOf DataBaseGridView_ColumnDisplayChanged
        End Sub

        Public Sub SetDataBaseGridViewEventEnable(ByVal Enable As Boolean)
            If Enable Then
                AddHandler DataBaseGridView.ColumnDisplayChanged, AddressOf DataBaseGridView_ColumnDisplayChanged
            Else
                RemoveHandler DataBaseGridView.ColumnDisplayChanged, AddressOf DataBaseGridView_ColumnDisplayChanged
            End If
        End Sub

        Private Sub InitializeLiteratureTabControl()
            LiteratureTabControl = New _FormMain.LiteratureTabControl(Configuration)
            LiteratureToolStrip = New _FormMain.SecondaryToolStrip()

            SplitContainerSecondary.Panel2.Controls.Add(LiteratureTabControl)
            SplitContainerSecondary.Panel2.Controls.Add(LiteratureToolStrip)

            AddHandler LiteratureToolStrip.ToolStripButtonHideClick, AddressOf LiteratureToolStrip_ToolStripButtonHideClick
            AddHandler LiteratureToolStrip.ToolStripButtonMoveUpClick, AddressOf LiteratureToolStrip_ToolStripButtonMoveUpClick
            AddHandler LiteratureToolStrip.ToolStripButtonMoveDownClick, AddressOf LiteratureToolStrip_ToolStripButtonMoveDownClick
        End Sub

        Private Sub DataBaseGridView_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
            RaiseEvent DataBaseGridViewDoubleClick(Me)

            If DataBaseGridView.SelectedRows.Count = 0 Then
                Exit Sub
            End If

            LiteratureTabControl.Load(CType(DataBaseGridView.SelectedRows(0).Tag, _BibTeX.Literature))
            SplitContainerSecondary.Panel2Collapsed = False
        End Sub

        Private Sub LiteratureToolStrip_ToolStripButtonHideClick()
            SplitContainerSecondary.Panel2Collapsed = True
        End Sub

        Private Sub LiteratureToolStrip_ToolStripButtonMoveUpClick()
            If DataBaseGridView.SelectedRows.Count = 0 Then
                Exit Sub
            End If

            MsgBox(DataBaseGridView.SelectedRows(0).Index)
        End Sub

        Private Sub LiteratureToolStrip_ToolStripButtonMoveDownClick()

        End Sub

        Private Sub DataBaseGridView_ColumnDisplayChanged(ByVal sender As Object)
            Configuration.SetConfig(CType(sender, _FormMain.DataGridView).GetColumnDisplayConfiguration)
            Configuration.Save()
            RaiseEvent ColumnDisplayChanged(Me)
        End Sub

        Private Sub DataBaseGridView_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs)
            If DataBaseGridView.SelectedRows.Count = 0 Then
                Exit Sub
            End If

            Dim MinIndex As Integer = DataBaseGridView.Rows.Count
            Dim MaxIndex As Integer = 0

            For Each Row As DataGridViewRow In DataBaseGridView.SelectedRows
                If MinIndex > Row.Index Then
                    MinIndex = Row.Index
                End If

                If MaxIndex < Row.Index Then
                    MaxIndex = Row.Index
                End If
            Next

            LiteratureToolStrip.SetButtonEnable(MinIndex = 0, MaxIndex = DataBaseGridView.Rows.Count - 1)

            If Not SplitContainerSecondary.Panel2Collapsed Then
                LiteratureTabControl.Load(CType(DataBaseGridView.SelectedRows(0).Tag, _BibTeX.Literature))
                DataBaseGridView.Focus()
            End If

        End Sub

        Public Sub SetSplitContainerEventEnable(ByVal Enable As Boolean)
            If Enable Then
                AddHandler SplitContainerPrimary.SplitterMoved, AddressOf SplitContainer_SplitterMoved
                AddHandler SplitContainerSecondary.SplitterMoved, AddressOf SplitContainer_SplitterMoved
            Else
                RemoveHandler SplitContainerPrimary.SplitterMoved, AddressOf SplitContainer_SplitterMoved
                RemoveHandler SplitContainerSecondary.SplitterMoved, AddressOf SplitContainer_SplitterMoved
            End If
        End Sub

        Public Sub DataBaseGridViewRefresh(ByVal Configuration As _FormConfiguration.Configuration)
            Dim DataTable As New DataTable
            If Configuration.GetConfig(TableName.DataBaseGridViewColunmConfiguration, DataTable) Then
                DataBaseGridView.DisplayRefresh(DataTable)
            End If
        End Sub

        Public Sub LiteratureTabControlRefresh(ByVal Configuration As _FormConfiguration.Configuration)
            Dim DataTable As New DataTable
            If Configuration.GetConfig(TableName.LiteratureDetailDisplayConfiguration, DataTable) Then
                LiteratureTabControl.DisplayRefresh(DataTable)
            End If
        End Sub
    End Class
End Namespace

