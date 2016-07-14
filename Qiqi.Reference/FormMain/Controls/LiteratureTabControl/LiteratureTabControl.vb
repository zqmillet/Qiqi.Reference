Namespace _FormMain
    Public Class LiteratureTabControl
        Inherits System.Windows.Forms.TabControl

        Dim DataTable As DataTable
        Dim AbsoluteRowHeight As Integer = 21
        Dim AbsoluteColumnWidth As Integer = 75
        Dim ControlMargin As Padding = New Padding(5)
        Dim Literature As _BibTeX.Literature

        Public Sub New(ByVal Configuration As _FormConfiguration.Configuration)
            With Me
                .Dock = DockStyle.Fill
                .DataTable = New DataTable
                .Literature = New _BibTeX.Literature
            End With
            Configuration.GetConfig(TableName.LiteratureDetailDisplayConfiguration, Me.DataTable)
        End Sub

        Public Sub Load(ByVal Configuration As _FormConfiguration.Configuration)
            Configuration.GetConfig(TableName.LiteratureDetailDisplayConfiguration, Me.DataTable)
        End Sub

        Public Sub Load(ByVal Literature As _BibTeX.Literature)
            If Literature Is Nothing Then
                Exit Sub
            End If

            Me.Literature = Literature

            Dim ErrorMessage As String = "There exist an error when find the configuration of literature type " & Literature.Type

            ' Find the display configuration of this literature
            Dim LiteratureDetailDisplayConfiguration As _FormConfiguration.LiteratureDetailDisplay.Configuration = GetLiteratureDetailDisplayConfiguration(Literature.Type)
            If LiteratureDetailDisplayConfiguration.LiteratureType = "" Then
                MsgBox(ErrorMessage, MsgBoxStyle.OkOnly, "Error")
                Exit Sub
            End If

            Dim TabPageIndex As Integer = 0
            For Each TabPageConfiguration As _FormConfiguration.LiteratureDetailDisplay.TabPageConfiguration In LiteratureDetailDisplayConfiguration.TabPages
                Dim TableLayoutPanel As New TableLayoutPanel
                With TableLayoutPanel
                    .Dock = DockStyle.Fill
                    .RowCount = TabPageConfiguration.Properties.Count
                    .CellBorderStyle = TableLayoutPanelCellBorderStyle.Single

                    If CType(TabPageConfiguration.Properties.Item(0), _FormConfiguration.LiteratureDetailDisplay.PropertyConfiguration).DisplayLabel Then
                        .ColumnCount = 2
                    Else
                        .ColumnCount = 1
                    End If


                    If .ColumnCount = 1 And .RowCount = 1 Then
                        ' 1*1
                        .ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100))
                        .RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100))
                        .Controls.Add(GenerateControl(Literature.GetProperty(CType(TabPageConfiguration.Properties.Item(0), _FormConfiguration.LiteratureDetailDisplay.PropertyConfiguration).PropertyName)))
                    ElseIf .ColumnCount = 2 And .RowCount > 0 Then
                        ' n*2
                        .ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, AbsoluteColumnWidth))
                        .ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100))

                        Dim Index As Integer = 0
                        For Each PropertyConfiguration As _FormConfiguration.LiteratureDetailDisplay.PropertyConfiguration In TabPageConfiguration.Properties
                            If PropertyConfiguration.Height = 0 Then
                                Index = .RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, AbsoluteRowHeight))
                            ElseIf PropertyConfiguration.Height > 0 Then
                                Index = .RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, PropertyConfiguration.Height))
                            Else
                                MsgBox(ErrorMessage, MsgBoxStyle.OkOnly, "Error")
                                Exit Sub
                            End If

                            Dim Label As New Label
                            With Label
                                .Text = PropertyConfiguration.DisplayName
                                .Margin = ControlMargin
                            End With

                            .Controls.Add(Label, 0, Index)
                            .Controls.Add(GenerateControl(Literature.GetProperty(PropertyConfiguration.PropertyName)), 1, Index)
                        Next
                    Else
                        MsgBox(ErrorMessage, MsgBoxStyle.OkOnly, "Error")
                        Exit Sub
                    End If


                    For Each PropertyConfiguration As _FormConfiguration.LiteratureDetailDisplay.PropertyConfiguration In TabPageConfiguration.Properties

                    Next
                End With


                If TabPageIndex <= Me.TabCount - 1 Then
                    With Me.TabPages(TabPageIndex)
                        .Controls.Add(TableLayoutPanel)
                        .Text = TabPageConfiguration.TabPageName
                        If .Controls.Count > 1 Then
                            .Controls.RemoveAt(0)
                        End If

                        If .Controls.Count >= 2 Then
                            MsgBox("Error!")
                        End If
                    End With
                Else
                    Dim TabPage As New TabPage
                    With TabPage
                        .Text = TabPageConfiguration.TabPageName
                        .Controls.Add(TableLayoutPanel)
                    End With

                    Me.TabPages.Add(TabPage)
                End If
                TabPageIndex += 1
            Next

            For Index As Integer = Me.TabPages.Count - 1 To TabPageIndex Step -1
                Me.TabPages.RemoveAt(Index)
            Next
        End Sub

        Private Function GenerateControl(ByVal LiteratureProperty As _BibTeX.LiteratureProperty) As Object
            Select Case LiteratureProperty.Name
                Case "BibTeXKey"
                    Return New _FormMain._LiteratureTabControl.SingleLineTextBox(LiteratureProperty.Value)
                Case "SourceCode", "Abstract", "Review"
                    Return New _FormMain._LiteratureTabControl.SourceCodeTextBox(LiteratureProperty.Value)
                Case Else
                    Return New _FormMain._LiteratureTabControl.MultiLineTextBox(LiteratureProperty.Value)
            End Select
        End Function


        Private Function GetLiteratureDetailDisplayConfiguration(ByVal LiteratureType As String) As _FormConfiguration.LiteratureDetailDisplay.Configuration
            Dim DefaultLiteratureDetailDisplayConfiguration As New _FormConfiguration.LiteratureDetailDisplay.Configuration

            For Each Row As DataRow In DataTable.Rows
                Dim LiteratureDetailDisplayConfiguration As New _FormConfiguration.LiteratureDetailDisplay.Configuration(Row(0))

                If LiteratureDetailDisplayConfiguration.LiteratureType.ToLower.Trim = LiteratureType.ToLower.Trim Then
                    Return LiteratureDetailDisplayConfiguration
                End If

                If LiteratureDetailDisplayConfiguration.LiteratureType.ToLower.Trim = "Default".ToLower.Trim Then
                    DefaultLiteratureDetailDisplayConfiguration = LiteratureDetailDisplayConfiguration
                End If
            Next
            Return DefaultLiteratureDetailDisplayConfiguration
        End Function

        Public Sub DisplayRefresh(ByVal DataTableColumnConfiguration As DataTable)
            Me.DataTable = DataTableColumnConfiguration

            If Me.Literature.Type = "" Then
                Exit Sub
            End If

            Load(Literature)
        End Sub

    End Class
End Namespace
