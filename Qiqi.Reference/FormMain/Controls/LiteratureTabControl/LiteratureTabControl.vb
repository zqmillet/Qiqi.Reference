Namespace _FormMain
    Public Class LiteratureTabControl
        Inherits System.Windows.Forms.TabControl

        Dim DataTable As DataTable
        Dim AbsoluteRowHeight As Integer = 21
        Dim AbsoluteColumnWidth As Integer = 75
        Dim ControlMargin As Padding = New Padding(5)
        Dim Literature As _BibTeX.Literature

        Private Const TabPageConfigurationString = "Default{Required Fields{Author{Author,True,1,True,False},Title{Title,True,1,True,False},Journal{Journal,True,1,True,False},Year{Year,True,0,True,False},BibTeXKey{BibTeXKey,True,0,True,False}},Source Code{SourceCode{Source Code,False,0,False,True}}}"

        Public WriteOnly Property HighlightColor As _LiteratureTabControl.HighlightColor
            Set(Value As _LiteratureTabControl.HighlightColor)
                For Each TabPage As TabPage In Me.TabPages
                    If TabPage.Controls(0) Is Nothing Then
                        Exit Property
                    End If

                    If Not TypeOf TabPage.Controls(0) Is TableLayoutPanel Then
                        Exit Property
                    End If

                    With CType(TabPage.Controls(0), TableLayoutPanel)
                        For Index As Integer = 0 To .RowCount - 1
                            With CType(.GetControlFromPosition(.ColumnCount - 1, Index), _FormMain._LiteratureTabControl.MultiLineTextBox)
                                If .SyntaxHighlight Then
                                    .TextBox.HighlightColor = Value
                                    .TextBox.TextSyntaxHighLight()
                                End If
                            End With
                        Next
                    End With
                Next
            End Set
        End Property

        Public WriteOnly Property HighlightStyle As _LiteratureTabControl.HighlightStyle
            Set(Value As _LiteratureTabControl.HighlightStyle)
                For Each TabPage As TabPage In Me.TabPages
                    If TabPage.Controls(0) Is Nothing Then
                        Exit Property
                    End If

                    If Not TypeOf TabPage.Controls(0) Is TableLayoutPanel Then
                        Exit Property
                    End If

                    With CType(TabPage.Controls(0), TableLayoutPanel)
                        For Index As Integer = 0 To .RowCount - 1
                            With CType(.GetControlFromPosition(.ColumnCount - 1, Index), _FormMain._LiteratureTabControl.MultiLineTextBox)
                                If .SyntaxHighlight Then
                                    .TextBox.HighlightStyle = Value
                                    .TextBox.TextSyntaxHighLight()
                                End If
                            End With
                        Next
                    End With
                Next
            End Set
        End Property

        Public WriteOnly Property DetailFont As Font
            Set(Value As Font)
                For Each TabPage As TabPage In Me.TabPages
                    If TabPage.Controls(0) Is Nothing Then
                        Exit Property
                    End If

                    If Not TypeOf TabPage.Controls(0) Is TableLayoutPanel Then
                        Exit Property
                    End If

                    With CType(TabPage.Controls(0), TableLayoutPanel)
                        For Index As Integer = 0 To .RowCount - 1
                            With CType(.GetControlFromPosition(.ColumnCount - 1, Index), _FormMain._LiteratureTabControl.MultiLineTextBox)
                                If Not .SyntaxHighlight Then
                                    .TextBox.TextFont = Value
                                    .TextBox.TextFontUpdate()
                                End If
                            End With
                        Next
                    End With
                Next
            End Set
        End Property

        Public WriteOnly Property HighlightFont As Font
            Set(Value As Font)
                For Each TabPage As TabPage In Me.TabPages
                    If TabPage.Controls(0) Is Nothing Then
                        Exit Property
                    End If

                    If Not TypeOf TabPage.Controls(0) Is TableLayoutPanel Then
                        Exit Property
                    End If

                    With CType(TabPage.Controls(0), TableLayoutPanel)
                        For Index As Integer = 0 To .RowCount - 1
                            With CType(.GetControlFromPosition(.ColumnCount - 1, Index), _FormMain._LiteratureTabControl.MultiLineTextBox)
                                If .SyntaxHighlight Then
                                    .TextBox.TextFont = Value
                                    .TextBox.TextSyntaxHighLight()
                                End If
                            End With
                        Next
                    End With
                Next
            End Set
        End Property

        Public Sub New(ByVal Configuration As _FormConfiguration.Configuration)
            With Me
                .Dock = DockStyle.Fill
                .DataTable = New DataTable
                .Literature = New _BibTeX.Literature
                AddHandler .GotFocus, AddressOf Me_GotFocus
            End With

            If Not Configuration.GetConfig(TableName.LiteratureDetailDisplayConfiguration, Me.DataTable) Then
                Exit Sub
            End If
        End Sub

        'Public Sub Load(ByVal Configuration As _FormConfiguration.Configuration)
        '    Configuration.GetConfig(TableName.LiteratureDetailDisplayConfiguration, Me.DataTable)
        'End Sub

        Private Sub Me_GotFocus()
            Me.Enabled = False
            Me.Enabled = True
        End Sub

        Public Sub ResetTabPages()
            Me.DataTable.Rows.Clear()
            Me.DataTable.Rows.Add(TabPageConfigurationString)
        End Sub

        Public Sub Load(ByVal Literature As _BibTeX.Literature, Optional ByVal TabStop As Boolean = True)
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
                        ' .Controls.Add(GenerateControl(Literature.GetProperty(CType(TabPageConfiguration.Properties.Item(0), _FormConfiguration.LiteratureDetailDisplay.PropertyConfiguration).PropertyName)))
                        Dim PropertyConfiguration As _FormConfiguration.LiteratureDetailDisplay.PropertyConfiguration = TabPageConfiguration.Properties(0)
                        .Controls.Add(GenerateControl(PropertyConfiguration, TabStop))
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
                            ' .Controls.Add(GenerateControl(Literature.GetProperty(PropertyConfiguration.PropertyName)), 1, Index)
                            .Controls.Add(GenerateControl(PropertyConfiguration, TabStop), 1, Index)
                        Next
                    Else
                        MsgBox(ErrorMessage, MsgBoxStyle.OkOnly, "Error")
                        Exit Sub
                    End If
                End With

                If TabPageIndex <= Me.TabCount - 1 Then
                    With Me.TabPages(TabPageIndex)
                        .Controls.Add(TableLayoutPanel)
                        .Text = TabPageConfiguration.TabPageName
                        .TabStop = TabStop
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

        Private Function GenerateControl(ByVal PropertyConfiguration As _FormConfiguration.LiteratureDetailDisplay.PropertyConfiguration, Optional ByVal TabStop As Boolean = True) As Object
            Dim Control As Control

            If PropertyConfiguration.Height = 0 Then
                Control = New _FormMain._LiteratureTabControl.SingleLineTextBox()
            Else
                Control = New _FormMain._LiteratureTabControl.MultiLineTextBox()
            End If

            If RowCount = 1 Then
                Control = New _FormMain._LiteratureTabControl.MultiLineTextBox()
            End If

            CType(Control, _FormMain._LiteratureTabControl.MultiLineTextBox).SyntaxHighlight = PropertyConfiguration.SyntaxHighlight
            CType(Control, _FormMain._LiteratureTabControl.MultiLineTextBox).Text = Literature.GetProperty(PropertyConfiguration.PropertyName).Value
            CType(Control, _FormMain._LiteratureTabControl.MultiLineTextBox).TextBox.TabStop = TabStop

            Return Control
        End Function

        Private Function GenerateControl(ByVal LiteratureProperty As _BibTeX.LiteratureProperty) As Object
            Select Case LiteratureProperty.Name
                Case "BibTeXKey"
                    Return New _FormMain._LiteratureTabControl.SingleLineTextBox(LiteratureProperty.Value)
                Case "SourceCode", "Abstract", "Review"
                    Return New _FormMain._LiteratureTabControl.MultiLineTextBox(LiteratureProperty.Value)
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
