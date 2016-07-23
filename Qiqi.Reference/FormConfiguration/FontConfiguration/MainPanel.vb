Namespace _FormConfiguration
    Namespace FontConfiguration
        Public Class MainPanel
            Inherits System.Windows.Forms.Panel

            Public Modified As Boolean
            Dim Configuration As _FormConfiguration.Configuration

            Dim FontFamilyList As ArrayList

            Dim TableLayoutPanel As TableLayoutPanel
            Const TableLayoutPanelWidth As Integer = 250
            Const FirstColumnWidth As Integer = 80
            Const SecondColumnWidth As Integer = TableLayoutPanelWidth - FirstColumnWidth
            Const RowHeight As Integer = 22

            Dim ControlList As New ArrayList
            Dim GroupFontFamilyComboBox As FontFamilyComboBox
            Dim GroupFontSizeComboBox As FontSizeComboBox
            Dim ListFontFamilyComboBox As FontFamilyComboBox
            Dim ListFontSizeComboBox As FontSizeComboBox
            Dim DetailFontFamilyComboBox As FontFamilyComboBox
            Dim DetailFontSizeComboBox As FontSizeComboBox
            Dim HighlightFontFamilyComboBox As FontFamilyComboBox
            Dim HighlightFontSizeComboBox As FontSizeComboBox
            Dim EntryTypeFontColorComboBox As FontColorComboBox
            Dim BibTeXKeyFontColorComboBox As FontColorComboBox
            Dim TagNameFontColorComboBox As FontColorComboBox
            Dim TagValueFontColorComboBox As FontColorComboBox

            Public Sub New(ByVal Configuration As _FormConfiguration.Configuration)
                With Me
                    .Modified = False
                    .Configuration = Configuration
                End With

                InitializeFontFamilyList()
                InitializeTableLayoutPanel()
                InitializeControlValue()
            End Sub

            Private Sub InitializeControlValue()
                ' Fill the ListView
                Dim DataTable As New DataTable
                Dim Index As Integer = 0

                RemoveHandlerConfigurationChanged()
                If Configuration.GetConfig(TableName.FontConfigurationConfiguration, DataTable) Then
                    For Each Row As DataRow In DataTable.Rows
                        For Each Control As Control In ControlList
                            If Control.Name = Row(0) Then
                                If TypeOf Control Is FontFamilyComboBox Then
                                    CType(Control, FontFamilyComboBox).SelectedText = Row(1)
                                ElseIf TypeOf Control Is FontSizeComboBox Then
                                    CType(Control, FontSizeComboBox).SelectedText = Row(1)
                                ElseIf TypeOf Control Is FontColorComboBox Then
                                    CType(Control, FontColorComboBox).SelectedColor = Row(1)
                                    CType(Control, FontColorComboBox).Bold = Row(2)
                                Else
                                    Continue For
                                End If
                            End If
                        Next
                    Next
                End If
                AddHaddlerConfigurationChanged()
            End Sub

            Public Function GetConfiguration() As DataTable
                Dim DataTable As New DataTable

                With DataTable
                    .TableName = TableName.FontConfigurationConfiguration
                    .Columns.Add("Name")
                    .Columns.Add("Value")
                    .Columns.Add("Bold")
                End With

                Dim Row(2) As String
                For Each Control In ControlList
                    Row(0) = Control.Name
                    If TypeOf Control Is FontFamilyComboBox Then
                        Row(1) = CType(Control, FontFamilyComboBox).SelectedText
                        Row(2) = ""
                    ElseIf TypeOf Control Is FontSizeComboBox Then
                        Row(1) = CType(Control, FontSizeComboBox).SelectedText
                        Row(2) = ""
                    ElseIf TypeOf Control Is FontColorComboBox Then
                        Row(1) = CType(Control, FontColorComboBox).SelectedColor
                        Row(2) = CType(Control, FontColorComboBox).Bold
                    Else
                        Continue For
                    End If
                    DataTable.Rows.Add(Row)
                Next

                Return DataTable
            End Function

            Private Sub ConfigurationChanged()
                'MsgBox(BibTeXKeyFontColorComboBox.Name)
                GetConfiguration()
            End Sub

            Private Sub InitializeTableLayoutPanel()
                TableLayoutPanel = New TableLayoutPanel
                With TableLayoutPanel
                    .Dock = DockStyle.None
                    ' .CellBorderStyle = TableLayoutPanelCellBorderStyle.Single
                    .Location = New Point(0, 0)
                    .Size = New Size(TableLayoutPanelWidth, Me.ClientSize.Height - 2)
                    .Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Bottom
                    .ColumnCount = 1
                    .ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, TableLayoutPanelWidth))
                    .RowCount = 5
                    .RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                    .RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                    .RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                    .RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                    .RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))

                End With

                GroupFontFamilyComboBox = New FontFamilyComboBox(TableLayoutPanelWidth)
                GroupFontFamilyComboBox.Name = NameOf(GroupFontFamilyComboBox)

                GroupFontSizeComboBox = New FontSizeComboBox(TableLayoutPanelWidth)
                GroupFontSizeComboBox.Name = NameOf(GroupFontSizeComboBox)

                ListFontFamilyComboBox = New FontFamilyComboBox(TableLayoutPanelWidth)
                ListFontFamilyComboBox.Name = NameOf(ListFontFamilyComboBox)

                ListFontSizeComboBox = New FontSizeComboBox(TableLayoutPanelWidth)
                ListFontSizeComboBox.Name = NameOf(ListFontSizeComboBox)

                DetailFontFamilyComboBox = New FontFamilyComboBox(TableLayoutPanelWidth)
                DetailFontFamilyComboBox.Name = NameOf(DetailFontFamilyComboBox)

                DetailFontSizeComboBox = New FontSizeComboBox(TableLayoutPanelWidth)
                DetailFontSizeComboBox.Name = NameOf(DetailFontSizeComboBox)

                HighlightFontFamilyComboBox = New FontFamilyComboBox(TableLayoutPanelWidth)
                HighlightFontFamilyComboBox.Name = NameOf(HighlightFontFamilyComboBox)

                HighlightFontSizeComboBox = New FontSizeComboBox(TableLayoutPanelWidth)
                HighlightFontSizeComboBox.Name = NameOf(HighlightFontSizeComboBox)

                EntryTypeFontColorComboBox = New FontColorComboBox("Entry Type", TableLayoutPanelWidth)
                EntryTypeFontColorComboBox.Name = NameOf(EntryTypeFontColorComboBox)

                BibTeXKeyFontColorComboBox = New FontColorComboBox("BibTeX Key", TableLayoutPanelWidth)
                BibTeXKeyFontColorComboBox.Name = NameOf(BibTeXKeyFontColorComboBox)

                TagNameFontColorComboBox = New FontColorComboBox("Tag Name", TableLayoutPanelWidth)
                TagNameFontColorComboBox.Name = NameOf(TagNameFontColorComboBox)

                TagValueFontColorComboBox = New FontColorComboBox("Tag Value", TableLayoutPanelWidth)
                TagValueFontColorComboBox.Name = NameOf(TagValueFontColorComboBox)

                ControlList = New ArrayList
                Dim MarginBetweenZones As Integer = 8
                ControlList.Add(New FormSeparator("Group Font Configuration", TableLayoutPanelWidth))
                ControlList.Add(GroupFontFamilyComboBox)
                ControlList.Add(GroupFontSizeComboBox)
                ControlList.Add(New FormSeparator("List Font Configuration", TableLayoutPanelWidth, MarginBetweenZones))
                ControlList.Add(ListFontFamilyComboBox)
                ControlList.Add(ListFontSizeComboBox)
                ControlList.Add(New FormSeparator("Detail Font Configuration", TableLayoutPanelWidth, MarginBetweenZones))
                ControlList.Add(DetailFontFamilyComboBox)
                ControlList.Add(DetailFontSizeComboBox)
                ControlList.Add(New FormSeparator("Syntax Hightlight Font Configuration", TableLayoutPanelWidth, MarginBetweenZones))
                ControlList.Add(HighlightFontFamilyComboBox)
                ControlList.Add(HighlightFontSizeComboBox)
                ControlList.Add(New FormSeparator("Syntax Hightlight Color Configuration", TableLayoutPanelWidth, MarginBetweenZones))
                ControlList.Add(EntryTypeFontColorComboBox)
                ControlList.Add(BibTeXKeyFontColorComboBox)
                ControlList.Add(TagNameFontColorComboBox)
                ControlList.Add(TagValueFontColorComboBox)


                Dim Index As Integer = 0
                For Each Control As Control In ControlList
                    TableLayoutPanel.Controls.Add(Control, 0, Index)
                    Index += 1
                Next

                Me.Controls.Add(TableLayoutPanel)
            End Sub

            Private Sub InitializeFontFamilyList()
                FontFamilyList = New ArrayList
                For Each Font As FontFamily In System.Drawing.FontFamily.Families
                    FontFamilyList.Add(Font)
                Next
            End Sub

            Private Sub AddHaddlerConfigurationChanged()
                For Each Control In ControlList
                    If TypeOf Control Is FontFamilyComboBox Then
                        AddHandler CType(Control, FontFamilyComboBox).SelectedChanged, AddressOf ConfigurationChanged
                    ElseIf TypeOf Control Is FontSizeComboBox Then
                        AddHandler CType(Control, FontSizeComboBox).SelectedChanged, AddressOf ConfigurationChanged
                    ElseIf TypeOf Control Is FontColorComboBox Then
                        AddHandler CType(Control, FontColorComboBox).SelectedChanged, AddressOf ConfigurationChanged
                    End If
                Next
            End Sub

            Private Sub RemoveHandlerConfigurationChanged()
                For Each Control In ControlList
                    If TypeOf Control Is FontFamilyComboBox Then
                        RemoveHandler CType(Control, FontFamilyComboBox).SelectedChanged, AddressOf ConfigurationChanged
                    ElseIf TypeOf Control Is FontSizeComboBox Then
                        RemoveHandler CType(Control, FontSizeComboBox).SelectedChanged, AddressOf ConfigurationChanged
                    ElseIf TypeOf Control Is FontColorComboBox Then
                        RemoveHandler CType(Control, FontColorComboBox).SelectedChanged, AddressOf ConfigurationChanged
                    End If
                Next
            End Sub

        End Class
    End Namespace
End Namespace

