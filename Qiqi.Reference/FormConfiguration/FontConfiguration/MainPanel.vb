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

            Dim GroupFontFamilyComboBox As FontFamilyComboBox
            Dim GroupFontSizeComboBox As FontSizeComboBox
            Dim ListFontFamilyComboBox As FontFamilyComboBox
            Dim ListFontSizeComboBox As FontSizeComboBox
            Dim DetailFontFamilyComboBox As FontFamilyComboBox
            Dim DetailFontSizeComboBox As FontSizeComboBox
            Dim HighlightFontFamilyComboBox As FontFamilyComboBox
            Dim HighlightFontSizeComboBox As FontSizeComboBox


            Public Sub New(ByVal Configuration As _FormConfiguration.Configuration)
                With Me
                    .Modified = False
                    .Configuration = Configuration
                End With

                InitializeFontFamilyList()
                InitializeTableLayoutPanel()
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
                GroupFontSizeComboBox = New FontSizeComboBox(TableLayoutPanelWidth)
                ListFontFamilyComboBox = New FontFamilyComboBox(TableLayoutPanelWidth)
                ListFontSizeComboBox = New FontSizeComboBox(TableLayoutPanelWidth)
                DetailFontFamilyComboBox = New FontFamilyComboBox(TableLayoutPanelWidth)
                DetailFontSizeComboBox = New FontSizeComboBox(TableLayoutPanelWidth)
                HighlightFontFamilyComboBox = New FontFamilyComboBox(TableLayoutPanelWidth)
                HighlightFontSizeComboBox = New FontSizeComboBox(TableLayoutPanelWidth)

                Dim ControlList As New ArrayList
                ControlList.Add(New FormSeparator("Group Font Configuration", TableLayoutPanelWidth))
                ControlList.Add(GroupFontFamilyComboBox)
                ControlList.Add(GroupFontSizeComboBox)
                ControlList.Add(New FormSeparator("List Font Configuration", TableLayoutPanelWidth))
                ControlList.Add(ListFontFamilyComboBox)
                ControlList.Add(ListFontSizeComboBox)
                ControlList.Add(New FormSeparator("Detail Font Configuration", TableLayoutPanelWidth))
                ControlList.Add(DetailFontFamilyComboBox)
                ControlList.Add(DetailFontSizeComboBox)
                ControlList.Add(New FormSeparator("Syntax Hightlight Font Configuration", TableLayoutPanelWidth))
                ControlList.Add(HighlightFontFamilyComboBox)
                ControlList.Add(HighlightFontSizeComboBox)

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
        End Class
    End Namespace
End Namespace

