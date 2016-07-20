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
            Dim GroupFontFamilyComboBox As ComboBox
            Dim GroupFontFamilyLabel As Label

            Dim GroupFontSizeComboBox As ComboBox
            Dim GroupFontSizeLabel As Label

            Dim ListFontFamilyComboBox As ComboBox
            Dim ListFontFamilyLabel As Label

            Dim ListFontSizeComboBox As ComboBox
            Dim ListFontSizeLabel As Label

            Dim DetailFontFamilyComboBox As ComboBox
            Dim DetailFontFamilyLabel As Label

            Dim DetailFontSizeComboBox As ComboBox
            Dim DetailFontSizeLabel As Label

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
                    .CellBorderStyle = TableLayoutPanelCellBorderStyle.Single
                    .Location = New Point(0, 0)
                    .Size = New Size(TableLayoutPanelWidth, Me.ClientSize.Height - 4)
                    .ColumnCount = 2
                    .ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, FirstColumnWidth))
                    .ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, SecondColumnWidth))
                    .RowCount = 2
                    .RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, RowHeight))

                End With

                GroupFontFamilyLabel = New Label
                GroupFontFamilyLabel.Text = "Font Family"
                GroupFontFamilyLabel.Margin = New Padding(0, 4, 0, 0)
                TableLayoutPanel.Controls.Add(GroupFontFamilyLabel, 0, 0)
                GroupFontFamilyComboBox = New FontFamilyComboBox(FontFamilyList)
                TableLayoutPanel.Controls.Add(GroupFontFamilyComboBox, 1, 0)



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

