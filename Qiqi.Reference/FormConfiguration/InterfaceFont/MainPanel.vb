Namespace _FormConfiguration
    Namespace InterfaceFont
        ''' <summary>
        ''' This class is the panel of the interface font configuration.
        ''' </summary>
        Public Class MainPanel
            Inherits System.Windows.Forms.Panel

            ' All configuration of the whole software.
            Dim Configuration As _FormConfiguration.Configuration

            ' This TableLayoutPanel is used to contain all ComboBoxex of this panel.
            Dim TableLayoutPanel As TableLayoutPanel

            ' The following constants are used to draw the interface of this panel.
            ' The width of the TableLayoutPanel.
            Const TableLayoutPanelWidth As Integer = 240
            ' The width and the height of the GroupTreeView.
            Const GroupTreeViewWidth As Integer = 150
            Const GroupTreeViewHeight As Integer = 200
            ' This panel is separated into two parts:
            ' - left part is used to gain the input parameters,
            ' - right part is used to show the preview.
            ' This constant is the distance between the left part and the right part.
            Const SpaceBetweenLeftAndRight As Integer = 15
            ' The left part is separated into several input zones,
            ' and this constant is the distance between two input zones.
            Const MarginBetweenZones As Integer = 21
            ' The file Template.bib is the sample which is shown in the preview interface.
            ' This constant is the path of the file Template.bib
            Dim TamplateBibTeXPath As String = Application.StartupPath & "\Template.bib"

            ' The following ComboBoxes are used to gain the user's input data.
            ' The following two ComboBoxes are used to gain the font famility and size of the GroupTreeView.
            Dim GroupFontFamilyComboBox As FontFamilyComboBox
            Dim GroupFontSizeComboBox As FontSizeComboBox
            ' The following two ComboBoxes are used to gain the font famility and size of the DateGridView.
            Dim ListFontFamilyComboBox As FontFamilyComboBox
            Dim ListFontSizeComboBox As FontSizeComboBox
            ' The following two ComboBoxes are used to gain the font famility and size of the LiteratureControlTab.
            Dim DetailFontFamilyComboBox As FontFamilyComboBox
            Dim DetailFontSizeComboBox As FontSizeComboBox
            ' The following two ComboBoxes are used to gain the font famility and size of the BibTeX codes.
            Dim HighlightFontFamilyComboBox As FontFamilyComboBox
            Dim HighlightFontSizeComboBox As FontSizeComboBox
            ' The following four ComboBoxes are used to gain the colors of the the BibTeX codes.
            Dim EntryTypeFontColorComboBox As FontColorComboBox
            Dim BibTeXKeyFontColorComboBox As FontColorComboBox
            Dim TagNameFontColorComboBox As FontColorComboBox
            Dim TagValueFontColorComboBox As FontColorComboBox
            ' The ControlList is ArrayList which is used to contain the above ComboBoxes.
            ' With this ArrayList, the above ComboBoxes can be initialize in a loop.
            Dim ControlList As New ArrayList

            ' The FormSaparator is a label of the preview.
            Dim FormSaparator As FormSeparator
            ' The following three controls consist the preview interface.
            Dim GroupTreeView As _FormMain.GroupTreeView
            Dim DataGridView As _FormMain.DataGridView
            Dim LiteratureTabControl As _FormMain.LiteratureTabControl

            ' The following four panel are transparent. 
            ' They are put over other controls, which should not be click by mouse.
            ' They have another function:
            ' we can draw red rectangles on them to emphasize the controls which are under them. 
            Dim CoverPanel4GroupTreeView As TransparentPanel
            Dim CoverPanel4DataGridView As TransparentPanel
            Dim CoverPanel4LiteratureTabControl As TransparentPanel
            Dim CoverPanel4AllControls As TransparentPanel
            ' The CoverPanelList is ArrayList which is used to contain the above panels.
            ' With this ArrayList, the above panels can be initialize in a loop.
            Dim CoverPanelList As ArrayList

            ''' <summary>
            ''' This is the constructor.
            ''' </summary>
            ''' <param name="Configuration">All configuration of the whole software.</param>
            Public Sub New(ByVal Configuration As _FormConfiguration.Configuration)
                With Me
                    .Configuration = Configuration
                    ' We set the size of this panel is (500, 500) to avoid the situation that some controls' sizes are negatives.
                    ' The real size of this panel should be assigned after this panel is constructed.
                    .ClientSize = New Size(500, 500)
                End With

                InitializeTableLayoutPanel()
                InitializeInputComboBoxes()
                InitializePreview()
            End Sub

            ''' <summary>
            ''' This sub is used to initialize the preview interface.
            ''' </summary>
            Private Sub InitializePreview()
                ' Create the label of the preview interface.
                FormSaparator = New FormSeparator("Preview", Me.ClientSize.Width - TableLayoutPanelWidth - SpaceBetweenLeftAndRight)
                With FormSaparator
                    .Location = New Point(TableLayoutPanelWidth + SpaceBetweenLeftAndRight, 0)
                    .Anchor = AnchorStyles.Right Or AnchorStyles.Left Or AnchorStyles.Top
                End With
                Me.Controls.Add(FormSaparator)

                ' Create the panel to cover all the controls.
                CoverPanel4AllControls = New TransparentPanel
                With CoverPanel4AllControls
                    .BorderStyle = BorderStyle.None
                    .Width = Me.ClientSize.Width - TableLayoutPanelWidth
                    .Height = Me.ClientSize.Height
                    .Location = New Point(TableLayoutPanelWidth, 0)
                    ' HideBorder is a custom parameter,
                    ' the TransparentPanel has a red border with 6pt width by default,
                    ' if set the HideBorder parameter is True,
                    ' this red border will be hidden.
                    .HideBorder = True
                    .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
                End With
                Me.Controls.Add(CoverPanel4AllControls)

                ' Greate there panels to cover GroupTreeView, DataGridView, and LiteratureTabControl.
                CoverPanel4GroupTreeView = New TransparentPanel()
                CoverPanel4DataGridView = New TransparentPanel()
                CoverPanel4LiteratureTabControl = New TransparentPanel()
                ' Then put these three panels in to an ArrayList.
                CoverPanelList = New ArrayList
                CoverPanelList.Add(CoverPanel4GroupTreeView)
                CoverPanelList.Add(CoverPanel4DataGridView)
                CoverPanelList.Add(CoverPanel4LiteratureTabControl)
                ' Make these three panels invisible.
                ' At last, add them on the MainPanel.
                For Each Cover As TransparentPanel In CoverPanelList
                    Cover.Visible = False
                    Me.Controls.Add(Cover)
                Next

                ' Create the GroupTreeView, and put it on the MainPanel.
                GroupTreeView = New _FormMain.GroupTreeView(Configuration)
                With GroupTreeView
                    .Width = GroupTreeViewWidth
                    .Height = GroupTreeViewHeight
                    .Location = New Point(TableLayoutPanelWidth + SpaceBetweenLeftAndRight, FormSaparator.Height + 4)
                    .Anchor = AnchorStyles.Top Or AnchorStyles.Left
                    .TabStop = False
                    .Scrollable = False
                End With
                Me.Controls.Add(GroupTreeView)

                ' Create the DataGridView, and put it on the MainPanel.
                DataGridView = New _FormMain.DataGridView(Configuration)
                With DataGridView
                    .Width = Me.ClientSize.Width - TableLayoutPanelWidth - SpaceBetweenLeftAndRight - GroupTreeViewWidth - 4
                    .Height = GroupTreeViewHeight
                    .Location = New Point(TableLayoutPanelWidth + SpaceBetweenLeftAndRight + GroupTreeViewWidth + 4, FormSaparator.Height + 4)
                    .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
                    .TabStop = False
                    .ResetColumn()
                    .ScrollBars = ScrollBars.None

                    ' Load the Template.bib file and show the data on the DataGridView.
                    Dim ErrorMessage As New _BibTeX.ErrorMessage
                    If Not .DataBaseLoading(TamplateBibTeXPath, ErrorMessage) Then
                        MsgBox(ErrorMessage.GetErrorMessage)
                    End If

                    ' Show the group tree on the GroupTreeView.
                    GroupTreeView.Loading(.DataBase)
                    ' Select the first node.
                    If Not GroupTreeView.Nodes(0).Nodes(0).Nodes(0) Is Nothing Then
                        With GroupTreeView.Nodes(0).Nodes(0).Nodes(0)
                            .BackColor = System.Drawing.SystemColors.Highlight
                            .ForeColor = System.Drawing.SystemColors.HighlightText
                        End With
                    End If
                End With
                Me.Controls.Add(DataGridView)

                ' Create the LiteratureTabControl, and put it on the MainPanel.
                LiteratureTabControl = New _FormMain.LiteratureTabControl(Configuration)
                With LiteratureTabControl
                    .Width = Me.ClientSize.Width - TableLayoutPanelWidth - SpaceBetweenLeftAndRight
                    .Height = Me.ClientSize.Height - GroupTreeViewHeight - FormSaparator.Height - 4 - 4
                    .Location = New Point(TableLayoutPanelWidth + SpaceBetweenLeftAndRight, GroupTreeViewHeight + 4 + FormSaparator.Height + 4)
                    .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
                    .TabStop = False
                    .ResetTabPages()
                    .Load(DataGridView.DataBase.LiteratureList(0), False)
                End With
                Me.Controls.Add(LiteratureTabControl)

                ' Add the GroupTreeView, DataGridView, and LiteratureTabControl into the Tags of their cover panels.
                ' With this map, we can handle them more easily.
                CoverPanel4GroupTreeView.Tag = GroupTreeView
                CoverPanel4DataGridView.Tag = DataGridView
                CoverPanel4LiteratureTabControl.Tag = LiteratureTabControl

                ' Initialize the fonts and colors of the preview controls.
                GroupTreeView.Font = New Font(GroupFontFamilyComboBox.SelectedText, Val(GroupFontSizeComboBox.SelectedText))
                DataGridView.DefaultCellStyle.Font = New Font(ListFontFamilyComboBox.SelectedText, Val(ListFontSizeComboBox.SelectedText))
                LiteratureTabControl.DetailFont = New Font(DetailFontFamilyComboBox.SelectedText, Val(DetailFontSizeComboBox.SelectedText))
                LiteratureTabControl.HighlightFont = New Font(HighlightFontFamilyComboBox.SelectedText, Val(HighlightFontSizeComboBox.SelectedText))
                LiteratureTabControl.HighlightColor = New _FormMain._LiteratureTabControl.HighlightColor(
                    EntryTypeFontColorComboBox.SelectedColor,
                    BibTeXKeyFontColorComboBox.SelectedColor,
                    TagNameFontColorComboBox.SelectedColor,
                    TagValueFontColorComboBox.SelectedColor)
                LiteratureTabControl.HighlightStyle = New _FormMain._LiteratureTabControl.HighlightStyle(
                    EntryTypeFontColorComboBox.Bold,
                    BibTeXKeyFontColorComboBox.Bold,
                    TagNameFontColorComboBox.Bold,
                    TagValueFontColorComboBox.Bold)
            End Sub

            ''' <summary>
            ''' Override the sub OnResize, when MainPanel's size is changed, this sub will be called.
            ''' Because the FormConfiguration's size is fixed, so this sub just be called twice.
            ''' This sub is firstly called when this MainPanel is created.
            ''' This sub is secondly called when this MainPanel's size is set.
            ''' The Function of this sub is to adjust the sizes of three cover panels.
            ''' At the beginning, I just set the size, location, and the anchor of each cover panels after they are created.
            ''' At this moment, the MainPanel's size is not set, so there are some bugs.
            ''' So, I set the size, location, and the anchor of each cover panels when the MainPanel has been initialized.
            ''' </summary>
            ''' <param name="eventargs"></param>
            Protected Overrides Sub OnResize(eventargs As EventArgs)
                ' When the MainPanel is created, this sub will be called,
                ' but at this time, the CoverPanelList has not been initialized.
                ' So, if the CoverPanelList is nothing, exit this sub.
                If CoverPanelList Is Nothing Then
                    Exit Sub
                End If

                ' If this sub runs here, it means that the CoverPanelList has been initialized,
                ' and it means that the MainPanel has been initialized.
                For Each Cover As TransparentPanel In CoverPanelList
                    ' For each cover panel, set its location and size.
                    With Cover
                        .Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top Or AnchorStyles.Bottom
                        .Height = .Tag.Height + 1
                        .Width = .Tag.Width + 1
                        .Location = New Point(.Tag.Location.X - 1, .Tag.Location.Y - 1)
                    End With
                Next
            End Sub

            ''' <summary>
            ''' This sub is used to fill the input ComboBoxes in the MainPanel according to the configuration file.
            ''' </summary>
            Private Sub InitializeInputComboBoxes()
                ' This data table is used to save the initial values of the ComboBoxes.
                Dim DataTable As New DataTable

                ' Remove the handlers of all ComboBoxes.
                ' So, when their values is changed, the corresponding evnets will not be triggered.
                RemoveHandlerConfigurationChanged()
                ' Read the initial values of ComboBoxed from the configuration.
                If Configuration.GetConfig(TableName.InterfaceFontConfiguration, DataTable) Then
                    ' Fill the ComboBoxes.
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

                ' Re-add the handlers of all ComboBoxes.
                AddHaddlerConfigurationChanged()
            End Sub


            ''' <summary>
            ''' This function is used to convert the values of all ComboBoxes into a DataTable,
            ''' which is easy to be saved.
            ''' </summary>
            ''' <returns>The DataTable which contains the values of all ComboBoxes.</returns>
            Public Function GetConfiguration() As DataTable
                Dim DataTable As New DataTable

                With DataTable
                    .TableName = TableName.InterfaceFontConfiguration
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

            ''' <summary>
            ''' If any ComboBox's value is changed, this sub will be called.
            ''' This sub is used to upadate the preview interface.
            ''' </summary>
            Private Sub ConfigurationChanged(ByVal sender As Object)
                If sender Is GroupFontFamilyComboBox Or sender Is GroupFontSizeComboBox Then
                    GroupTreeView.Font = New Font(GroupFontFamilyComboBox.SelectedText, Val(GroupFontSizeComboBox.SelectedText))
                    Exit Sub
                End If

                If sender Is ListFontFamilyComboBox Or sender Is ListFontSizeComboBox Then
                    DataGridView.DefaultCellStyle.Font = New Font(ListFontFamilyComboBox.SelectedText, Val(ListFontSizeComboBox.SelectedText))
                    Exit Sub
                End If

                If sender Is DetailFontFamilyComboBox Or sender Is DetailFontSizeComboBox Then
                    LiteratureTabControl.DetailFont = New Font(DetailFontFamilyComboBox.SelectedText, Val(DetailFontSizeComboBox.SelectedText))
                    Exit Sub
                End If

                If sender Is HighlightFontFamilyComboBox Or sender Is HighlightFontSizeComboBox Then
                    LiteratureTabControl.HighlightFont = New Font(HighlightFontFamilyComboBox.SelectedText, Val(HighlightFontSizeComboBox.SelectedText))
                    Exit Sub
                End If

                If sender Is EntryTypeFontColorComboBox Or
                   sender Is BibTeXKeyFontColorComboBox Or
                   sender Is TagNameFontColorComboBox Or
                   sender Is TagValueFontColorComboBox Then
                    LiteratureTabControl.HighlightColor = New _FormMain._LiteratureTabControl.HighlightColor(
                        EntryTypeFontColorComboBox.SelectedColor,
                        BibTeXKeyFontColorComboBox.SelectedColor,
                        TagNameFontColorComboBox.SelectedColor,
                        TagValueFontColorComboBox.SelectedColor)
                    LiteratureTabControl.HighlightStyle = New _FormMain._LiteratureTabControl.HighlightStyle(
                        EntryTypeFontColorComboBox.Bold,
                        BibTeXKeyFontColorComboBox.Bold,
                        TagNameFontColorComboBox.Bold,
                        TagValueFontColorComboBox.Bold)
                End If
            End Sub

            ''' <summary>
            ''' This sub is used to initialize the TableLayoutPanel.
            ''' </summary>
            Private Sub InitializeTableLayoutPanel()
                ' Create the TableLayoutPanel.
                TableLayoutPanel = New TableLayoutPanel
                ' Configurate the TableLayoutPanel.
                With TableLayoutPanel
                    .Dock = DockStyle.None
                    .Location = New Point(0, 0)
                    .Size = New Size(TableLayoutPanelWidth, Me.ClientSize.Height - 2)
                    .Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Bottom
                    .ColumnCount = 1
                    .ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, TableLayoutPanelWidth))
                End With

                ' Create all the ComboBoxes, and intialize their names, which are used to identify themselves.
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

                ' Add all the ComboBoxes into an ArrayList,
                ' so we can configurate them in a loop.
                ControlList = New ArrayList
                ControlList.Add(New FormSeparator("Group Font Configuration", TableLayoutPanelWidth))
                ControlList.Add(GroupFontFamilyComboBox)
                ControlList.Add(GroupFontSizeComboBox)
                ControlList.Add(New FormSeparator("List Font Configuration", TableLayoutPanelWidth, MarginBetweenZones))
                ControlList.Add(ListFontFamilyComboBox)
                ControlList.Add(ListFontSizeComboBox)
                ControlList.Add(New FormSeparator("Detail Font Configuration", TableLayoutPanelWidth, MarginBetweenZones))
                ControlList.Add(DetailFontFamilyComboBox)
                ControlList.Add(DetailFontSizeComboBox)
                ControlList.Add(New FormSeparator("Code Font Configuration", TableLayoutPanelWidth, MarginBetweenZones))
                ControlList.Add(HighlightFontFamilyComboBox)
                ControlList.Add(HighlightFontSizeComboBox)
                ControlList.Add(EntryTypeFontColorComboBox)
                ControlList.Add(BibTeXKeyFontColorComboBox)
                ControlList.Add(TagNameFontColorComboBox)
                ControlList.Add(TagValueFontColorComboBox)

                ' For each ComboBox or Label in the ArrayList,
                ' we add it on the TableLayoutPanel,
                ' and bind its event SubControlMouseEnter to the sub InputControls_MouseEnter.
                ' So, if mouse enters these controls, the sub InputControls_MouseEnter will be called.
                Dim Index As Integer = 0
                For Each Control As Control In ControlList
                    TableLayoutPanel.Controls.Add(Control, 0, Index)

                    If TypeOf Control Is FontFamilyComboBox Then
                        AddHandler CType(Control, FontFamilyComboBox).SubControlMouseEnter, AddressOf InputControls_MouseEnter
                    ElseIf TypeOf Control Is FontSizeComboBox Then
                        AddHandler CType(Control, FontSizeComboBox).SubControlMouseEnter, AddressOf InputControls_MouseEnter
                    ElseIf TypeOf Control Is FontColorComboBox Then
                        AddHandler CType(Control, FontColorComboBox).SubControlMouseEnter, AddressOf InputControls_MouseEnter
                    Else
                        AddHandler Control.MouseEnter, AddressOf InputControls_MouseEnter
                    End If

                    Index += 1
                Next

                ' At last, we add the TableLayoutPanel on the MainPanel.
                Me.Controls.Add(TableLayoutPanel)
            End Sub

            ''' <summary>
            ''' When mouse enters any ComboxBox or Label in the TableLayoutPanel, this sub will be called.
            ''' This sub set the Visible attributes of cover panels to emphasize the zones of the preview interface,
            ''' according which control the mouse enters.
            ''' </summary>
            ''' <param name="sender"></param>
            ''' <param name="e"></param>
            Private Sub InputControls_MouseEnter(sender As Object, e As System.EventArgs)
                Select Case sender.Name
                    Case "GroupFontConfiguration", GroupFontFamilyComboBox.Name, GroupFontSizeComboBox.Name
                        CoverPanel4GroupTreeView.Visible = True
                        CoverPanel4DataGridView.Visible = False
                        CoverPanel4LiteratureTabControl.Visible = False
                    Case "ListFontConfiguration", ListFontFamilyComboBox.Name, ListFontSizeComboBox.Name
                        CoverPanel4GroupTreeView.Visible = False
                        CoverPanel4DataGridView.Visible = True
                        CoverPanel4LiteratureTabControl.Visible = False
                    Case "DetailFontConfiguration", DetailFontFamilyComboBox.Name, DetailFontSizeComboBox.Name
                        LiteratureTabControl.SelectedIndex = 0
                        CoverPanel4LiteratureTabControl.Refresh()
                        CoverPanel4GroupTreeView.Visible = False
                        CoverPanel4DataGridView.Visible = False
                        CoverPanel4LiteratureTabControl.Visible = True
                    Case "CodeFontConfiguration", HighlightFontFamilyComboBox.Name,
                     HighlightFontSizeComboBox.Name, EntryTypeFontColorComboBox.Name,
                     BibTeXKeyFontColorComboBox.Name, TagNameFontColorComboBox.Name,
                     TagValueFontColorComboBox.Name
                        LiteratureTabControl.SelectedIndex = 1
                        CoverPanel4LiteratureTabControl.Refresh()
                        CoverPanel4GroupTreeView.Visible = False
                        CoverPanel4DataGridView.Visible = False
                        CoverPanel4LiteratureTabControl.Visible = True
                End Select
            End Sub

            ''' <summary>
            ''' This sub is used to bind the event SelectedChanged of all ComboBoxes to the sub ConfigurationChanged.
            ''' </summary>
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

            ''' <summary>
            ''' This sub is used to unbind the event SelectedChanged of all ComboBoxes from the sub ConfigurationChanged.
            ''' </summary>
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

