Namespace _FormConfiguration
    Namespace LiteratureDetailDisplay
        Public Class MainPanel
            Inherits System.Windows.Forms.Panel

            Dim SelectedItemBuffer As String = ""

            Private Const ListViewWidth As Integer = 150
            Private SelectedColor As Color = Color.FromArgb(51, 153, 255)

            Dim ToolStrip As ToolStrip
            Dim ListView As System.Windows.Forms.ListView

            Dim TextBoxLiteratureType As _FormConfiguration.ToolStripTextBox
            Dim ButtonNew As ToolStripButton
            Dim ButtonDelete As ToolStripButton

            Dim InheritList As ArrayList


            Dim Configuration As _FormConfiguration.Configuration
            Public Modified As Boolean

            Public Sub New(ByVal Configuration As _FormConfiguration.Configuration)
                With Me
                    .Modified = False
                    .Configuration = Configuration
                End With

                InitializeToolStrip()
                InitializeListView()
            End Sub

            Private Sub InitializeToolStrip()
                ToolStrip = New ToolStrip
                With ToolStrip
                    .Dock = DockStyle.None
                    .Size = New Size(ListViewWidth, .Size.Height)
                    .Location = New Point(0, 0)
                    .GripStyle = ToolStripGripStyle.Hidden
                    .RenderMode = ToolStripRenderMode.Professional
                    .BackColor = Me.BackColor
                End With

                TextBoxLiteratureType = New _FormConfiguration.ToolStripTextBox("Literature Type")
                With TextBoxLiteratureType
                    .Size = New Size(93, .Size.Height)
                    AddHandler .TextChanged, AddressOf ToolStripTextBox_TextChanged
                End With

                ToolStrip.Items.Add(TextBoxLiteratureType)
                ToolStrip.Items.Add(New ToolStripSeparator)

                ButtonNew = New ToolStripButton
                With ButtonNew
                    .Name = "ButtonNew"
                    .ToolTipText = "New Literature Type"
                    .Image = Resource.Icon.ToolStripButtonNew
                    .Enabled = False
                    .Visible = True
                    AddHandler .Click, AddressOf ToolStripButton_Click
                End With
                ToolStrip.Items.Add(ButtonNew)

                ButtonDelete = New ToolStripButton
                With ButtonDelete
                    .Name = "ButtonDelete"
                    .ToolTipText = "Delete Literature Type"
                    .Image = Resource.Icon.ToolStripButtonDelete
                    .Enabled = False
                    .Visible = True
                    AddHandler .Click, AddressOf ToolStripButton_Click
                End With
                ToolStrip.Items.Add(ButtonDelete)

                Me.Controls.Add(ToolStrip)
            End Sub

            Private Sub ToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
                Select Case CType(sender, ToolStripButton).Name
                    Case "ButtonNew"
                        ButtonNew_Click()
                    Case "ButtonDelete"
                        ButtonDelete_Click()
                End Select
            End Sub

            Private Sub ButtonNew_Click()
                ' If TextBoxLiteratureType is empty
                If TextBoxLiteratureType.IsEmpty Then
                    ' Exit sub
                    Exit Sub
                End If

                ' If there exist this literature type
                If ExistLiteratureType(TextBoxLiteratureType.Text) Then
                    ' Show the messagebox let the user make a choice 
                    If MsgBox("The literature type """ & TextBoxLiteratureType.Text & """ exists." & vbCr & vbCr & _
                           "Replace the existed literature type?", MsgBoxStyle.OkCancel, "") = MsgBoxResult.Ok Then
                        ' If user want to replace the existing literature type
                        ' Add a new literature type in ListView
                        InheritList.Add(NewLiteratureType(TextBoxLiteratureType.Text).Tag)
                        ' Clear the TextBoxLiteratureType
                        TextBoxLiteratureType.SetText("")
                        ' Exit sub
                        Exit Sub
                    Else
                        ' Exit sub
                        Exit Sub
                    End If
                End If

                ' If there is no this literature type
                ' Add a new literature type in ListView
                InheritList.Add(NewLiteratureType(TextBoxLiteratureType.Text).Tag)
                ' Clear the TextBoxLiteratureType
                TextBoxLiteratureType.SetText("")
            End Sub


            Private Function NewLiteratureType(ByVal LiteratureType As String) As ListViewItem
                For Each ListViewItem As ListViewItem In ListView.Items
                    If ListViewItem.Text.Trim.ToLower = LiteratureType.Trim.ToLower Then
                        ListViewItem.Tag = LiteratureType & "{{{}}}"
                        ListView.SelectedItems.Clear()
                        ListViewItem.Selected = True

                        Return ListViewItem
                    End If
                Next

                Dim NewListViewItem As ListViewItem = ListView.Items.Add(LiteratureType)
                NewListViewItem.Tag = LiteratureType & "{{{}}}"
                ListView.SelectedItems.Clear()
                NewListViewItem.Selected = True

                Return NewListViewItem
            End Function

            Private Function ExistLiteratureType(ByVal LiteratureType As String) As Boolean
                For Each Item As ListViewItem In ListView.Items
                    If Item.Text.Trim.ToLower = LiteratureType.Trim.ToLower Then
                        Return True
                    End If
                Next

                Return False
            End Function

            Private Sub ButtonDelete_Click()
                If ListView.SelectedItems.Count = 0 Then
                    Exit Sub
                End If

                Dim Message As String = ""
                If ListView.SelectedItems.Count = 1 Then
                    Message = "Delete this " & ListView.SelectedItems.Count & " literature type?"
                Else
                    Message = "Delete these " & ListView.SelectedItems.Count & " literature types?"
                End If

                If MsgBox(Message, vbOKCancel, "Delete") = MsgBoxResult.Ok Then
                    For Each ListViewItem As ListViewItem In ListView.SelectedItems
                        ListView.Items.Remove(ListViewItem)
                        InheritList.Remove(ListViewItem.Tag)
                    Next
                    ListView.SelectedItems.Clear()
                    ButtonDelete.Enabled = False
                End If
            End Sub

            Private Sub InitializeListView()
                ListView = New ListView

                With ListView
                    .FullRowSelect = True
                    .View = View.Details
                    .Location = New Point(0, ToolStrip.Height + 4)
                    .Size = New Size(ListViewWidth, Me.ClientSize.Height - ToolStrip.Height - 4)
                    .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Top
                    .GridLines = True
                    .Scrollable = True
                    .MultiSelect = True

                    .Columns.Add("Literature Type")
                    .Columns.Item(0).Width = 120

                    AddHandler ListView.ItemSelectionChanged, AddressOf ListView_ItemSelectionChanged

                End With

                Dim DataTable As New DataTable
                If Not Configuration.GetConfig(TableName.LiteratureDetailDisplayConfiguration, DataTable) Then
                    Exit Sub
                End If

                InheritList = New ArrayList
                For Each Row As DataRow In DataTable.Rows
                    Dim ListViewItem As New ListViewItem
                    ListViewItem.Text = New _FormConfiguration.LiteratureDetailDisplay.Configuration(Row(0)).LiteratureType
                    ListViewItem.Tag = Row(0)
                    InheritList.Add(ListViewItem.Tag)
                    ListView.Items.Add(ListViewItem)
                Next

                If ListView.Items.Count > 0 Then
                    ListView.Items(0).Selected = True
                    ListView.SelectedItems.Clear()
                    ListView.Items(0).Selected = True
                End If

                Me.Controls.Add(ListView)


            End Sub

            Private Sub ListView_ItemSelectionChanged(ByVal sender As Object, ByVal e As ListViewItemSelectionChangedEventArgs)
                If Not e.IsSelected Then
                    Exit Sub
                End If

                For Each ListViewItem As ListViewItem In ListView.Items
                    If (Not ListViewItem.BackColor = SelectedColor) And ListViewItem.Selected Then
                        ListViewItem.ForeColor = Color.White
                        ListViewItem.BackColor = SelectedColor
                    End If

                    If ListViewItem.BackColor = SelectedColor And (Not ListViewItem.Selected) Then
                        ListViewItem.ForeColor = Color.Black
                        ListViewItem.BackColor = Color.White
                    End If
                Next

                ButtonDelete.Enabled = Not ListView.SelectedItems(0).Text = "Default"
                SelectedItemBuffer = ListView.SelectedItems(0).Tag

                Dim LiteratureTypeItem As New _FormConfiguration.LiteratureDetailDisplay.LiteratureTypeItem(SelectedItemBuffer, InheritList)

                With LiteratureTypeItem
                    .Location = New Point(ListViewWidth + 4, 0)
                    .Size = New Size(Me.ClientSize.Width - 4 - ListViewWidth, Me.ClientSize.Height)
                    AddHandler .SaveCode, AddressOf LiteratureTypeItem_SaveCode
                End With

                Me.Controls.Add(LiteratureTypeItem)

                For Each Control As Object In Me.Controls
                    If Control Is LiteratureTypeItem Then
                        Continue For
                    End If

                    If TypeOf Control Is _FormConfiguration.LiteratureDetailDisplay.LiteratureTypeItem Then
                        'CType(Control, _FormConfiguration.LiteratureDetailDisplay.LiteratureTypeItem).ReInitializeTabPageItems(SelectedItemBuffer, InheritList)
                        'Exit Sub
                        Me.Controls.Remove(Control)
                    End If
                Next
            End Sub


            Private Sub LiteratureTypeItem_SaveCode(ByVal Code As String)
                If ListView.SelectedItems.Count = 0 Then
                    Exit Sub
                End If
                ListView.SelectedItems(0).Tag = Code

                InheritList.Clear()
                For Each ListViewItem As ListViewItem In ListView.Items
                    InheritList.Add(ListViewItem.Tag)
                Next
            End Sub

            Private Sub ToolStripTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
                ButtonNew.Enabled = Not TextBoxLiteratureType.IsEmpty
            End Sub

            ''' <summary>
            ''' Get the DataTable of this panel
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function GetConfiguration() As DataTable
                Dim DataTable As New DataTable

                With DataTable
                    .TableName = TableName.LiteratureDetailDisplayConfiguration
                    .Columns.Add("LiteratureDisplayBuffer")
                End With

                For Each ListViewItem As ListViewItem In ListView.Items
                    Dim Row(0) As String
                    Row(0) = ListViewItem.Tag

                    DataTable.Rows.Add(Row)
                Next

                Return DataTable
            End Function
        End Class
    End Namespace
End Namespace
