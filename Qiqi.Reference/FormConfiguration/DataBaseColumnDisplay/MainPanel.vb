Namespace _FormConfiguration
    Namespace DataBaseColumnDisplay
        Public Class MainPanel
            Inherits System.Windows.Forms.Panel

            Private SelectedColor As Color = Color.FromArgb(51, 153, 255)

            Dim ToolStrip As ToolStrip
            Dim TextBoxPropertyName As _FormConfiguration.ToolStripTextBox
            Dim TextBoxDisplayName As _FormConfiguration.ToolStripTextBox
            Dim TextBoxColumnWidth As _FormConfiguration.ToolStripTextBox

            Dim ButtonNew As ToolStripButton
            Dim ButtonDelete As ToolStripButton
            Dim ButtonMoveUp As ToolStripButton
            Dim ButtonMoveDown As ToolStripButton
            Dim ButtonOpen As ToolStripButton
            Dim ButtonSave As ToolStripButton
            Dim ButtonOK As ToolStripButton
            Dim ButtonCancel As ToolStripButton

            Dim ToolStripSeparator1 As ToolStripSeparator
            Dim ToolStripSeparator2 As ToolStripSeparator

            Dim ListView As _FormConfiguration.ListView
            Dim ListViewModifyItem As ListViewItem
            Dim ModifyMode As Boolean = False
            Dim MouseOverColumn As String = ""



            Dim Configuration As _FormConfiguration.Configuration

            ' This variable is used to fix the multiselect bug of list view
            Dim FirstChange As Boolean = True

            Public Modified As Boolean

            ''' <summary>
            ''' Constructor with parameter
            ''' </summary>
            ''' <param name="Configuration">Configuration</param>
            ''' <remarks></remarks>
            Public Sub New(ByVal Configuration As _FormConfiguration.Configuration)
                With Me
                    .Modified = False
                    .Configuration = Configuration
                    ' .BorderStyle = Windows.Forms.BorderStyle.FixedSingle
                End With
                InitializeToolStrip()
                InitializeListView()
            End Sub

            ''' <summary>
            ''' Initialize ToolStrip
            ''' </summary>
            ''' <remarks></remarks>
            Public Sub InitializeToolStrip()
                ' Create a new ToolStrip
                ToolStrip = New ToolStrip
                ' Create two new ToolStripSeparators
                ToolStripSeparator1 = New ToolStripSeparator
                ToolStripSeparator2 = New ToolStripSeparator

                ' Set default style of ToolStrip
                With ToolStrip
                    .GripStyle = ToolStripGripStyle.Hidden
                    .RenderMode = ToolStripRenderMode.Professional
                    .BackColor = Me.BackColor
                End With

                ' Add "Property Name" TextBox
                TextBoxPropertyName = New _FormConfiguration.ToolStripTextBox("Property Name")
                With TextBoxPropertyName
                    .Size = New Size(100, .Height)
                    AddHandler .TextChanged, AddressOf ToolStripTextBox_TextChanged
                End With
                ToolStrip.Items.Add(TextBoxPropertyName)
                ToolStrip.Items.Add(New ToolStripSeparator)

                ' Add "Display Name" TextBox
                TextBoxDisplayName = New _FormConfiguration.ToolStripTextBox("Display Name")
                With TextBoxDisplayName
                    .Size = New Size(100, .Height)
                    AddHandler .TextChanged, AddressOf ToolStripTextBox_TextChanged
                End With
                ToolStrip.Items.Add(TextBoxDisplayName)
                ToolStrip.Items.Add(New ToolStripSeparator)

                ' Add "Column Width" TextBox
                TextBoxColumnWidth = New _FormConfiguration.ToolStripTextBox("Column Width")
                With TextBoxColumnWidth
                    .Size = New Size(100, .Height)
                    AddHandler .TextChanged, AddressOf ToolStripTextBox_TextChanged
                End With
                ToolStrip.Items.Add(TextBoxColumnWidth)
                ToolStrip.Items.Add(New ToolStripSeparator)

                ' Add "New" Button 
                ButtonNew = New ToolStripButton
                With ButtonNew
                    .Name = "ButtonNew"
                    .ToolTipText = "New Property"
                    .Image = Resource.Icon.ToolStripButtonNew
                    .Enabled = False
                    .Visible = True
                    AddHandler .Click, AddressOf ToolStripButton_Click
                End With
                ToolStrip.Items.Add(ButtonNew)

                ' Add "Delete" Button
                ButtonDelete = New ToolStripButton
                With ButtonDelete
                    .Name = "ButtonDelete"
                    .ToolTipText = "Delete Property"
                    .Image = Resource.Icon.ToolStripButtonDelete
                    .Enabled = False
                    .Visible = True
                    AddHandler .Click, AddressOf ToolStripButton_Click
                End With
                ToolStrip.Items.Add(ButtonDelete)
                ToolStrip.Items.Add(ToolStripSeparator1)

                ' Add "Move Up" Button
                ButtonMoveUp = New ToolStripButton
                With ButtonMoveUp
                    .Name = "ButtonMoveUp"
                    .ToolTipText = "Move Up"
                    .Image = Resource.Icon.ToolStripButtonMoveUp
                    .Enabled = False
                    .Visible = True
                    AddHandler .Click, AddressOf ToolStripButton_Click
                End With
                ToolStrip.Items.Add(ButtonMoveUp)

                ' Add "Move Down" Button
                ButtonMoveDown = New ToolStripButton
                With ButtonMoveDown
                    .Name = "ButtonMoveDown"
                    .ToolTipText = "Move Down"
                    .Image = Resource.Icon.ToolStripButtonMoveDown
                    .Enabled = False
                    .Visible = True
                    AddHandler .Click, AddressOf ToolStripButton_Click
                End With
                ToolStrip.Items.Add(ButtonMoveDown)
                ToolStrip.Items.Add(ToolStripSeparator2)

                ' Add "Open" Button
                ButtonOpen = New ToolStripButton
                With ButtonOpen
                    .Name = "ButtonOpen"
                    .ToolTipText = "Open Configuration"
                    .Image = Resource.Icon.ToolStripButtonOpen
                    .Enabled = True
                    .Visible = True
                    AddHandler .Click, AddressOf ToolStripButton_Click
                End With
                ToolStrip.Items.Add(ButtonOpen)

                ' Add "Save" Button
                ButtonSave = New ToolStripButton
                With ButtonSave
                    .Name = "ButtonSave"
                    .ToolTipText = "Save Configuration"
                    .Image = Resource.Icon.ToolStripButtonSave
                    .Enabled = True
                    .Visible = True
                    AddHandler .Click, AddressOf ToolStripButton_Click
                End With
                ToolStrip.Items.Add(ButtonSave)

                ' Add "OK" Button
                ButtonOK = New ToolStripButton
                With ButtonOK
                    .Name = "ButtonOK"
                    .ToolTipText = "OK"
                    .Image = Resource.Icon.ToolStripButtonOK
                    .Enabled = True
                    .Visible = False
                    AddHandler .Click, AddressOf ToolStripButton_Click
                End With
                ToolStrip.Items.Add(ButtonOK)

                ' Add "Cancel" Button
                ButtonCancel = New ToolStripButton
                With ButtonCancel
                    .Name = "ButtonCancel"
                    .ToolTipText = "Cancel"
                    .Image = Resource.Icon.ToolStripButtonCancel
                    .Enabled = True
                    .Visible = False
                    AddHandler .Click, AddressOf ToolStripButton_Click
                End With
                ToolStrip.Items.Add(ButtonCancel)

                ' Add this ToolStrip on the panel
                Me.Controls.Add(ToolStrip)
            End Sub

            ''' <summary>
            ''' If the button is pushed, select a corresponding sub to run
            ''' </summary>
            ''' <param name="sender"></param>
            ''' <param name="e"></param>
            ''' <remarks></remarks>
            Private Sub ToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
                Select Case CType(sender, ToolStripButton).Name
                    Case "ButtonNew"
                        ButtonNew_Click()
                    Case "ButtonDelete"
                        ButtonDelete_Click()
                    Case "ButtonMoveUp"
                        ButtonMoveUp_Click()
                    Case "ButtonMoveDown"
                        ButtonMoveDown_Click()
                    Case "ButtonOpen"
                        ButtonOpen_Click()
                    Case "ButtonSave"
                        ButtonSave_Click()
                    Case "ButtonOK"
                        ButtonOK_Click()
                    Case "ButtonCancel"
                        ButtonCancel_Click()
                End Select
            End Sub

            Private Sub ButtonNew_Click()
                Modified = True

                With ListView
                    Dim Index = .Items.Count
                    .Items.Add("")
                    With .Items(Index)
                        .Checked = True
                        .SubItems.Add(Index + 1)
                        .SubItems.Add(TextBoxPropertyName.Text)
                        .SubItems.Add(TextBoxDisplayName.Text)
                        .SubItems.Add(TextBoxColumnWidth.Text)
                        .Selected = True
                    End With
                End With

                TextBoxPropertyName.SetText("")
                TextBoxDisplayName.SetText("")
                TextBoxColumnWidth.SetText("")

                TextBoxPropertyName.Focus()
            End Sub

            Private Sub ButtonDelete_Click()
                Modified = True

                Dim Message As String = ""
                Dim SelectedItemCount As Integer = ListView.SelectedItems.Count

                If SelectedItemCount = 1 Then
                    Message = "Detele this 1 item?"
                ElseIf SelectedItemCount > 1 Then
                    Message = "Delete these " & SelectedItemCount & " items?"
                End If


                If Not MsgBox(Message, MsgBoxStyle.OkCancel, "Delete") = MsgBoxResult.Ok Then
                    Exit Sub
                End If

                For Each ListViewItem As ListViewItem In ListView.SelectedItems
                    ListView.Items.Remove(ListViewItem)
                Next

                Dim Index As Integer = 0

                For Each ListViewItem As ListViewItem In ListView.Items
                    Index += 1
                    ListViewItem.SubItems(1).Text = Index
                Next
            End Sub

            Private Sub ButtonMoveUp_Click()
                Modified = True

                Dim Index As Integer
                For Each ListViewItem As ListViewItem In ListView.SelectedItems
                    Index = ListView.Items.IndexOf(ListViewItem)
                    ListView.Items.Remove(ListViewItem)
                    ListView.Items.Insert(Index - 1, ListViewItem)
                Next

                Index = 1
                For Each ListViewItem As ListViewItem In ListView.Items
                    ListViewItem.SubItems(1).Text = Index
                    Index += 1
                Next
            End Sub

            Private Sub ButtonMoveDown_Click()
                Modified = True

                Dim Index As Integer
                Dim ListViewItem As New ListViewItem
                Dim Count As Integer = ListView.SelectedItems.Count
                For i As Integer = Count - 1 To 0 Step -1
                    ListViewItem = ListView.SelectedItems(i)
                    Index = ListView.Items.IndexOf(ListViewItem)
                    ListView.Items.Remove(ListViewItem)
                    ListView.Items.Insert(Index + 1, ListViewItem)
                Next

                Index = 1
                For Each ListViewItem In ListView.Items
                    ListViewItem.SubItems(1).Text = Index
                    Index += 1
                Next
            End Sub

            Private Sub ButtonOpen_Click()
                ' Modified = True

            End Sub

            Private Sub ButtonSave_Click()

            End Sub

            Private Sub ButtonOK_Click()
                Modified = True

                With ListViewModifyItem
                    .SubItems(2).Text = TextBoxPropertyName.Text
                    .SubItems(3).Text = TextBoxDisplayName.Text
                    .SubItems(4).Text = TextBoxColumnWidth.Text
                End With

                ButtonCancel_Click()
            End Sub

            Private Sub ButtonCancel_Click()
                ' Set three TextBoxes' text
                TextBoxPropertyName.SetText("")
                TextBoxDisplayName.SetText("")
                TextBoxColumnWidth.SetText("")

                ' Recover the color of modified item
                ListView.SelectedItems.Clear()
                ListView.Enabled = True
                ListViewModifyItem.Selected = True

                ' Show these button
                ButtonNew.Visible = True
                ButtonDelete.Visible = True
                ButtonMoveUp.Visible = True
                ButtonMoveDown.Visible = True
                ButtonSave.Visible = True
                ButtonOpen.Visible = True
                ToolStripSeparator1.Visible = True
                ToolStripSeparator2.Visible = True

                ' Hide these two button
                ButtonOK.Visible = False
                ButtonCancel.Visible = False
            End Sub


            ''' <summary>
            ''' If the texts of three textbox is changed, set the enable of the button
            ''' </summary>
            ''' <param name="sender"></param>
            ''' <param name="e"></param>
            ''' <remarks></remarks>
            Private Sub ToolStripTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
                Dim Enable As Boolean = Not (TextBoxDisplayName.IsEmpty Or TextBoxPropertyName.IsEmpty Or TextBoxColumnWidth.IsEmpty)
                ButtonNew.Enabled = Enable
                ButtonOK.Enabled = Enable
            End Sub

            ''' <summary>
            ''' Initialize ListView
            ''' </summary>
            ''' <remarks></remarks>
            Public Sub InitializeListView()
                ' Create a new ListView
                ListView = New ListView

                ' Set default style of ListView
                With ListView
                    .FullRowSelect = True
                    .View = View.Details
                    .Location = New Point(0, ToolStrip.Size.Height + 4)
                    .Size = New Size(Me.ClientSize.Width, Me.ClientSize.Height - 4 - ToolStrip.Size.Height)
                    .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
                    .GridLines = True
                    .Scrollable = True
                    .MultiSelect = True
                    .CheckBoxes = True
                End With

                ' Set columns of ListView
                With ListView
                    .Columns.Add("Display", 60, HorizontalAlignment.Center)
                    .Columns.Add("Index", 60)
                    .Columns.Add("Property Name", 120)
                    .Columns.Add("Display Name", 120)
                    .Columns.Add("Column Width", 100)
                End With

                ' Fill the ListView
                Dim DataTable As New DataTable
                Dim Index As Integer = 0
                If Configuration.GetConfig(TableName.DataBaseGridViewColunmConfiguration, DataTable) Then
                    For Each Row As DataRow In DataTable.Rows
                        Index += 1
                        ListView.Items.Add("")
                        With ListView.Items(Index - 1)
                            .Checked = Row.Item(0)
                            .SubItems.Add(Index)
                            .SubItems.Add(Row.Item(1))
                            .SubItems.Add(Row.Item(2))
                            .SubItems.Add(Row.Item(3))
                        End With
                    Next
                End If

                ' Add the ListView on the panel
                Me.Controls.Add(ListView)

                ' If the item of ListView is double clicked, run this sub
                AddHandler ListView.DoubleClick, AddressOf ListView_DoubleClick

                AddHandler ListView.ItemSelectionChanged, AddressOf ListView_ItemSelectionChanged
                AddHandler ListView.MouseMove, AddressOf ListView_MouseMoved

                ' This event is used to fix the multiselect bug of list view
                AddHandler ListView.ItemChecked, AddressOf ListView_ItemChecked
            End Sub

            ''' <summary>
            ''' If the selection of item is change, run this sub
            ''' </summary>
            ''' <param name="sender"></param>
            ''' <param name="e"></param>
            ''' <remarks></remarks>
            Private Sub ListView_ItemSelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs)
                For Each ListViewItem As ListViewItem In ListView.Items
                    If ListViewItem.Selected Then
                        ListViewItem.BackColor = SelectedColor
                        ListViewItem.ForeColor = Color.White
                    Else
                        ListViewItem.BackColor = Color.White
                        ListViewItem.ForeColor = Color.Black
                    End If
                Next

                If ListView.SelectedItems.Count = 0 Then
                    ButtonDelete.Enabled = False
                    ButtonMoveDown.Enabled = False
                    ButtonMoveUp.Enabled = False
                    Exit Sub
                Else
                    ButtonDelete.Enabled = True
                End If

                Dim MaxIndex As Integer = 0
                Dim MinIndex As Integer = ListView.Items.Count
                Dim Index As Integer = 0

                For Each ListViewItem As ListViewItem In ListView.SelectedItems
                    Index = ListView.Items.IndexOf(ListViewItem)

                    If MaxIndex < Index Then
                        MaxIndex = Index
                    End If

                    If MinIndex > Index Then
                        MinIndex = Index
                    End If
                Next


                ButtonMoveUp.Enabled = Not (MinIndex = 0)
                ButtonMoveDown.Enabled = Not (MaxIndex = ListView.Items.Count - 1)

            End Sub

            Private Sub ListView_MouseMoved(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
                Dim LeftPosition As Integer = 0
                For Each Column As ColumnHeader In ListView.Columns
                    If e.X < Column.Width + LeftPosition And e.X > LeftPosition Then
                        MouseOverColumn = Column.Text
                    End If
                    LeftPosition += Column.Width
                Next
            End Sub

            ''' <summary>
            ''' If item of listview is double clicked, enter the modify mode
            ''' </summary>
            ''' <param name="sender"></param>
            ''' <param name="e"></param>
            ''' <remarks></remarks>
            Private Sub ListView_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs)
                ' If select nothing, exit sub
                If ListView.SelectedItems.Count = 0 Then
                    Exit Sub
                End If

                ' Record the selected item
                ListViewModifyItem = ListView.SelectedItems(0)

                ' Set colors of items
                For Each item As ListViewItem In ListView.Items
                    item.BackColor = Color.White
                Next
                ListViewModifyItem.BackColor = Color.Red

                ' Enter the modify mode
                ModifyMode = True

                ' Fill there text box
                TextBoxPropertyName.SetText(ListView.SelectedItems(0).SubItems(2).Text)
                TextBoxDisplayName.SetText(ListView.SelectedItems(0).SubItems(3).Text)
                TextBoxColumnWidth.SetText(ListView.SelectedItems(0).SubItems(4).Text)

                ' Set the focus
                Select Case MouseOverColumn
                    Case "Display Name"
                        TextBoxDisplayName.Focus()
                    Case "Column Width"
                        TextBoxColumnWidth.Focus()
                    Case Else
                        TextBoxPropertyName.Focus()
                End Select

                ' Hide these button
                ButtonNew.Visible = False
                ButtonDelete.Visible = False
                ButtonMoveUp.Visible = False
                ButtonMoveDown.Visible = False
                ButtonSave.Visible = False
                ButtonOpen.Visible = False
                ToolStripSeparator1.Visible = False
                ToolStripSeparator2.Visible = False

                ' Show these two button
                ButtonOK.Visible = True
                ButtonCancel.Visible = True

                ' Set the enable of ButtonOK
                ToolStripTextBox_TextChanged(Nothing, Nothing)

                ListView.Enabled = False
            End Sub

            ''' <summary>
            ''' Get the DataTable of this panel
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function GetConfiguration() As DataTable
                Dim DataTable As New DataTable

                With DataTable
                    .TableName = TableName.DataBaseGridViewColunmConfiguration
                    .Columns.Add("Display")
                    .Columns.Add("Property Name")
                    .Columns.Add("Display Name")
                    .Columns.Add("Column Width")
                End With

                For Each ListViewItem As ListViewItem In ListView.Items
                    Dim Row(3) As String
                    Row(0) = ListViewItem.Checked
                    Row(1) = ListViewItem.SubItems(2).Text
                    Row(2) = ListViewItem.SubItems(3).Text
                    Row(3) = ListViewItem.SubItems(4).Text

                    DataTable.Rows.Add(Row)
                Next

                Return DataTable
            End Function

            ''' <summary>
            ''' This event is used to fix the multiselect bug of list view
            ''' </summary>
            ''' <param name="sender"></param>
            ''' <param name="e"></param>
            ''' <remarks></remarks>
            Private Sub ListView_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs)
                If CType(sender, ListView).SelectedItems.Count > 1 Then
                    If (FirstChange) Then
                        FirstChange = False
                        e.Item.Checked = Not e.Item.Checked
                    Else
                        FirstChange = True
                    End If
                End If
            End Sub
        End Class
    End Namespace
End Namespace


