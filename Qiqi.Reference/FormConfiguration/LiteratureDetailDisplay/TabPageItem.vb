Namespace _FormConfiguration
    Namespace LiteratureDetailDisplay
        Public Class TabPageItem
            Inherits System.Windows.Forms.Panel

            Public Checked As Boolean

            Dim CheckBox As CheckBox
            Dim TextBox As _FormConfiguration.TextBox
            Dim DataGridView As _FormConfiguration.LiteratureDetailDisplay.DataGridView
            Dim TabPageConfiguration As _FormConfiguration.LiteratureDetailDisplay.TabPageConfiguration

            Public Completed As Boolean
            Public TabPageName As String


            Public Event CheckedChanged()
            Public Event ContentChanged(ByVal sender As Object)

            Public Sub New(ByVal TabPageConfiguration As _FormConfiguration.LiteratureDetailDisplay.TabPageConfiguration)
                With Me
                    .Padding = New Padding(0)
                    .Margin = New Padding(0)
                    .TabPageConfiguration = TabPageConfiguration
                    .Checked = False
                    .Completed = False
                    .TabPageName = TabPageConfiguration.TabPageName
                End With

                InitializeTextBox()
                InitializeDataGridView()
            End Sub

            Private Sub InitializeDataGridView()
                DataGridView = New _FormConfiguration.LiteratureDetailDisplay.DataGridView
                With DataGridView
                    .Location = New Point(0, 16)
                    .Size = New Size(Me.ClientSize.Width, Me.ClientSize.Height - 20)
                    .SelectionMode = DataGridViewSelectionMode.CellSelect
                    AddHandler .RowsAdded, AddressOf DataGridView_RowsAdded
                    AddHandler .RowsRemoved, AddressOf DataGridView_RowsRemoved
                    AddHandler .ContentChanged, AddressOf DataGridView_ContentChanged
                End With
                Me.Controls.Add(DataGridView)

                If TabPageConfiguration.Properties.Count = 0 Then
                    DataGridView.Rows.Add()
                Else
                    For Each PropertyConfiguration As _FormConfiguration.LiteratureDetailDisplay.PropertyConfiguration In TabPageConfiguration.Properties
                        Dim Index As Integer = DataGridView.Rows.Add()
                        With DataGridView.Rows(Index)
                            For Each ColumnHeader As DataGridViewColumn In DataGridView.Columns
                                Select Case ColumnHeader.Name.ToLower
                                    Case "PropertyNameColumn".ToLower
                                        .Cells(ColumnHeader.Name).Value = CType(TabPageConfiguration.Properties.Item(Index), PropertyConfiguration).PropertyName
                                    Case "HeightColumn".ToLower
                                        .Cells(ColumnHeader.Name).Value = CType(TabPageConfiguration.Properties.Item(Index), PropertyConfiguration).Height
                                    Case "DisplayNameColumn".ToLower
                                        .Cells(ColumnHeader.Name).Value = CType(TabPageConfiguration.Properties.Item(Index), PropertyConfiguration).DisplayName
                                    Case "DisplayLabelColumn".ToLower
                                        .Cells(ColumnHeader.Name).Value = CType(TabPageConfiguration.Properties.Item(Index), PropertyConfiguration).DisplayLabel
                                    Case "IsRequiredColumn".ToLower
                                        .Cells(ColumnHeader.Name).Value = CType(TabPageConfiguration.Properties.Item(Index), PropertyConfiguration).IsRequired
                                    Case "SyntaxHighlightColumn".ToLower
                                        .Cells(ColumnHeader.Name).Value = CType(TabPageConfiguration.Properties.Item(Index), PropertyConfiguration).SyntaxHighlight
                                End Select
                            Next
                        End With
                    Next
                End If



                SetHeight()
                DataGridView.ClearSelection()
            End Sub

            Private Sub InitializeTextBox()
                CheckBox = New CheckBox
                With CheckBox
                    .Checked = False
                    .Text = ""
                    .Location = New Point(0, 0)
                    .Padding = New Padding(0)
                    .Margin = New Padding(0)
                    .Size = New Size(14, 14)
                    AddHandler .CheckedChanged, AddressOf CheckBox_CheckedChanged
                End With
                Me.Controls.Add(CheckBox)

                TextBox = New _FormConfiguration.TextBox
                With TextBox
                    .Text = TabPageConfiguration.TabPageName
                    .Location = New Point(18, 0)
                    .TextAlign = HorizontalAlignment.Left
                    .BorderStyle = Windows.Forms.BorderStyle.None
                    .Size = New Size(.Size.Width, 30)
                    AddHandler .TextChanged, AddressOf TextBox_TextChanged
                End With
                Me.Controls.Add(TextBox)
            End Sub

            Private Sub DataGridView_ContentChanged()
                Me.Completed = IsCompleted()
                RaiseEvent ContentChanged(Me)
            End Sub

            Private Sub DataGridView_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs)
                SetHeight()
            End Sub

            Private Sub DataGridView_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs)
                SetHeight()
            End Sub

            Private Sub SetHeight()
                Me.Size = New Size(Me.Width, 30 + 20 * (1 + DataGridView.Rows.Count))
            End Sub

            Private Sub TextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
                Me.TabPageName = TextBox.Text
                Completed = IsCompleted()
                RaiseEvent ContentChanged(Me)
            End Sub

            Private Function IsCompleted() As Boolean
                If TextBox.Text.Trim = "" Then
                    Return False
                End If
                Return DataGridView.Completed
            End Function

            Private Sub CheckBox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
                Me.Checked = CheckBox.Checked
                RaiseEvent CheckedChanged()
            End Sub

            Public Function GenerateCode() As String
                Dim Code As String = ""

                For i As Integer = 0 To DataGridView.Rows.Count - 2
                    Code &= DataGridView.Rows(i).Cells("PropertyNameColumn").Value & "{"
                    Code &= DataGridView.Rows(i).Cells("DisplayNameColumn").Value & ","
                    Code &= CType(DataGridView.Rows(i).Cells("DisplayLabelColumn"), DataGridViewCheckBoxCell).EditedFormattedValue & ","
                    Code &= DataGridView.Rows(i).Cells("HeightColumn").Value & ","
                    Code &= CType(DataGridView.Rows(i).Cells("IsRequiredColumn"), DataGridViewCheckBoxCell).EditedFormattedValue & ","
                    Code &= CType(DataGridView.Rows(i).Cells("SyntaxHighlightColumn"), DataGridViewCheckBoxCell).EditedFormattedValue
                    Code &= "},"
                Next

                Code = Code.Remove(Code.LastIndexOf(","))
                Return Code
            End Function
        End Class
    End Namespace
End Namespace

