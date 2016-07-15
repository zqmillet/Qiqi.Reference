Namespace _FormConfiguration
    Namespace LiteratureDetailDisplay
        Public Class DataGridView
            Inherits System.Windows.Forms.DataGridView

            Dim MenuStrip As ContextMenuStrip
            Dim EditingCell As DataGridViewCell

            Public Completed As Boolean
            Public Event ContentChanged()

            Private SelectedColor As Color = Color.FromArgb(51, 153, 255)

            Public Sub New()
                InitializeMenuStrip()

                With Me
                    Dim PropertyNameColumn As New DataGridViewTextBoxColumn

                    With PropertyNameColumn
                        .HeaderText = "Property Name"
                        .Name = "PropertyNameColumn"
                        .Width = 90
                        .Frozen = True
                        .SortMode = DataGridViewColumnSortMode.NotSortable
                        .Resizable = False
                    End With
                    .Columns.Add(PropertyNameColumn)

                    Dim HeightColumn As New DataGridViewTextBoxColumn
                    With HeightColumn
                        .HeaderText = "Height"
                        .Name = "HeightColumn"
                        .Width = 50
                        .Frozen = True
                        .SortMode = DataGridViewColumnSortMode.NotSortable
                        .Resizable = False
                    End With
                    .Columns.Add(HeightColumn)

                    Dim DisplayNameColumn As New DataGridViewTextBoxColumn
                    With DisplayNameColumn
                        .HeaderText = "Display Name"
                        .Name = "DisplayNameColumn"
                        .Width = 85
                        .Frozen = True
                        .SortMode = DataGridViewColumnSortMode.NotSortable
                        .Resizable = False
                    End With
                    .Columns.Add(DisplayNameColumn)

                    Dim DisplayLabelColumn As New _FormConfiguration.DataGridViewDisableCheckBoxColumn
                    With DisplayLabelColumn
                        .HeaderText = "Display Label"
                        .Name = "DisplayLabelColumn"
                        .Width = 90
                        .Frozen = True
                        .SortMode = DataGridViewColumnSortMode.NotSortable
                        .Resizable = False
                    End With
                    .Columns.Add(DisplayLabelColumn)

                    Dim IsRequiredColumn As New DataGridViewCheckBoxColumn
                    With IsRequiredColumn
                        .HeaderText = "Required"
                        .Name = "IsRequiredColumn"
                        .Width = 70
                        .Frozen = True
                        .SortMode = DataGridViewColumnSortMode.NotSortable
                        .Resizable = False
                    End With
                    .Columns.Add(IsRequiredColumn)

                    Dim SyntaxHighlightColumn As New DataGridViewCheckBoxColumn
                    With SyntaxHighlightColumn
                        .HeaderText = "Syntax Highlight"
                        .Name = "SyntaxHighLightColumn"
                        .Width = 110
                        .Frozen = True
                        .SortMode = DataGridViewColumnSortMode.NotSortable
                        .Resizable = False
                    End With
                    .Columns.Add(SyntaxHighlightColumn)

                    .ColumnHeadersHeight = 25
                    .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
                    .RowTemplate.Height = 20
                    .RowHeadersVisible = False

                    For Each Row As DataGridViewRow In Me.Rows
                        Row.Height = 20
                    Next
                    ' .Dock = DockStyle.Fill
                    .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
                    .BorderStyle = Windows.Forms.BorderStyle.None
                    .BackgroundColor = Color.FromArgb(240, 240, 240)
                    .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
                    .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
                    .AllowUserToResizeRows = False
                    .AllowUserToResizeColumns = False

                    AddHandler .EditingControlShowing, AddressOf Me_EditingControlShowing
                    AddHandler .CellValueChanged, AddressOf Me_CellValueChanged
                    AddHandler .CellContentClick, AddressOf Me_CellContentClick
                    AddHandler .LostFocus, AddressOf Me_LostFocus
                    AddHandler .RowsAdded, AddressOf Me_RowsAdded
                    AddHandler .RowsRemoved, AddressOf Me_RowsRemoved
                    AddHandler .CellMouseClick, AddressOf Me_CellMouseClick
                End With

                EnableCheckBox()
            End Sub

            Private Sub Me_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs)
                If Me.SelectedCells.Count = 0 Then
                    Exit Sub
                End If

                If e.Button = Windows.Forms.MouseButtons.Right Then
                    MenuStrip.Show(System.Windows.Forms.Cursor.Position.X, System.Windows.Forms.Cursor.Position.Y)
                End If
            End Sub

            Private Sub MenuStripItem_Clicked(ByVal sender As Object, ByVal e As System.EventArgs)
                Select Case CType(sender, ToolStripMenuItem).Name
                    Case "MenuItemDelete"
                        Dim SelectedCells As New ArrayList
                        For Each Cell As DataGridViewCell In Me.SelectedCells
                            SelectedCells.Add(Cell)
                        Next

                        Dim RowList As New ArrayList
                        For Each Row As DataGridViewRow In Me.Rows
                            For Each Cell As DataGridViewCell In Me.SelectedCells
                                If Row.Cells.Contains(Cell) Then
                                    If Not Row.Index = Me.Rows.Count - 1 Then
                                        Row.DefaultCellStyle.BackColor = SelectedColor
                                        Row.DefaultCellStyle.ForeColor = Color.White
                                        RowList.Add(Row)
                                    End If
                                    Exit For
                                End If
                            Next
                        Next

                        If RowList.Count = 0 Then
                            Exit Sub
                        End If

                        Dim Message As String = ""
                        If RowList.Count = 1 Then
                            Message = "Delete this property?"
                        Else
                            Message = "Delete these " & RowList.Count & " properties?"
                        End If

                        If MsgBox(Message, MsgBoxStyle.OkCancel, "Delete") = MsgBoxResult.Ok Then
                            For Each Row As DataGridViewRow In RowList
                                Me.Rows.Remove(Row)
                            Next
                        Else
                            For Each Cell As DataGridViewCell In SelectedCells
                                Cell.Selected = True
                            Next
                        End If

                        For Each Row As DataGridViewRow In Me.Rows
                            Row.DefaultCellStyle.BackColor = Color.White
                            Row.DefaultCellStyle.ForeColor = Color.Black
                        Next

                        IsCompleted()
                End Select
            End Sub

            Private Sub InitializeMenuStrip()
                MenuStrip = New ContextMenuStrip

                Dim MenuItemDelete As New ToolStripMenuItem
                With MenuItemDelete
                    .Text = "Delete"
                    .Name = "MenuItemDelete"
                    .Image = Resource.Icon.ToolStripButtonDelete
                    AddHandler .Click, AddressOf MenuStripItem_Clicked
                End With
                MenuStrip.Items.Add(MenuItemDelete)
            End Sub

            Private Sub Me_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs)
                IsCompleted()
                EnableCheckBox()
                RaiseEvent ContentChanged()
            End Sub

            Private Sub Me_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs)
                IsCompleted()
                EnableCheckBox()
                CType(Me.Rows(Me.NewRowIndex).Cells("DisplayLabelColumn"), _FormConfiguration.DataGridViewDisableCheckBoxCell).Value = False
                RaiseEvent ContentChanged()
            End Sub

            Private Sub EnableCheckBox()
                If Me.Rows.Count > 2 Then
                    Me.Columns("DisplayLabelColumn").ReadOnly = True
                Else
                    Me.Columns("DisplayLabelColumn").ReadOnly = False
                End If

                For i As Integer = 0 To Me.Rows.Count - 1
                    CType(Me.Rows(i).Cells("DisplayLabelColumn"), _FormConfiguration.DataGridViewDisableCheckBoxCell).Enabled = (Me.Rows.Count > 2)
                    If Me.Rows.Count > 2 Then
                        CType(Me.Rows(i).Cells("DisplayLabelColumn"), _FormConfiguration.DataGridViewDisableCheckBoxCell).Value = True
                    End If
                Next
            End Sub

            Private Sub Me_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs)
                If TypeOf (e.Control) Is DataGridViewTextBoxEditingControl Then
                    With CType(e.Control, DataGridViewTextBoxEditingControl)
                        AddHandler .TextChanged, AddressOf TextBox_TextChanged
                        .Margin = New Padding(0)
                        .Padding = New Padding(0)
                    End With

                End If
            End Sub

            Private Sub TextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

                With CType(sender, DataGridViewTextBoxEditingControl)
                    Me.CurrentCell.Value = .Text
                    IsCompleted()
                    RaiseEvent ContentChanged()

                    If Me.Completed Then
                        .BackColor = Color.White
                    Else
                        .BackColor = Color.Yellow
                    End If
                End With

            End Sub

            Private Sub Me_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
                If Me.Columns(e.ColumnIndex).Name = "DisplayLabelColumn" Then
                    RaiseEvent ContentChanged()
                End If
            End Sub

            Public Sub IsCompleted()
                Dim ExistErrorInRow As Boolean = False
                Completed = True
                For i As Integer = 0 To Me.Rows.Count - 2
                    ExistErrorInRow = False
                    For j As Integer = 0 To Me.Rows(i).Cells.Count - 3
                        If Me.Rows(i).Cells(j).Value Is Nothing Then
                            ExistErrorInRow = True
                            Me.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                            Exit For
                        End If

                        If Me.Rows(i).Cells(j).Value.ToString.Trim = "" Then
                            ExistErrorInRow = True
                            Me.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                            Exit For
                        End If

                        Select Case Me.Columns(j).Name
                            Case "PropertyNameColumn"
                                ExistErrorInRow = ExistErrorInPropertyName(Me.Rows(i).Cells(j).Value.ToString)
                                If ExistErrorInRow Then
                                    Me.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                                    Exit For
                                End If
                            Case "HeightColumn"
                                ExistErrorInRow = ExistErrorInHeight(Me.Rows(i).Cells(j).Value.ToString)
                                If ExistErrorInRow Then
                                    Me.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                                    Exit For
                                End If
                            Case "DisplayNameColumn"
                                ExistErrorInRow = ExistErrorInDisplayName(Me.Rows(i).Cells(j).Value.ToString)
                                If ExistErrorInRow Then
                                    Me.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                                    Exit For
                                End If
                        End Select
                    Next

                    If Not ExistErrorInRow Then
                        Me.Rows(i).DefaultCellStyle.BackColor = Color.White
                    Else
                        Completed = False
                    End If
                Next
            End Sub

            Private Function ExistErrorInPropertyName(ByVal PropertyName As String) As Boolean
                For Each c As Char In PropertyName
                    If ",{}".Contains(c) Then
                        Return True
                    End If
                Next

                Return False
            End Function

            Private Function ExistErrorInDisplayName(ByVal DisplayName As String) As Boolean
                For Each c As Char In DisplayName
                    If ",{}".Contains(c) Then
                        Return True
                    End If
                Next

                Return False
            End Function

            Private Function ExistErrorInHeight(ByVal Height As String) As Boolean
                If Not IsNumeric(Height) Then
                    Return True
                End If

                If Val(Height) < 0 Then
                    Return True
                End If

                Return False
            End Function

            Private Sub Me_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
                IsCompleted()
                RaiseEvent ContentChanged()
            End Sub

            Private Sub Me_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
                Me.ClearSelection()
            End Sub
        End Class
    End Namespace
End Namespace