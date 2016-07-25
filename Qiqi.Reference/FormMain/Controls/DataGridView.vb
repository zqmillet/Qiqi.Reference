Namespace _FormMain
    Public Class DataGridView
        Inherits System.Windows.Forms.DataGridView

        Dim BibTeXFullName As String
        Dim DataTableColumnConfiguration As DataTable
        Dim MenuStrip As ContextMenuStrip

        Public Event ProgressUpdate(ByVal Progress As Double)

        Public Delegate Sub DelegateShowDataBase()

        Public Event ColumnDisplayChanged(ByVal sender As Object)

        Public DataBase As _BibTeX.DataBase

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New(ByVal Configuration As _FormConfiguration.Configuration)
            ' Style Configuration
            StyleConfiguration()

            If Not Configuration.GetConfig(TableName.DataBaseGridViewColunmConfiguration, Me.DataTableColumnConfiguration) Then
                Exit Sub
            End If

            With Me
                .Dock = DockStyle.Fill
                .MenuStrip = New ContextMenuStrip
            End With

            ' Column configuration
            ColumnConfiguration()

            AddHandler Me.RowPostPaint, AddressOf Me_RowPostPaint
            AddHandler Me.MouseMove, AddressOf Me_MouseMove
            AddHandler Me.ColumnWidthChanged, AddressOf Me_ColumnWidthChanged
            AddHandler Me.ColumnDisplayIndexChanged, AddressOf Me_ColumnDisplayIndexChanged
            AddHandler Me.ColumnHeaderMouseClick, AddressOf Me_ColumnHeaderMouseClick
            AddHandler MenuStrip.ItemClicked, AddressOf MenuStrip_ItemClicked
        End Sub

        Private Sub DataBase_ProgressUpdate(ByVal Progress As Double)
            RaiseEvent ProgressUpdate(Progress)
        End Sub

        Public Function DataBaseLoading(ByVal BibTeXFullName As String, ByRef ErrorMessage As _BibTeX.ErrorMessage) As Boolean
            ' If there is no BibTeX file, exit sub
            If Not My.Computer.FileSystem.FileExists(BibTeXFullName) Then
                Return False
            End If

            Me.BibTeXFullName = BibTeXFullName

            ' Create a database of BibTeX
            DataBase = New _BibTeX.DataBase(BibTeXFullName)

            ' Add progress update event
            AddHandler DataBase.ProgressUpdate, AddressOf DataBase_ProgressUpdate

            ' If sub DataBaseLoading is not successful, show the ErrorMessage, and exit sub
            If Not DataBase.DataBaseLoading(ErrorMessage) Then
                ' MsgBox(ErrorMessage)
                Return False
            End If

            If Me.InvokeRequired Then
                Me.Invoke(New DelegateShowDataBase(AddressOf ShowDataBase))
            Else
                ShowDataBase()
            End If

            Return True
        End Function

        Private Sub ShowDataBase()
            ' Clear the GridView
            Me.Rows.Clear()
            ' Show the DataBase in GridView
            For Each Literature As _BibTeX.Literature In DataBase.LiteratureList
                Dim Index As Integer = Me.Rows.Add()
                With Me.Rows(Index)
                    .Tag = Literature
                    .Height = 20
                    .ReadOnly = True

                    For Each Column As DataGridViewColumn In Me.Columns
                        If TypeOf Column Is DataGridViewImageColumn Then
                            Continue For
                        End If
                        .Cells(Column.Name).Value = Literature.GetPropertyValue(Column.Name)
                    Next

                    If Not ExistColumn("ExistURL") Then
                        Continue For
                    End If

                    If Literature.ExistURL() Then
                        .Cells("ExistURL").Value = Resource.Icon.ExistURL
                    Else
                        .Cells("ExistURL").Value = Resource.Icon.NotExistURL
                    End If

                    If Not ExistColumn("ExistFile") Then
                        Continue For
                    End If

                    If Literature.ExistFile() Then
                        .Cells("ExistFile").Value = Resource.Icon.ExistFile
                    Else
                        .Cells("ExistFile").Value = Resource.Icon.NotExistFile
                    End If
                End With
            Next
        End Sub

        Private Sub StyleConfiguration()
            With Me
                ' Fill the parent
                .Dock = DockStyle.Fill
                ' Set the background color
                .BackgroundColor = Color.FromArgb(240, 240, 240)
                ' Set the style of border
                .BorderStyle = BorderStyle.None
                ' Not allow user to modify
                .AllowUserToAddRows = False
                .ReadOnly = True
                ' Allow user to order columns
                .AllowUserToOrderColumns = True
                ' Make row header visible
                .RowHeadersVisible = True
                ' Remove all column
                .Columns.Clear()
                ' Set the border style of column header
                .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
                ' Not allow user to modify the height of colunm
                .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
                ' Not allow user to modify the height of row
                .AllowUserToResizeRows = False

                .RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            End With
        End Sub

        Private Sub ColumnConfiguration()
            ' Config the column
            With Me
                .Columns.Clear()
                .ColumnCount = 0

                ' The first column shows whether does the file exist
                Dim ImageColumnFileExist As New DataGridViewImageColumn
                With ImageColumnFileExist
                    .HeaderText = ""
                    .Name = "ExistFile"
                    .ImageLayout = DataGridViewImageCellLayout.Zoom
                    .Width = 20
                    .Frozen = True
                    .Resizable = False
                End With
                .Columns.Add(ImageColumnFileExist)

                ' The second column shows whether does the URL exist
                Dim ImageColumnURLExist As New DataGridViewImageColumn
                With ImageColumnURLExist
                    .HeaderText = ""
                    .Name = "ExistURL"
                    .ImageLayout = DataGridViewImageCellLayout.Zoom
                    .Width = 20
                    .Frozen = True
                    .Resizable = False
                End With
                .Columns.Add(ImageColumnURLExist)

                ' The rest of columns
                For index As Integer = 0 To DataTableColumnConfiguration.Rows.Count - 1
                    .ColumnCount += 1
                    With .Columns(.ColumnCount - 1)
                        .Visible = DataTableColumnConfiguration.Rows(index).Item(0)
                        .Name = DataTableColumnConfiguration.Rows(index).Item(1)
                        .HeaderText = DataTableColumnConfiguration.Rows(index).Item(2)
                        .Width = DataTableColumnConfiguration.Rows(index).Item(3)
                    End With
                Next
            End With
        End Sub

        Public Sub ResetColumn()
            ' Config the column
            With Me
                .Columns.Clear()
                .ColumnCount = 2
                .RowHeadersVisible = False

                With .Columns(0)
                    .Visible = True
                    .Name = "EntryType"
                    .HeaderText = "Type"
                    .Width = 60
                End With

                With .Columns(1)
                    .Visible = True
                    .Name = "Title"
                    .HeaderText = "Title"
                    .Width = 230
                End With
            End With
        End Sub

        Private Sub Me_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs)
            If e.Button = Windows.Forms.MouseButtons.Right Then
                With MenuStrip
                    .Items.Clear()
                    For Each Row As DataRow In DataTableColumnConfiguration.Rows
                        Dim MenuItem As New ToolStripMenuItem
                        With MenuItem
                            .Text = Row("Display Name")
                            .Checked = Row("Display")
                            .ToolTipText = Row("Property Name")
                            .Tag = Me.Columns(Row("Property Name"))
                        End With
                        .Items.Add(MenuItem)
                    Next
                    .Show(System.Windows.Forms.Cursor.Position.X, System.Windows.Forms.Cursor.Position.Y)
                End With
            End If
        End Sub

        Private Sub MenuStrip_ItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs)
            CType(e.ClickedItem.Tag, DataGridViewColumn).Visible = Not CType(e.ClickedItem, ToolStripMenuItem).Checked
            DataTableColumnConfiguration = GetColumnDisplayConfiguration()
            RaiseEvent ColumnDisplayChanged(Me)
        End Sub

        Private Sub Me_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs)
            If Not Me.RowHeadersVisible Then
                Exit Sub
            End If

            Dim Rectangle As New System.Drawing.Rectangle(e.RowBounds.Location.X,
                                                          e.RowBounds.Location.Y,
                                                          Me.RowHeadersWidth - 4,
                                                          e.RowBounds.Height)
            TextRenderer.DrawText(e.Graphics, _
                                  (e.RowIndex + 1).ToString, _
                                  Me.RowHeadersDefaultCellStyle.Font, _
                                  Rectangle, Me.RowHeadersDefaultCellStyle.ForeColor, _
                                  TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
        End Sub


        Private Sub Me_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
            Dim MinLocationX As Integer = 0
            Dim MaxLocationX As Integer = Me.RowHeadersWidth + 7
            For Each Column As Object In Me.Columns
                If TypeOf Column Is DataGridViewImageColumn Then
                    MaxLocationX += CType(Column, DataGridViewImageColumn).Width
                End If
            Next

            If e.X < MaxLocationX Then
                Me.AllowUserToResizeColumns = False
            Else
                Me.AllowUserToResizeColumns = True
            End If
        End Sub

        Private Sub Me_ColumnWidthChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs)
            DataTableColumnConfiguration = GetColumnDisplayConfiguration()
            RaiseEvent ColumnDisplayChanged(Me)
        End Sub

        Private Sub Me_ColumnDisplayIndexChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs)
            DataTableColumnConfiguration = GetColumnDisplayConfiguration()
            RaiseEvent ColumnDisplayChanged(Me)
        End Sub

        Public Function GetColumnDisplayConfiguration() As DataTable
            Dim DataTable As New DataTable

            With DataTable
                .TableName = TableName.DataBaseGridViewColunmConfiguration
                .Columns.Add("Display")
                .Columns.Add("Property Name")
                .Columns.Add("Display Name")
                .Columns.Add("Column Width")
            End With

            For Index As Integer = 0 To Me.Columns.Count - 1
                For Each Column As DataGridViewColumn In Me.Columns
                    If TypeOf Me.Columns(Index) Is DataGridViewImageColumn Then
                        Continue For
                    End If

                    If Column.DisplayIndex = Index Then
                        Dim Row(3) As String
                        Row(0) = Column.Visible.ToString
                        Row(1) = Column.Name
                        Row(2) = Column.HeaderText
                        Row(3) = Column.Width

                        DataTable.Rows.Add(Row)
                    End If
                Next
            Next

            Return DataTable
        End Function

        Private Function ExistColumn(ByVal ColumnName As String) As Boolean
            For Each Column As DataGridViewColumn In Me.Columns
                If Column.Name.Trim.ToLower = ColumnName.Trim.ToLower Then
                    Return True
                End If
            Next

            Return False
        End Function

        ''' <summary>
        ''' Refresh the columns of DataGridView according to the configuration
        ''' </summary>
        ''' <param name="DataTableColumnConfiguration"></param>
        ''' <remarks></remarks>
        Public Sub DisplayRefresh(ByVal DataTableColumnConfiguration As DataTable)
            ' Update the configuration
            Me.DataTableColumnConfiguration = DataTableColumnConfiguration

            ' Display index of column
            Dim DisplayIndex As Integer = 2
            ' Read the configuration, if there is no corresponding columns in DataGridView, add a new column into DataGridView
            For Each Row As DataRow In DataTableColumnConfiguration.Rows
                ' If there exists this property name in configuration
                If ExistColumn(Row("Property Name")) Then
                    ' It needn't to create a new column
                    With CType(Me.Columns(Row("Property Name")), DataGridViewColumn)
                        .Visible = Row("Display")
                        .Name = Row("Property Name")
                        .HeaderText = Row("Display Name")
                        .Width = Row("Column Width")
                        .DisplayIndex = DisplayIndex
                    End With
                Else
                    ' Else, create a new column and add it into DataGridView
                    Dim NewColumn As New DataGridViewTextBoxColumn
                    With NewColumn
                        .Visible = Row("Display")
                        .Name = Row("Property Name")
                        .HeaderText = Row("Display Name")
                        .Width = Row("Column Width")
                        .DisplayIndex = DisplayIndex
                    End With
                    Me.Columns.Add(NewColumn)

                    ' Then read the value of each row to fill the new column
                    For Each DataGridViewRow As DataGridViewRow In Me.Rows
                        DataGridViewRow.Cells(Row("Property Name")).Value _
                            = CType(DataGridViewRow.Tag, _BibTeX.Literature).GetPropertyValue(Row("Property Name"))
                    Next
                End If

                DisplayIndex += 1
            Next

            ' Read the column of DataGridView, if there exist a column which does not exist in configuration, remove it from the DataGridView
            Dim RemoveColumnList As New ArrayList
            For Each Column As DataGridViewColumn In Me.Columns
                If TypeOf Column Is DataGridViewImageColumn Then
                    Continue For
                End If

                Dim ExistColumn As Boolean = False
                For Each Row As DataRow In DataTableColumnConfiguration.Rows
                    If Column.Name = Row("Property Name") Then
                        ExistColumn = True
                        Exit For
                    End If
                Next

                If Not ExistColumn Then
                    RemoveColumnList.Add(Column)
                End If
            Next

            For Each Column As DataGridViewColumn In RemoveColumnList
                Me.Columns.Remove(Column)
            Next
        End Sub
    End Class
End Namespace


