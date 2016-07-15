Namespace _FormConfiguration
    Namespace LiteratureDetailDisplay
        Public Class LiteratureTypeItem
            Inherits System.Windows.Forms.Panel
            Public LiteratureType As String

            Dim Configuration As _FormConfiguration.LiteratureDetailDisplay.Configuration

            Dim ToolStrip As ToolStrip
            Dim TextBoxTabPageLabel As _FormConfiguration.ToolStripTextBox

            Dim FlowLayoutPanel As FlowLayoutPanel

            Dim ButtonNew As ToolStripButton
            Dim ButtonDelete As ToolStripButton
            Dim ButtonMoveUp As ToolStripButton
            Dim ButtonMoveDown As ToolStripButton
            'Dim ButtonSave As ToolStripButton
            Dim ComboBoxInherit As ToolStripComboBox

            Dim InheritList As ArrayList

            Public Event SaveCode(ByVal Code As String)

            Public Sub New(ByVal Buffer As String, ByVal InheritList As ArrayList)
                With Me
                    .Configuration = New _FormConfiguration.LiteratureDetailDisplay.Configuration(Buffer)
                    .LiteratureType = Configuration.LiteratureType
                    .Anchor = AnchorStyles.Top Or AnchorStyles.Right Or AnchorStyles.Bottom Or AnchorStyles.Left
                    .InheritList = InheritList
                End With

                InitializeToolStrip()
                InitializeTabPageItems()
            End Sub

            Public Sub ReInitializeTabPageItems(ByVal Buffer As String, ByVal InheritList As ArrayList)
                With Me
                    .Configuration = New _FormConfiguration.LiteratureDetailDisplay.Configuration(Buffer)
                    .LiteratureType = Configuration.LiteratureType
                    .Anchor = AnchorStyles.Top Or AnchorStyles.Right Or AnchorStyles.Bottom Or AnchorStyles.Left
                    .InheritList = InheritList
                End With

                ReInitializeTabPageItems(ComboBoxInherit.Text)
            End Sub

            Public Sub ReInitializeTabPageItems(ByVal ComboBoxText As String)
                If Not ComboBoxInherit Is Nothing Then
                    ComboBoxInherit.Items.Clear()
                    For Each LiteratureBuffer As String In InheritList
                        Dim LiteratureDisplayConfiguration As New _FormConfiguration.LiteratureDetailDisplay.Configuration(LiteratureBuffer)
                        If Not LiteratureDisplayConfiguration.LiteratureType = Me.LiteratureType Then
                            ComboBoxInherit.Items.Add(LiteratureDisplayConfiguration.LiteratureType)
                        End If
                    Next
                End If

                RemoveHandler ComboBoxInherit.SelectedIndexChanged, AddressOf ComboBoxInherit_SelectedIndexChanged
                ComboBoxInherit.Text = ComboBoxText
                AddHandler ComboBoxInherit.SelectedIndexChanged, AddressOf ComboBoxInherit_SelectedIndexChanged

                With FlowLayoutPanel
                    .Controls.Clear()
                    .FlowDirection = FlowDirection.TopDown
                    .WrapContents = False
                    .AutoScroll = True
                    .HorizontalScroll.Visible = False
                    .Margin = New Padding(0)

                    For Each TabPageConfiguration As _FormConfiguration.LiteratureDetailDisplay.TabPageConfiguration In Configuration.TabPages
                        Dim TabPageItem As New TabPageItem(TabPageConfiguration)

                        With TabPageItem
                            .Size = New Size(FlowLayoutPanel.ClientSize.Width - 20, .Size.Height)
                            ' .Anchor = AnchorStyles.Left Or AnchorStyles.Right
                            AddHandler .CheckedChanged, AddressOf TabPageItem_CheckedChanged
                            AddHandler .ContentChanged, AddressOf TabPageItem_ContentChanged
                        End With
                        .Controls.Add(TabPageItem)
                    Next

                    ' .BorderStyle = Windows.Forms.BorderStyle.FixedSingle
                    .Location = New Point(0, ToolStrip.Height + 4)
                    .Size = New Size(Me.Width, Me.Height - ToolStrip.Height - 4)
                    .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
                    AddHandler .Resize, AddressOf FlowLayoutPanel_Resize
                End With

                Me.Controls.Add(FlowLayoutPanel)
                ' Set enable of ButtonSave
                TabPageItem_ContentChanged(Nothing)
            End Sub

            Private Sub InitializeTabPageItems()
                FlowLayoutPanel = New FlowLayoutPanel

                With FlowLayoutPanel
                    .FlowDirection = FlowDirection.TopDown
                    .WrapContents = False
                    .AutoScroll = True
                    .HorizontalScroll.Visible = False


                    .Margin = New Padding(0)
                    For Each TabPageConfiguration As _FormConfiguration.LiteratureDetailDisplay.TabPageConfiguration In Configuration.TabPages
                        Dim TabPageItem As New TabPageItem(TabPageConfiguration)

                        With TabPageItem
                            .Size = New Size(FlowLayoutPanel.ClientSize.Width - 20, .Size.Height)
                            ' .Anchor = AnchorStyles.Left Or AnchorStyles.Right
                            AddHandler .CheckedChanged, AddressOf TabPageItem_CheckedChanged
                            AddHandler .ContentChanged, AddressOf TabPageItem_ContentChanged
                        End With
                        .Controls.Add(TabPageItem)
                    Next

                    ' .BorderStyle = Windows.Forms.BorderStyle.FixedSingle
                    .Location = New Point(0, ToolStrip.Height + 4)
                    .Size = New Size(Me.Width, Me.Height - ToolStrip.Height - 4)
                    .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
                    AddHandler .Resize, AddressOf FlowLayoutPanel_Resize
                End With

                Me.Controls.Add(FlowLayoutPanel)
                ' Set enable of ButtonSave
                TabPageItem_ContentChanged(Nothing)
            End Sub

            Private Sub TabPageItem_CheckedChanged()
                Dim DeleteEnable As Boolean = False
                Dim MoveDownEnable As Boolean = False
                Dim MoveUpEnable As Boolean = False

                Dim MaxIndex As Integer = 0
                Dim MinIndex As Integer = Integer.MaxValue
                Dim Index As Integer = 0
                Dim CheckCount As Integer = 0
                Dim TabPageCount As Integer = 0
                For Each Control As Object In FlowLayoutPanel.Controls
                    If TypeOf (Control) Is TabPageItem Then
                        TabPageCount += 1
                        DeleteEnable = DeleteEnable Or CType(Control, TabPageItem).Checked

                        If CType(Control, TabPageItem).Checked Then
                            CheckCount += 1
                            If MaxIndex < Index Then
                                MaxIndex = Index
                            End If

                            If MinIndex > Index Then
                                MinIndex = Index
                            End If
                        End If

                        Index += 1
                    End If
                Next

                If CheckCount = 0 Then
                    ButtonDelete.Enabled = False
                    ButtonMoveUp.Enabled = False
                    ButtonMoveDown.Enabled = False
                Else
                    ButtonDelete.Enabled = DeleteEnable
                    ButtonMoveUp.Enabled = Not (MinIndex = 0)
                    ButtonMoveDown.Enabled = Not (MaxIndex = TabPageCount - 1)
                End If

            End Sub

            Private Sub TabPageItem_ContentChanged(ByVal sender As Object)
                Dim TabPageItemCount As Integer = 0
                For Each Control As Object In FlowLayoutPanel.Controls
                    If TypeOf Control Is TabPageItem Then
                        TabPageItemCount += 1
                    End If
                Next

                If TabPageItemCount = 0 Then
                    ' ButtonSave.Enabled = False
                    Exit Sub
                End If

                ' ButtonSave.Enabled = True
                For Each Control As Object In FlowLayoutPanel.Controls
                    If TypeOf (Control) Is _FormConfiguration.LiteratureDetailDisplay.TabPageItem Then
                        If Not CType(Control, _FormConfiguration.LiteratureDetailDisplay.TabPageItem).Completed Then
                            ' ButtonSave.Enabled = False
                            Exit Sub
                        End If
                    End If
                Next

                RaiseEvent SaveCode(GenerateCode)
            End Sub

            Private Sub FlowLayoutPanel_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs)
                For Each TabPageItem As Object In FlowLayoutPanel.Controls
                    If TypeOf (TabPageItem) Is _FormConfiguration.LiteratureDetailDisplay.TabPageItem Then
                        With CType(TabPageItem, _FormConfiguration.LiteratureDetailDisplay.TabPageItem)
                            .Size = New Size(FlowLayoutPanel.ClientSize.Width - 20, .Size.Height)
                        End With
                    End If
                Next

            End Sub

            Private Sub InitializeToolStrip()
                ToolStrip = New ToolStrip
                With ToolStrip
                    .GripStyle = ToolStripGripStyle.Hidden
                    .RenderMode = ToolStripRenderMode.Professional
                    .BackColor = Me.BackColor
                End With

                TextBoxTabPageLabel = New _FormConfiguration.ToolStripTextBox("Tab Page Label")
                With TextBoxTabPageLabel
                    .Size = New Size(100, .Size.Height)
                    AddHandler .TextChanged, AddressOf TextBoxTabPageLabel_TextChanged
                End With
                ToolStrip.Items.Add(TextBoxTabPageLabel)


                ' Add "New" Button 
                ButtonNew = New ToolStripButton
                With ButtonNew
                    .Name = "ButtonNew"
                    .ToolTipText = "New TabPage"
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
                    .ToolTipText = "Delete TabPage"
                    .Image = Resource.Icon.ToolStripButtonDelete
                    .Enabled = False
                    .Visible = True
                    AddHandler .Click, AddressOf ToolStripButton_Click
                End With
                ToolStrip.Items.Add(ButtonDelete)
                ToolStrip.Items.Add(New ToolStripSeparator)

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
                ToolStrip.Items.Add(New ToolStripSeparator)

                ' Add "Save" Button
                'ButtonSave = New ToolStripButton
                'With ButtonSave
                '    .Name = "ButtonSave"
                '    .ToolTipText = "Save"
                '    .Image = Resource.Icon.ToolStripButtonSave
                '    .Visible = True
                '    .Enabled = False
                '    AddHandler .Click, AddressOf ToolStripButton_Click
                'End With
                'ToolStrip.Items.Add(ButtonSave)
                'ToolStrip.Items.Add(New ToolStripSeparator)

                ToolStrip.Items.Add(New ToolStripLabel("Inherit"))
                ComboBoxInherit = New ToolStripComboBox
                With ComboBoxInherit
                    .Name = "ComboBoxInherit"
                    .Size = New Size(150, 14)
                    .FlatStyle = FlatStyle.Standard
                    .DropDownStyle = ComboBoxStyle.DropDownList

                    For Each Buffer As String In InheritList
                        Dim LiteratureDisplayConfiguration As New _FormConfiguration.LiteratureDetailDisplay.Configuration(Buffer)
                        If Not LiteratureDisplayConfiguration.LiteratureType = Me.LiteratureType Then
                            .Items.Add(LiteratureDisplayConfiguration.LiteratureType)
                        End If
                    Next
                    AddHandler .SelectedIndexChanged, AddressOf ComboBoxInherit_SelectedIndexChanged
                End With
                ToolStrip.Items.Add(ComboBoxInherit)

                Me.Controls.Add(ToolStrip)
            End Sub

            Private Sub ComboBoxInherit_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
                With CType(sender, ToolStripComboBox)
                    For Each Buffer As String In InheritList
                        Dim LiteratureDisplayConfiguration As New _FormConfiguration.LiteratureDetailDisplay.Configuration(Buffer)
                        If .Text = LiteratureDisplayConfiguration.LiteratureType Then
                            Me.Configuration = LiteratureDisplayConfiguration
                            ReInitializeTabPageItems(ComboBoxInherit.Text)
                            RaiseEvent SaveCode(GenerateCode)
                            Exit Sub
                        End If
                    Next
                End With

            End Sub

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
                        'Case "ButtonSave"
                        '    ButtonSave_Click()
                End Select
            End Sub

            Private Sub ButtonNew_Click()
                ' Create a new TabPageItem
                Dim TabPageItem As New _FormConfiguration.LiteratureDetailDisplay.TabPageItem(New _FormConfiguration.LiteratureDetailDisplay.TabPageConfiguration(TextBoxTabPageLabel.Text))
                ' Add events
                With TabPageItem
                    AddHandler .CheckedChanged, AddressOf TabPageItem_CheckedChanged
                    AddHandler .ContentChanged, AddressOf TabPageItem_ContentChanged
                End With
                ' Clear the TextBox
                TextBoxTabPageLabel.SetText("")
                ' Add this TabPageItem into the panel
                FlowLayoutPanel.Controls.Add(TabPageItem)
                ' Resize the TabPageItem
                FlowLayoutPanel_Resize(Nothing, Nothing)
                ' Reset the enable of buttons
                TabPageItem_CheckedChanged()
                ' Reset the enable of ButtonSave
                TabPageItem_ContentChanged(Nothing)
            End Sub

            Private Sub ButtonDelete_Click()
                Dim SelectedList As New ArrayList
                For Each Control As Object In FlowLayoutPanel.Controls
                    If TypeOf Control Is _FormConfiguration.LiteratureDetailDisplay.TabPageItem Then
                        With CType(Control, _FormConfiguration.LiteratureDetailDisplay.TabPageItem)
                            If .Checked Then
                                SelectedList.Add(Control)
                            End If
                        End With
                    End If
                Next

                If SelectedList.Count = 0 Then
                    Exit Sub
                End If

                Dim Message As String = ""
                If SelectedList.Count = 1 Then
                    Message = "Delete this " & SelectedList.Count & " tab page?"
                Else
                    Message = "Delete these " & SelectedList.Count & " tab pages?"
                End If

                If MsgBox(Message, vbOKCancel, "Delete") = MsgBoxResult.Ok Then
                    For Each Control As Object In SelectedList
                        FlowLayoutPanel.Controls.Remove(Control)
                    Next
                    ButtonDelete.Enabled = False
                    ' ButtonSave.Enabled = False
                    ButtonMoveDown.Enabled = False
                    ButtonMoveUp.Enabled = False
                    TabPageItem_CheckedChanged()
                    TabPageItem_ContentChanged(Nothing)
                End If

                RaiseEvent SaveCode(GenerateCode)
            End Sub

            Private Sub ButtonMoveUp_Click()
                If FlowLayoutPanel.Controls.Count < 2 Then
                    Exit Sub
                End If

                For Each TabPageItem As Object In FlowLayoutPanel.Controls
                    If Not TypeOf TabPageItem Is _FormConfiguration.LiteratureDetailDisplay.TabPageItem Then
                        Exit Sub
                    End If
                Next

                For Index As Integer = 1 To FlowLayoutPanel.Controls.Count - 1
                    If CType(FlowLayoutPanel.Controls(Index), _FormConfiguration.LiteratureDetailDisplay.TabPageItem).Checked Then
                        FlowLayoutPanel.Controls.SetChildIndex(FlowLayoutPanel.Controls(Index), Index - 1)
                    End If
                Next

                TabPageItem_CheckedChanged()

                RaiseEvent SaveCode(GenerateCode)
            End Sub

            Private Sub ButtonMoveDown_Click()
                If FlowLayoutPanel.Controls.Count < 2 Then
                    Exit Sub
                End If

                For Each TabPageItem As Object In FlowLayoutPanel.Controls
                    If Not TypeOf TabPageItem Is _FormConfiguration.LiteratureDetailDisplay.TabPageItem Then
                        Exit Sub
                    End If
                Next

                For Index As Integer = FlowLayoutPanel.Controls.Count - 2 To 0 Step -1
                    If CType(FlowLayoutPanel.Controls(Index), _FormConfiguration.LiteratureDetailDisplay.TabPageItem).Checked Then
                        FlowLayoutPanel.Controls.SetChildIndex(FlowLayoutPanel.Controls(Index), Index + 1)
                    End If
                Next

                TabPageItem_CheckedChanged()

                RaiseEvent SaveCode(GenerateCode)
            End Sub

            'Private Sub ButtonSave_Click()
            '    RaiseEvent SaveCode(GenerateCode)
            'End Sub

            Private Sub TextBoxTabPageLabel_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
                ButtonNew.Enabled = Not TextBoxTabPageLabel.IsEmpty
            End Sub

            Private Function GenerateCode() As String
                Dim Code As String = ""

                Code &= LiteratureType & "{"

                For Each Control As Object In FlowLayoutPanel.Controls
                    If TypeOf Control Is TabPageItem Then
                        With CType(Control, TabPageItem)
                            Code &= .TabPageName & "{" & .GenerateCode & "},"
                        End With
                    End If
                Next
                Code = Code.Remove(Code.LastIndexOf(","))
                Code &= "}"

                Return Code
            End Function
        End Class

    End Namespace
End Namespace

