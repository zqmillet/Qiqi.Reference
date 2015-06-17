Namespace _FormMain
    Public Class PrimaryToolStrip
        Inherits System.Windows.Forms.ToolStrip

        Public Sub New()
            MyBase.New()

            With Me
                .Anchor = AnchorStyles.Left Or AnchorStyles.Top
            End With
        End Sub
    End Class

    Public Class SecondaryToolStrip
        Inherits System.Windows.Forms.ToolStrip

        Dim ToolStripButtonHide As ToolStripButton
        Dim ToolStripButtonMoveUp As ToolStripButton
        Dim ToolStripButtonMoveDown As ToolStripButton

        Public Event ToolStripButtonHideClick()
        Public Event ToolStripButtonMoveUpClick()
        Public Event ToolStripButtonMoveDownClick()

        Public Sub New()
            MyBase.New()

            ToolStripButtonHide = New ToolStripButton
            AddHandler ToolStripButtonHide.Click, AddressOf ToolStripButton_Click
            With ToolStripButtonHide
                .Margin = New Padding(3, 3, 0, 0)
                .ToolTipText = "Hide"
                .Name = "ToolStripButtonHide"
                .Image = Resource.Icon.ToolStripButtonHide
                .Enabled = True
                .Visible = True
            End With

            ToolStripButtonMoveUp = New ToolStripButton
            AddHandler ToolStripButtonMoveUp.Click, AddressOf ToolStripButton_Click
            With ToolStripButtonMoveUp
                .Margin = New Padding(3, 3, 0, 0)
                .ToolTipText = "Move Up"
                .Name = "ToolStripButtonMoveUp"
                .Image = Resource.Icon.ToolStripButtonMoveUp
                .Enabled = True
                .Visible = True
            End With

            ToolStripButtonMoveDown = New ToolStripButton
            AddHandler ToolStripButtonMoveDown.Click, AddressOf ToolStripButton_Click
            With ToolStripButtonMoveDown
                .Margin = New Padding(3, 3, 0, 0)
                .ToolTipText = "Move Down"
                .Name = "ToolStripButtonMoveDown"
                .Image = Resource.Icon.ToolStripButtonMoveDown
                .Enabled = True
                .Visible = True
            End With

            With Me
                .Dock = DockStyle.Left
                .GripStyle = ToolStripGripStyle.Hidden
                .RenderMode = ToolStripRenderMode.Professional
                .Items.Add(ToolStripButtonHide)
                .Items.Add(New ToolStripSeparator)
                .Items.Add(ToolStripButtonMoveUp)
                .Items.Add(ToolStripButtonMoveDown)
            End With
        End Sub

        Private Sub ToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Select Case CType(sender, ToolStripButton).Name
                Case "ToolStripButtonHide"
                    RaiseEvent ToolStripButtonHideClick()
                Case "ToolStripButtonMoveUp"
                    RaiseEvent ToolStripButtonMoveUpClick()
                Case "ToolStripButtonMoveDown"
                    RaiseEvent ToolStripButtonMoveDownClick()
            End Select
        End Sub

        Public Sub SetButtonEnable(ByVal Top As Boolean, ByVal Bottom As Boolean)
            If Top Then
                ToolStripButtonMoveUp.Enabled = False
            Else
                ToolStripButtonMoveUp.Enabled = True
            End If

            If Bottom Then
                ToolStripButtonMoveDown.Enabled = False
            Else
                ToolStripButtonMoveDown.Enabled = True
            End If

        End Sub
    End Class

End Namespace