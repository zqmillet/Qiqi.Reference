Namespace _FormConfiguration
    Namespace FontConfiguration
        Public Class ColorComboBox
            Inherits Windows.Forms.Panel

            Dim Label As New Label
            Dim ComboBox As New ColorPicker
            Dim CheckBox As New CheckBox

            Public Const HeadLabelWidth As Integer = 75
            Public Const CheckBoxWidth As Integer = 50


            Public Sub New(ByVal Text As String, ByVal Width As Integer)
                With ComboBox
                    .DropDownStyle = ComboBoxStyle.DropDownList
                    .Margin = New Padding(0)
                    .FlatStyle = FlatStyle.System
                    .Location = New Point(HeadLabelWidth, 0)
                    .Size = New Size(Width - HeadLabelWidth - CheckBoxWidth, Me.Height)
                    .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
                    .AddStandardColors()
                End With

                With Label
                    .Text = Text
                    .AutoSize = True
                    .Location = New Point(0, 4)
                End With

                With CheckBox
                    .Text = "Bold"
                    .AutoSize = False
                    .Location = New Point(Width - CheckBoxWidth + 4, -1)
                End With

                With Me
                    .Height = 20
                    .Width = Width
                    .Margin = New Padding(0, 0, 0, 4)
                    .Controls.Add(Label)
                    .Controls.Add(ComboBox)
                    .Controls.Add(CheckBox)
                End With

                FillComboBox()
            End Sub

            Public Overridable Sub FillComboBox()
                'ComboBox.Items.Add("red")
                'ComboBox.Items.Add("blue")
            End Sub

        End Class

    End Namespace
End Namespace
