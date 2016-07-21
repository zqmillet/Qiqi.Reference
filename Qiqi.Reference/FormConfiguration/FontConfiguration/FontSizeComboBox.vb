Namespace _FormConfiguration
    Namespace FontConfiguration
        Public Class FontSizeComboBox
            Inherits Windows.Forms.Panel

            Public ComboBox As New ComboBox
            Public HeadLabel As New Label
            Public TailLabel As New Label
            Public Const HeadLabelWidth As Integer = 75
            Public Const TailLabelWidth As Integer = 20

            Public Sub New(ByVal Width As Integer)
                With ComboBox
                    .DropDownStyle = ComboBoxStyle.DropDownList
                    .Margin = New Padding(0)
                    .FlatStyle = FlatStyle.System
                    .Location = New Point(HeadLabelWidth, 0)
                    .Size = New Size(Width - HeadLabelWidth - TailLabelWidth, Me.Height)
                    .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
                End With

                With HeadLabel
                    .Text = "Font Size"
                    .AutoSize = True
                    .Location = New Point(0, 4)
                End With

                With TailLabel
                    .Text = "pt"
                    .AutoSize = False
                    .Location = New Point(Width - TailLabelWidth + 4, 4)
                End With

                With Me
                    .Height = 20
                    .Width = Width
                    .Margin = New Padding(0, 0, 0, 4)
                    .Controls.Add(HeadLabel)
                    .Controls.Add(ComboBox)
                    .Controls.Add(TailLabel)
                End With

                FillComboBox()
            End Sub

            Public Overridable Sub FillComboBox()
                For Each FontSize As Integer In {5, 5.5, 6.5, 7.5, 9, 10.5, 12, 14, 15, 16, 18, 22, 24, 26, 36, 42, 48, 54, 63, 72}
                    ComboBox.Items.Add(FontSize)
                Next
            End Sub
        End Class
    End Namespace
End Namespace