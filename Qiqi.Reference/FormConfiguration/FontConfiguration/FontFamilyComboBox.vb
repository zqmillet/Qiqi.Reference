Namespace _FormConfiguration
    Namespace FontConfiguration
        Public Class FontFamilyComboBox
            Inherits FontSizeComboBox

            Public Sub New(ByVal Width As Integer)
                MyBase.New(Width)

                Me.TailLabel.Visible = False
                Me.HeadLabel.Text = "Font Family"
                Me.ComboBox.Size = New Size(Me.Width - HeadLabelWidth, Me.Height)
            End Sub

            Public Overrides Sub FillComboBox()
                For Each Font As FontFamily In System.Drawing.FontFamily.Families
                    Me.ComboBox.Items.Add(Font.Name)
                Next
            End Sub
        End Class
    End Namespace
End Namespace