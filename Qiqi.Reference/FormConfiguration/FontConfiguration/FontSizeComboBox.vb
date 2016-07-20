Namespace _FormConfiguration
    Namespace FontConfiguration
        Public Class FontSizeComboBox
            Inherits Windows.Forms.ComboBox

            Dim FontSizeList() As Integer = {5, 5.5, 6.5, 7.5, 9, 10.5, 12, 14, 15, 16, 18, 22, 24, 26, 36, 42, 48, 54, 63, 72}

            Private Sub New()
                Me.DropDownStyle = ComboBoxStyle.DropDown
                Me.Dock = DockStyle.Fill
                For Each FontSize As Integer In FontSizeList
                    Me.Items.Add(FontSize)
                Next
            End Sub
        End Class
    End Namespace
End Namespace