Namespace _FormConfiguration
    Namespace FontConfiguration
        Public Class FontFamilyComboBox
            Inherits Windows.Forms.ComboBox

            Public Sub New()
                Me.DropDownStyle = ComboBoxStyle.DropDownList
                Me.Dock = DockStyle.Fill
                Me.Margin = New Padding(0)
                Me.FlatStyle = FlatStyle.Popup

                For Each FontFamily As FontFamily In System.Drawing.FontFamily.Families
                    Me.Items.Add(FontFamily.Name)
                Next
            End Sub

            Public Sub New(ByVal FontList As ArrayList)
                Me.DropDownStyle = ComboBoxStyle.DropDownList
                Me.Dock = DockStyle.Fill
                Me.Margin = New Padding(0)
                Me.FlatStyle = FlatStyle.Popup

                If FontList Is Nothing Then
                    Exit Sub
                End If

                For Each FontFamily As FontFamily In FontList
                    Me.Items.Add(FontFamily.Name)
                Next
            End Sub
        End Class
    End Namespace
End Namespace