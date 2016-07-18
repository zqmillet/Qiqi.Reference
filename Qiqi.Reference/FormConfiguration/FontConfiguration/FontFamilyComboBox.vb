Namespace _FormConfiguration
    Namespace FontConfiguration
        Public Class FontFamilyComboBox
            Inherits Windows.Forms.ComboBox

            Public Sub New()
                Me.DropDownStyle = ComboBoxStyle.DropDownList
                For Each FontFamily As FontFamily In System.Drawing.FontFamily.Families
                    Me.Items.Add(FontFamily.Name)
                Next
            End Sub

            Public Sub New(ByVal FontList As ArrayList)
                Me.DropDownStyle = ComboBoxStyle.DropDownList

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