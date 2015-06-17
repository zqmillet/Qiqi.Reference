Namespace _FormMain
    Namespace _LiteratureTabControl
        Public Class SourceCodeTextBox
            Inherits _FormMain._LiteratureTabControl.MultiLineTextBox

            Public Sub New(ByVal Text As String)
                AddHandler TextBox.TextChanged, AddressOf TextBox_TextChanged

                TextBox.Text = Text
                TextBox.Font = New Font(TextBox.Font.FontFamily, TextBox.Font.Size)

            End Sub

            Private Sub TextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)

                With TextBox
                    .Font = New Font(TextBox.Font.FontFamily, TextBox.Font.Size)
                End With

            End Sub

        End Class
    End Namespace
End Namespace


