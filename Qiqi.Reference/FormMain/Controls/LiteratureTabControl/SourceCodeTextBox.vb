Namespace _FormMain
    Namespace _LiteratureTabControl
        Public Class SourceCodeTextBox
            Inherits _FormMain._LiteratureTabControl.MultiLineTextBox

            Public Sub New(ByVal Text As String)
                MyBase.New(Text)

                AddHandler TextBox.Paint, AddressOf TextBox_Paint
                

            End Sub


            Private Sub TextBox_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs)
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


