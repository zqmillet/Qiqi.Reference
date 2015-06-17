Namespace _FormMain
    Namespace _LiteratureTabControl
        Public Class SingleLineTextBox
            Inherits _FormMain._LiteratureTabControl.MultiLineTextBox

            Public Sub New(ByVal Text As String)
                With Me
                    .TextBox.Multiline = False
                    .TextBox.Text = Text
                    .Margin = New Padding(2)
                End With
            End Sub
        End Class
    End Namespace
End Namespace


