Namespace _FormMain
    Namespace _LiteratureTabControl
        Public Class SingleLineTextBox
            Inherits _FormMain._LiteratureTabControl.MultiLineTextBox

            Public Sub New()
                With Me
                    .TextBox.Multiline = False
                    .Padding = New Padding(4)
                    .BackColor = Color.White
                End With
            End Sub

            Public Sub New(ByVal Text As String)
                Me.New
                Me.TextBox.Text = Text
            End Sub
        End Class
    End Namespace
End Namespace


