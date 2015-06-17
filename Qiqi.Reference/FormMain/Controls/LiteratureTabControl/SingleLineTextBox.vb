Namespace _FormMain
    Namespace _LiteratureTabControl
        Public Class SingleLineTextBox
            Inherits _FormMain._LiteratureTabControl.MultiLineTextBox

            Public Sub New(ByVal Text As String)
                MyBase.New(Text)
                With Me
                    .TextBox.Multiline = False
                    .Margin = New Padding(2)
                End With
            End Sub
        End Class
    End Namespace
End Namespace


