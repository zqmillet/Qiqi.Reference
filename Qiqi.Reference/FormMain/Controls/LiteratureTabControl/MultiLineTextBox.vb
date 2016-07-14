Namespace _FormMain
    Namespace _LiteratureTabControl
        Public Class MultiLineTextBox
            Inherits System.Windows.Forms.Panel

            Public TextBox As DerivedRichTextBox

            Public Sub New()
                TextBox = New DerivedRichTextBox

                With TextBox
                    .Text = ""
                    .Multiline = True
                    .Padding = New Padding(0)
                    .Dock = DockStyle.Fill
                    .BorderStyle = Windows.Forms.BorderStyle.None
                    .ScrollBars = ScrollBars.Vertical
                End With

                With Me
                    .Padding = New Padding(3)
                    .Margin = New Padding(0)
                    .Dock = DockStyle.Fill
                    .BorderStyle = Windows.Forms.BorderStyle.None
                    .BackColor = Color.White
                    .Controls.Add(TextBox)
                End With
            End Sub

            Public Sub New(ByVal Text As String)
                TextBox = New DerivedRichTextBox

                With TextBox
                    .Text = Text
                    .Multiline = True
                    .Padding = New Padding(0)
                    .Dock = DockStyle.Fill
                    .BorderStyle = Windows.Forms.BorderStyle.None
                    .ScrollBars = ScrollBars.Vertical
                End With

                With Me
                    .Padding = New Padding(3)
                    .Margin = New Padding(0)
                    .Dock = DockStyle.Fill
                    .BorderStyle = Windows.Forms.BorderStyle.None
                    .BackColor = Color.White
                    .Controls.Add(TextBox)
                End With
            End Sub
        End Class
    End Namespace
End Namespace