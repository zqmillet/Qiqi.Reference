Namespace _FormConfiguration
    Namespace InterfaceFont
        ''' <summary>
        ''' This class is a label with a long tail.
        ''' </summary>
        Public Class FormSeparator
            Inherits Windows.Forms.Panel

            ' The label which is used to show the text.
            Dim Label As New Label
            ' The line which is regarded as a long tail.
            Dim Line As New Label

            ''' <summary>
            ''' Constructor.
            ''' </summary>
            ''' <param name="Text">The text of the Label.</param>
            ''' <param name="Width">The width of the whole Label.</param>
            ''' <param name="TopMargin">The top margin of the Label.</param>
            Public Sub New(ByVal Text As String, ByVal Width As Integer, Optional ByVal TopMargin As Integer = 0)
                With Line
                    .Text = ""
                    .AutoSize = False
                    .Height = 2
                    .Width = Width
                    .BorderStyle = BorderStyle.Fixed3D
                    .Margin = New Padding(0, 4, 0, 4)
                    .Location = New Point(0, 10)
                    .Anchor = AnchorStyles.Right Or AnchorStyles.Left
                End With

                With Label
                    .Text = Text
                    .AutoSize = True
                    .Margin = New Padding(0)
                    .Location = New Point(0, 4)
                End With

                With Me
                    .Width = Width
                    .Height = 20
                    .Controls.Add(Label)
                    .Controls.Add(Line)
                    .Margin = New Padding(0, TopMargin, 0, 4)
                    .Name = Text.Replace(" ", "")
                End With
            End Sub
        End Class
    End Namespace
End Namespace
