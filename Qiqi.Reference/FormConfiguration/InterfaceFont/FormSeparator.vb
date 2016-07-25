Namespace _FormConfiguration
    Namespace InterfaceFont
        Public Class FormSeparator
            Inherits Windows.Forms.Panel

            Dim Label As New Label
            Dim Line As New Label

            Public Sub New(ByVal Text As String, ByVal Width As Integer)
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
                    ' .Font = New Font(Label.Font, FontStyle.Bold)
                    .AutoSize = True
                    .Margin = New Padding(0)
                    .Location = New Point(0, 4)
                End With

                With Me
                    .Width = Width
                    .Height = 20
                    .Controls.Add(Label)
                    .Controls.Add(Line)
                    .Margin = New Padding(0, 0, 0, 4)
                    .Name = Text.Replace(" ", "")
                End With
            End Sub

            Public Sub New(ByVal Text As String, ByVal Width As Integer, ByVal TopMarigin As Integer)
                Me.New(Text, Width)
                Me.Margin = New Padding(0, TopMarigin, 0, 4)
            End Sub

        End Class
    End Namespace
End Namespace
