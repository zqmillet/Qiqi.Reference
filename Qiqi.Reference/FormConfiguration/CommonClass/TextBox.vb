Namespace _FormConfiguration
    Public Class TextBox
        Inherits System.Windows.Forms.TextBox

        Public Sub New()
            AddHandler Me.TextChanged, AddressOf Me_TextChanged

            Me.Me_TextChanged(Nothing, Nothing)
        End Sub

        Private Sub Me_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim Size As SizeF = TextRenderer.MeasureText(Me.Text, Me.Font)
            Me.Size = New Size(Size.Width + 8, Me.Size.Height)

            If Me.Text.Trim = "" Then
                Me.BackColor = Color.Yellow
            Else
                Me.BackColor = Color.FromArgb(240, 240, 240)
            End If
        End Sub
    End Class
End Namespace


