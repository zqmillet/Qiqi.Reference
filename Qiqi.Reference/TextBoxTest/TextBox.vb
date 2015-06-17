Namespace _Test
    Public Class TextBox
        Inherits System.Windows.Forms.TextBox

        Public Sub New()
            With Me
                .Multiline = True
                .ScrollBars = Windows.Forms.ScrollBars.None
                .Dock = DockStyle.Fill
                .WordWrap = True

                AddHandler Me.TextChanged, AddressOf Me_TextChanged

            End With


        End Sub

        Private Sub Me_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
            If Me.Text.Length > 100 Then
                Me.ScrollBars = Windows.Forms.ScrollBars.Vertical
            Else
                Me.ScrollBars = Windows.Forms.ScrollBars.None
            End If
        End Sub
    End Class
End Namespace

