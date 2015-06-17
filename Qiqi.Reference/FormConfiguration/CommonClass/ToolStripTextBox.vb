Namespace _FormConfiguration
    Public Class ToolStripTextBox
        Inherits System.Windows.Forms.ToolStripTextBox

        Dim Label As String
        Public IsEmpty As Boolean

        Public Sub New(ByVal Label As String)
            With Me
                .Label = Label
                .IsEmpty = True
                .ForeColor = Color.Gray
                .Text = Label
                .BorderStyle = Windows.Forms.BorderStyle.FixedSingle
            End With


            AddHandler Me.TextChanged, AddressOf ToolStripTextBox_TextChanged
            AddHandler Me.Enter, AddressOf ToolStripTextBox_Enter
            AddHandler Me.Leave, AddressOf ToolStripTextBox_Leave

        End Sub

        Public Sub SetText(ByVal Text As String)
            If Text.Trim = "" Then
                Me.Text = ""
                Me.IsEmpty = True
                Me.ForeColor = Color.Gray
                Me.Text = Label
            Else
                Me.Text = Text
                Me.IsEmpty = False
                Me.ForeColor = Color.Black
            End If

        End Sub

        Private Sub ToolStripTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
            If Me.Text.Trim = "" Or ForeColor = Color.Gray Then
                Me.IsEmpty = True
            Else
                Me.IsEmpty = False
            End If
        End Sub


        Private Sub ToolStripTextBox_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
            With Me
                .ForeColor = Color.Black
                .SelectAll()
                If IsEmpty Then
                    .Text = ""
                End If
            End With
        End Sub

        Private Sub ToolStripTextBox_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
            With Me
                If IsEmpty Then
                    .ForeColor = Color.Gray
                    .Text = Label

                End If
            End With
        End Sub
    End Class
End Namespace

