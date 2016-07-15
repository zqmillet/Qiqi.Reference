Imports Qiqi.Reference._FormMain._LiteratureTabControl

Public Class FormTest
    Private TextBox As New _FormMain._LiteratureTabControl.RichTextBox

    Private Sub FormTest_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Controls.Add(TextBox)
        Me.WindowState = FormWindowState.Maximized
        TextBox.TextFont = New Font("Consolas", 10)
        TextBox.SyntaxHighLight = True

        TextBox.BorderStyle = BorderStyle.None
        TextBox.Text = "@TechReport{Alliance-2013-p-,Title={312312},313={132312}}"

    End Sub
End Class