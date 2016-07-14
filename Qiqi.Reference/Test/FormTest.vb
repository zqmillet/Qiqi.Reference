Imports SyntaxHighlighter


Public Class FormTest
    Private TextBox As New _Test.SyntaxRTB

    Private Sub FormTest_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Controls.Add(TextBox)
        Me.WindowState = FormWindowState.Maximized
        TextBox.Dock = DockStyle.Fill

        TextBox.BorderStyle = BorderStyle.None
        TextBox.Text = "@TechReport{Alliance-2013-p-,
  Title                    = {Designed-In Cyber Security for Cyber-Physical Systems},
  Author                   = {Cyber Security Research Alliance},
  Institution              = {Cyber Security Research Alliance},
  Year                     = {2013},

  File                     = {Alliance-2013-p-.pdf:References\\Alliance-2013-p-.pdf:PDF}
}"
        ' Me.Size = New Size(800, 600)

        '' TextBox.LanguageOption = RichTextBoxLanguageOptions.DualFont
        'TextBox.Font = New Font("Consolas", 10)

        'For i As Integer = 0 To TextBox.TextLength - 2
        '    TextBox.Select(i, i + 1)
        '    TextBox.SelectionColor = Color.DarkRed
        'Next

        'TextBox.Select(0, 0)

        'AddHandler TextBox.TextChanged, AddressOf TextBox_TextChanged
    End Sub


    'Private Sub TextBox_TextChanged()
    '    TextBox.Font = New Font("Consolas", 10)

    '    For i As Integer = 0 To TextBox.TextLength - 2
    '        TextBox.Select(i, i + 1)
    '        TextBox.SelectionColor = Color.DarkRed
    '    Next

    '    TextBox.Select(0, 0)
    'End Sub
End Class