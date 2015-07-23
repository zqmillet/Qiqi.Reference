Namespace _Test
    Public Class TextBox
        Inherits System.Windows.Forms.TextBox

        Public Sub New()
            With Me
                .Multiline = True
                .ScrollBars = Windows.Forms.ScrollBars.None
                .Dock = DockStyle.Fill
                .WordWrap = True
                .BorderStyle = Windows.Forms.BorderStyle.None
                .ScrollBars = RichTextBoxScrollBars.Vertical

            End With

            Dim Reader As New IO.StreamReader(Application.StartupPath & "\TestDataBase\SPChar.bib")
            Me.Text = Reader.ReadToEnd
            Reader.Close()
        End Sub

        'Private Sub Me_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.TextInput
        '    Me.Font = New Font(Me.Font.FontFamily, Me.Font.Size)
        'End Sub




        'Private Sub Me_Invalidated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Invalidated
        '    RemoveHandler Me.Invalidated, AddressOf Me_Invalidated
        '    Me.Font = New Font(Me.Font.FontFamily, Me.Font.Size)
        'End Sub
    End Class
End Namespace

