Public Class FormTest

    Public Sub New()
        InitializeComponent()

        For Each Font As FontFamily In System.Drawing.FontFamily.Families

        Next
    End Sub

    Private Sub FormTest_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ColorDialog1.ShowDialog()
    End Sub


End Class