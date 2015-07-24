Public Class FormTest

    Public Sub New()

        InitializeComponent()

        'Me.Controls.Add(New _Test.TextBox)
        'Me.Size = New Size(800, 600)

        Dim d As New Qiqi.DataBase("D:\iCloud\iPaper\Test.bib")
        d.Load()
    End Sub
End Class