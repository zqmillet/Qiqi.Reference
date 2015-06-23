Public Class FormTest

    'Public Sub New()

    '    InitializeComponent()

    '    Me.Controls.Add(New _Test.TextBox)
    '    Me.Size = New Size(800, 600)
    'End Sub

    Public Sub New()

        ' 此调用是设计器所必需的。
        InitializeComponent()
        Me.Size = New Size(800, 600)
        SplitContainer1.IsSplitterFixed = False
        SplitContainer1.FixedPanel = FixedPanel.Panel1
        SplitContainer1.SplitterDistance = 100


    End Sub
End Class