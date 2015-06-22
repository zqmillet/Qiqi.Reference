Namespace _FormMain
    Public Module StatusMessage
        Public Const DataBaseLoading As String = "DataBase Loading"
        Public Const DataBaseLoaded As String = "DataBase Loaded"
    End Module

    Public Class StatusStrip
        Inherits System.Windows.Forms.StatusStrip

        Public StatusLabel As ToolStripStatusLabel

        Public Sub New()
            StatusLabel = New ToolStripStatusLabel

            With Me
                .Visible = False
            End With

            With StatusLabel
                .Text = ""
            End With

            Me.Items.Add(StatusLabel)
        End Sub

        Public Sub DataBaseLoading(ByVal Progress As Double)
            StatusLabel.Text = _FormMain.StatusMessage.DataBaseLoading & " " & Math.Round(Progress * 100, 2) & "%"
        End Sub

        Public Sub DataBaseLoaded()
            StatusLabel.Text = _FormMain.StatusMessage.DataBaseLoaded
        End Sub

        Public Sub ShowErrorMessage(ByVal ErrorMessage As String)
            StatusLabel.Text = ErrorMessage
        End Sub
    End Class
End Namespace


