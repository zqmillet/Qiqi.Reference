Namespace _FormMain
    Public Class DataBaseTabControl
        Inherits System.Windows.Forms.TabControl

        Public Event TabPageChanged(ByVal sender As Object)

        Public Sub New()
            With Me
                .Dock = DockStyle.Fill

                AddHandler .ControlAdded, AddressOf Me_ControlAdded
                AddHandler .ControlRemoved, AddressOf Me_ControlRemoved
                AddHandler .SelectedIndexChanged, AddressOf Me_SelectedIndexChanged
            End With
        End Sub

        Private Sub Me_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
            RaiseEvent TabPageChanged(Me)
        End Sub

        Private Sub Me_ControlAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.ControlEventArgs)
            RaiseEvent TabPageChanged(Me)
        End Sub

        Private Sub Me_ControlRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.ControlEventArgs)
            RaiseEvent TabPageChanged(Me)
        End Sub

        Public Function Exist(ByVal BibTeXFullName As String) As Boolean
            For Each TabPage As _FormMain.DataBaseTabPage In Me.TabPages
                If TabPage.Name.ToLower.Trim = BibTeXFullName.Trim.ToLower Then
                    Return True
                End If
            Next

            Return False
        End Function
    End Class
End Namespace