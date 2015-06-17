Namespace _FormConfiguration
    Public Class ListView
        Inherits System.Windows.Forms.ListView

        Public Msg As UInt32
        Const WM_LBUTTONDBLCLK = &H203

        Protected Overloads Overrides Sub WndProc(ByRef m As Message)
            Msg = m.Msg
            If m.Msg = WM_LBUTTONDBLCLK Then
                Dim p As Point = PointToClient(New Point(Cursor.Position.X, Cursor.Position.Y))
                Dim lvi As ListViewItem = GetItemAt(p.X, p.Y)
                If Not lvi Is Nothing Then
                    lvi.Selected = True
                End If
                OnDoubleClick(New EventArgs)
            Else
                MyBase.WndProc(m)
            End If
        End Sub

    End Class
End Namespace
