Namespace _FormConfiguration
    Namespace InterfaceFont
        Public Class TransparentPanel
            Inherits Panel

            Public HideBorder As Boolean = False

            Protected Overrides ReadOnly Property CreateParams() As System.Windows.Forms.CreateParams
                Get
                    Dim cp As CreateParams = MyBase.CreateParams
                    cp.ExStyle = cp.ExStyle Or &H20 ''#WS_EX_TRANSPARENT
                    Return cp
                End Get
            End Property

            Protected Overrides Sub OnPaintBackground(ByVal e As System.Windows.Forms.PaintEventArgs)
                ' MyBase.OnPaintBackground(e)
            End Sub

            Protected Overrides Sub OnPaint(e As PaintEventArgs)
                MyBase.OnPaint(e)

                If HideBorder Then
                    Exit Sub
                End If

                Dim gr As Graphics = Me.CreateGraphics
                gr.DrawRectangle(New Pen(Color.Red, 6), New Rectangle(New Point(0, 0), Me.Size))
            End Sub
        End Class
    End Namespace
End Namespace
