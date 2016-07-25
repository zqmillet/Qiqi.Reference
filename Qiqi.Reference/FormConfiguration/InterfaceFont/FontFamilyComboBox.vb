Namespace _FormConfiguration
    Namespace InterfaceFont
        Public Class FontFamilyComboBox
            Inherits FontSizeComboBox

            Public Sub New(ByVal Width As Integer)
                MyBase.New(Width)

                Me.TailLabel.Visible = False
                Me.HeadLabel.Text = "Font Family"
                Me.ComboBox.Size = New Size(Me.Width - HeadLabelWidth, Me.Height)
            End Sub

            Public Overrides Sub FillComboBox()
                For Each Font As FontFamily In System.Drawing.FontFamily.Families
                    Me.ComboBox.Items.Add(Font.Name)
                Next
            End Sub

            Public Overrides Sub ComboBox_DrawItem(ByVal sender As Object, ByVal e As DrawItemEventArgs)
                If (e.Index >= 0) Then
                    e.DrawBackground()

                    Dim Brush As Brush
                    If ((e.State And DrawItemState.Selected) _
                        <> DrawItemState.None) Then
                        Brush = SystemBrushes.HighlightText
                    Else
                        Brush = SystemBrushes.WindowText
                    End If

                    Dim ItemFont As New Font(ComboBox.Items(e.Index).ToString, ComboBox.Font.Size, FontStyle.Regular)
                    e.Graphics.DrawString(ComboBox.Items(e.Index), ItemFont, Brush, e.Bounds.X + 3, e.Bounds.Y + 1)
                    ' Draw the focus Rectangle if appropriate
                    If ((e.State And DrawItemState.NoFocusRect) _
                        = DrawItemState.None) Then
                        e.DrawFocusRectangle()
                    End If
                End If
            End Sub
        End Class
    End Namespace
End Namespace