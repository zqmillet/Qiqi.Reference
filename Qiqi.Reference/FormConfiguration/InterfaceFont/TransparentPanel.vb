Namespace _FormConfiguration
    Namespace InterfaceFont

        ''' <summary>
        ''' This class is a panel which is tansparent.
        ''' This class is used to cover the other control, which should be click by mouse.
        ''' Configurate the Enable attribute of a control to prevent it from being clicked is not ideal sometime,
        ''' because if the Enable = False, the color of this control is gray.
        ''' </summary>
        Public Class TransparentPanel
            Inherits Panel

            ' This class has a red border with 6pt width by default.
            ' If you want to remove this red border, configurate HideBorder = False
            Public HideBorder As Boolean = False

            ''' <summary>
            ''' CreateParams is a property of Control object.
            ''' CreateParams is used to gets the required creation parameters when the control handle is created.
            ''' In this class, we override the property CreateParams,
            ''' and set the ExStyle = ExStyle Or &H20 to make the background is trasparent.
            ''' </summary>
            ''' <returns></returns>
            Protected Overrides ReadOnly Property CreateParams() As System.Windows.Forms.CreateParams
                Get
                    Dim cp As CreateParams = MyBase.CreateParams
                    cp.ExStyle = cp.ExStyle Or &H20
                    Return cp
                End Get
            End Property

            ''' <summary>
            ''' We override the sub OnPaintBackground to prevent the background of this panel from being painted.
            ''' </summary>
            ''' <param name="e"></param>
            Protected Overrides Sub OnPaintBackground(ByVal e As System.Windows.Forms.PaintEventArgs)
                ' Do nothing.
            End Sub

            ''' <summary>
            ''' We override this sub to draw a red border with 6pt width on this panel.
            ''' </summary>
            ''' <param name="e"></param>
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
