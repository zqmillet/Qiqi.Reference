Namespace _FormConfiguration
    Public Class DataGridViewDisableCheckBoxColumn
        Inherits DataGridViewCheckBoxColumn

        Sub New()
            Me.CellTemplate = New _FormConfiguration.DataGridViewDisableCheckBoxCell()
        End Sub
    End Class

    Public Class DataGridViewDisableCheckBoxCell
        Inherits DataGridViewCheckBoxCell

        Dim enabledValue As Boolean

        Public Property Enabled() As Boolean
            Get
                Return enabledValue
            End Get
            Set(ByVal value As Boolean)
                enabledValue = value
            End Set
        End Property

        Public Overrides Function Clone() As Object
            Dim cell As New DataGridViewDisableCheckBoxCell
            cell = MyBase.Clone
            Return cell
        End Function

        Protected Overrides Sub Paint(ByVal graphics As System.Drawing.Graphics, ByVal clipBounds As System.Drawing.Rectangle, ByVal cellBounds As System.Drawing.Rectangle, ByVal rowIndex As Integer, ByVal elementState As System.Windows.Forms.DataGridViewElementStates, ByVal value As Object, ByVal formattedValue As Object, ByVal errorText As String, ByVal cellStyle As System.Windows.Forms.DataGridViewCellStyle, ByVal advancedBorderStyle As System.Windows.Forms.DataGridViewAdvancedBorderStyle, ByVal paintParts As System.Windows.Forms.DataGridViewPaintParts)
            MyBase.Paint(graphics, clipBounds, cellBounds, rowIndex, elementState, value, formattedValue, errorText, cellStyle, advancedBorderStyle, paintParts)

            If (Me.enabledValue) Then
                Dim cellBackground As New SolidBrush(cellStyle.BackColor)

                graphics.FillRectangle(cellBackground, cellBounds)
                cellBackground.Dispose()

                PaintBorder(graphics, clipBounds, cellBounds, cellStyle, advancedBorderStyle)
                Dim checkBoxArea As Rectangle = cellBounds
                Dim buttonAdjustment As Rectangle = Me.BorderWidths(advancedBorderStyle)

                checkBoxArea.X += buttonAdjustment.X
                checkBoxArea.Y += buttonAdjustment.Y

                checkBoxArea.Height -= buttonAdjustment.Height
                checkBoxArea.Width -= buttonAdjustment.Width

                Dim drawInPoint As New Point(cellBounds.X + cellBounds.Width / 2 - 7, cellBounds.Y + cellBounds.Height / 2 - 7)

                CheckBoxRenderer.DrawCheckBox(graphics, drawInPoint, System.Windows.Forms.VisualStyles.CheckBoxState.CheckedDisabled)
            End If
        End Sub
    End Class
End Namespace

