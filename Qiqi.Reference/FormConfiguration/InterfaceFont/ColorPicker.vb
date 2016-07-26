Namespace _FormConfiguration
    Namespace InterfaceFont
        ''' <summary>
        ''' This class is ComboBox to help user to pick a color.
        ''' The characteristic of this class is that it can show the color block of each dropped item,
        ''' </summary>
        Public Class ColorPicker
            Inherits ComboBox

            ' Data for each color in the list
            Private _ButtonArea As New Rectangle

            Public Sub New()
                MyBase.New
                Me.DropDownStyle = ComboBoxStyle.DropDownList
                Me.DrawMode = Windows.Forms.DrawMode.OwnerDrawFixed
                Me.BackColor = System.Drawing.SystemColors.Control
                ' Cache the button's modified ClientRectangle (see Note)

                'Me.ForeColor = System.Drawing.SystemColors.AppWorkspace
                With _ButtonArea
                    .X = Me.ClientRectangle.X - 1
                    .Y = Me.ClientRectangle.Y - 1
                    .Width = Me.Width + 2
                    .Height = Me.Height + 2
                End With

                AddHandler Me.DrawItem, AddressOf Me_DrawItem
            End Sub


            ' Populate control with standard colors
            Public Sub AddStandardColors()
                Items.Clear()
                Items.Add(New ColorInfo("Black"))
                Items.Add(New ColorInfo("Blue"))
                Items.Add(New ColorInfo("Lime"))
                Items.Add(New ColorInfo("Cyan"))
                Items.Add(New ColorInfo("Red"))
                Items.Add(New ColorInfo("Fuchsia"))
                Items.Add(New ColorInfo("Yellow"))
                Items.Add(New ColorInfo("White"))
                Items.Add(New ColorInfo("Navy"))
                Items.Add(New ColorInfo("Green"))
                Items.Add(New ColorInfo("Teal"))
                Items.Add(New ColorInfo("Maroon"))
                Items.Add(New ColorInfo("Purple"))
                Items.Add(New ColorInfo("Olive"))
                Items.Add(New ColorInfo("Gray"))
            End Sub

            ' Draw list item
            Protected Sub Me_DrawItem(ByVal sender As Object, ByVal e As DrawItemEventArgs)
                If (e.Index >= 0) Then
                    ' Get this color
                    Dim Color As ColorInfo = Items(e.Index)
                    ' Fill background
                    e.DrawBackground()
                    ' Draw color box
                    Dim Rectangle As Rectangle = New Rectangle
                    Rectangle.X = (e.Bounds.X + 1)
                    Rectangle.Y = (e.Bounds.Y + 1)
                    Rectangle.Width = 30
                    Rectangle.Height = (e.Bounds.Height - 3)
                    e.Graphics.FillRectangle(New SolidBrush(Color.Color), Rectangle)
                    ' e.Graphics.DrawRectangle(SystemPens.WindowText, Rectangle)
                    ' Write color name
                    Dim brush As Brush
                    If ((e.State And DrawItemState.Selected) _
                        <> DrawItemState.None) Then
                        brush = SystemBrushes.HighlightText
                    Else
                        brush = SystemBrushes.WindowText
                    End If

                    e.Graphics.DrawString(Color.Text, Font, brush, (e.Bounds.X + (Rectangle.X + (Rectangle.Width + 2))),
                                                                   (e.Bounds.Y + ((e.Bounds.Height - Font.Height) / 2)))
                    ' Draw the focus Rectangle if appropriate
                    If ((e.State And DrawItemState.NoFocusRect) _
                        = DrawItemState.None) Then
                        e.DrawFocusRectangle()
                    End If
                End If
            End Sub

            ''' <summary>
            ''' Gets or sets the currently selected item.
            ''' </summary>
            Public Shadows Property SelectedItem As Integer
                Get
                    If MyBase.SelectedItem Is Nothing Then
                        Return 0
                    Else
                        Return CType(MyBase.SelectedItem, ColorInfo).Color.ToArgb
                    End If
                End Get
                Set(ByVal Argb As Integer)
                    Dim i As Integer = 0
                    Do While (i < Items.Count)
                        If (CType(Items(i), ColorInfo).Color.ToArgb = Argb) Then
                            SelectedIndex = i
                            Exit Property
                        End If
                        i += 1
                    Loop

                    SelectedIndex = 0
                End Set
            End Property
        End Class

        Public Class ColorInfo

            Public Text As String
            Public Color As Color

            Public Sub New(ByVal Text As String)
                MyBase.New
                Me.Text = Text
                Me.Color = Drawing.Color.FromName(Text)
            End Sub
        End Class
    End Namespace
End Namespace
