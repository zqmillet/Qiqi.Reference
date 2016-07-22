Public Class ColorPicker
    Inherits ComboBox

    ' Data for each color in the list

    Public Sub New()
        MyBase.New
        DropDownStyle = ComboBoxStyle.DropDownList
        DrawMode = DrawMode.OwnerDrawFixed
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
            Dim color As ColorInfo = Items(e.Index)
            ' Fill background
            e.DrawBackground()
            ' Draw color box
            Dim rect As Rectangle = New Rectangle
            rect.X = (e.Bounds.X + 1)
            rect.Y = (e.Bounds.Y + 1)
            rect.Width = 30
            rect.Height = (e.Bounds.Height - 1)
            e.Graphics.FillRectangle(New SolidBrush(color.Color), rect)
            ' e.Graphics.DrawRectangle(SystemPens.WindowText, rect)
            ' Write color name
            Dim brush As Brush
            If ((e.State And DrawItemState.Selected) _
                        <> DrawItemState.None) Then
                brush = SystemBrushes.HighlightText
            Else
                brush = SystemBrushes.WindowText
            End If

            e.Graphics.DrawString(color.Text, Font, brush, (e.Bounds.X _
                            + (rect.X _
                            + (rect.Width + 2))), (e.Bounds.Y _
                            + ((e.Bounds.Height - Font.Height) _
                            / 2)))
            ' Draw the focus rectangle if appropriate
            If ((e.State And DrawItemState.NoFocusRect) _
                        = DrawItemState.None) Then
                e.DrawFocusRectangle()
            End If

        End If

    End Sub

    ''' <summary>
    ''' Gets or sets the currently selected item.
    ''' </summary>
    Public Shadows Property SelectedItem As ColorInfo
        Get
            Return CType(MyBase.SelectedItem, ColorInfo)
        End Get
        Set
            MyBase.SelectedItem = Value
        End Set
    End Property

    ''' <summary>
    ''' Gets the text of the selected item, or sets the selection to
    ''' the item with the specified text.
    ''' </summary>
    Public Shadows Property SelectedText As String
        Get
            If (SelectedIndex >= 0) Then
                Return Me.SelectedItem.Text
            End If

            Return String.Empty
        End Get
        Set
            Dim i As Integer = 0
            Do While (i < Items.Count)
                If (CType(Items(i), ColorInfo).Text = Value) Then
                    SelectedIndex = i
                    Exit Do
                End If

                i = (i + 1)
            Loop

        End Set
    End Property

    ''' <summary>
    ''' Gets the value of the selected item, or sets the selection to
    ''' the item with the specified value.
    ''' </summary>
    Public Shadows Property SelectedValue As Color
        Get
            If (SelectedIndex >= 0) Then
                Return Me.SelectedItem.Color
            End If

            Return Color.White
        End Get
        Set
            Dim i As Integer = 0
            Do While (i < Items.Count)
                If (CType(Items(i), ColorInfo).Color = Value) Then
                    SelectedIndex = i
                    Exit Do
                End If

                i = (i + 1)
            Loop

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