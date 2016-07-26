Namespace _FormConfiguration
    Namespace InterfaceFont
        ''' <summary>
        ''' This class is a ComboBox with a head label and a tail label.
        ''' It can be used to let the user select a number which is the size of a font.
        ''' </summary>
        Public Class FontSizeComboBox
            Inherits Windows.Forms.Panel

            ' This is the mainbody, the ComboBox.
            Public ComboBox As New ComboBox
            ' This is the head label.
            Public HeadLabel As New Label
            ' This is the tail label.
            Public TailLabel As New Label
            ' These two widths are width of the head label and the width of the tail width.
            Public Const HeadLabelWidth As Integer = 75
            Public Const TailLabelWidth As Integer = 20

            ' The following event will be triggered when the text of ComboBox is changed.
            Public Event SelectedChanged()
            ' The following event will be triggered when the mouse enters the ComboBox, the head label, or the tail label.
            Public Event SubControlMouseEnter(sender As Object, e As System.EventArgs)

            ''' <summary>
            ''' This property is the text of the ComboBox.
            ''' When this property is assigned, the ComboBox's selected index will be changed.
            ''' </summary>
            ''' <returns></returns>
            Public Property SelectedText As String
                Get
                    If ComboBox.SelectedItem Is Nothing Then
                        Return ""
                    Else
                        Return ComboBox.SelectedItem.ToString
                    End If
                End Get
                Set(ByVal Value As String)
                    Dim Index As Integer = 0
                    For Each Item As Object In ComboBox.Items
                        If Item.ToString = Value Then
                            ComboBox.SelectedIndex = Index
                            Exit Property
                        End If
                        Index += 1
                    Next
                End Set
            End Property

            ''' <summary>
            ''' Constructor.
            ''' </summary>
            ''' <param name="Width">The width of the whole panel.</param>
            Public Sub New(ByVal Width As Integer)
                With ComboBox
                    .DropDownStyle = ComboBoxStyle.DropDownList
                    .DrawMode = DrawMode.OwnerDrawFixed
                    .BackColor = System.Drawing.SystemColors.Control
                    .Margin = New Padding(0)
                    .Padding = New Padding(0)
                    .FlatStyle = FlatStyle.System
                    .Location = New Point(HeadLabelWidth, 0)
                    .Size = New Size(Width - HeadLabelWidth - TailLabelWidth, Me.Height)
                    .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
                    AddHandler .DrawItem, AddressOf ComboBox_DrawItem
                    AddHandler .SelectedIndexChanged, AddressOf ComboBox_SelectedIndexChanged
                    AddHandler .MouseEnter, AddressOf ComboBox_MouseEnter
                End With

                With HeadLabel
                    .Text = "Font Size"
                    .AutoSize = True
                    .Location = New Point(0, 4)
                    AddHandler .MouseEnter, AddressOf Label_MouseEnter
                End With

                With TailLabel
                    .Text = "pt"
                    .AutoSize = False
                    .Location = New Point(Width - TailLabelWidth + 4, 4)
                    AddHandler .MouseEnter, AddressOf Label_MouseEnter
                End With

                With Me
                    .Height = 22
                    .Width = Width
                    .Margin = New Padding(0, 0, 0, 4)
                    .Controls.Add(HeadLabel)
                    .Controls.Add(ComboBox)
                    .Controls.Add(TailLabel)
                    AddHandler .MouseEnter, AddressOf Me_MouseEnter
                End With

                FillComboBox()
            End Sub

            ''' <summary>
            ''' When mouse enters the panel, this sub will be called.
            ''' </summary>
            Private Sub Me_MouseEnter()
                RaiseEvent SubControlMouseEnter(Me, Nothing)
            End Sub

            ''' <summary>
            ''' When mouse enters the two labels, this sub will be called.
            ''' </summary>
            Private Sub Label_MouseEnter()
                RaiseEvent SubControlMouseEnter(Me, Nothing)
            End Sub

            ''' <summary>
            ''' When mouse enters the ComboBox, this sub will be called.
            ''' </summary>
            Private Sub ComboBox_MouseEnter()
                RaiseEvent SubControlMouseEnter(Me, Nothing)
            End Sub

            ''' <summary>
            ''' When the selected index of ComboBox is changed, this sub will be called.
            ''' </summary>
            Public Sub ComboBox_SelectedIndexChanged()
                Me.SelectedText = ComboBox.SelectedItem
                RaiseEvent SelectedChanged()
            End Sub

            ''' <summary>
            ''' I override this sub to control the dropped items more freely.
            ''' </summary>
            ''' <param name="sender"></param>
            ''' <param name="e"></param>
            Public Overridable Sub ComboBox_DrawItem(ByVal sender As Object, ByVal e As DrawItemEventArgs)
                If (e.Index >= 0) Then
                    e.DrawBackground()

                    Dim Brush As Brush
                    If ((e.State And DrawItemState.Selected) _
                        <> DrawItemState.None) Then
                        Brush = SystemBrushes.HighlightText
                    Else
                        Brush = SystemBrushes.WindowText
                    End If

                    e.Graphics.DrawString(ComboBox.Items(e.Index), Font, Brush, e.Bounds.X + 3, e.Bounds.Y + 1)
                    ' Draw the focus Rectangle if appropriate
                    If ((e.State And DrawItemState.NoFocusRect) _
                        = DrawItemState.None) Then
                        e.DrawFocusRectangle()
                    End If
                End If
            End Sub

            ''' <summary>
            ''' This sub is used to initialize the dropped items of the ComboBox.
            ''' </summary>
            Public Overridable Sub FillComboBox()
                For Each FontSize As Integer In {5, 5.5, 6.5, 7.5, 9, 10.5, 12, 14, 15, 16, 18, 22, 24, 26, 36, 42, 48, 54, 63, 72}
                    ComboBox.Items.Add(FontSize)
                Next
            End Sub
        End Class
    End Namespace
End Namespace