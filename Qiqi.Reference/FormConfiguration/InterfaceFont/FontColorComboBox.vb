Namespace _FormConfiguration
    Namespace InterfaceFont
        ''' <summary>
        ''' This class is ComboBox with a head label.
        ''' It can be used to let the user select a color.
        ''' </summary>
        Public Class FontColorComboBox
            Inherits Windows.Forms.Panel

            ' This is the head label.
            Dim Label As New Label
            ' This is a color picker which is used to let the user select a color.
            Dim ComboBox As New ColorPicker
            ' This is a CheckBox which is used to let the user select the normal font or the bold font.
            Dim CheckBox As New CheckBox

            ' These two widths are width of the head label and the width of the tail width.
            Public Const HeadLabelWidth As Integer = 75
            Public Const CheckBoxWidth As Integer = 50

            ' The following event will be triggered when the mouse enters the ComboBox, the head label, or the CheckBox.
            Public Event SubControlMouseEnter(sender As Object, e As System.EventArgs)

            ''' <summary>
            ''' This property is the selected color of the ComboBox.
            ''' When this property is assigned, the ComboBox's selected index will be changed.
            ''' In this class, I used a integer to repesent a RGB color.
            ''' </summary>
            ''' <returns></returns>
            Public Property SelectedColor As Integer
                Get
                    Return ComboBox.SelectedItem
                End Get
                Set(ByVal Argb As Integer)
                    ComboBox.SelectedItem = Argb
                End Set
            End Property

            ''' <summary>
            ''' This property is the value of the CheckBox.
            ''' </summary>
            ''' <returns></returns>
            Public Property Bold As Boolean
                Get
                    Return CheckBox.Checked
                End Get
                Set(ByVal Value As Boolean)
                    CheckBox.Checked = Value
                End Set
            End Property

            ''' <summary>
            ''' The following event will be triggered when the selected color of ComboBox is changed,
            ''' or the checked value of the CheckBox is changed.
            ''' </summary>
            Public Event SelectedChanged()

            ''' <summary>
            ''' Constructor.
            ''' </summary>
            ''' <param name="Text">The text of the label.</param>
            ''' <param name="Width">The width of the whole panel.</param>
            Public Sub New(ByVal Text As String, ByVal Width As Integer)
                With ComboBox
                    .DropDownStyle = ComboBoxStyle.DropDownList
                    .Margin = New Padding(0)
                    .FlatStyle = FlatStyle.System
                    .Location = New Point(HeadLabelWidth, 0)
                    .Size = New Size(Width - HeadLabelWidth - CheckBoxWidth, 15)
                    .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
                    .AddStandardColors()
                    AddHandler .SelectedIndexChanged, AddressOf ComboBox_SelectedIndexChanged
                    AddHandler .MouseEnter, AddressOf ComboBox_MouseEnter
                End With

                With Label
                    .Text = Text
                    .AutoSize = True
                    .Location = New Point(0, 4)
                    AddHandler .MouseEnter, AddressOf Label_MouseEnter
                End With

                With CheckBox
                    .Text = "Bold"
                    .AutoSize = False
                    .Location = New Point(Width - CheckBoxWidth + 8, -1)
                    AddHandler .CheckedChanged, AddressOf CheckedChanged_CheckedChanged
                    AddHandler .MouseEnter, AddressOf CheckBox_MouseEnter
                End With

                With Me
                    .Height = 22
                    .Width = Width
                    .Margin = New Padding(0, 0, 0, 4)
                    .Controls.Add(Label)
                    .Controls.Add(ComboBox)
                    .Controls.Add(CheckBox)
                    AddHandler .MouseEnter, AddressOf Me_MouseEnter
                End With
            End Sub

            ''' <summary>
            ''' When mouse enters the panel, this sub will be called.
            ''' </summary>
            Private Sub Me_MouseEnter()
                RaiseEvent SubControlMouseEnter(Me, Nothing)
            End Sub

            ''' <summary>
            ''' When mouse enters the label, this sub will be called.
            ''' </summary>
            Private Sub Label_MouseEnter()
                RaiseEvent SubControlMouseEnter(Me, Nothing)
            End Sub

            ''' <summary>
            ''' When mouse enters the CheckBox, this sub will be called.
            ''' </summary>
            Private Sub CheckBox_MouseEnter()
                RaiseEvent SubControlMouseEnter(Me, Nothing)
            End Sub

            ''' <summary>
            ''' When mouse enters the ComboBox, this sub will be called.
            ''' </summary>
            Private Sub ComboBox_MouseEnter()
                RaiseEvent SubControlMouseEnter(Me, Nothing)
            End Sub

            ''' <summary>
            ''' When mouse enters the panel, this sub will be called.
            ''' </summary>
            Private Sub ComboBox_SelectedIndexChanged()
                RaiseEvent SelectedChanged()
            End Sub

            ''' <summary>
            ''' When the checked value of the CheckBox is changed, this sub will be called.
            ''' </summary>
            Private Sub CheckedChanged_CheckedChanged()
                RaiseEvent SelectedChanged()
            End Sub
        End Class
    End Namespace
End Namespace
