Namespace _FormConfiguration
    Namespace FontConfiguration
        Public Class FontColorComboBox
            Inherits Windows.Forms.Panel

            Dim Label As New Label
            Dim ComboBox As New ColorPicker
            Dim CheckBox As New CheckBox

            Public Const HeadLabelWidth As Integer = 75
            Public Const CheckBoxWidth As Integer = 50

            Public Property SelectedColor As Integer
                Get
                    Return ComboBox.SelectedItem
                End Get
                Set(ByVal Argb As Integer)
                    ComboBox.SelectedItem = Argb
                End Set
            End Property

            Public Property Bold As Boolean
                Get
                    Return CheckBox.Checked
                End Get
                Set(ByVal Value As Boolean)
                    CheckBox.Checked = Value
                End Set
            End Property


            Public Event SelectedChanged()


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
                End With

                With Label
                    .Text = Text
                    .AutoSize = True
                    .Location = New Point(0, 4)
                End With

                With CheckBox
                    .Text = "Bold"
                    .AutoSize = False
                    .Location = New Point(Width - CheckBoxWidth + 8, -1)
                    AddHandler .CheckedChanged, AddressOf CheckedChanged_CheckedChanged
                End With

                With Me
                    .Height = 25
                    .Width = Width
                    .Margin = New Padding(0, 0, 0, 4)
                    .Controls.Add(Label)
                    .Controls.Add(ComboBox)
                    .Controls.Add(CheckBox)
                End With
            End Sub

            Private Sub ComboBox_SelectedIndexChanged()
                RaiseEvent SelectedChanged()
            End Sub

            Private Sub CheckedChanged_CheckedChanged()
                RaiseEvent SelectedChanged()
            End Sub
        End Class
    End Namespace
End Namespace
