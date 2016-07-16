Imports System.ComponentModel

Namespace _FormMain
    Namespace _LiteratureTabControl
        Public Class MultiLineTextBox
            Inherits System.Windows.Forms.Panel
            Implements INotifyPropertyChanged

            Public TextBox As _FormMain._LiteratureTabControl.RichTextBox
            ' Private SyntaxHighlightValue As Boolean = False

            Public Event PropertyChanged As PropertyChangedEventHandler _
                Implements INotifyPropertyChanged.PropertyChanged

            Public Property SyntaxHighlight() As String
                Get
                    Return Me.TextBox.SyntaxHighLight
                End Get

                Set(ByVal Value As String)
                    Me.TextBox.SyntaxHighLight = Value
                End Set
            End Property

            Public Sub New()
                TextBox = New RichTextBox

                With TextBox
                    .Text = ""
                    .Multiline = True
                    .Padding = New Padding(0)
                    .Dock = DockStyle.Fill
                    .BorderStyle = Windows.Forms.BorderStyle.None
                    .ScrollBars = ScrollBars.Vertical
                End With

                With Me
                    .Padding = New Padding(3)
                    .Margin = New Padding(0)
                    .Dock = DockStyle.Fill
                    .BorderStyle = Windows.Forms.BorderStyle.None
                    .BackColor = Color.White
                    .Controls.Add(TextBox)
                End With
            End Sub

            Public Sub New(ByVal Text As String)
                TextBox = New RichTextBox

                With TextBox
                    .Text = Text
                    .Multiline = True
                    .Padding = New Padding(0)
                    .Dock = DockStyle.Fill
                    .BorderStyle = Windows.Forms.BorderStyle.None
                    .ScrollBars = ScrollBars.Vertical
                End With

                With Me
                    .Padding = New Padding(3)
                    .Margin = New Padding(0)
                    .Dock = DockStyle.Fill
                    .BorderStyle = Windows.Forms.BorderStyle.None
                    .BackColor = Color.White
                    .Controls.Add(TextBox)
                End With
            End Sub
        End Class
    End Namespace
End Namespace