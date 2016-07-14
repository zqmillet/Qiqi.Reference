Imports Qiqi.Reference._BibTeX

Namespace _FormMain
    Namespace _LiteratureTabControl
        Public Class SourceCodeTextBox
            Inherits _FormMain._LiteratureTabControl.MultiLineTextBox

            Private FontSize As Integer = 10

            Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
            Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Integer) As Integer
            Public HighLight As Boolean = False

            Private Enum EditMessages
                LineIndex = 187
                LineFromChar = 201
                GetFirstVisibleLine = 206
                CharFromPos = 215
                PosFromChar = 1062
            End Enum

            Public Sub New()
                Me.TextBox.AcceptsTab = True
                AddHandler TextBox.TextChanged, AddressOf TextBox_TextChanged
            End Sub

            Public Sub New(ByVal Text As String)
                Me.TextBox.Text = Text
                AddHandler TextBox.TextChanged, AddressOf TextBox_TextChanged
                TextBox_TextChanged(Nothing, Nothing)
            End Sub

            Public Sub New(ByVal Text As String, ByVal FontSize As Integer)
                Me.TextBox.Text = Text
                Me.FontSize = FontSize
                AddHandler TextBox.TextChanged, AddressOf TextBox_TextChanged
                TextBox_TextChanged(Nothing, Nothing)
            End Sub

            Public Function GetFirstVisibleLine() As Integer
                Return SendMessage(Me.TextBox.Handle, EditMessages.GetFirstVisibleLine, 0, 0)
            End Function

            Protected Sub TextBox_TextChanged(sender As Object, e As EventArgs)

                Dim SelectionAt As Integer = TextBox.SelectionStart
                Dim Index As Integer = 0
                Dim StartIndex As Integer = 0
                Dim State As LexicalAnalysisStatus = LexicalAnalysisStatus.IdleMode
                Dim BracketNumber As Integer = 0
                Dim FirstLineIndex As Integer = GetFirstVisibleLine()

                ' Lock the update
                LockWindowUpdate(TextBox.Handle.ToInt32)

                Dim BoldFont As New Font("Consolas", FontSize, FontStyle.Bold)
                Dim NormalFont As New Font("Consolas", FontSize)

                TextBox.SelectionStart = 0
                TextBox.SelectionLength = TextBox.TextLength
                TextBox.SelectionFont = NormalFont
                TextBox.SelectionColor = Color.Black

                If HighLight Then
                    ' Syntax highlight
                    For Each c As Char In Me.TextBox.Text
                        Select Case State
                            Case LexicalAnalysisStatus.IdleMode
                                Select Case c
                                    Case "@"
                                        State = LexicalAnalysisStatus.ReadType
                                        StartIndex = Index
                                    Case Else
                                        ' do nothing
                                End Select
                            Case LexicalAnalysisStatus.ReadType
                                Select Case c
                                    Case "{"
                                        State = LexicalAnalysisStatus.ReadBibTeXKey
                                        Me.SelectText(StartIndex + 1, Index - 1)
                                        TextBox.SelectionColor = Color.Blue
                                        TextBox.SelectionFont = BoldFont
                                        StartIndex = Index
                                    Case Else
                                        ' do nothing
                                End Select
                            Case LexicalAnalysisStatus.ReadBibTeXKey
                                Select Case c
                                    Case ","
                                        State = LexicalAnalysisStatus.ReadName
                                        Me.SelectText(StartIndex + 1, Index - 1)
                                        TextBox.SelectionColor = Color.Red
                                        StartIndex = Index
                                    Case Else
                                        ' do nothing
                                End Select
                            Case LexicalAnalysisStatus.ReadName
                                Select Case c
                                    Case "="
                                        State = LexicalAnalysisStatus.BracketMode
                                        Me.SelectText(StartIndex + 1, Index - 1)
                                        TextBox.SelectionColor = Color.Black
                                        TextBox.SelectionFont = BoldFont
                                        StartIndex = Index
                                    Case Else
                                        ' do nothing
                                End Select
                            Case LexicalAnalysisStatus.BracketMode
                                Select Case c
                                    Case "{"
                                        State = LexicalAnalysisStatus.BeginReadValue
                                        StartIndex = Index
                                        BracketNumber = 1
                                    Case Else
                                        ' do thing
                                End Select
                            Case LexicalAnalysisStatus.BeginReadValue
                                Select Case c
                                    Case "{"
                                        BracketNumber += 1
                                    Case "}"
                                        BracketNumber -= 1
                                End Select

                                If BracketNumber = 0 Then
                                    State = LexicalAnalysisStatus.WaitComma
                                    Me.SelectText(StartIndex + 1, Index - 1)
                                    TextBox.SelectionColor = Color.Green
                                End If
                            Case LexicalAnalysisStatus.WaitComma
                                Select Case c
                                    Case ","
                                        State = LexicalAnalysisStatus.ReadName
                                        StartIndex = Index
                                    Case "", vbCr, vbLf, vbCrLf
                                        ' do nothing
                                    Case Else
                                        ' finish syntax highlight
                                End Select
                        End Select
                        Index += 1
                    Next
                End If

                ' Restore the selectionstart
                Me.TextBox.SelectionStart = Me.TextBox.Find(Me.TextBox.Lines(FirstLineIndex), RichTextBoxFinds.NoHighlight)
                Me.TextBox.ScrollToCaret()

                TextBox.SelectionStart = SelectionAt
                TextBox.SelectionLength = 0

                ' Unlock the update
                LockWindowUpdate(0)
            End Sub

            Private Sub SelectText(ByVal BeginIndex As Integer, ByVal EndIndex As Integer)
                Dim SpaceString As String = " " & vbCr & vbLf & vbCrLf
                For i As Integer = BeginIndex To EndIndex
                    If Not SpaceString.Contains(Me.TextBox.Text(i)) Then
                        TextBox.SelectionStart = i
                        Exit For
                    End If
                Next

                For i As Integer = EndIndex To BeginIndex Step -1
                    If Not SpaceString.Contains(Me.TextBox.Text(i)) Then
                        TextBox.SelectionLength = i - TextBox.SelectionStart + 1
                        Exit For
                    End If
                Next
            End Sub
        End Class
    End Namespace
End Namespace


