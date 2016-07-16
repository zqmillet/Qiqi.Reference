Imports Qiqi.Reference._BibTeX

Namespace _FormMain
    Namespace _LiteratureTabControl
        Public Class RichTextBox
            Inherits Windows.Forms.RichTextBox

            ' Member variables
            Public SyntaxHighLight As Boolean = False
            Public FontSize As Integer = 10
            Public BoldFont As Boolean = False
            Public TextFont As Font = Me.Font

            Public EntryTypeColor As Color = Color.Blue
            Public BibTeXKeyColor As Color = Color.Red
            Public TagNameColor As Color = Color.Green
            Public TagValueColor As Color = Color.Black

            ''' <summary>
            ''' Messages
            ''' </summary>
            Private Enum EditMessages
                LineIndex = 187
                LineFromChar = 201
                GetFirstVisibleLine = 206
                CharFromPos = 215
                PosFromChar = 1062
            End Enum

            ' Windows APIs
            Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
            Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Integer) As Integer

            ''' <summary>
            ''' This function is used to disable the zoom function
            ''' </summary>
            ''' <param name="m"></param>
            Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
                ' Windows message constant for scrollwheel moving
                Const WM_SCROLLWHEEL As Integer = &H20A

                Dim ScrollingAndPressingControl As Boolean = m.Msg = WM_SCROLLWHEEL AndAlso Control.ModifierKeys = Keys.Control

                ' If scolling and pressing control then do nothing (don't let the base class know), 
                ' otherwise send the info down to the base class as normal
                If (Not ScrollingAndPressingControl) Then
                    MyBase.WndProc(m)
                End If
            End Sub

            ''' <summary>
            ''' Construction function
            ''' </summary>
            Public Sub New()
                InitializeStyle()
            End Sub

            ''' <summary>
            ''' Construction function with parameter
            ''' </summary>
            ''' <param name="Text"></param>
            Public Sub New(ByVal Text As String)
                Me.Text = Text
                InitializeStyle()
            End Sub

            ''' <summary>
            ''' This sub is used to initialize the style of the RichTextBox
            ''' </summary>
            Private Sub InitializeStyle()
                With Me
                    .BorderStyle = BorderStyle.None
                    .Dock = DockStyle.Fill
                End With
            End Sub

            ''' <summary>
            ''' This sub 
            ''' </summary>
            ''' <param name="BeginIndex"></param>
            ''' <param name="EndIndex"></param>
            Private Sub SelectText(ByVal BeginIndex As Integer, ByVal EndIndex As Integer)
                Me.SelectionStart = 0
                Me.SelectionLength = 0

                Dim SpaceString As String = " " & vbCr & vbLf & vbCrLf
                For i As Integer = BeginIndex To EndIndex
                    If Not SpaceString.Contains(Me.Text(i)) Then
                        Me.SelectionStart = i
                        Exit For
                    End If
                Next

                For i As Integer = EndIndex To BeginIndex Step -1
                    If Not SpaceString.Contains(Me.Text(i)) Then
                        Me.SelectionLength = i - Me.SelectionStart + 1
                        Exit For
                    End If
                Next
            End Sub

            ''' <summary>
            ''' This function is used to get the number of the first visible line of the RichTextBox
            ''' </summary>
            ''' <returns></returns>
            Public Function GetFirstVisibleLine() As Integer
                Return SendMessage(Me.Handle, EditMessages.GetFirstVisibleLine, 0, 0)
            End Function

            ''' <summary>
            ''' Overrides the event TextChanged
            ''' </summary>
            ''' <param name="e"></param>
            Protected Overrides Sub OnTextChanged(ByVal e As System.EventArgs)
                If SyntaxHighLight Then
                    TextSyntaxHighLight()
                Else
                    TextFontUpdate()
                End If
            End Sub

            Public Sub TextFontUpdate()
                If Me.Text.Trim = "" Then
                    Exit Sub
                End If

                Dim SelectionAt As Integer = Me.SelectionStart
                Dim FirstLineIndex As Integer = GetFirstVisibleLine()

                Dim NormalFont As New Font(Me.TextFont.FontFamily, Me.FontSize)

                LockWindowUpdate(Me.Handle.ToInt32)

                Me.SelectionStart = 0
                Me.SelectionLength = Me.TextLength
                Me.SelectionFont = NormalFont
                Me.SelectionColor = Color.Black

                ' Restore the selectionstart
                Me.SelectionStart = Me.Find(Me.Lines(FirstLineIndex), RichTextBoxFinds.NoHighlight)
                Me.ScrollToCaret()

                Me.SelectionStart = SelectionAt
                Me.SelectionLength = 0

                ' Unlock the update
                LockWindowUpdate(0)
            End Sub

            Public Sub TextSyntaxHighLight()
                If Me.Text.Trim = "" Then
                    Exit Sub
                End If

                Dim SelectionAt As Integer = Me.SelectionStart
                Dim FirstLineIndex As Integer = GetFirstVisibleLine()

                Dim StartIndex As Integer = 0
                Dim EndIndex As Integer = 0
                Dim State As LexicalAnalysisStatus = LexicalAnalysisStatus.IdleMode
                Dim BracketNumber As Integer = 0

                ' Lock the update
                LockWindowUpdate(Me.Handle.ToInt32)

                Dim BoldFont As New Font(Me.TextFont.FontFamily, Me.FontSize, FontStyle.Bold)
                Dim NormalFont As New Font(Me.TextFont.FontFamily, Me.FontSize)

                Me.SelectionStart = 0
                Me.SelectionLength = Me.TextLength
                Me.SelectionFont = NormalFont
                Me.SelectionColor = Color.Black

                ' Syntax highlight
                ' High light entry type
                StartIndex = Me.Text.IndexOf("@") + 1
                EndIndex = Me.Text.Substring(StartIndex).IndexOf("{")
                SelectText(StartIndex, EndIndex)
                Me.SelectionColor = EntryTypeColor
                ' High light BibTeXKey
                StartIndex = EndIndex + 2
                EndIndex = Me.Text.Substring(StartIndex).IndexOf(",") + StartIndex - 1
                SelectText(StartIndex, EndIndex)
                Me.SelectionColor = BibTeXKeyColor

                EndIndex += 1
                Do  ' High light tag name
                    StartIndex = EndIndex + 1

                    EndIndex = Me.Text.Substring(StartIndex).IndexOf("=") + StartIndex - 1
                    SelectText(StartIndex, EndIndex)
                    Me.SelectionColor = TagNameColor
                    ' High light tag value
                    If GetBracketIndexes(StartIndex, EndIndex) Then
                        SelectText(StartIndex + 1, EndIndex - 1)
                        Me.SelectionColor = Color.Red
                    Else
                        Exit Do
                    End If

                    ' Wait comma

                    If Me.Text.Substring(EndIndex).IndexOf("=") < 0 Then
                        Exit Do
                    End If
                    EndIndex = Me.Text.Substring(EndIndex).IndexOf(",") + EndIndex
                Loop While True

                ' Restore the selectionstart
                Me.SelectionStart = Me.Find(Me.Lines(FirstLineIndex), RichTextBoxFinds.NoHighlight)
                Me.ScrollToCaret()

                Me.SelectionStart = SelectionAt
                Me.SelectionLength = 0

                ' Unlock the update
                LockWindowUpdate(0)
            End Sub

            Private Function BraceMatching(ByVal Code As String)
                Return Code.Replace("{", "").Length = Code.Replace("}", "").Length
            End Function

            Private Function GetBracketIndexes(ByRef StartIndex As Integer, ByRef EndIndex As Integer) As Boolean
                StartIndex = Me.Text.Substring(StartIndex).IndexOf("{") + StartIndex
                EndIndex = StartIndex

                Do
                    Dim Delta As Integer = Me.Text.Substring(EndIndex + 1).IndexOf("}") + 1
                    EndIndex = Delta + EndIndex
                    If Delta < 0 Or Me.Text.Substring(EndIndex + 1).Trim.Length = 0 Then
                        Return False
                    End If

                    If StartIndex < 0 Or EndIndex < 0 Then
                        Return False
                    End If
                Loop While Not BraceMatching(Me.Text.Substring(StartIndex, EndIndex - StartIndex + 1))
                Return True
            End Function

        End Class
    End Namespace
End Namespace
