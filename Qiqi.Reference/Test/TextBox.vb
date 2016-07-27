Imports Qiqi.Reference._BibTeX

Namespace _Test
    Public Class QiqiTextBox
        Inherits System.Windows.Forms.RichTextBox

        Public Sub New()
            With Me
                .Multiline = True
                .ScrollBars = Windows.Forms.ScrollBars.None
                .Dock = DockStyle.Fill
                .WordWrap = True
                .BorderStyle = Windows.Forms.BorderStyle.None
                .ScrollBars = RichTextBoxScrollBars.Vertical

            End With

            Dim Reader As New IO.StreamReader(Application.StartupPath & "\TestDataBase\SPChar.bib")
            Me.Text = Reader.ReadToEnd
            Me.Font = New Font("Consolas", 10)
            Reader.Close()
        End Sub

        'Private Sub Me_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.TextInput
        '    Me.Font = New Font(Me.Font.FontFamily, Me.Font.Size)
        'End Sub




        'Private Sub Me_Invalidated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Invalidated
        '    RemoveHandler Me.Invalidated, AddressOf Me_Invalidated
        '    Me.Font = New Font(Me.Font.FontFamily, Me.Font.Size)
        'End Sub
    End Class

    Public Class SyntaxRTB

        Inherits System.Windows.Forms.RichTextBox



        Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer



        Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Integer) As Integer



        Private _SyntaxHighlight_CaseSensitive As Boolean = False

        Private Words As New DataTable



        'Contains Windows Messages for the SendMessage API call

        Private Enum EditMessages

            LineIndex = 187

            LineFromChar = 201

            GetFirstVisibleLine = 206

            CharFromPos = 215

            PosFromChar = 1062

        End Enum



        Protected Overrides Sub OnTextChanged(ByVal e As System.EventArgs)

            ColorVisibleLines()

        End Sub



        Public Sub ColorRtb()

            Dim FirstVisibleChar As Integer

            Dim i As Integer = 0

            While i < Me.Lines.Length

                FirstVisibleChar = GetCharFromLineIndex(i)

                ColorLineNumber(i, FirstVisibleChar)

                i += 1

            End While

        End Sub



        Public Sub ColorVisibleLines()

            Dim FirstLine As Integer = FirstVisibleLine()

            Dim LastLine As Integer = LastVisibleLine()

            Dim FirstVisibleChar As Integer

            If (FirstLine = 0) And (LastLine = 0) Then

                'If there is no text it will error, so exit the sub

                Exit Sub

            Else

                While FirstLine < LastLine

                    FirstVisibleChar = GetCharFromLineIndex(FirstLine)

                    ColorLineNumber(FirstLine, FirstVisibleChar)

                    FirstLine += 1

                End While

            End If

        End Sub



        Public Sub ColorLineNumber(ByVal LineIndex As Integer, ByVal lStart As Integer)

            Dim i As Integer = 0

            'Dim Instance As Integer

            'Dim LeadingChar, TrailingChar As String

            Dim SelectionAt As Integer = Me.SelectionStart

            Dim MyRow As DataRow

            Dim Line() As String, MyI As Integer, MyStr As String

            ' Lock the update

            LockWindowUpdate(Me.Handle.ToInt32)

            MyI = lStart

            If CaseSensitive Then

                Line = Split(Me.Lines(LineIndex).ToString, " ")

            Else

                Line = Split(Me.Lines(LineIndex).ToLower, " ")

            End If

            For Each MyStr In Line

                Me.SelectionStart = MyI

                Me.SelectionLength = MyStr.Length

                If Words.Rows.Contains(MyStr) Then

                    MyRow = Words.Rows.Find(MyStr)

                    If (Not CaseSensitive) Or (CaseSensitive And MyRow("Word") = MyStr) Then

                        Me.SelectionColor = Color.FromName(MyRow("Color"))

                    End If

                Else

                    Me.SelectionColor = Color.Black

                End If

                MyI += MyStr.Length + 1

            Next

            ' Restore the selectionstart

            Me.SelectionStart = SelectionAt

            Me.SelectionLength = 0

            Me.SelectionColor = Color.Black

            ' Unlock the update

            LockWindowUpdate(0)

        End Sub



        Public Function GetCharFromLineIndex(ByVal LineIndex As Integer) As Integer

            Return SendMessage(Me.Handle, EditMessages.LineIndex, LineIndex, 0)

        End Function



        Public Function FirstVisibleLine() As Integer

            Return SendMessage(Me.Handle, EditMessages.GetFirstVisibleLine, 0, 0)

        End Function



        Public Function LastVisibleLine() As Integer

            Dim LastLine As Integer = FirstVisibleLine() + (Me.Height / Me.Font.Height)

            If LastLine > Me.Lines.Length Or LastLine = 0 Then

                LastLine = Me.Lines.Length

            End If

            Return LastLine

        End Function



        Public Sub New()

            Dim MyRow As DataRow

            Dim arrKeyWords() As String, strKW As String

            Me.AcceptsTab = True

            ''Load all the keywords and the colors to make them

            Words.Columns.Add("Word")

            Words.PrimaryKey = New DataColumn() {Words.Columns(0)}

            Words.Columns.Add("Color")

            arrKeyWords = New String() {"select", "insert", "delete",
"truncate", "from", "where", "into", "inner", "update",
"outer", "on", "is", "declare", "set", "use", "values", "as",
"order", "by", "drop", "view", "go", "trigger", "cube",
"binary", "varbinary", "image", "char", "varchar", "text",
"datetime", "smalldatetime", "decimal", "numeric", "float",
"real", "bigint", "int", "smallint", "tinyint", "money",
"smallmoney", "bit", "cursor", "timestamp", "uniqueidentifier",
"sql_variant", "table", "nchar", "nvarchar", "ntext", "left",
"right", "like", "and", "all", "in", "null", "join", "not", "or"}

            For Each strKW In arrKeyWords

                MyRow = Words.NewRow()

                MyRow("Word") = strKW

                MyRow("Color") = Color.LightCoral.Name

                Words.Rows.Add(MyRow)

            Next

        End Sub



        Public Property CaseSensitive() As Boolean

            Get

                Return _SyntaxHighlight_CaseSensitive

            End Get

            Set(ByVal Value As Boolean)

                _SyntaxHighlight_CaseSensitive = Value

            End Set

        End Property





    End Class


End Namespace

