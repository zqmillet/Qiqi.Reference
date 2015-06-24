Imports System.IO

Namespace _BibTeX
    Public Module DataBaseError
        Public Const IdleMode As String = "Idle mode error."
        Public Const CommentModel As String = "Comment model error."
        Public Const ReadType As String = "Read type error."
        Public Const ReadBibTeXKey As String = "Read BibTeXKey error."
        Public Const ReadBuffer As String = "Read buffer error."
    End Module

    ''' <summary>
    ''' Class DataBase represents a BibTeX database 
    ''' </summary>
    ''' <remarks></remarks>
    Public Class DataBase
        ' The file name and directory of BibTeX file
        Public FileFullName As String
        ' Paper List
        Public LiteratureList As ArrayList

        ' Progress
        Public Progress As Double

        Public Event ProgressUpdate(ByVal Progress As Double)

        Public ExistGroup As Boolean = False
        Public GroupBuffer As String = ""
        Public GroupVersion As Integer = -1


        ''' <summary>
        ''' The constructor with parameter
        ''' </summary>
        ''' <param name="FileFullName">The full name of BibTeX file</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal FileFullName As String)
            With Me
                .FileFullName = FileFullName
                .LiteratureList = New ArrayList
                .ExistGroup = False
                .GroupBuffer = ""
            End With
        End Sub

        ''' <summary>
        ''' Read the BibTeX File, load the information into the paper list
        ''' </summary>
        ''' <returns>If successful, return True, else, return False</returns>
        ''' <remarks></remarks>
        Public Function DataBaseLoading(ByRef ErrorMessage As _BibTeX.ErrorMessage) As Boolean
            ' If there is no file, return False
            If Not My.Computer.FileSystem.FileExists(FileFullName) Then
                Return False
            End If

            ' Generate a text reader
            Dim BibTeXReader As New StreamReader(FileFullName)
            ' Buffer of BibTeX
            Dim BibTeXBuffer As String = ""

            ' Read all BibTeX and write into the BibTeXBuffer
            BibTeXBuffer = BibTeXReader.ReadToEnd

            ' State is used to mark the state of lexical analysis process
            Dim State As Integer = LexicalAnalysisStatus.IdleMode

            ' Parameters of literature
            Dim LiteratureType As String = ""
            Dim LiteratureID As String = ""
            Dim LiteratureBuffer As String = ""

            ' Parameter of comment
            Dim CommentBuffer As String = ""

            ' Counter of brackets, if c = "{", add 1, if c = "}", sub 1
            Dim BracketNumber As Integer = 0

            Dim RaiseEventInterval As Double = (BibTeXBuffer.Length - 1) / 10

            Dim LineNumber As Integer = 0

            ' Lexical analysis of BibTeXBuffer
            For i As Integer = 0 To BibTeXBuffer.Length - 1
                If (vbCr).Contains(BibTeXBuffer(i)) Then
                    LineNumber += 1
                End If

                If i = BibTeXBuffer.Length - 1 Then
                    Me.Progress = 1
                    RaiseEvent ProgressUpdate(Me.Progress)
                ElseIf i >= RaiseEventInterval Then
                    Me.Progress = i / (BibTeXBuffer.Length - 1)
                    RaiseEventInterval += (BibTeXBuffer.Length - 1) / 10
                    RaiseEvent ProgressUpdate(Me.Progress)
                End If

                Select Case State
                    Case LexicalAnalysisStatus.IdleMode
                        Select Case BibTeXBuffer(i)
                            Case "%"
                                State = LexicalAnalysisStatus.CommentModel
                            Case "@"
                                LiteratureType = ""
                                LiteratureID = ""
                                LiteratureBuffer = ""
                                BracketNumber = 0
                                State = LexicalAnalysisStatus.ReadType
                                ErrorMessage.LineNumber = LineNumber
                            Case " "
                                ' Do nothing
                            Case vbCr
                                ' Do nothing
                            Case ""
                                ' Do nothing
                            Case Chr(10)
                                ' Do nothing
                            Case Else
                                ' DEBUG
                                ErrorMessage.SetErrorMessage(DataBaseError.IdleMode)
                                Return False
                        End Select
                    Case LexicalAnalysisStatus.CommentModel
                        Select Case BibTeXBuffer(i)
                            Case vbCr
                                State = LexicalAnalysisStatus.IdleMode
                        End Select
                    Case LexicalAnalysisStatus.ReadType
                        Select Case BibTeXBuffer(i)
                            Case " "
                                ErrorMessage.SetErrorMessage(DataBaseError.ReadType)
                                Return False
                            Case ","
                                ErrorMessage.SetErrorMessage(DataBaseError.ReadType)
                                Return False
                            Case vbCr
                                ErrorMessage.SetErrorMessage(DataBaseError.ReadType)
                                Return False
                            Case "}"
                                ErrorMessage.SetErrorMessage(DataBaseError.ReadType)
                                Return False
                            Case "@"
                                ErrorMessage.SetErrorMessage(DataBaseError.ReadType)
                                Return False
                            Case "{"
                                BracketNumber += 1
                                If LiteratureType.ToLower.Trim = "comment" Then
                                    State = LexicalAnalysisStatus.CommentType
                                    CommentBuffer = ""
                                Else
                                    State = LexicalAnalysisStatus.ReadBibTeXKey
                                End If
                            Case Else
                                LiteratureType &= BibTeXBuffer(i)
                        End Select
                    Case LexicalAnalysisStatus.ReadBibTeXKey
                        Select Case BibTeXBuffer(i)
                            Case " "
                                ErrorMessage.SetErrorMessage(DataBaseError.ReadBibTeXKey)
                                Return False
                            Case ","
                                State = LexicalAnalysisStatus.ReadBuffer
                            Case vbCr
                                ErrorMessage.SetErrorMessage(DataBaseError.ReadBibTeXKey)
                                Return False
                            Case "}"
                                ErrorMessage.SetErrorMessage(DataBaseError.ReadBibTeXKey)
                                Return False
                            Case "@"
                                ErrorMessage.SetErrorMessage(DataBaseError.ReadBibTeXKey)
                                Return False
                            Case "{"
                                ErrorMessage.SetErrorMessage(DataBaseError.ReadBibTeXKey)
                                Return False
                            Case Else
                                LiteratureID &= BibTeXBuffer(i)
                        End Select
                    Case LexicalAnalysisStatus.ReadBuffer
                        Select Case BibTeXBuffer(i)
                            Case "{"
                                BracketNumber += 1
                                LiteratureBuffer &= BibTeXBuffer(i)
                            Case "}"
                                BracketNumber -= 1
                                If BracketNumber = 0 Then
                                    Dim Literature As New _BibTeX.Literature(LiteratureType, LiteratureID, LiteratureBuffer, ErrorMessage)
                                    If ErrorMessage.HasError Then
                                        ErrorMessage.SetErrorMessage(DataBaseError.ReadBuffer)
                                        Return False
                                    End If

                                    Me.LiteratureList.Add(Literature)
                                    State = LexicalAnalysisStatus.IdleMode
                                Else
                                    LiteratureBuffer &= BibTeXBuffer(i)
                                End If
                            Case Else
                                LiteratureBuffer &= BibTeXBuffer(i)
                        End Select
                    Case LexicalAnalysisStatus.CommentType
                        Select Case BibTeXBuffer(i)
                            Case "{"
                                BracketNumber += 1
                            Case "}"
                                BracketNumber -= 1
                                If BracketNumber = 0 Then
                                    ' Do nothing
                                    State = LexicalAnalysisStatus.IdleMode
                                    CommentAnalysis(CommentBuffer)
                                End If
                            Case Else
                                CommentBuffer &= BibTeXBuffer(i)
                        End Select

                End Select
            Next

            If State = LexicalAnalysisStatus.IdleMode Then
                Return True
            Else
                ErrorMessage.SetErrorMessage(DataBaseError.IdleMode)
                Return False
            End If
        End Function

        Private Sub CommentAnalysis(ByRef CommentBuffer As String)
            If Not CommentBuffer.Trim.ToLower.IndexOf("jabref-meta:") = 0 Then
                Exit Sub
            End If

            CommentBuffer = CommentBuffer.Remove(0, CommentBuffer.IndexOf(":") + 1).Trim

            If CommentBuffer.ToLower.IndexOf("groupsversion:") = 0 Then
                CommentBuffer = CommentBuffer.Remove(0, CommentBuffer.IndexOf(":") + 1).Trim
                CommentBuffer = CommentBuffer.Remove(CommentBuffer.IndexOf(";"))
                If IsNumeric(CommentBuffer) Then
                    GroupVersion = Val(CommentBuffer)
                End If
            End If

            If CommentBuffer.ToLower.IndexOf("groupstree:") = 0 Then
                GroupBuffer = CommentBuffer.Remove(0, CommentBuffer.IndexOf(":") + 1).Trim
            End If

            If GroupVersion >= 0 And GroupBuffer <> "" Then
                ExistGroup = True
            Else
                ExistGroup = False
            End If
        End Sub


        ''' <summary>
        ''' Judge the existance of whole literature
        ''' </summary>
        ''' <param name="BibTeXBuffer">BibTeX buffer</param>
        ''' <returns>If there is a whole literature, return True, else, return False</returns>
        ''' <remarks></remarks>
        Private Function ExistWholeLiterature(ByRef BibTeXBuffer As String) As Boolean
            ' If there is no char "@" in the buffer
            If Not BibTeXBuffer.Contains("@") Then
                ' There is no a whole literature, return False
                Return False
            End If



            ' Remove all chars before the first char "@"
            BibTeXBuffer = BibTeXBuffer.Remove(0, BibTeXBuffer.IndexOf("@"))

            ' BracketCounter is used to count the number of brackets
            Dim BracketCounter As Integer = 0
            ' MeetFirstLeftBracket is used to record if there a char "{"
            Dim MeetFirstLeftBracket As Boolean = False

            ' For each c in BibTeXBuffer
            For Each c As Char In BibTeXBuffer
                ' Count the number of the brackets
                If c = "{" Then
                    BracketCounter += 1
                    MeetFirstLeftBracket = True
                ElseIf c = "}" Then
                    BracketCounter -= 1
                End If

                ' If there exist at least one char "{" and the numbers of "{" and "}" is equal
                If BracketCounter = 0 And MeetFirstLeftBracket Then
                    ' There exists a literature, return True
                    Return True
                End If
            Next

            ' There is no literature, return False
            Return False
        End Function

        Private Function ExtractWholeLiterature(ByRef BibTeXBuffer As String) As String
            ' BracketCounter is used to count the number of brackets
            Dim BracketCounter As Integer = 0
            ' MeetFirstLeftBracket is used to record if there a char "{"
            Dim MeetFirstLeftBracket As Boolean = False
            ' Recode the position of the end of the first literature
            Dim Index As Integer = 0

            ' For each c in BibTeXBuffer
            For Each c As Char In BibTeXBuffer
                ' Count the number of the brackets
                If c = "{" Then
                    BracketCounter += 1
                    MeetFirstLeftBracket = True
                ElseIf c = "}" Then
                    BracketCounter -= 1
                End If

                ' Update the position
                Index += 1

                ' If there exist at least one char "{" and the numbers of "{" and "}" is equal
                If BracketCounter = 0 And MeetFirstLeftBracket Then
                    ' Record the return value
                    Dim TempReturn As String = BibTeXBuffer.Remove(Index)
                    ' Remove the return value from buffer
                    BibTeXBuffer = BibTeXBuffer.Remove(0, TempReturn.Length)
                    ' Return literaturn
                    Return TempReturn
                End If
            Next

            'If there is no literature, return False
            Return ""
        End Function
    End Class
End Namespace


