Namespace _BibTeX
    ''' <summary>
    ''' This modele is used to contain error message of LiteratureBuffer
    ''' </summary>
    ''' <remarks></remarks>
    Public Module LiteratureError
        Public Const Comment As String = "This is not a literature, but a comment."
        Public Const HeadError As String = "The format of buffer's head is wrong."
        Public Const BeginAnalysis As String = "Begin analysis error."
        Public Const ReadName As String = "Read name error."
        Public Const WaitEqualSign As String = "Wait equal sign error."
        Public Const BeginReadValue As String = "Begin read value error."
        Public Const BracketMode As String = "Bracket mode error."
        Public Const EndReadValue As String = "End read value error."
        Public Const QuotationMode As String = "Quotation mode error."
        Public Const WaitComma As String = "Wait comma error."
    End Module

    Public Enum LexicalAnalysisStatus
        BeginAnalysis
        ReadName
        WaitEqualSign
        BeginReadValue
        BracketMode
        EndReadValue
        QuotationMode
        WaitComma
        IdleMode
        CommentModel
        ReadType
        ReadBibTeXKey
        ReadBuffer
        CommentType
    End Enum


    ''' <summary>
    ''' Class Literature represents a literature and its information about type, bibtexkey, etc.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Literature
        ' Type of this literature
        Public Type As String
        ' BibTeXKey of this literature
        Public ID As String
        ' Detail information of this literature
        Public PropertyList As ArrayList

        ''' <summary>
        ''' The constructor without parameter
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            With Me
                .Type = ""
                .ID = ""
                .PropertyList = New ArrayList
            End With
        End Sub

        ''' <summary>
        ''' The constructor with parameters
        ''' </summary>
        ''' <param name="Type">Type of </param>
        ''' <param name="ID"></param>
        ''' <remarks></remarks>
        Public Sub New(ByVal Type As String, ByVal ID As String, ByVal LiteratureBuffer As String, ByRef ErrorMessage As _BibTeX.ErrorMessage)
            With Me
                .Type = Type
                .ID = ID
                .PropertyList = New ArrayList
            End With

            LexicalAnalysis(LiteratureBuffer, ErrorMessage)
        End Sub

        Private Sub Add(ByVal LiteratureProperty As _BibTeX.LiteratureProperty)
            For Each PropertyItem As _BibTeX.LiteratureProperty In Me.PropertyList
                If PropertyItem.Name = LiteratureProperty.Name Then
                    PropertyItem.Value &= ";" & LiteratureProperty.Value
                    Exit Sub
                End If
            Next

            Me.PropertyList.Add(LiteratureProperty)
        End Sub

        ''' <summary>
        ''' Lexical analysis of literature buffer
        ''' </summary>
        ''' <param name="LiteratureBuffer">BibTeX of literature</param>
        ''' <param name="ErrorMessage">If there exist error in LiteratureBuffer, write the error message in ErrorMessage</param>
        ''' <returns>If there is no error, return True, else, return False</returns>
        ''' <remarks></remarks>
        Private Function LexicalAnalysis(ByVal LiteratureBuffer As String, ByRef ErrorMessage As _BibTeX.ErrorMessage) As Boolean
            ' State is used to mark the state of lexical analysis process
            Dim State As Integer = LexicalAnalysisStatus.BeginAnalysis
            ' Temporary variable
            Dim c As Char = ""

            ' The name of property
            Dim Name As String = ""
            ' The value of property
            Dim Value As String = ""
            ' Counter of brackets, if c = "{", add 1, if c = "}", sub 1
            Dim BracketNumber As Integer = 0

            ' Add a space in the end
            ' If not add this space, the last property will not be read
            LiteratureBuffer &= " "

            ' Read the LiteratureBuffer char-by-char
            For i As Integer = 0 To LiteratureBuffer.Length - 1
                If (vbCr & vbCrLf & Chr(10)).Contains(LiteratureBuffer(i)) Then
                    ErrorMessage.LineNumber += 1
                End If

                ' Save the i th char in c
                c = LiteratureBuffer(i)

                ' State machine
                Select Case State
                    Case LexicalAnalysisStatus.BeginAnalysis
                        Select Case c
                            Case " "
                                ' Do nothing
                            Case "="
                                ErrorMessage.Message = LiteratureError.BeginAnalysis
                                Return False
                            Case "{"
                                ErrorMessage.Message = LiteratureError.BeginAnalysis
                                Return False
                            Case "}"
                                ErrorMessage.Message = LiteratureError.BeginAnalysis
                                Return False
                            Case """"
                                ErrorMessage.Message = LiteratureError.BeginAnalysis
                                Return False
                            Case vbCr
                                ' Do nothing
                            Case Chr(10)
                                ' Do nothing
                            Case Else
                                State = LexicalAnalysisStatus.ReadName
                                Name = c
                        End Select
                    Case LexicalAnalysisStatus.ReadName
                        Select Case c
                            Case " "
                                State = LexicalAnalysisStatus.WaitEqualSign
                            Case "="
                                State = LexicalAnalysisStatus.BeginReadValue
                            Case "{"
                                ErrorMessage.Message = LiteratureError.ReadName
                                Return False
                            Case "}"
                                ErrorMessage.Message = LiteratureError.ReadName
                                Return False
                            Case """"
                                ErrorMessage.Message = LiteratureError.ReadName
                                Return False
                            Case vbCr
                                State = LexicalAnalysisStatus.WaitEqualSign
                            Case Chr(10)
                                State = LexicalAnalysisStatus.WaitEqualSign
                            Case Else
                                Name &= c
                        End Select
                    Case LexicalAnalysisStatus.WaitEqualSign
                        Select Case c
                            Case " "
                                ' Do nothing
                            Case "="
                                State = LexicalAnalysisStatus.BeginReadValue
                            Case "{"
                                ErrorMessage.Message = LiteratureError.WaitEqualSign
                                Return False
                            Case "}"
                                ErrorMessage.Message = LiteratureError.WaitEqualSign
                                Return False
                            Case """"
                                ErrorMessage.Message = LiteratureError.WaitEqualSign
                                Return False
                            Case vbCr
                                ' Do nothing
                            Case Chr(10)
                                ' Do nothing
                            Case Else
                                ErrorMessage.Message = LiteratureError.WaitEqualSign
                                Return False
                        End Select
                    Case LexicalAnalysisStatus.BeginReadValue
                        Select Case c
                            Case " "
                                ' Do nothing
                            Case "="
                                ErrorMessage.Message = LiteratureError.BeginReadValue
                                Return False
                            Case "{"
                                State = LexicalAnalysisStatus.BracketMode
                                BracketNumber = 1
                                Value = ""
                            Case "}"
                                ErrorMessage.Message = LiteratureError.BeginReadValue
                                Return False
                            Case """"
                                State = LexicalAnalysisStatus.QuotationMode
                                Value = ""
                            Case vbCr
                                ' Do nothing
                            Case Chr(10)
                                ' Do nothing
                            Case Else
                                ErrorMessage.Message = LiteratureError.BeginReadValue
                                Return False
                        End Select
                    Case LexicalAnalysisStatus.BracketMode
                        Select Case c
                            Case " "
                                Value &= c
                            Case "="
                                Value &= c
                            Case "{"
                                BracketNumber += 1
                            Case "}"
                                BracketNumber -= 1
                                If BracketNumber = 0 Then
                                    If Not Name.Trim = "" Then
                                        Me.Add(New LiteratureProperty(Name, Value))
                                        Name = ""
                                        Value = ""
                                    End If

                                    State = LexicalAnalysisStatus.EndReadValue
                                End If
                            Case """"
                                Value &= c
                            Case vbCr
                                Value &= c
                            Case Chr(10)
                                Value &= c
                            Case Else
                                Value &= c
                        End Select
                    Case LexicalAnalysisStatus.QuotationMode
                        Select Case c
                            Case """"
                                'State = LexicalAnalysisStatus.WaitComma
                                If IsBufferEnd(LiteratureBuffer, i + 1) Then
                                    Me.Add(New LiteratureProperty(Name, Value))
                                    Return True
                                End If

                                If NextIsComma(LiteratureBuffer, i + 1) Then
                                    Dim LeftBuffer As String = LiteratureBuffer.Remove(0, i + 2)
                                    If LexicalAnalysis(LeftBuffer, ErrorMessage) Then
                                        Me.Add(New LiteratureProperty(Name, Value))
                                        ErrorMessage.Clear()
                                        Return True
                                    Else
                                        ' MsgBox(LeftBuffer)
                                    End If
                                End If

                                Value &= c
                            Case Else
                                Value &= c
                        End Select
                    Case LexicalAnalysisStatus.EndReadValue
                        Select Case c
                            Case " "
                                ' Do nothing
                            Case vbCr
                                ' Do nothing
                            Case Chr(10)
                                ' Do nothing
                            Case ","
                                State = LexicalAnalysisStatus.BeginAnalysis
                            Case Else
                                ErrorMessage.Message = LiteratureError.EndReadValue
                                Return False
                        End Select
                    Case Else
                        ' Error state
                End Select
            Next

            'Select Case State
            '    Case LexicalAnalysisStatus.QuotationMode
            '        ErrorMessage = LiteratureError.QuotationMode
            '        Return False
            '    Case Else
            '        Return True
            'End Select

            Return True
        End Function

        Private Function NextIsComma(ByVal LiteratureBuffer As String, ByVal BeginIndex As Integer) As Boolean
            If BeginIndex > LiteratureBuffer.Length - 1 Then
                Return False
            End If

            For i As Integer = BeginIndex To LiteratureBuffer.Length - 1
                Select Case LiteratureBuffer(i)
                    Case " "
                        ' Do nothing
                    Case vbCr
                        ' Do nothing
                    Case ","
                        Return True
                    Case Else
                        Return False
                End Select
            Next

            Return False
        End Function

        Private Function IsBufferEnd(ByVal LiteratureBuffer As String, ByVal BeginIndex As Integer) As Boolean
            If BeginIndex > LiteratureBuffer.Length - 1 Then
                Return True
            End If

            For i As Integer = BeginIndex To LiteratureBuffer.Length - 1
                If LiteratureBuffer(i) <> " " And LiteratureBuffer(i) <> vbCr And LiteratureBuffer(i) <> Chr(10) Then
                    Return False
                End If
            Next

            Return True
        End Function


        ''' <summary>
        ''' Get the value according to the name
        ''' </summary>
        ''' <param name="Name">Property name</param>
        ''' <returns>If there exist the name, return corresponding value, else, return empty string</returns>
        ''' <remarks></remarks>
        Public Function GetPropertyValue(ByVal Name As String) As String
            If Name.ToLower.Trim = "BibTeXKey".ToLower.Trim Then
                Return Me.ID
            End If

            If Name.ToLower.Trim = "EntryType".ToLower.Trim Then
                Return Me.Type
            End If

            For Each LiteratureProperty As LiteratureProperty In Me.PropertyList
                If LiteratureProperty.Name.ToLower.Trim = Name.ToLower.Trim Then
                    Return LiteratureProperty.Value
                End If
            Next

            Return ""
        End Function

        Public Function GetProperty(ByVal Name As String) As _BibTeX.LiteratureProperty
            If Name.ToLower.Trim = "BibTeXKey".ToLower.Trim Then
                Return New LiteratureProperty("BibTeXKey", Me.ID)
            End If

            If Name.ToLower.Trim = "EntryType".ToLower.Trim Then
                Return New LiteratureProperty("EntryType", Me.Type)
            End If

            If Name.ToLower.Trim = "SourceCode".ToLower.Trim Then
                Return New LiteratureProperty("SourceCode", GenerateCode)
            End If

            For Each LiteratureProperty As LiteratureProperty In Me.PropertyList
                If LiteratureProperty.Name.ToLower.Trim = Name.ToLower.Trim Then
                    Return LiteratureProperty
                End If
            Next

            Return New LiteratureProperty("", "")
        End Function

        Private Function GenerateCode() As String
            Dim MaxPropertyNameLength As Integer = 0
            For Each LiteratureProperty As _BibTeX.LiteratureProperty In Me.PropertyList
                If LiteratureProperty.Name.Length > MaxPropertyNameLength Then
                    MaxPropertyNameLength = LiteratureProperty.Name.Length
                End If
            Next

            Dim Code As String = ""

            Code &= "@" & Me.Type & "{" & Me.ID & ","
            Dim Index As Integer = 0
            For Each LiteratureProperty As _BibTeX.LiteratureProperty In Me.PropertyList
                Code &= vbCrLf & "    " & LiteratureProperty.Name.PadRight(MaxPropertyNameLength) & " = {" & LiteratureProperty.Value & "}"
                Index += 1
                If Index = Me.PropertyList.Count Then
                    Code &= vbCrLf & "}"
                Else
                    Code &= ","
                End If
            Next

            Return Code
        End Function

        Public Function Exist(ByVal Name As String) As Boolean
            For Each LiteratureProperty As LiteratureProperty In Me.PropertyList
                If LiteratureProperty.Name.ToLower.Trim = Name.ToLower.Trim Then
                    Return True
                End If
            Next

            Return False
        End Function

        Public Function ExistURL() As Boolean
            For Each LiteratureProperty As LiteratureProperty In Me.PropertyList
                If LiteratureProperty.Name.ToLower.Trim = "url" Then
                    If Not LiteratureProperty.Value.Trim = "" Then
                        Return True
                    End If
                End If
            Next

            Return False
        End Function

        Public Function ExistFile() As Boolean
            For Each LiteratureProperty As LiteratureProperty In Me.PropertyList
                If LiteratureProperty.Name.ToLower.Trim = "file" Then
                    If Not LiteratureProperty.Value.Trim = "" Then
                        Return True
                    End If
                End If
            Next

            Return False
        End Function

    End Class

    ''' <summary>
    ''' Class PaperProperty represents class Literature's property 
    ''' </summary>
    ''' <remarks></remarks>
    Public Class LiteratureProperty
        ' The name of this property
        Public Name As String
        ' The value of this property
        Public Value As String

        Public Sub New(ByVal Name As String, ByVal Value As String)
            With Me
                .Name = Name
                .Value = Value
            End With
        End Sub
    End Class


End Namespace

