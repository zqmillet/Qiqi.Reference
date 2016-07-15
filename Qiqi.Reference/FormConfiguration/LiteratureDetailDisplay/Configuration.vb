Namespace _FormConfiguration
    Namespace LiteratureDetailDisplay
        Public Module LexicalAnalysisError
            Public Const Idle As String = "Idle error."
            Public Const ReadLiteratureType As String = "Read literature type error."
            Public Const ReadTabPageName As String = "Read tab page name error."
            Public Const ReadPropertyName As String = "Read property name error."
            Public Const ReadPropertyValue As String = "Read property value error."
            Public Const WaitPropertyComma As String = "Wait property comma error."
            Public Const WaitTabPageComma As String = "Wait tab page comma error."

        End Module

        Public Module LexicalAnalysisStatus
            Public Const Idle As Integer = 0
            Public Const ReadLiteratureType As Integer = 1
            Public Const ReadTabPageName As Integer = 2
            Public Const ReadPropertyName As Integer = 3
            Public Const ReadPropertyValue As Integer = 4
            Public Const EndPropertyValue As Integer = 5
            Public Const EndPropertyName As Integer = 6
            Public Const EndTabPageName As Integer = 7
            Public Const WaitPropertyComma As Integer = 8
            Public Const WaitTabPageComma As Integer = 9
        End Module

        Public Class Configuration
            Public LiteratureType As String

            Public TabPages As ArrayList

            Public Sub New()
                With Me
                    .LiteratureType = ""
                    .TabPages = New ArrayList
                End With
            End Sub

            Public Sub New(ByVal Buffer As String)
                With Me
                    .LiteratureType = ""
                    .TabPages = New ArrayList
                End With

                Dim ErrorMessage As String = ""

                If Not LexicalAnalysis(Buffer, ErrorMessage) Then
                    Me.LiteratureType = ""
                    Me.TabPages = New ArrayList
                    MsgBox(ErrorMessage, MsgBoxStyle.OkOnly, "Error")
                End If

                RemoveEmptyTabPages()
            End Sub

            Private Sub RemoveEmptyTabPages()
                Dim RemoveList As New ArrayList
                For Each TabPage As TabPageConfiguration In TabPages
                    If TabPage.Properties.Count = 0 Then
                        RemoveList.Add(TabPage)
                        Continue For
                    End If

                    If TabPage.TabPageName.Trim = "" Then
                        RemoveList.Add(TabPage)
                        Continue For
                    End If
                Next

                For Each TabPage As TabPageConfiguration In RemoveList
                    TabPages.Remove(TabPage)
                Next
            End Sub

            Private Function LexicalAnalysis(ByVal Buffer As String, ByRef ErrorMessage As String) As Boolean
                Dim State As Integer = LexicalAnalysisStatus.Idle

                Dim LiteratureType As String = ""
                Dim TabPageName As String = ""
                Dim PropertyName As String = ""
                Dim PropertyValue As String = ""

                Dim TabPageConfiguration As New TabPageConfiguration("")
                Dim PropertyConfiguration As New PropertyConfiguration("")

                Dim Index As Integer = 0

                For Each c As Char In Buffer
                    Select Case State
                        Case LexicalAnalysisStatus.Idle
                            Select Case c
                                Case " "
                                    ' Do nothing
                                Case ","
                                    ErrorMessage = LexicalAnalysisError.Idle
                                    Return False
                                Case "{"
                                    ErrorMessage = LexicalAnalysisError.Idle
                                    Return False
                                Case "}"
                                    ErrorMessage = LexicalAnalysisError.Idle
                                    Return False
                                Case Else
                                    LiteratureType = c
                                    State = LexicalAnalysisStatus.ReadLiteratureType
                            End Select
                        Case LexicalAnalysisStatus.ReadLiteratureType
                            Select Case c
                                Case " "
                                    LiteratureType &= c
                                Case ","
                                    LiteratureType &= c
                                Case "{"
                                    Me.LiteratureType = LiteratureType.Trim
                                    TabPageName = ""
                                    State = LexicalAnalysisStatus.ReadTabPageName
                                Case "}"
                                    ErrorMessage = LexicalAnalysisError.ReadLiteratureType
                                    Return False
                                Case Else
                                    LiteratureType &= c
                            End Select
                        Case LexicalAnalysisStatus.ReadTabPageName
                            Select Case c
                                Case " "
                                    TabPageName &= c
                                Case ","
                                    TabPageName &= c
                                Case "{"
                                    TabPageName = TabPageName.Trim
                                    TabPageConfiguration = New TabPageConfiguration(TabPageName)
                                    Me.TabPages.Add(TabPageConfiguration)
                                    PropertyName = ""
                                    State = LexicalAnalysisStatus.ReadPropertyName
                                Case "}"
                                    ErrorMessage = LexicalAnalysisError.ReadTabPageName
                                    Return False
                                Case Else
                                    TabPageName &= c
                            End Select
                        Case LexicalAnalysisStatus.ReadPropertyName
                            Select Case c
                                Case " "
                                    PropertyName &= c
                                Case ","
                                    ErrorMessage = LexicalAnalysisError.ReadPropertyName
                                    Return False
                                Case "{"
                                    PropertyName = PropertyName.Trim
                                    PropertyValue = ""
                                    PropertyConfiguration = New PropertyConfiguration(PropertyName)
                                    TabPageConfiguration.Properties.Add(PropertyConfiguration)
                                    Index = 0
                                    State = LexicalAnalysisStatus.ReadPropertyValue
                                Case "}"
                                    ErrorMessage = LexicalAnalysisError.ReadPropertyName
                                    Return False
                                Case Else
                                    PropertyName &= c
                            End Select
                        Case LexicalAnalysisStatus.ReadPropertyValue
                            Select Case c
                                Case " "
                                    PropertyValue &= c
                                Case ","
                                    ' Add this property value
                                    Select Case Index
                                        Case 0
                                            PropertyConfiguration.DisplayName = PropertyValue
                                        Case 1
                                            PropertyConfiguration.DisplayLabel = PropertyValue
                                        Case 2
                                            PropertyConfiguration.Height = Val(PropertyValue)
                                        Case 3
                                            PropertyConfiguration.IsRequired = PropertyValue
                                    End Select
                                    Index += 1
                                    PropertyValue = ""
                                Case "{"
                                    ErrorMessage = LexicalAnalysisError.ReadPropertyValue
                                    Return False
                                Case "}"
                                    ' Add this property
                                    PropertyConfiguration.SyntaxHighlight = PropertyValue
                                    State = LexicalAnalysisStatus.WaitPropertyComma
                                Case Else
                                    PropertyValue &= c
                            End Select
                        Case LexicalAnalysisStatus.WaitPropertyComma
                            Select Case c
                                Case ","
                                    ' Add this property name
                                    PropertyName = ""
                                    State = LexicalAnalysisStatus.ReadPropertyName
                                Case "}"
                                    ' Add this tab page
                                    State = LexicalAnalysisStatus.WaitTabPageComma
                                Case Else
                                    ErrorMessage = LexicalAnalysisError.WaitPropertyComma
                                    Return False
                            End Select
                        Case LexicalAnalysisStatus.WaitTabPageComma
                            Select Case c
                                Case ","
                                    ' Add this 
                                    TabPageName = ""
                                    State = LexicalAnalysisStatus.ReadTabPageName
                                Case "}"
                                    Return True
                                Case Else
                                    ErrorMessage = LexicalAnalysisError.WaitTabPageComma
                                    Return False
                            End Select
                    End Select


                Next

                Return False
            End Function
        End Class

        Public Class TabPageConfiguration
            Public TabPageName As String

            Public Properties As ArrayList

            Public Sub New(ByVal TabPageName As String)
                With Me
                    .TabPageName = TabPageName
                    .Properties = New ArrayList
                End With
            End Sub
        End Class

        Public Class PropertyConfiguration
            Public PropertyName As String
            Public Height As Integer
            Public DisplayName As String
            Public DisplayLabel As Boolean
            Public IsRequired As Boolean
            Public SyntaxHighlight As Boolean

            Public Sub New(ByVal PropertyName As String)
                With Me
                    .PropertyName = PropertyName
                    .Height = 0
                    .DisplayName = ""
                    .DisplayLabel = False
                    .IsRequired = False
                    .SyntaxHighlight = False
                End With
            End Sub
        End Class
    End Namespace
End Namespace

