Namespace Qiqi
    Namespace Compiler
        Public Enum BibTeXAnalysisState
            Idle
            ReadType
            ReadBibTeXKey
            ReadBuffer
            ReadComment
            ReadAtComment

            ReadPropertyName
            ReadPropertyValue
        End Enum

        Public Module BibTeXErrorMessage
            Public Const FileNotExist As String = "File do not Exist!"
            Public Const IdleSyntaxError As String = "Idle Mode Syntax Error!"
            Public Const TypeSyntaxError As String = "Type Syntax Error!"
            Public Const BibTeXKeySyntaxError As String = "BibTeXKey Syntax Error!"

            Public Const PropertyNameSyntaxError As String = "Property Name Syntax Error!"
        End Module

        Public Class BibTeX
            Public Shared Sub Compile(ByRef DataBase As Qiqi.DataBase)
                ' If there do not exist the database, exit sub
                If Not My.Computer.FileSystem.FileExists(DataBase.FullFileName) Then
                    DataBase.SetErrorMessage(BibTeXErrorMessage.FileNotExist)
                    Exit Sub
                End If

                Dim Reader As New IO.StreamReader(DataBase.FullFileName, System.Text.Encoding.Default)

                ' The buffer of database
                Dim DataBaseBuffer As String = Reader.ReadToEnd
                ' State of lexical analysis
                Dim AnalysisState As Qiqi.Compiler.BibTeXAnalysisState = Qiqi.Compiler.BibTeXAnalysisState.Idle

                ' The parameters of a literature
                Dim LiteratureType As String = ""
                Dim LiteratureID As String = ""
                Dim LiteratureBuffer As String = ""

                Dim BraketNumber As Integer = 0

                ' Initialize the line number
                DataBase.CompileResult.LineNumber = 1

                For Index As Integer = 0 To DataBaseBuffer.Length - 1
                    Dim c As Char = DataBaseBuffer(Index)

                    ' Update the line number
                    If c = vbCr Then
                        DataBase.CompileResult.LineNumber += 1
                    End If

                    Select Case AnalysisState
                        Case BibTeXAnalysisState.Idle
                            Select Case c
                                Case "%"
                                    AnalysisState = BibTeXAnalysisState.ReadComment
                                Case "@"
                                    ' If c = "@", it means the begin of a literature,
                                    ' then initialize the information of a new literature.
                                    LiteratureBuffer = ""
                                    LiteratureID = ""
                                    LiteratureType = ""

                                    AnalysisState = BibTeXAnalysisState.ReadType
                                Case " ", vbCr, vbLf
                                    ' Do nothing
                                Case Else
                                    DataBase.SetErrorMessage(BibTeXErrorMessage.IdleSyntaxError)
                                    Exit Sub
                            End Select
                        Case BibTeXAnalysisState.ReadType
                            ' The type should not have the following char:
                            ' % @ { } vbCr vbLf ,
                            ' If c = "{", it means that the type has been read.
                            Select Case c
                                Case "%", "@", ",", "}", vbCr, vbLf
                                    DataBase.SetErrorMessage(BibTeXErrorMessage.TypeSyntaxError)
                                    Exit Sub
                                Case "{"
                                    BraketNumber = 1

                                    If LiteratureType.Trim.ToLower = "comment" Then
                                        AnalysisState = BibTeXAnalysisState.ReadAtComment
                                    Else
                                        AnalysisState = BibTeXAnalysisState.ReadBibTeXKey
                                    End If
                                Case Else
                                    LiteratureType &= c
                            End Select
                        Case BibTeXAnalysisState.ReadBibTeXKey
                            ' The BibTeXKey should not have the following char:
                            ' % @ { } vbCr vbLf ,
                            ' If c = ",", it means that the type has been read.
                            Select Case c
                                Case "%", "@", "{", "}", vbCr, vbLf
                                    DataBase.SetErrorMessage(BibTeXErrorMessage.BibTeXKeySyntaxError)
                                    Exit Sub
                                Case ","
                                    AnalysisState = BibTeXAnalysisState.ReadBuffer
                                Case Else
                                    LiteratureID &= c
                            End Select
                        Case BibTeXAnalysisState.ReadBuffer
                            Select Case c
                                Case "{"
                                    BraketNumber += 1
                                    LiteratureBuffer &= c
                                Case "}"
                                    BraketNumber -= 1
                                    If Not BraketNumber = 0 Then
                                        LiteratureBuffer &= c
                                    End If
                                Case Else
                                    LiteratureBuffer &= c
                            End Select

                            If BraketNumber = 0 Then
                                LiteratureBuffer = LiteratureBuffer.Trim
                                AnalysisState = BibTeXAnalysisState.Idle

                                Dim Literature As New Literature
                                Literature.ID = LiteratureID
                                Literature.Type = LiteratureType
                                Literature.InformationList = GetLiteratureInformation(LiteratureBuffer, DataBase.CompileResult)
                                ' GetLiteratureInformation(LiteratureBuffer, Literature, DataBase.CompileResult)

                                If DataBase.CompileResult.ExistError Then
                                    Exit Sub
                                Else
                                    DataBase.LiteratureList.Add(Literature)
                                End If
                            End If
                        Case BibTeXAnalysisState.ReadComment
                            Select Case c
                                Case vbCr
                                    AnalysisState = BibTeXAnalysisState.Idle
                                Case Else
                                    ' Do nithing
                            End Select
                        Case BibTeXAnalysisState.ReadAtComment
                            Select Case c
                                Case "{"
                                    BraketNumber += 1
                                    LiteratureBuffer &= c
                                Case "}"
                                    BraketNumber -= 1
                                    If Not BraketNumber = 0 Then
                                        LiteratureBuffer &= c
                                    End If
                                Case Else
                                    LiteratureBuffer &= c
                            End Select

                            If BraketNumber = 0 Then
                                AnalysisState = BibTeXAnalysisState.Idle
                            End If
                    End Select
                Next
                Reader.Close()
            End Sub

            Private Shared Function GetLiteratureInformation(ByVal LiteratureBuffer As String, _
                                                             ByRef CompileResult As Qiqi.CompileResult) As ArrayList
                Dim InformationList As New ArrayList
                Dim AnalysisState As Qiqi.Compiler.BibTeXAnalysisState = Qiqi.Compiler.BibTeXAnalysisState.Idle

                Dim PropertyName As String = ""
                Dim PropertyValue As String = ""

                For Index As Integer = 0 To LiteratureBuffer.Length - 1
                    Dim c As Char = LiteratureBuffer(Index)
                    Select Case AnalysisState
                        Case BibTeXAnalysisState.Idle
                            Select Case c
                                Case " ", vbCr, vbLf
                                    ' Do nothing
                                Case "%", "{", "}", "?", ";", "(", ")", "[", "]", "-", "=", "+", ",", ".", "<", ">", "@", "!", "#"
                                    ' If there is illegal character, exit sub
                                    CompileResult.SetErrorMessage(BibTeXErrorMessage.PropertyNameSyntaxError)
                                    Return InformationList
                                Case Else
                                    ' Begin to read the property name
                                    PropertyName = c
                                    AnalysisState = BibTeXAnalysisState.ReadPropertyName
                            End Select
                        Case BibTeXAnalysisState.ReadPropertyName
                            Select Case c
                                Case "="
                                    AnalysisState = BibTeXAnalysisState.ReadPropertyValue
                                Case Else
                                    PropertyName &= c
                            End Select
                        Case BibTeXAnalysisState.ReadPropertyValue
                    End Select
                Next


                Return InformationList
            End Function
        End Class
    End Namespace
End Namespace


