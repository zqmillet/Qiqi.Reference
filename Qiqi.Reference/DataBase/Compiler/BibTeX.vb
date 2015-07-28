Namespace Qiqi
    Namespace Compiler
        Public Enum BibTeXAnalysisState
            Idle
            ReadType
            ReadBibTeXKey
            ReadBuffer
            ReadComment
        End Enum

        Public Module BibTeXErrorMessage
            Public Const FileNotExist As String = "File do not Exist!"
            Public Const IdleSyntaxError As String = "Idle Mode Syntax Error!"
            Public Const TypeSyntaxError As String = "Type Syntax Error!"
            Public Const BibTeXKeySyntaxError As String = "BibTeXKey Syntax Error!"

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
                                    AnalysisState = BibTeXAnalysisState.ReadBibTeXKey
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
                                Case "}"
                                    BraketNumber -= 1
                                Case Else
                                    LiteratureBuffer &= c
                            End Select

                            If BraketNumber = 0 Then
                                LiteratureBuffer = LiteratureBuffer.Trim
                                AnalysisState = BibTeXAnalysisState.Idle
                            End If
                        Case BibTeXAnalysisState.ReadComment
                            Select Case c
                                Case vbCr
                                    AnalysisState = BibTeXAnalysisState.Idle
                                Case Else
                                    ' Do nithing
                            End Select
                    End Select
                Next
                Reader.Close()
            End Sub
        End Class
    End Namespace
End Namespace


