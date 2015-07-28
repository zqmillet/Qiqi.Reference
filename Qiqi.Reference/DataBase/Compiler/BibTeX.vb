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
            Public Const BibTeXSyntaxError As String = "BibTeX Syntax Error!"

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
                                    AnalysisState = BibTeXAnalysisState.Idle
                                Case "@"
                                    ' If c = "@", it means the begin of a literature,
                                    ' then initialize the information of a new literature.


                                    AnalysisState = BibTeXAnalysisState.ReadType
                                Case " "
                                    ' Do nothing
                                Case vbCr
                                    ' Do nothing
                                Case vbLf
                                    ' Do nothing
                                Case Else
                                    DataBase.SetErrorMessage(BibTeXErrorMessage.IdleSyntaxError)
                                    Exit Sub
                            End Select
                        Case BibTeXAnalysisState.ReadType
                            ' The type should not have the following char:
                            ' % @ { } vbCr vbLf
                            ' If c = "{", it means that the type has been read.
                            Select Case c
                                Case "%"
                                    DataBase.SetErrorMessage(BibTeXErrorMessage.TypeSyntaxError)
                                    Exit Sub
                                Case "@"
                                    DataBase.SetErrorMessage(BibTeXErrorMessage.TypeSyntaxError)
                                    Exit Sub
                                Case "}"
                                    DataBase.SetErrorMessage(BibTeXErrorMessage.TypeSyntaxError)
                                    Exit Sub
                                Case vbCr
                                    DataBase.SetErrorMessage(BibTeXErrorMessage.TypeSyntaxError)
                                    Exit Sub
                                Case vbLf
                                    DataBase.SetErrorMessage(BibTeXErrorMessage.TypeSyntaxError)
                                    Exit Sub
                                Case "{"

                                Case Else

                            End Select
                        Case BibTeXAnalysisState.ReadBibTeXKey
                        Case BibTeXAnalysisState.ReadBuffer
                        Case BibTeXAnalysisState.ReadComment
                    End Select
                Next
                Reader.Close()
            End Sub
        End Class
    End Namespace
End Namespace


