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
            Public Const SyntaxError As String = "Syntax Error!"

        End Module

        Public Class BibTeX
            Public Shared Sub Compile(ByRef DataBase As Qiqi.DataBase)
                If Not My.Computer.FileSystem.FileExists(DataBase.FullFileName) Then
                    DataBase.SetErrorMessage(BibTeXErrorMessage.FileNotExist)
                    Exit Sub
                End If

                Dim Reader As New IO.StreamReader(DataBase.FullFileName, System.Text.Encoding.Default)
                Dim DataBaseBuffer As String = Reader.ReadToEnd
                Dim AnalysisState As Qiqi.Compiler.BibTeXAnalysisState = Qiqi.Compiler.BibTeXAnalysisState.Idle


                DataBase.CompileResult.LineNumber = 1

                For Index As Integer = 0 To DataBaseBuffer.Length - 1
                    Dim c As Char = DataBaseBuffer(Index)

                    If c = vbCr Then
                        DataBase.CompileResult.LineNumber += 1
                    End If

                    Select Case AnalysisState
                        Case BibTeXAnalysisState.Idle
                            Select Case c
                                Case "%"
                                    AnalysisState = BibTeXAnalysisState.Idle
                                Case "@"
                                    AnalysisState = BibTeXAnalysisState.ReadType
                                Case " "
                                    ' Do nothing
                                Case vbCr
                                    ' Do nothing
                                Case vbLf
                                    ' Do nothing
                                Case Else
                                    DataBase.SetErrorMessage(BibTeXErrorMessage.SyntaxError)
                                    Exit Sub
                            End Select
                        Case BibTeXAnalysisState.ReadType
                            Select Case c
                                Case "%"
                                Case "@"


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


