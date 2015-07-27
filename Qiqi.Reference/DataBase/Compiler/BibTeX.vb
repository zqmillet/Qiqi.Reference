Namespace Qiqi
    Namespace Compiler
        Public Module BibTeXError
            Public Const FileNotExist As String = "The file does not exist."

        End Module

        Public Enum BibTeXAnalysisState
            Idle
            ReadType
            ReadBibTeXKey
            ReadBuffer
            ReadComment
        End Enum

        Public Class BibTeX
            Public Shared Function Compile(ByRef DataBase As Qiqi.DataBase) As Boolean
                If Not My.Computer.FileSystem.FileExists(DataBase.FullFileName) Then
                    DataBase.CompileResult.ErrorMessage = Qiqi.Compiler.BibTeXError.FileNotExist
                    Return False
                End If

                Dim Reader As New IO.StreamReader(DataBase.FullFileName, System.Text.Encoding.Default)
                Dim DataBaseBuffer As String = Reader.ReadToEnd
                Dim AnalysisState As Qiqi.Compiler.BibTeXAnalysisState = Qiqi.Compiler.BibTeXAnalysisState.Idle



                For Index As Integer = 0 To DataBaseBuffer.Length - 1
                    Dim c As Char = DataBaseBuffer(Index)
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
                                    Return False
                            End Select
                        Case BibTeXAnalysisState.ReadType
                        Case BibTeXAnalysisState.ReadBibTeXKey
                        Case BibTeXAnalysisState.ReadBuffer
                        Case BibTeXAnalysisState.ReadComment
                    End Select
                Next
                Reader.Close()
            End Function
        End Class
    End Namespace
End Namespace


