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
        End Enum

        Public Class BibTeX
            Public Shared Sub Compile(ByRef DataBase As Qiqi.DataBase)
                If Not My.Computer.FileSystem.FileExists(DataBase.FullFileName) Then
                    DataBase.CompileResult.ErrorMessage = Qiqi.Compiler.BibTeXError.FileNotExist
                    Exit Sub
                End If

                Dim Reader As New IO.StreamReader(DataBase.FullFileName, System.Text.Encoding.Default)
                Dim DataBaseBuffer As String = Reader.ReadToEnd
                Dim AnalysisState As Qiqi.Compiler.BibTeXAnalysisState = Qiqi.Compiler.BibTeXAnalysisState.Idle


                For Index As Integer = 0 To DataBaseBuffer.Length - 1
                    Select Case AnalysisState
                        Case BibTeXAnalysisState.Idle
                        Case BibTeXAnalysisState.ReadType
                        Case BibTeXAnalysisState.ReadBibTeXKey
                        Case BibTeXAnalysisState.ReadBuffer
                    End Select
                Next
                Reader.Close()
            End Sub
        End Class
    End Namespace
End Namespace


