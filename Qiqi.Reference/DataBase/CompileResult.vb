Namespace Qiqi
    Public Class CompileResult
        Public LineNumber As Integer
        Private ErrorMessage As String
        Private ExistError As Boolean

        Public Sub New()
            With Me
                .ExistError = False
                .LineNumber = 0
                .ErrorMessage = ""
            End With
        End Sub
    End Class
End Namespace
