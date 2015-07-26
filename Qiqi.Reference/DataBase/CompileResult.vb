Namespace Qiqi
    Public Class CompileResult
        Public LineNumber As Integer
        Public ErrorMessage As String
        Public ExistError As Boolean

        Public Sub New()
            With Me
                .ExistError = False
                .LineNumber = 0
                .ErrorMessage = ""
            End With
        End Sub
    End Class
End Namespace
