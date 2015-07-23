Namespace Qiqi
    Public Class ErrorMessage
        Public LineNumber As Integer
        Private Message As String
        Private ExistError As Boolean

        Public Sub New()
            With Me
                .ExistError = False
                .LineNumber = 0
                .Message = ""
            End With
        End Sub
    End Class
End Namespace


