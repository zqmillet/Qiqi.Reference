Namespace _BibTeX
    Public Class ErrorMessage
        Private ExistError As Boolean
        Private Message As String
        Public LineNumber As Integer

        Public Sub New()
            With Me
                .ExistError = False
                .LineNumber = 0
                .Message = ""
            End With
        End Sub

        Public Sub Clear()
            With Me
                .ExistError = False
                .LineNumber = 0
                .Message = ""
            End With
        End Sub

        Public Function HasError() As Boolean
            Return ExistError
        End Function

        Public Function GetErrorMessage() As String
            Return Me.Message
        End Function

        Public Sub SetErrorMessage(ByVal ErrorMessage As String)
            With Me
                .ExistError = True
                .Message = ErrorMessage
            End With
        End Sub
    End Class
End Namespace

