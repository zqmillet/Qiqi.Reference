Namespace _BibTeX
    Public Class ErrorMessage
        Public ExistError As Boolean
        Public Message As String
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
    End Class
End Namespace

