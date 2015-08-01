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

        Public Sub SetErrorMessage(ByVal ErrorMessage As String)
            With Me
                .ErrorMessage = ErrorMessage
                .ExistError = True
            End With

            ' Debug
            MsgBox("Error Message = " & Me.ErrorMessage & vbCrLf & "Line Number = " & Me.LineNumber)
        End Sub

        Public Sub ClearErrorMessage()
            With Me
                .ErrorMessage = ""
                .ExistError = False
            End With
        End Sub
    End Class
End Namespace
