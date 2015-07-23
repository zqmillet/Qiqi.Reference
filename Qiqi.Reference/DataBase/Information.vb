Namespace Qiqi
    Public Class Information
        Public Name As String
        Public Value As String

        Public Sub New()
            With Me
                .Name = ""
                .Value = ""
            End With
        End Sub

        Public Sub New(ByVal Name As String, ByVal Value As String)
            With Me
                .Name = Name
                .Value = Value
            End With
        End Sub
    End Class
End Namespace

