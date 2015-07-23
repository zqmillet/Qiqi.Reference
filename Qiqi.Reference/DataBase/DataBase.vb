Namespace Qiqi
    Public Class DataBase
        Public LiteratureList As ArrayList
        Public Group As Qiqi.Group
        Public FullFileName As String

        Public Sub New()
            With Me
                .LiteratureList = New ArrayList
                .Group = New Qiqi.Group
                .FullFileName = ""
            End With
        End Sub

        Public Sub New(ByVal FullFileName As String)
            With Me
                .LiteratureList = New ArrayList
                .Group = New Qiqi.Group
                .FullFileName = FullFileName
            End With
        End Sub

        Public Sub Load()

        End Sub
    End Class
End Namespace
