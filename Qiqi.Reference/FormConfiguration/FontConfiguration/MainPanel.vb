Namespace _FormConfiguration
    Namespace FontConfiguration
        Public Class MainPanel
            Inherits System.Windows.Forms.Panel

            Public Modified As Boolean
            Dim Configuration As _FormConfiguration.Configuration

            Public Sub New(ByVal Configuration As _FormConfiguration.Configuration)
                With Me
                    .Modified = False
                    .Configuration = Configuration
                End With
            End Sub


        End Class
    End Namespace
End Namespace

