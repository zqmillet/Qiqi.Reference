﻿Namespace _FormConfiguration
    Namespace FontConfiguration
        Public Class FormSeparator
            Inherits Windows.Forms.Panel

            Dim Label As New Label
            Dim Line As New Label

            Public Sub New(ByVal Text As String, ByVal Width As Integer)
                With Line
                    .Text = ""
                    .AutoSize = False
                    .Height = 2
                    .Width = Width
                    .BorderStyle = BorderStyle.Fixed3D
                    .Margin = New Padding(0, 4, 0, 4)
                    .Location = New Point(0, 10)
                End With

                With Label
                    .Text = Text
                    .AutoSize = True
                    .Margin = New Padding(0)
                    .Location = New Point(0, 4)
                End With

                With Me
                    .Width = Width
                    .Height = 20
                    .Controls.Add(Label)
                    .Controls.Add(Line)
                    .Margin = New Padding(0, 0, 0, 4)
                End With
            End Sub

        End Class
    End Namespace
End Namespace
