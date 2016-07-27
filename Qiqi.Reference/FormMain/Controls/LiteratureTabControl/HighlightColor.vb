Namespace _FormMain
    Namespace _LiteratureTabControl
        Public Class HighlightColor
            Public EntryType As Integer
            Public BibTeXKey As Integer
            Public TagName As Integer
            Public TagValue As Integer

            Public Sub New()
                With Me
                    .EntryType = Color.Black.ToArgb
                    .BibTeXKey = Color.Red.ToArgb
                    .TagName = Color.Blue.ToArgb
                    .TagValue = Color.Green.ToArgb
                End With
            End Sub

            Public Sub New(ByVal EntryType As Integer, ByVal BibTeXKey As Integer, ByVal TagName As Integer, ByVal TagValue As Integer)
                With Me
                    .EntryType = EntryType
                    .BibTeXKey = BibTeXKey
                    .TagName = TagName
                    .TagValue = TagValue
                End With
            End Sub
        End Class
    End Namespace
End Namespace