Namespace _FormMain
    Namespace _LiteratureTabControl
        Public Class HighlightStyle
            Public EntryType As FontStyle
            Public BibTeXKey As FontStyle
            Public TagName As FontStyle
            Public TagValue As FontStyle

            Public Sub New()
                With Me
                    .EntryType = FontStyle.Regular
                    .BibTeXKey = FontStyle.Regular
                    .TagName = FontStyle.Regular
                    .TagValue = FontStyle.Regular
                End With
            End Sub

            Public Sub New(ByVal EntryType As FontStyle, ByVal BibTeXKey As FontStyle, ByVal TagName As FontStyle, ByVal TagValue As FontStyle)
                With Me
                    .EntryType = EntryType
                    .BibTeXKey = BibTeXKey
                    .TagName = TagName
                    .TagValue = TagValue
                End With
            End Sub

            Public Sub New(ByVal EntryType As Boolean, ByVal BibTeXKey As Boolean, ByVal TagName As Boolean, ByVal TagValue As Boolean)
                With Me
                    .EntryType = Boolean2FontStyle(EntryType)
                    .BibTeXKey = Boolean2FontStyle(BibTeXKey)
                    .TagName = Boolean2FontStyle(TagName)
                    .TagValue = Boolean2FontStyle(TagValue)
                End With
            End Sub

            Private Function Boolean2FontStyle(ByVal Bold As Boolean) As FontStyle
                If Bold Then
                    Return FontStyle.Bold
                Else
                    Return FontStyle.Regular
                End If
            End Function
        End Class
    End Namespace
End Namespace
