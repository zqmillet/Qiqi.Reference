﻿Namespace Qiqi
    Public Enum DataBaseType
        ' This is not a currect database type
        ErrorType
        ' the database has no extension
        Unknown
        ' the type of database is BibTeX 
        BibTeX
    End Enum

    Public Class DataBase
        Public LiteratureList As ArrayList
        Public Group As Qiqi.Group
        Public FullFileName As String
        Public CompileResult As Qiqi.CompileResult

        Public Sub New()
            With Me
                .LiteratureList = New ArrayList
                .Group = New Qiqi.Group
                .FullFileName = ""
                .CompileResult = New Qiqi.CompileResult
            End With
        End Sub

        Public Sub New(ByVal FullFileName As String)
            With Me
                .LiteratureList = New ArrayList
                .Group = New Qiqi.Group
                .FullFileName = FullFileName
                .CompileResult = New Qiqi.CompileResult
            End With
        End Sub

        Public Sub Load()
            GetDataBaseType()
        End Sub


        ''' <summary>
        ''' 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetDataBaseType() As Qiqi.DataBaseType
            If FullFileName.Trim = "" Then
                Return Qiqi.DataBaseType.ErrorType
            End If

            If Not FullFileName.Contains("\") Then
                Return Qiqi.DataBaseType.ErrorType
            End If

            Dim Extension As String = FullFileName.Remove(0, FullFileName.LastIndexOf("\") + 1)

            If Not Extension.Contains(".") Then
                Return Qiqi.DataBaseType.Unknown
            End If

            Extension = Extension.Remove(0, Extension.LastIndexOf(".") + 1)

            Select Case Extension.Trim.ToLower
                Case "bib"
                    Return Qiqi.DataBaseType.BibTeX
                Case Else
                    Return Qiqi.DataBaseType.Unknown
            End Select
        End Function
    End Class
End Namespace
