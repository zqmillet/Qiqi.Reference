Namespace _FormMain
    Public Class GroupTreeView
        Inherits System.Windows.Forms.TreeView

        Public Delegate Sub DelegateShowGroup(ByVal GroupBuffer As String)


        Public Sub Loading(ByVal DataBase As _BibTeX.DataBase)
            Select Case DataBase.GroupVersion
                Case 3
                    If Me.InvokeRequired Then
                        Me.Invoke(New DelegateShowGroup(AddressOf ShowGroupVersion3), DataBase.GroupBuffer)
                    Else
                        ShowGroupVersion3(DataBase.GroupBuffer)
                    End If

                Case Else

            End Select
        End Sub

        Private Sub ShowGroupVersion3(ByVal GroupBuffer As String)
            Dim AnalysisState As Integer = GroupAnalysisStatus.Idle

            Dim GroupLevel As String = ""
            Dim GroupTitle As String = ""
            Dim GroupType As String = ""

            For Index As Integer = 0 To GroupBuffer.Length - 1
                Dim c As Char = GroupBuffer(Index)

                If (vbCr & Chr(10) & vbCrLf).Contains(c) Then
                    Continue For
                End If


                Select Case AnalysisState
                    Case GroupAnalysisStatus.Idle
                        Select Case c
                            Case " "
                                ' Do nothing
                            Case vbCr
                                ' Do nothing
                            Case Chr(10)
                                ' Do nothing
                            Case ":"
                                MsgBox(GroupAnalysisError.Idle)
                                Exit Sub
                            Case "\"
                                MsgBox(GroupAnalysisError.Idle)
                                Exit Sub
                            Case ";"
                                MsgBox(GroupAnalysisError.Idle)
                                Exit Sub
                            Case Else
                                GroupLevel &= c
                                AnalysisState = GroupAnalysisStatus.ReadGroupLevel
                        End Select
                    Case GroupAnalysisStatus.ReadGroupLevel
                        Select Case c
                            Case " "
                                AnalysisState = GroupAnalysisStatus.ReadGroupType
                            Case vbCr
                                ' Do nothing
                            Case Chr(10)
                                ' Do nothing
                            Case ":"
                                MsgBox(GroupAnalysisError.ReadGroupLevel)
                                Exit Sub
                            Case "\"
                                MsgBox(GroupAnalysisError.ReadGroupLevel)
                                Exit Sub
                            Case ";"
                                MsgBox(GroupAnalysisError.ReadGroupLevel)
                                Exit Sub
                            Case Else
                        End Select
                    Case GroupAnalysisStatus.ReadGroupType
                        Select Case c
                            Case " "
                            Case vbCr
                            Case Chr(10)
                            Case ":"
                            Case "\"
                            Case ";"
                            Case Else
                        End Select
                    Case GroupAnalysisStatus.ReadGroupTitle
                        Select Case c
                            Case " "
                            Case vbCr
                            Case Chr(10)
                            Case ":"
                            Case "\"
                            Case ";"
                            Case Else
                        End Select
                    Case GroupAnalysisStatus.ReadHierarchicalContext
                        Select Case c
                            Case " "
                            Case vbCr
                            Case Chr(10)
                            Case ":"
                            Case "\"
                            Case ";"
                            Case Else
                        End Select
                    Case GroupAnalysisStatus.ReadBibTeXKey
                        Select Case c
                            Case " "
                            Case vbCr
                            Case Chr(10)
                            Case ":"
                            Case "\"
                            Case ";"
                            Case Else
                        End Select
                End Select

            Next
        End Sub
    End Class

    Public Module HierarchicalContext
        Public Const IndependentGroup As Integer = 0
        Public Const RefineSuperGroup As Integer = 1
        Public Const IncludeSubGroups As Integer = 2
    End Module

    Public Enum GroupAnalysisStatus
        Idle
        ReadGroupLevel
        ReadGroupTitle
        ReadGroupType
        ReadHierarchicalContext
        ReadBibTeXKey
    End Enum

    Public Module GroupAnalysisError
        Public Const Idle As String = "Idle error."
        Public Const ReadGroupLevel As String = "Read group level error."
        Public Const ReadGroupTitle As String = "Read group title error."
        Public Const ReadGroupType As String = "Read group type error."
        Public Const ReadHierarchicalContext As String = "Read hierarchical context error."
        Public Const ReadBibTeXKey As String = "Read BibTeXKey error."
    End Module

    Public Class GroupTreeNode
        Inherits System.Windows.Forms.TreeNode

        Public HierarchicalContext As Integer
        Public BibTeXKeyList As ArrayList

        Public Sub New(ByVal HierarchicalContext As Integer)
            With Me
                .BibTeXKeyList = New ArrayList
                .HierarchicalContext = HierarchicalContext
            End With
        End Sub
    End Class
End Namespace
