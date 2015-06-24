Namespace _FormMain
    Public Class GroupTreeView
        Inherits System.Windows.Forms.TreeView

        Public Delegate Sub DelegateShowGroup(ByVal GroupBuffer As String)


        Public Sub New()
            Me.ShowRootLines = False
        End Sub

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
            Dim AnalysisStatus As Integer = GroupAnalysisStatus.Idle

            Dim GroupLevel As String = ""
            Dim GroupTitle As String = ""
            Dim GroupType As String = ""
            Dim GroupHierarchicalContext As String = ""

            Dim LastTreeViewNodeList As New ArrayList

            For Index As Integer = 0 To GroupBuffer.Length - 1
                Dim c As Char = GroupBuffer(Index)

                If (vbCr & Chr(10) & vbCrLf).Contains(c) Then
                    Continue For
                End If


                Select Case AnalysisStatus
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
                                AnalysisStatus = GroupAnalysisStatus.ReadGroupLevel
                        End Select
                    Case GroupAnalysisStatus.ReadGroupLevel
                        Select Case c
                            Case " "
                                If Not IsNumeric(GroupLevel) Then
                                    MsgBox(GroupAnalysisError.ReadGroupLevel)
                                    Exit Sub
                                End If
                                AnalysisStatus = GroupAnalysisStatus.ReadGroupType
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
                                GroupLevel &= c
                        End Select
                    Case GroupAnalysisStatus.ReadGroupType
                        Select Case c
                            Case " "
                                MsgBox(GroupAnalysisError.ReadGroupType)
                                Exit Sub
                            Case vbCr
                                MsgBox(GroupAnalysisError.ReadGroupType)
                                Exit Sub
                            Case Chr(10)
                                MsgBox(GroupAnalysisError.ReadGroupType)
                                Exit Sub
                            Case ":"
                                Select Case GroupType.ToLower
                                    Case "AllEntriesGroup".ToLower
                                        AnalysisStatus = GroupAnalysisStatus.ReadRootNode
                                    Case "ExplicitGroup".ToLower
                                        AnalysisStatus = GroupAnalysisStatus.ReadGroupTitle
                                    Case Else
                                        MsgBox(GroupAnalysisError.ReadGroupType)
                                        Exit Sub
                                End Select
                            Case "\"
                                MsgBox(GroupAnalysisError.ReadGroupType)
                                Exit Sub
                            Case ";"
                                MsgBox(GroupAnalysisError.ReadGroupType)
                                Exit Sub
                            Case Else
                                GroupType &= c
                        End Select
                    Case GroupAnalysisStatus.ReadRootNode
                        Select Case c
                            Case " "
                                ' Do nothing
                            Case vbCr
                                ' Do nothing
                            Case Chr(10)
                                ' Do nothing
                            Case ":"
                                MsgBox(GroupAnalysisError.ReadRootNode)
                                Exit Sub
                            Case "\"
                                MsgBox(GroupAnalysisError.ReadRootNode)
                                Exit Sub
                            Case ";"
                                If Not LastTreeViewNodeList.Count = 0 Then
                                    MsgBox(GroupAnalysisError.ReadRootNode)
                                    Exit Sub
                                End If

                                Dim Node As New _FormMain.GroupTreeNode(GroupTitle, 0)
                                Me.Nodes.Add(Node)
                                LastTreeViewNodeList.Add(Node)
                                GroupType = ""
                                GroupLevel = ""
                                GroupTitle = ""
                                GroupHierarchicalContext = ""
                                AnalysisStatus = GroupAnalysisStatus.Idle
                            Case Else
                                MsgBox(GroupAnalysisError.ReadRootNode)
                                Exit Sub
                        End Select
                    Case GroupAnalysisStatus.ReadGroupTitle
                        Select Case c
                            Case " "
                                GroupTitle &= c
                            Case vbCr
                                MsgBox(GroupAnalysisError.ReadGroupTitle)
                                Exit Sub
                            Case Chr(10)
                                MsgBox(GroupAnalysisError.ReadGroupTitle)
                                Exit Sub
                            Case ":"
                                GroupTitle &= c
                            Case "\"
                                If NextIsSemicolon(GroupBuffer, Index) Then
                                    Index += 1
                                    AnalysisStatus = GroupAnalysisStatus.ReadHierarchicalContext
                                Else
                                    GroupTitle &= c
                                End If
                            Case ";"
                                GroupTitle &= c
                            Case Else
                                GroupTitle &= c
                        End Select
                    Case GroupAnalysisStatus.ReadHierarchicalContext
                        Select Case c
                            Case " "
                                ' Do nothing
                            Case vbCr
                                ' Do nothing
                            Case Chr(10)
                                ' Do nothing
                            Case ":"
                                MsgBox(GroupAnalysisError.ReadHierarchicalContext)
                                Exit Sub
                            Case "\"
                                If NextIsSemicolon(GroupBuffer, Index) Then
                                    Index += 1

                                    AnalysisStatus = GroupAnalysisStatus.EndGroupProperty
                                Else
                                    MsgBox(GroupAnalysisError.ReadHierarchicalContext)
                                    Exit Sub
                                End If
                            Case ";"
                                MsgBox(GroupAnalysisError.ReadHierarchicalContext)
                                Exit Sub
                            Case Else
                                GroupHierarchicalContext &= c
                                If Not IsNumeric(GroupHierarchicalContext) Then
                                    MsgBox(GroupAnalysisError.ReadHierarchicalContext)
                                    Exit Sub
                                End If
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

        Private Function NextIsSemicolon(ByVal GroupBuffer As String, ByVal Index As Integer) As Boolean
            If Index = GroupBuffer.Length - 1 Then
                Return False
            End If

            If GroupBuffer(Index + 1) = ";" Then
                Return True
            Else
                Return False
            End If

        End Function

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
        EndGroupProperty
        ReadRootNode
        ReadBibTeXKey
    End Enum

    Public Module GroupAnalysisError
        Public Const Idle As String = "Idle error."
        Public Const ReadGroupLevel As String = "Read group level error."
        Public Const ReadGroupTitle As String = "Read group title error."
        Public Const ReadGroupType As String = "Read group type error."
        Public Const ReadHierarchicalContext As String = "Read hierarchical context error."
        Public Const EndGroupProperty As String = "End group property error."
        Public Const ReadRootNode As String = "Read root node error."
        Public Const ReadBibTeXKey As String = "Read BibTeXKey error."
    End Module

    Public Class GroupTreeNode
        Inherits System.Windows.Forms.TreeNode

        Public HierarchicalContext As Integer
        Public BibTeXKeyList As ArrayList

        Public Sub New(ByVal GroupTitle As String, ByVal HierarchicalContext As Integer)
            With Me
                .Text = GroupTitle
                .BibTeXKeyList = New ArrayList
                .HierarchicalContext = HierarchicalContext
            End With
        End Sub
    End Class
End Namespace
