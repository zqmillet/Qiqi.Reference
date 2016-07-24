Namespace _FormMain
    Public Class GroupTreeView
        Inherits System.Windows.Forms.TreeView

        Public Delegate Sub DelegateShowGroup(ByVal GroupBuffer As String)

        Public Sub New()
            Me.ShowRootLines = False
            Me.ItemHeight = 18
            AddHandler Me.NodeMouseDoubleClick, AddressOf Me_NodeMouseDoubleClick
        End Sub

        ''' <summary>
        ''' This sub is used to stop auto expand/collapse on double-click event
        ''' </summary>
        ''' <param name="m"></param>
        Protected Overrides Sub DefWndProc(ByRef m As Message)
            If m.Msg = 515 Then
            Else
                MyBase.DefWndProc(m)
            End If
        End Sub

        ''' <summary>
        ''' This sub is used to expand/collapse on double-click event except root node
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub Me_NodeMouseDoubleClick(ByVal sender As Object, ByVal e As TreeNodeMouseClickEventArgs)
            If Not Me.SelectedNode.Parent Is Nothing Then
                If Me.SelectedNode.IsExpanded Then
                    Me.SelectedNode.Collapse()
                Else
                    Me.SelectedNode.Expand()
                End If
            End If
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
            Dim BibTeXKey As String = ""

            Dim CurrentLevel As Integer = 0
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
                                GroupType = ""
                                GroupLevel = ""
                                GroupTitle = ""
                                GroupHierarchicalContext = ""
                                AnalysisStatus = GroupAnalysisStatus.ReadGroupLevel
                                GroupLevel &= c
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

                                Dim Node As New _FormMain.GroupTreeNode("Root", 0)
                                Me.Nodes.Add(Node)
                                LastTreeViewNodeList.Add(Node)
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
                                    Dim Node As New _FormMain.GroupTreeNode(GroupTitle, GroupHierarchicalContext)
                                    If Val(GroupLevel) <= 0 Then
                                        MsgBox(GroupAnalysisError.ReadHierarchicalContext)
                                        Exit Sub
                                    End If

                                    If Val(GroupLevel) > LastTreeViewNodeList.Count - 1 Then
                                        LastTreeViewNodeList.Add(Node)
                                    Else
                                        LastTreeViewNodeList.Item(Val(GroupLevel)) = Node
                                    End If

                                    CType(LastTreeViewNodeList.Item(Val(GroupLevel) - 1), _FormMain.GroupTreeNode).Nodes.Add(Node)
                                    CurrentLevel = Val(GroupLevel)
                                    GroupHierarchicalContext = ""
                                    GroupLevel = ""
                                    GroupTitle = ""
                                    GroupType = ""
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
                    Case GroupAnalysisStatus.EndGroupProperty
                        Select Case c
                            Case " "
                                ' Do nothing
                            Case vbCr
                                ' Do nothing
                            Case Chr(10)
                                ' Do nothing
                            Case ":"
                                MsgBox(GroupAnalysisError.EndGroupProperty)
                                Exit Sub
                            Case "\"
                                MsgBox(GroupAnalysisError.EndGroupProperty)
                                Exit Sub
                            Case ";"
                                'GroupType = ""
                                'GroupLevel = ""
                                'GroupTitle = ""
                                'GroupHierarchicalContext = ""
                                AnalysisStatus = GroupAnalysisStatus.ReadGroupLevel
                            Case Else
                                AnalysisStatus = GroupAnalysisStatus.ReadBibTeXKey
                                BibTeXKey &= c
                        End Select
                    Case GroupAnalysisStatus.ReadBibTeXKey
                        Select Case c
                            Case " "
                                MsgBox(GroupAnalysisError.ReadBibTeXKey)
                                Exit Sub
                            Case vbCr
                                MsgBox(GroupAnalysisError.ReadBibTeXKey)
                                Exit Sub
                            Case Chr(10)
                                MsgBox(GroupAnalysisError.ReadBibTeXKey)
                                Exit Sub
                            Case "\"
                                If Not NextIsSemicolon(GroupBuffer, Index) Then
                                    MsgBox(GroupAnalysisError.ReadBibTeXKey)
                                    Exit Sub
                                End If

                                CType(LastTreeViewNodeList.Item(CurrentLevel), _FormMain.GroupTreeNode).BibTeXKeyList.Add(BibTeXKey)
                                BibTeXKey = ""
                                AnalysisStatus = GroupAnalysisStatus.EndBibTeXKey
                            Case ";"
                                MsgBox(GroupAnalysisError.ReadBibTeXKey)
                                Exit Sub
                            Case Else
                                BibTeXKey &= c
                        End Select
                    Case GroupAnalysisStatus.EndBibTeXKey
                        Select Case c
                            Case " "
                                ' Don othing
                            Case vbCr
                                ' Don othing
                            Case Chr(10)
                                ' Don othing
                            Case ":"
                                MsgBox(GroupAnalysisError.EndBibTeXKey)
                                Exit Sub
                            Case "\"
                                MsgBox(GroupAnalysisError.EndBibTeXKey)
                                Exit Sub
                            Case ";"
                                AnalysisStatus = GroupAnalysisStatus.Idle
                            Case Else
                                BibTeXKey &= c
                                AnalysisStatus = GroupAnalysisStatus.ReadBibTeXKey
                        End Select
                End Select
            Next

            Me.ExpandAll()
        End Sub

        Private Function NextIsSemicolon(ByVal GroupBuffer As String, ByRef Index As Integer) As Boolean
            If Index = GroupBuffer.Length - 1 Then
                Return False
            End If

            For i As Integer = Index + 1 To GroupBuffer.Length - 1
                Select Case GroupBuffer(i)
                    Case vbCr
                        ' Do nothing
                    Case Chr(10)
                        ' Do nothing
                    Case ";"
                        Index = i
                        Return True
                    Case Else
                        Return False
                End Select
            Next

            Return False
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
        EndBibTeXKey
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
        Public Const EndBibTeXKey As String = "End BibTeXKey error."
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

        Public Function ExistBibTeXKey(ByVal BibTeXKey As String) As Boolean
            For Each Key As String In Me.BibTeXKeyList
                If BibTeXKey.ToLower.Trim = Key.ToLower.Trim Then
                    Return True
                End If
            Next

            Return False
        End Function
    End Class
End Namespace
