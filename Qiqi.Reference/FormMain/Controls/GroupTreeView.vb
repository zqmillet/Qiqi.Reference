Namespace _FormMain
    Public Class GroupTreeView
        Inherits System.Windows.Forms.TreeView

        Public Delegate Sub DelegateShowGroup(ByVal GroupBuffer As String)


        Public Sub Loading(ByVal GroupVersion As Integer, ByVal GroupBuffer As String)
            Select Case GroupVersion
                Case 3
                    If Me.InvokeRequired Then
                        Me.Invoke(New DelegateShowGroup(AddressOf ShowGroupVersion3), GroupBuffer)
                    Else
                        ShowGroupVersion3(GroupBuffer)
                    End If

                Case Else

            End Select
        End Sub

        Private Sub ShowGroupVersion3(ByVal GroupBuffer As String)
            Dim AnalysisState As Integer = GroupAnalysisStatus.Idle

            Dim GroupLevel As String = ""


            For Index As Integer = 0 To GroupBuffer.Length - 1
                Dim c As Char = GroupBuffer(Index)

                If (vbCr & Chr(10) & vbCrLf).Contains(c) Then
                    Continue For
                End If


                Select Case AnalysisState
                    Case GroupAnalysisStatus.Idle

                    Case GroupAnalysisStatus.ReadGroupLevel
                    Case GroupAnalysisStatus.ReadGroupTitle
                    Case GroupAnalysisStatus.ReadGroupType
                    Case GroupAnalysisStatus.ReadHierarchicalContext
                    Case GroupAnalysisStatus.ReadBibTeXKey
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
