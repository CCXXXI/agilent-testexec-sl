' Implements the manual sorting of items by columns.
Class ListViewItemComparer
    Implements IComparer

    Public Column As Integer

    Public Sub New()
        Column = 0
    End Sub

    Public Sub New(ByVal col As Integer)
        Me.Column = col
    End Sub

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
       Implements IComparer.Compare
        Return [String].Compare(CType(x, ListViewItem).SubItems(Column).Text, CType(y, ListViewItem).SubItems(Column).Text)
    End Function
End Class
