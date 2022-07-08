
Public Class FormCollection
    Inherits System.Collections.CollectionBase

    ' Restricts to Form types, items that can be added to the collection.
    Public Sub Add(ByVal aform As Form)
        ' Invokes Add method of the List object to add a widget.
        List.Add(aform)
    End Sub

    Public Sub Remove(ByVal index As Integer)
        ' Check to see if there is a form at the supplied index.
        If index > Count - 1 Or index < 0 Then
            ' If no form exists, a messagebox is shown and the operation is 
            ' cancelled.
            System.Windows.Forms.MessageBox.Show("Index not valid!")
        Else
            ' Invokes the RemoveAt method of the List object.
            List.RemoveAt(index)
        End If
    End Sub

    Public ReadOnly Property Item(ByVal index As Integer) As Form
        Get
            ' The appropriate item is retrieved from the List object and 
            ' explicitly cast to the Form type, then returned to the 
            ' caller.
            Return CType(List.Item(index), Form)
        End Get
    End Property

End Class
