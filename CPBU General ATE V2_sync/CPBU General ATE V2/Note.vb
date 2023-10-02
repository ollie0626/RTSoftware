Public Class Note

    Private Sub Note_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub


    Private Sub Note_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class