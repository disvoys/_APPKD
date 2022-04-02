Public Class Window_UnitsDraw


    Private Sub CancelButton_Click(sender As Object, e As RoutedEventArgs)
        Me.Hide()
    End Sub


    Private Sub OKButton_Click(sender As Object, e As RoutedEventArgs)

        Dim bmm As Boolean = False
        Dim binch As Boolean = False
        Dim ball As Boolean = False

        If Cmm.IsChecked = True Then bmm = True
        If Cinch.IsChecked = True Then binch = True
        If C_all.IsChecked = True Then ball = True


        ChangeUnitsDrawing.MainUnitsDraw(bmm, binch, ball)
        Me.Hide()

    End Sub



End Class
