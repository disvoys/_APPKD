Public Class WindowRenameDassault
    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)

        Dim s As String = TextMA.Text
        If s = "" Then s = "-vide-"
        If GetEnv() = "DASSAULT AVIATION" Then
            FctionCATIA.ArboDassault(s)
        ElseIf GetEnv = "AIRBUS" Then
            FctionCATIA.ArboAirbus(s)
        Else
            Dim m As New MessageErreur("La création d'arborescence n'est pas pris en charge avec l'environnement client sélectionné", Notifications.Wpf.NotificationType.Warning)
        End If

        Me.Hide()
    End Sub

    Private Sub Button_PreviewKeyDown(sender As Object, e As KeyEventArgs)

    End Sub

    Private Sub Button_Click_1(sender As Object, e As RoutedEventArgs)
        Me.Hide()
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)

    End Sub

    Private Sub Window_MouseMove(sender As Object, e As MouseEventArgs)

    End Sub

    Dim dernierChar As String = Nothing
    Private Sub TextMA_TextChanged(sender As Object, e As TextChangedEventArgs)
        If Env = "[DASSAULT AVIATION]" And dernierChar <> "-" Then
            Dim s As String = TextMA.Text
            If TextMA.Text.Length = 10 And Strings.Right(TextMA.Text, 1) <> "-" Then 'MA12300Z00
                TextMA.Text = TextMA.Text & "-"
                TextMA.Select(TextMA.Text.Length + 1, 1)
            End If
        End If
        dernierChar = Strings.Right(TextMA.Text, 1)
    End Sub
End Class
