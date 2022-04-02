Public Class Window_PDxF
    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)

        Dim path As String = ""
        Dim BrowserFolder As New Ookii.Dialogs.Wpf.VistaFolderBrowserDialog


        If BrowserFolder.ShowDialog() = True Then
            TextPath.Text = BrowserFolder.SelectedPath
        End If

    End Sub

    Private Sub CancelButton_Click(sender As Object, e As RoutedEventArgs)
        Me.Hide()
    End Sub

    Private Sub OKButton_Click(sender As Object, e As RoutedEventArgs)

        Dim bdxf As Boolean = False
        Dim bpdf As Boolean = False
        Dim bdwg As Boolean = False

        If Cpdf.IsChecked = True Then bpdf = True
        If CDxf.IsChecked = True Then bdxf = True
        If Cdwg.IsChecked = True Then bdwg = True

        If System.IO.Directory.Exists(TextPath.Text) Then
            GoToPDF(TextPath.Text, bpdf, bdxf, bdwg)

            Me.Hide()
        Else

            Dim mm As New MessageErreur("Séléctionner un dossier. Vérifier que le chemin séléctionné soit correct.", Notifications.Wpf.NotificationType.Warning)
        End If


    End Sub


End Class