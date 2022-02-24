Class Page2D
    Dim ListFiles2D As New List(Of String)

    Private Sub ButtonADFILES_Click(sender As Object, e As RoutedEventArgs)

        Dim path As String = ""
        If Not MonActiveDoc Is Nothing Then
            path = MonActiveDoc.Path
        End If
        Dim BrowserFolder As New Ookii.Dialogs.Wpf.VistaOpenFileDialog With {
         .Multiselect = True,
         .InitialDirectory = path,
         .Filter = "Fichier 2D CATIA (.CATDrawing)|*.CATDrawing",
         .Title = "Sélection des plans"
        }


        If BrowserFolder.ShowDialog() = True Then
            For Each item In BrowserFolder.FileNames
                ListFiles2D.Add(item)
            Next
        End If

        ListFiles.Items.Clear()
        For Each item In ListFiles2D
            Dim str() As String = Strings.Split(item, "\")
            Dim f As String = str(UBound(str))
            ListFiles.Items.Add(f)
        Next

        If ListFiles2D.Count > 0 Then
            Group1.IsEnabled = True
            Group2.IsEnabled = True
        End If


        If ListFiles2D.Count = 0 Then
            CountFiles.Content = "Aucun fichier séléctionné"
        ElseIf ListFiles2D.count = 1 Then
            CountFiles.Content = "1 fichier séléctionné"
        ElseIf ListFiles2D.count > 1 Then
            CountFiles.Content = ListFiles2D.Count & " fichiers séléctionnés"
        End If

    End Sub

    Private Sub ButtonDeletefile_Click(sender As Object, e As RoutedEventArgs)
        ListFiles.Items.Clear()
        ListFiles2D.Clear()
        Group1.IsEnabled = False
        Group2.IsEnabled = False

        CountFiles.Content = "Aucun fichier séléctionné"
    End Sub

    Private Sub Page_Loaded(sender As Object, e As RoutedEventArgs)
        Date1.Text = Date.Now
        ' Main.ChangeStatus("coucou")
    End Sub

    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)

        Dim bdxf As Boolean = False
        Dim bpdf As Boolean = False
        Dim bdwg As Boolean = False

        If Cpdf.IsChecked = True Then bpdf = True
        If CDxf.IsChecked = True Then bdxf = True
        If Cdwg.IsChecked = True Then bdwg = True

        GoToPDF("AppActivate", bpdf, bdxf, bdwg)
    End Sub

    Private Sub ButtonAddSuffixe_Click(sender As Object, e As RoutedEventArgs)

    End Sub

    Private Sub ButtonMajCartouche_Click(sender As Object, e As RoutedEventArgs)

        Dim GObOM As Boolean = False '=boolbom.ischecked


        Dim Da As String = Date1.Text.ToString
        Dim d() As String = Strings.Split(Da, "/")
        Dim Jour As Integer = d(0)
        Dim Mois As Integer = d(1)
        Dim annee As String = d(2)
        Da = Format(Jour, "00") & "-" & FctionCATIA.GetMonth(Mois) & "-" & Right(annee, 2)



        GoMajPlan(ListFiles2D, BoolDate.IsChecked, BoolNumOutillage.IsChecked, BoolDessinateur.IsChecked, GObOM, BoolSite.IsChecked, BoolTitre.IsChecked, BoolProgram.IsChecked, Da.ToString, TextPartnumber.Text.ToString, DESS.Text.ToString, SITE1.Text.ToString, Titre1.Text.ToString, Titre2.Text.ToString, Program1.Text.ToString, Program2.Text.ToString)

    End Sub
End Class
