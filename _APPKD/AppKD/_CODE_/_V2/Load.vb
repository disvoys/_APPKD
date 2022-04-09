Imports System.ComponentModel

Public Class Load

    Delegate Sub MyDelegate()

    Sub Main()



        Bgw = New BackgroundWorker With {
            .WorkerReportsProgress = True,
            .WorkerSupportsCancellation = True
        }

        AddHandler Bgw.DoWork, AddressOf bgw_doWork
        AddHandler Bgw.ProgressChanged, AddressOf bgw_progresschange
        AddHandler Bgw.RunWorkerCompleted, AddressOf bgw_runworkercompleted


        Bgw.ReportProgress(0)
        Bgw.RunWorkerAsync()



    End Sub

#Region "MultiThread"

    Private Sub bgw_doWork(ByVal sender As Object, ByVal e As DoWorkEventArgs)

        TypeActiveDoc = Nothing

        Try
            Bgw.ReportProgress(0, "Initilisation des données...")
        Catch ex As Exception
            Dim m As New MessageErreur("Erreur dans les données d'entrée. Réinstaller l'application", Notifications.Wpf.NotificationType.Error)
        End Try
        Try
            Bgw.ReportProgress(0, "Chargement des noeuds, patienter...")
            FctionCATIA.GetCATIA()
        Catch ex As Exception
            Dim m As New MessageErreur("Une erreur s'est produite. Impossible de lier l'application à Catia. Vérifier que Catia soit ouvert", Notifications.Wpf.NotificationType.Error)
            Exit Sub

        End Try
        Try
            FctionCATIA.CheckActiveDoc()
        Catch ex As Exception
            MsgBox("Aucun document ouvert. Ouvrir un product", vbCritical)
        End Try
        If TypeActiveDoc = "PRODUCT" Then
            Try
                Bgw.ReportProgress(0, "Récupération de l'arbre CATIA...")
                FctionCATIA.SauvegardeTreeCatiaToTxtFile(FichierTreeTxt, TreeSauv)
            Catch ex As Exception
                MsgBox("Erreur lors de l'import de l'arbre dans l'application.", vbCritical)
            End Try

            Try
                FctionCATIA.GoListDocuments()
            Catch ex As Exception
                MsgBox("Erreur lors du listing des documents ouverts. " & ex.Message, vbCritical)
            End Try


        ElseIf TypeActiveDoc = "PART" Then

            ListPropertiesPart.Clear()
            FctionCATIA.GoPart()

        Else
            Exit Sub
        End If

        Try
            FctionCATIA.getMATERIAL()
        Catch ex As Exception
            MsgBox("Erreur lors de la lecture du fichier CATMaterial. Si vous utilisez CATIA R21, vérifier que le fichier CATMaterial ne soit pas en R27.", vbCritical)
        End Try

        Bgw.ReportProgress(100, "Calcul à 100%, finalisation des propriétés graphiques.")



    End Sub 'Start du backgroundworker


    Private Sub bgw_runworkercompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs)

        If NbElements > 0 Then
            If TypeActiveDoc = "PRODUCT" Then
                SHowFinal()
                MonMainV3.MonTV.Visibility = Visibility.Visible
                MonMainV3.ButtonCatiaTest.Visibility = Visibility.Visible
                MonMainV3.LoadProgressBar.Visibility = Visibility.Hidden

                If MonMainV3.GridLoaded.Visibility = Visibility.Visible Then
                    MonMainV3.GridContentBDD.Visibility = Visibility.Visible
                End If
                MonMainV3.GridLoaded.Visibility = Visibility.Collapsed
                MonMainV3.ButtonHome.IsEnabled = True

            End If
            If TypeActiveDoc = "PART" Then

                FctionCATIA.RemplirTVPart()
                ColDoc = New ListCollectionView(ListPropertiesPart)
                MonMainV3.DataGridPropertiesPart.ItemsSource = ColDoc
                MonMainV3.TVCatpart.Visibility = Visibility.Visible
                MonMainV3.GridContentBDD.Visibility = Visibility.Hidden
                MonMainV3.ButtonCatiaTest.Visibility = Visibility.Visible
                MonMainV3.LoadProgressBar.Visibility = Visibility.Hidden
                MonMainV3.GridLoaded.Visibility = Visibility.Collapsed
                MonMainV3.GridContentCATPART.Visibility = Visibility.Visible
            End If
        Else
            Try
                FctionCATIA.GetCATIA()
                If MsgBox("Aucun document ne semble ouvert. Créer une nouvelle arborescence ?", MsgBoxStyle.YesNo + MsgBoxStyle.Information) = MsgBoxResult.Yes Then
                    WindowRenameDassault_.ShowDialog()
                    Main()
                Else
                    MonMainV3.GridContentBDD.Visibility = Visibility.Hidden
                    MonMainV3.LabelLoad.Visibility = Visibility.Hidden
                    MonMainV3.ButtonCatiaTest.Visibility = Visibility.Visible
                    MonMainV3.LoadProgressBar.Visibility = Visibility.Hidden

                End If
            Catch ex As Exception
                MonMainV3.GridContentBDD.Visibility = Visibility.Hidden
                MonMainV3.LabelLoad.Visibility = Visibility.Hidden
                MonMainV3.ButtonCatiaTest.Visibility = Visibility.Visible
                MonMainV3.LoadProgressBar.Visibility = Visibility.Hidden
            End Try
        End If





    End Sub 'Complete


    Private Sub bgw_progresschange(ByVal sender As Object, ByVal e As ProgressChangedEventArgs)


        Dim s As String
        Dim i As String = Math.Round(e.ProgressPercentage)

        s = "Chargement ... | " & i & " %"
        MonMainV3.LabelLoad.Content = s


    End Sub 'ProgressBar

    Sub RemplirMenuMaterialsTTS()

        For Each item As ItemFamille In ListFamilleMaterials
            Dim t As New Expander With {
                .Header = item.Name,
                .FontSize = 11
            }

            MonMainV3.MenuMaterial.Children.Add(t)
            Dim wp As New WrapPanel With {
                .Width = 310
            }

            For Each m As ItemMaterial In item.Materials
                Dim _t As New Button With {
                    .Content = m.Name,
                    .Foreground = Brushes.Black,
                    .FontSize = 11,
                    .BorderBrush = Brushes.Black,
                    .Width = 100,
                    .Background = Brushes.White,
                    .Margin = New Thickness(3, 0, 0, 3)
                }
                AddHandler _t.Click, AddressOf MonMainV3.ButtonListMaterial_click
                wp.Children.Add(_t)
                '  t.Items.Add(_t)
            Next
            t.Content = wp
        Next

        RemplirMenuTTS("TRAITEMENT ACIER", "TRAITEMENT_ACIER", 300)
        RemplirMenuTTS("TRAITEMENT ALU", "TRAITEMENT_ALU", 300)
        RemplirMenuTTS("PLANCHER ANTI-DERAPANT", "PLANCHER_ANTIDERAPANT", 300)
        RemplirMenuTTS("PEINTURE", "PEINTURE", 300)




    End Sub
    Sub SHowFinal()

        BoolStartBOM = False
        MonMainV3.ComboBOM.Items.Clear()


        MonMainV3.ComboBOM.Items.Add("Ensemble des éléments")
        MonMainV3.ComboBOM.SelectedIndex = 0

        If MonMainV3.MenuMaterial.Children.Count = 0 Then RemplirMenuMaterialsTTS()

        If TypeActiveDoc = "PRODUCT" Then
            Try
                FctionCATIA.RemplirTreeViewFromTxtFile(TreeSauv, MonMainV3.MonTV)
            Catch ex As Exception
                MsgBox("Erreur lors du remplissage de l'arbre. =>" & ex.Message, vbCritical)
                AfficherTableau()
            End Try
            Try
                CreerTV(MonMainV3.MonTV)
            Catch ex As Exception
                MsgBox("Erreur lors de la création de l'arbre =>" & ex.Message, vbCritical)
                AfficherTableau()
            End Try
            CreationNommenclature()
        End If
        AfficherTableau()

        FctionGetBOM.GoBOM("Ensemble des éléments")

    End Sub 'MaJ Visuelle

    Sub RemplirMenuTTS(Titre As String, Key As String, width As Integer)

        Dim e1 As New Expander With {
              .Header = Titre,
                .FontSize = 11
          }

        Dim wp1 As New WrapPanel With {
                .Width = 310
            }
        Dim str As String = INIFiles.GetString("TTS", Key, "")
        Dim s() As String = Split(str, "//")
        For Each item In s
            Dim _t As New Button With {
                    .Content = item,
                    .Width = 100,
                    .Foreground = Brushes.Black,
                    .FontSize = 11,
                    .Background = Brushes.White,
                    .BorderBrush = Brushes.Black,
                    .Margin = New Thickness(3, 0, 0, 3)
                }
            AddHandler _t.Click, AddressOf MonMainV3.ButtonListMenu
            wp1.Children.Add(_t)
        Next
        e1.Content = wp1

        MonMainV3.MenuTraitements.Children.Add(e1)


    End Sub
    Sub CreationNommenclature()

        Try
            Dim str() As String = Strings.Split(MonActiveDoc.Name, ".CATProduct")
            Dim str_ As String = MonActiveDoc.Product.Name 'str(0) 'kevin à corriger
            MonMainV3.ComboBOM.Items.Add(str_)


            Dim l As New List(Of String)

            For Each ic As ItemCatia In ListDocuments
                If ic.Type = "PRODUCT" Then
                    Dim strIC() = Strings.Split(ic.Doc.Name, ".CATProduct")
                    Dim strIC_ As String = ic.PartNumber  ' strIC(0)
                    l.Add(strIC_) 'kevin à corriger
                End If
            Next


            l.Sort()

            For Each item In l
                MonMainV3.ComboBOM.Items.Add(item)
            Next
        Catch ex As Exception
        End Try


    End Sub

    Sub AfficherTableau()
        ColDoc = New ListCollectionView(ListDocuments)
        ColDoc.SortDescriptions.Add(New SortDescription("Type", ListSortDirection.Descending))
        MonMainV3.MaDataGrid.ItemsSource = ColDoc

        Dim str_ As String = Nothing
        If NbElements - 1 > 1 Then
            str_ = " éléments trouvés"
        Else
            str_ = " élément trouvé"
        End If

        MonMainV3.TextNbElements.Text = NbElements - 1 & str_

    End Sub
    Private Sub CancelButton_Click(sender As Object, e As RoutedEventArgs)
        Bgw.CancelAsync()
        ListDocuments.Clear()
        AfficherTableau()


    End Sub
#End Region 'BackGround Worker

#Region "Treeview"


    Sub CreerTV(MonTV As TreeView)

        MonTV.Items.Add(ITRacine.TVitem)
        MonTV.Items(0).IsExpanded = True

    End Sub






#End Region



End Class
