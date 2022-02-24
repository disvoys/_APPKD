Imports System.ComponentModel
Imports System.IO
Imports System.Threading

Class MainWindow

#Region "VISUEL"
    Private Sub TopPanel_MouseDown(sender As Object, e As MouseButtonEventArgs)



        If e.ClickCount = 1 Then Me.DragMove()

        If e.ClickCount = 2 Then
            If Me.WindowState = WindowState.Maximized Then
                Me.WindowState = WindowState.Normal
            Else
                Me.WindowState = WindowState.Maximized
            End If
        End If
    End Sub

    Public Sub MaListeMenu_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        For Each item As ListViewItem In MaListeMenu.Items
            If item Is MaListeMenu.SelectedItem Then
                Dim t As TextBlock
                t = item.Content.children.item(0)
                t.Foreground = New SolidColorBrush(Color.FromRgb(46, 208, 113))
            Else
                Try
                    Dim t As TextBlock
                    t = item.Content.children.item(0)
                    t.Foreground = New SolidColorBrush(Color.FromRgb(255, 255, 255))
                Catch ex As Exception
                End Try
            End If
        Next

        Select Case MaListeMenu.SelectedIndex
            Case 0
                PageLoc.Content = PRenommage
            Case 1
                PageLoc.Content = PNommenclature
            Case 2
                PageLoc.Content = PStructure
            Case 3
                PageLoc.Content = PComplements
        End Select



    End Sub

    Private Sub CloseButton_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
        End
    End Sub

    Private Sub MaxButton_Click(sender As Object, e As RoutedEventArgs)
        If Me.WindowState = WindowState.Maximized Then
            Me.WindowState = WindowState.Normal
        Else
            Me.WindowState = WindowState.Maximized
        End If
    End Sub

    Private Sub MinButton_Click(sender As Object, e As RoutedEventArgs)
        Me.WindowState = WindowState.Minimized
    End Sub

    Private Sub ButtonCloseMenu_Click(sender As Object, e As RoutedEventArgs)
        ButtonOpenMenu.Visibility = Visibility.Visible
        ButtonCloseMenu.Visibility = Visibility.Collapsed
    End Sub

    Private Sub ButtonOpenMenu_Click(sender As Object, e As RoutedEventArgs)
        ButtonOpenMenu.Visibility = Visibility.Collapsed
        ButtonCloseMenu.Visibility = Visibility.Visible
    End Sub



    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)

        Me.MaxHeight = SystemParameters.MaximizedPrimaryScreenHeight
        Me.MinHeight = 500
        Me.MinWidth = 350
        Bgw = New BackgroundWorker With {
            .WorkerReportsProgress = True,
            .WorkerSupportsCancellation = True
        }


        AddHandler Bgw.DoWork, AddressOf bgw_doWork
        AddHandler Bgw.ProgressChanged, AddressOf bgw_progresschange
        AddHandler Bgw.RunWorkerCompleted, AddressOf bgw_runworkercompleted

        Bgw.RunWorkerAsync()





    End Sub

    Private Sub Bgw_doWork(ByVal sender As Object, ByVal e As DoWorkEventArgs)



        INIFiles = New GestionINIFiles(DossierBase & "\Données\Request.ini")
        FichierTreeTxt = DossierBase & "\Données\CatiaTreeTxt.txt"
        TreeSauv = DossierBase & "\Données\CatiaTreeSauv.txt"

        'CatalogueMatieres = INIFiles.GetString("CATMATERIAL", "FichierCatMaterial", "")

        'If Not File.Exists(CatalogueMatieres) Then
        CatalogueMatieres = DossierBase & "\Données\_CATALOGUE MATIERES.CATMaterial"
        ' End If



        TxtListTraitement = DossierBase & "\Données\ListTraitements.txt"
        FctionCATIA.GetCATIA()

        FctionCATIA.CheckActiveDoc()

        FctionCATIA.SauvegardeTreeCatiaToTxtFile(FichierTreeTxt, TreeSauv)
        FctionCATIA.RemplirTreeViewFromTxtFile(TreeSauv, PRenommage.MonTV)
        FctionCATIA.GetMATERIAL()


    End Sub 'Start du backgroundworker
    Private Sub Bgw_runworkercompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs)

        Try
            Dispatcher.Invoke(New MyDelegate(AddressOf SHowFinal))
        Catch ex As Exception
            'err aucun document n'est ouvert
        End Try

    End Sub 'Complete
    Private Sub Bgw_progresschange(ByVal sender As Object, ByVal e As ProgressChangedEventArgs)

        Dim s As String
        Dim i As Double = Math.Round(e.ProgressPercentage)

        s = "En chargement ... | " & i & " %"
        LabelPercent.Content = s
    End Sub 'ProgressBar
    Delegate Sub MyDelegate()

    Sub SHowFinal()
        On Error Resume Next

        CreerTV(PRenommage.MonTV)
        RazListDocuments()
        PRenommage.MonDTG.ItemsSource = ColDoc
        _PB.Visibility = Visibility.Hidden
        LabelPercent.Visibility = Visibility.Hidden
        MaListeMenu.SelectedItem = MaListeMenu.Items(0)
        On Error GoTo 0
        Mise_a_Jour_POPUPMATERIAL()
        Mise_a_Jour_POPUPTraitement()



        StatusText.Text = "[" & MonActiveDoc.Name & "]" & " | " & NbElements - 1 & " éléments trouvés"
        PageLoc.Visibility = Visibility.Visible


    End Sub 'MaJ Visuelle

#Region "PopupMatériau"
    Public Sub Mise_a_Jour_POPUPMATERIAL()

        For Each item As ItemFamille In ListFamilleMaterials
            Dim t As New Expander With {
                .Header = item.Name
            }
            PRenommage.MenuMaterial.Children.Add(t)
            Dim wp As New WrapPanel With {
                .Width = 240
            }

            For Each m As ItemMaterial In item.Materials
                Dim _t As New Button With {
                    .Content = m.Name,
                    .Width = 80,
                    .Background = Brushes.White
                }
                AddHandler _t.Click, AddressOf PRenommage.ButtonListMaterial_click
                wp.Children.Add(_t)
                '  t.Items.Add(_t)
            Next
            t.Content = wp
        Next
    End Sub
#End Region

#Region "PopupTraitement"

    Sub Mise_a_Jour_POPUPTraitement()

        Using sr As StreamReader = New StreamReader(TxtListTraitement)
            Dim Line As String
            Line = sr.ReadLine
            Do While (Not Line Is Nothing)

                Dim l() As String = Split(Line, ";")

                If l(0) = "[SEPARATION]" Then
                    Dim s As New Separator
                    PRenommage.MenuTraitements.Children.Add(s)
                Else
                    Dim t As New Button
                    t.Content = l(0)
                    t.Style = Application.Current.Resources("PopupButtonGrid")

                    AddHandler t.Click, AddressOf PRenommage.ButtonListTraitement_click

                    If l.Length > 0 And l.Length < 500 Then 'SUPPRIMER LE 500 - POUR TEST
                        For i As Integer = 1 To l.Length - 1
                            If l(i) = "[SEPARATION]" Then
                                Dim s_ As New Separator
                                ' t.Items.Add(s_)
                            Else
                                Dim t_ As New Button
                                t_.Content = l(i)

                                ' t.Items.Add(t_)
                                t.Style = FindResource("PopupButtonGrid")
                            End If

                        Next
                    End If

                    PRenommage.MenuTraitements.Children.Add(t)
                End If

                Line = sr.ReadLine
            Loop

        End Using

    End Sub
#End Region



#Region "Recursive Nodes Treeview"

    Function CreerNewNode(IC As ItemCatia) As TreeViewItem

        Dim Tn As New TreeViewItem
        Dim Lab As New Label
        Dim st As New StackPanel
        Dim Im As New Image With {
            .Width = 16,
            .Height = 16,
            .Source = New BitmapImage(New Uri(DossierImage & IC.Image))
        }

        Dim MonBind As New Binding("TextTv") With {
            .Source = IC
        }
        Lab.SetBinding(Label.ContentProperty, MonBind)
        ''   Lab.Content = MonBind

        st.Orientation = Orientation.Horizontal
        st.Children.Add(Im)
        st.Children.Add(Lab)
        Tn.Header = st
        Tn.DataContext = IC
        Return Tn

    End Function
    Sub CreerTV(monTV As TreeView)

        ListDocuments(0).TreeViewItem = CreerNewNode(ListDocuments(0))
        monTV.Items.Add(ListDocuments(0).TreeViewItem)
        RootCreerTV(ListDocuments(0))

        monTV.Items(0).IsExpanded = True

    End Sub
    Sub RootCreerTV(item As ItemCatia)

        Dim MonTvParent As TreeViewItem = item.TreeViewItem
        For Each enfant As ItemCatia In item.Enfants
            If ListItemDejaOK.Contains(enfant.PartNumber) Then
                For Each i As ItemCatia In ListDocuments
                    If i.PartNumber = enfant.PartNumber Then
                        i.Qte += 1
                        Exit For
                    End If
                Next
            Else
                ListItemDejaOK.Add(enfant.PartNumber)
            End If
            enfant.TreeViewItem = CreerNewNode(enfant)
            MonTvParent.Items.Add(enfant.TreeViewItem)
            If enfant.Enfants.Count > 0 Then
                RootCreerTV(enfant)
            End If
        Next

    End Sub

    Dim ListItemDejaOK As New List(Of String)

    Sub RazListDocuments()

        For Each item In ListDocuments
            ListItemCatia.Add(item)
        Next

        ListDocuments.Clear()

        Dim ListOfIdOK As New List(Of String)

        For Each i As ItemCatia In ListItemCatia
            If Not ListOfIdOK.Contains(i.PartNumber) Then
                ListDocuments.Add(i)
                ListOfIdOK.Add(i.PartNumber)
            End If
        Next

        ColDoc = New ListCollectionView(ListDocuments)
        ColDoc.SortDescriptions.Add(New SortDescription("Type", ListSortDirection.Descending))

    End Sub

    Private Sub MenuButton_Click(sender As Object, e As RoutedEventArgs)
        PopupMenu.IsOpen = True

    End Sub

    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)
        Process.Start(Application.ResourceAssembly.Location)
        Application.Current.Shutdown()
    End Sub

    Private Sub Button_Click_1(sender As Object, e As RoutedEventArgs)
        End
    End Sub


#End Region



#End Region





End Class
