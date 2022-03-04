
'couleur blanche : #F9F9F9
'couleur bleue : #0E1D31
'couleur rose : #EE4865
'couleur vert : #76C2AF

Imports System.ComponentModel
Imports System.IO
Imports System.Windows.Controls.Primitives
Imports KnowledgewareTypeLib
Imports Microsoft.Win32
Imports ProductStructureTypeLib

Public Class MainV3

    Dim bgwProgress As BackgroundWorker
    Dim NeedReload As Boolean = False
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)


        Getname()

        URLFolderSTEPReception.Text = Path.GetTempPath
        MonMainV3 = Me
        GridContentBDD.Visibility = Visibility.Collapsed
        GridChart.Visibility = Visibility.Collapsed
        GridBibliotheque.Visibility = Visibility.Collapsed
        GridAbout.Visibility = Visibility.Collapsed
        GridAdmin.Visibility = Visibility.Collapsed
        GridLoaded.Visibility = Visibility.Visible
        GridStep.Visibility = Visibility.Collapsed
        GridContentCATPART.Visibility = Visibility.Collapsed
        ListViewPDF.SelectedIndex = 0

        VerifDroitUseApp()

        ListeEnvironements()
        CreerCheckBoxEnvironnementSettings()

        Try
            INIFiles = New GestionINIFiles(DossierBase & "\Données\Request.ini")
            FichierTreeTxt = DossierBase & "\Données\CatiaTreeTxt.txt"
            TreeSauv = DossierBase & "\Données\CatiaTreeSauv.txt"
            CatalogueMatieres = DossierBase & "\Données\_CATALOGUE MATIERES_v2.CATMaterial"
        Catch ex As Exception
            Dim m As New MessageErreur("Des fichiers sont manquants pour que l'application fonctionne correctement. Réinstaller l'application.", Notifications.Wpf.NotificationType.Error)
        End Try

        CheckLeToogleClient(My.Settings.Client.ToString)

        TextPlant.Text = My.Settings.PLANT.ToString
        TextDessinateur.Text = My.Settings.DRN.ToString
        TextTitre1.Text = My.Settings.TITLE1.ToString
        TextTitre2.Text = My.Settings.TITLE2.ToString
        TextProgram.Text = My.Settings.PROGRAM.ToString
        TextCageCode.Text = My.Settings.CAGECODE.ToString
        ColDoc.Refresh()

        Try
            FctionCATIA.GetCATIA()
        Catch ex As Exception
        End Try
        If CATIA Is Nothing Then
            LabelNameCatiaProduct.Content = "[Impossible de lier l'application avec CATIA]"
            '    ButtonCatiaTest.IsEnabled = False
        Else
            Try
                LabelNameCatiaProduct.Content = "[" & CATIA.ActiveDocument.Name & "]"
            Catch ex As Exception
                LabelNameCatiaProduct.Content = "[Aucun élément ne semble ouvert dans CATIA]"
                '       ButtonCatiaTest.IsEnabled = False
            End Try

        End If

        aggrandirFenetre()

        Exit Sub
        'PARTIE ADMIN DESACTIVEE
        If cSQL.IsPortOpen("db4free.net", 3306) = True Then
            If My.Computer.Network.IsAvailable Then
                cSQL.ConnectionToDB()
                cSQL.CheckUser()
                LoadUsers()
            Else
                Dim merr As New MessageErreur("Impossible de vérifier les mises à jour, vérifier votre connexion internet", Notifications.Wpf.NotificationType.Error)
            End If
        Else
            Dim merr As New MessageErreur("Erreur : Ouvrir le port TCP 3306. Contacter l'administrateur.", Notifications.Wpf.NotificationType.Error)
        End If

    End Sub 'LOAD

    Sub VerifDroitUseApp()
        If FctionGetInfo.GetCarteMere() = "NBQ17110016440771E7200" Then

        Else
            If File.Exists("\\multilauncher\AppKDData\_ne pas supprimer.txt") Then
                'verif fichier droit utilisateur selon numero carte mere
            Else
                MsgBox("Vérifier la connexion avec le réseau. Impossible de récupérer les fichiers eXcent. Contacter l'administrateur.", MsgBoxStyle.Critical)
                End
            End If
        End If

        If DateTime.Compare(My.Settings.Shutdown, DateTime.Now) < 0 Then
            MsgBox("Vous n'avez plus l'autorisation d'utiliser l'application, contacter l'administrateur", MsgBoxStyle.Critical)
            End
        End If
    End Sub
    Function GetEnv() As String
        Dim n As String = Env
        n = Strings.Replace(n, "[", "")
        n = Strings.Replace(n, "]", "")
        Return n
    End Function
    Sub CheckLeToogleClient(s As String)
        For Each g As ToggleButton In ListToogleSettingsEnv
            If g.Tag = s Then
                g.IsChecked = True
                Env = g.Tag
                If g.Tag = "[ENVIRONNEMENT PERSO]" Then
                    ToolbarPerso.IsEnabled = True
                    TabOptions.AllowDrop = True
                Else
                    ToolbarPerso.IsEnabled = False
                    TabOptions.AllowDrop = False
                End If
                Dim f As String = GetAs(INIProperties.GetString(GetEnv, "FichierProperties", ""))

                RemplirTableauEnvPerso(f)
                labelClient.Text = GetEnv()
            End If
        Next
    End Sub
    Sub ListeEnvironements()
        INIProperties = New GestionINIFiles(DossierBase & "\Données\Environnements.ini")
        Using sr As New StreamReader(INIProperties.FileName)
            While Not sr.EndOfStream
                Dim l As String = sr.ReadLine
                If Strings.Left(l, 1) = "[" And Strings.Right(l, 1) = "]" Then
                    ListEnvironnements.Add(l)
                End If
            End While
        End Using
    End Sub

    Private Sub ToogleSettingsEnv_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)

        Dim s As String = GetEnv()

        For Each g As ToggleButton In ListToogleSettingsEnv
            If sender.tag = g.Tag Then
                g.IsChecked = True
                Env = g.Tag
                Dim n As String = Strings.Replace(g.Tag, "[", "")
                n = Strings.Replace(n, "]", "")
                labelClient.Text = n
                My.Settings.Client = g.Tag
                My.Settings.Save()
                Dim f As String = INIProperties.GetString(n, "FichierProperties", "")
                If f = "" Then
                    DataGridProperties.Items.Clear()
                    DataGridProperties.Items.Refresh()
                Else
                    f = GetAs(f)
                    RemplirTableauEnvPerso(f)
                End If
                If g.Tag = "[ENVIRONNEMENT PERSO]" Then
                    ToolbarPerso.IsEnabled = True
                    TabOptions.AllowDrop = True
                Else
                    ToolbarPerso.IsEnabled = False
                    TabOptions.AllowDrop = False
                End If

            Else
                g.IsChecked = False
            End If
            If GetEnv() <> s Then
                NeedReload = True
            End If
        Next
    End Sub

    Sub CreerCheckBoxEnvironnementSettings()

        ListToogleSettingsEnv.Clear()

        For Each n As String In ListEnvironnements
            Dim s As New StackPanel
            Dim g As New ToggleButton
            Dim t As New TextBlock

            s.Orientation = Orientation.Horizontal
            s.Margin = New Thickness(0, 0, 10, 10)

            g.IsChecked = False
            g.Background = New SolidColorBrush(Color.FromRgb(118, 194, 175))
            g.VerticalAlignment = VerticalAlignment.Top
            g.Margin = New Thickness(0, 0, 10, 0)
            g.Tag = n
            s.Tag = n
            ListToogleSettingsEnv.Add(g)

            n = Strings.Replace(n, "[", "")
            n = Strings.Replace(n, "]", "")
            t.Text = n
            t.VerticalAlignment = VerticalAlignment.Top

            s.Children.Add(g)
            s.Children.Add(t)
            StackPanelSettingsEnv.Children.Add(s)

            AddHandler g.Click, AddressOf ToogleSettingsEnv_Click
            AddHandler s.mousedown, AddressOf ToogleSettingsEnv_Click

        Next


        ' StackPanelSettingsEnv
    End Sub
    Sub LoadUsers()
        cSQLUsers.Load()
        ColDocUsers = New ListCollectionView(ListUsers)
        DataGridAdmin.ItemsSource = ColDocUsers
    End Sub
    Private Sub Window_Closed(sender As Object, e As EventArgs)
        cSQL.CLoseConnexion()
        If Directory.Exists("C:\Users\deske\Google Drive\TAFF\_APPKD\AppKD\bin\Debug\app.publish") Then Directory.Delete("C:\Users\deske\Google Drive\TAFF\_APPKD\AppKD\bin\Debug\app.publish")
        End
    End Sub 'CLOSE
    Private Sub Window_MouseMove(sender As Object, e As MouseEventArgs)

        Try
            FctionCATIA.GetCATIA()
        Catch ex As Exception
        End Try

        If CATIA Is Nothing Then
            LabelNameCatiaProduct.Content = "[Impossible de lier l'application avec CATIA]"
            '   ButtonCatiaTest.IsEnabled = False
        Else
            Try
                LabelNameCatiaProduct.Content = "[" & CATIA.ActiveDocument.Name & "]"
                '     ButtonCatiaTest.IsEnabled = True
            Catch ex As Exception
                LabelNameCatiaProduct.Content = "[Aucun élément ne semble ouvert dans CATIA]"
                '       ButtonCatiaTest.IsEnabled = False
            End Try


            Try
                If CATIA.ActiveDocument.Name <> ICRacine.Owner And GetPartOrProduct(CATIA.ActiveDocument.FullName) = True Then
                    TextProductActif.Text = ICRacine.Owner & " | [" & CATIA.ActiveDocument.Name & "]"
                    BadgeReload.Badge = "!"

                Else
                    TextProductActif.Text = ICRacine.Owner
                    If NeedReload = True Then
                        BadgeReload.Badge = "!"
                    Else
                        BadgeReload.Badge = ""
                    End If
                End If

            Catch ex As Exception

            End Try
        End If

        VerifFenetre()

    End Sub 'MOUSE MOVE


    Sub GoProgress(s As String)

        bgwProgress = New BackgroundWorker With {
            .WorkerReportsProgress = True,
            .WorkerSupportsCancellation = True
        }

        Select Case s
            Case "Fix"
                AddHandler bgwProgress.DoWork, AddressOf bgwProgressFix_doWork
            Case "Masse"
                AddHandler bgwProgress.DoWork, AddressOf bgwProgressMasse_doWork
            Case "Instance"
                AddHandler bgwProgress.DoWork, AddressOf bgwProgressInstance_doWork
            Case "Properties"
                AddHandler bgwProgress.DoWork, AddressOf bgwProgressProperties_doWork
        End Select

        AddHandler bgwProgress.RunWorkerCompleted, AddressOf bgwProgress_runworkercompleted

        bgwProgress.RunWorkerAsync()
    End Sub

    Private Sub bgwProgressProperties_doWork(ByVal sender As Object, ByVal e As DoWorkEventArgs)
        FctionCATIA.ResetPropoerties()
        Dim m As New MessageErreur("Les propriétés ont supprimées avec succès. Il est préférable de relancer un calcul de l'application pour mettre à jour le tableau.", Notifications.Wpf.NotificationType.Information)
    End Sub 'Start du backgroundworker
    Private Sub bgwProgressFix_doWork(ByVal sender As Object, ByVal e As DoWorkEventArgs)
        FixAll.CATMain(MonActiveDoc)
    End Sub 'Start du backgroundworker
    Private Sub bgwProgressMasse_doWork(ByVal sender As Object, ByVal e As DoWorkEventArgs)
        Dim MaMasse As Double = 0
        Try
            For Each ic As ItemCatia In ListDocuments
                MaMasse = Math.Round(ic.ProductCATIA.Analyze.Mass, 2)
                Dim sm As String = INIProperties.GetString(GetEnv, "ProprieteMASSE", "MASS")
                If Env = "[SPIRIT AEROSYSTEMS]" Then
                    If ic.Source = "Acheté" Then
                    Else
                        Try
                            If ic.ProductCATIA.UserRefProperties.Item(sm).ValueAsString = "" Or ic.ProductCATIA.UserRefProperties.Item(sm).ValueAsString = "-" Then
                                ic.ProductCATIA.UserRefProperties.Item(sm).ValuateFromString(MaMasse & " Kg")
                            Else
                                ic.ProductCATIA.UserRefProperties.Item(sm).ValuateFromString(ic.ProductCATIA.UserRefProperties.Item(sm).ValueAsString & " - " & MaMasse & " Kg")
                            End If
                        Catch ex As Exception
                            FctionCATIA.AddParamatres(ic.Owner, ic)
                            If ic.ProductCATIA.UserRefProperties.Item(sm).ValueAsString = "" Or ic.ProductCATIA.UserRefProperties.Item(sm).ValueAsString = "-" Then
                                ic.ProductCATIA.UserRefProperties.Item(sm).ValuateFromString(MaMasse & " Kg")
                            Else
                                ic.ProductCATIA.UserRefProperties.Item(sm).ValuateFromString(ic.ProductCATIA.UserRefProperties.Item(sm).ValueAsString & " - " & MaMasse & " Kg")
                            End If
                        End Try
                        ic.StockSize = ic.ProductCATIA.UserRefProperties.Item(sm).ValueAsString
                    End If
                Else
                    Try
                        ic.ProductCATIA.UserRefProperties.Item(sm).ValuateFromString(MaMasse)
                    Catch ex As Exception
                        FctionCATIA.AddParamatres(ic.Owner, ic)
                        ic.ProductCATIA.UserRefProperties.Item(sm).ValuateFromString(MaMasse)
                    End Try
                    If Env = "[DASSAULT AVIATION]" Then
                        ic.Perso5 = MaMasse
                    End If
                End If


            Next
        Catch ex As Exception
        End Try


        Dim m As New MessageErreur("Calul des masses terminé", Notifications.Wpf.NotificationType.Information)

    End Sub 'Start du backgroundworker
    Private Sub bgwProgressInstance_doWork(ByVal sender As Object, ByVal e As DoWorkEventArgs)
        FctionCATIA.BRename()
    End Sub 'Start du backgroundworker

    Private Sub bgwProgress_runworkercompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs)


        ProgressTableau.Visibility = Visibility.Hidden
        Me.MaDataGrid.Items.Refresh()

    End Sub 'Complete


    Private Sub bgw_progresschange(ByVal sender As Object, ByVal e As ProgressChangedEventArgs)


        Dim s As String
        Dim i As String = Math.Round(e.ProgressPercentage)

        s = "Chargement ... | " & i & " %"
        MonMainV3.LabelLoad.Content = s


    End Sub 'ProgressBar




    Function GetPartOrProduct(s As String) As Boolean


        If Right(s, 4) = "Part" Then Return True
        If Right(s, 7) = "Product" Then Return True
        If Right(s, 7) = "Drawing" Then Return False


        Return False
    End Function
    Private Sub Grid_MouseDown_1(sender As Object, e As MouseButtonEventArgs)
        If e.LeftButton = MouseButtonState.Pressed And e.ClickCount = 2 Then
            aggrandirFenetre()

        End If

        Try
            Me.DragMove()
        Catch ex As Exception
        End Try

    End Sub 'MOOVE WINDOW

    Sub aggrandirFenetre()
        Me.WindowStyle = WindowStyle.SingleBorderWindow
        If Me.WindowState = Forms.FormWindowState.Maximized Then
            Me.WindowState = WindowState.Normal
        Else
            Me.WindowState = Forms.FormWindowState.Maximized
        End If
        Me.WindowStyle = WindowStyle.None
    End Sub

    Private Sub ButtonMaximize_Click(sender As Object, e As RoutedEventArgs)

        aggrandirFenetre()



    End Sub 'window maximized
    Private Sub ButtonMinimize_Click(sender As Object, e As RoutedEventArgs)
        Me.WindowState = Forms.FormWindowState.Minimized
    End Sub 'Reduction de l'application
    Private Sub Logo_MouseUp(sender As Object, e As MouseButtonEventArgs)


        If Me.GridaDeplacer.Width = New GridLength(35) Then
            Me.GridaDeplacer.Width = New GridLength(180)
            Me.Logo.IsEnabled = False
        Else
            Me.GridaDeplacer.Width = New GridLength(35)
            Me.Logo.IsEnabled = True
        End If

    End Sub 'voir le menu left panel
    Private Sub ButtonClose_Click(sender As Object, e As RoutedEventArgs)
        cSQL.CLoseConnexion()
        End
    End Sub 'Close de l'application
    Private Sub MaListView_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        Dim index As Integer = MaListView.SelectedIndex
        TransitioningContentSlide.OnApplyTemplate()
        GridTOMove.Margin = New Thickness(0, (0 + (32 * index)), 0, 0)

        Select Case index
            Case 0 'CATIA
                If NbElements > 0 Then
                    If TypeActiveDoc = "PRODUCT" Then
                        GridContentBDD.Visibility = Visibility.Visible
                    ElseIf TypeActiveDoc = "PART" Then
                        GridContentCATPART.Visibility = Visibility.Visible
                    End If
                    GridLoaded.Visibility = Visibility.Collapsed
                Else
                    GridContentBDD.Visibility = Visibility.Hidden
                    GridContentCATPART.Visibility = Visibility.Hidden
                    GridLoaded.Visibility = Visibility.Visible
                End If

                GridChart.Visibility = Visibility.Collapsed
                GridBibliotheque.Visibility = Visibility.Collapsed
                GridAbout.Visibility = Visibility.Collapsed
                GridAdmin.Visibility = Visibility.Collapsed
                GridMacros.Visibility = Visibility.Collapsed
                GridPDF.Visibility = Visibility.Collapsed
                GridStep.Visibility = Visibility.Collapsed


            Case 1 'MACROS
                GridContentBDD.Visibility = Visibility.Collapsed
                GridChart.Visibility = Visibility.Collapsed
                GridBibliotheque.Visibility = Visibility.Collapsed
                GridAbout.Visibility = Visibility.Collapsed
                GridLoaded.Visibility = Visibility.Collapsed
                GridMacros.Visibility = Visibility.Visible
                GridPDF.Visibility = Visibility.Collapsed
                GridAdmin.Visibility = Visibility.Collapsed
                GridStep.Visibility = Visibility.Collapsed
                GridContentCATPART.Visibility = Visibility.Collapsed
            Case 2 'BIBLIOTHEQUE
                GridContentBDD.Visibility = Visibility.Collapsed
                GridChart.Visibility = Visibility.Collapsed
                GridBibliotheque.Visibility = Visibility.Visible
                GridAbout.Visibility = Visibility.Collapsed
                GridMacros.Visibility = Visibility.Collapsed
                GridLoaded.Visibility = Visibility.Collapsed
                GridPDF.Visibility = Visibility.Collapsed
                GridAdmin.Visibility = Visibility.Collapsed
                GridStep.Visibility = Visibility.Collapsed
                GridContentCATPART.Visibility = Visibility.Collapsed
            Case 3 'SEPARATOR

            Case 4 'PDF
                GridContentBDD.Visibility = Visibility.Collapsed
                GridChart.Visibility = Visibility.Collapsed
                GridBibliotheque.Visibility = Visibility.Collapsed
                GridAbout.Visibility = Visibility.Collapsed
                GridMacros.Visibility = Visibility.Collapsed
                GridLoaded.Visibility = Visibility.Collapsed
                GridPDF.Visibility = Visibility.Visible
                GridAdmin.Visibility = Visibility.Collapsed
                GridStep.Visibility = Visibility.Collapsed
                GridContentCATPART.Visibility = Visibility.Collapsed


            Case 5 'STEP
                GridContentBDD.Visibility = Visibility.Collapsed
                GridChart.Visibility = Visibility.Collapsed
                GridBibliotheque.Visibility = Visibility.Collapsed
                GridAbout.Visibility = Visibility.Collapsed
                GridMacros.Visibility = Visibility.Collapsed
                GridLoaded.Visibility = Visibility.Collapsed
                GridPDF.Visibility = Visibility.Collapsed
                GridAdmin.Visibility = Visibility.Collapsed
                GridStep.Visibility = Visibility.Visible
                GridContentCATPART.Visibility = Visibility.Collapsed
            Case 6 'SEPARATOR

            Case 7 'SETTINGS
                GridContentBDD.Visibility = Visibility.Collapsed
                GridChart.Visibility = Visibility.Visible
                GridBibliotheque.Visibility = Visibility.Collapsed
                GridAbout.Visibility = Visibility.Collapsed
                GridMacros.Visibility = Visibility.Collapsed
                GridLoaded.Visibility = Visibility.Collapsed
                GridPDF.Visibility = Visibility.Collapsed
                GridStep.Visibility = Visibility.Collapsed
                GridAdmin.Visibility = Visibility.Collapsed
                GridContentCATPART.Visibility = Visibility.Collapsed
            Case 8 'INFO
                GridContentBDD.Visibility = Visibility.Collapsed
                GridChart.Visibility = Visibility.Collapsed
                GridBibliotheque.Visibility = Visibility.Collapsed
                GridAbout.Visibility = Visibility.Visible
                GridMacros.Visibility = Visibility.Collapsed
                GridStep.Visibility = Visibility.Collapsed
                GridLoaded.Visibility = Visibility.Collapsed
                GridPDF.Visibility = Visibility.Collapsed
                GridAdmin.Visibility = Visibility.Collapsed
                GridContentCATPART.Visibility = Visibility.Collapsed

            Case 9 'ADMIN

                GridContentBDD.Visibility = Visibility.Collapsed
                GridChart.Visibility = Visibility.Collapsed
                GridBibliotheque.Visibility = Visibility.Collapsed
                GridAbout.Visibility = Visibility.Collapsed
                GridMacros.Visibility = Visibility.Collapsed
                GridLoaded.Visibility = Visibility.Collapsed
                GridPDF.Visibility = Visibility.Collapsed
                GridStep.Visibility = Visibility.Collapsed
                GridAdmin.Visibility = Visibility.Visible
                GridContentCATPART.Visibility = Visibility.Collapsed
        End Select


    End Sub 'navigation

    Private Sub MenuButton_Click(sender As Object, e As RoutedEventArgs)

        If Me.GridaDeplacer.Width = New GridLength(35) Then
            Me.GridaDeplacer.Width = New GridLength(180)
            Me.Logo.IsEnabled = False
        Else
            Me.GridaDeplacer.Width = New GridLength(35)
            Me.Logo.IsEnabled = True
        End If

    End Sub


    Private Sub ButtonProfilsBosch_Click(sender As Object, e As RoutedEventArgs)
        FctionCATIA.CreatefromBibliotheque(DossierBase & "\Bibliothèque\PROFILÉS\BOSCH\_TEMPLATE-PROFIL BOSCH_.CATPart")
    End Sub

    Private Sub MenuItem_Click(sender As Object, e As RoutedEventArgs)
        Dim tvi As TreeViewItem = TVSelected
        If tvi Is Nothing Then
            Exit Sub
        Else
            Try
                Dim it As ItemTV = TryCast(tvi.DataContext, ItemTV)
                FctionCATIA.SelectCATIA(it.ItemCATIA)
            Catch ex As Exception
                MsgBox("Erreur lors de la recherche du fichier dans CATIA.", vbCritical)
            End Try

        End If
    End Sub

    Private Sub MonTV_SelectedItemChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object))
        Dim tvi As TreeViewItem = MonTV.SelectedItem
        If tvi Is Nothing Then
            Exit Sub
        Else
            Dim it As ItemTV = TryCast(tvi.DataContext, ItemTV)
            For Each ic_ As ItemCatia In MaDataGrid.Items
                Try
                    If it.ItemCATIA IsNot Nothing Then
                        If ic_.PartNumber = it.ItemCATIA.PartNumber Then
                            MaDataGrid.SelectedItem = ic_
                            MaDataGrid.ScrollIntoView(ic_)
                            Exit For
                        End If
                    End If
                Catch ex As Exception

                End Try
            Next
        End If


        For Each tv As TreeViewItem In MonTV.Items
            majForegroundTV(tv)
        Next




    End Sub

    Sub majForegroundTV(tv As TreeViewItem)


        Dim sp As StackPanel = tv.Header
        Dim im As Image = sp.Children.Item(0)
        Dim lb As Label = sp.Children.Item(1)

        If tv Is MonTV.SelectedItem Then
            lb.Foreground = New SolidColorBrush(Color.FromRgb(255, 255, 255))
        Else
            lb.Foreground = New SolidColorBrush(Color.FromRgb(0, 0, 0))
        End If

        For Each _tv As TreeViewItem In tv.Items
            majForegroundTV(_tv)
        Next

    End Sub
    Private Sub TreeViewItem_PreviewMouseRightButtonDown(sender As Object, e As MouseButtonEventArgs)

        Dim tvi As TreeViewItem = TryCast(e.Source.parent.parent, TreeViewItem)
        If tvi Is Nothing Then
        Else
            tvi.IsSelected = True
            TVSelected = tvi
        End If
    End Sub

    Private Sub MenuItem_Click_1(sender As Object, e As RoutedEventArgs)
        Dim tvi As TreeViewItem = TVSelected
        If tvi Is Nothing Then
            Exit Sub
        Else
            Try
                Dim it As ItemTV = TryCast(tvi.DataContext, ItemTV)
                FctionCATIA.OpenFile(it.ItemCATIA.FileName)
            Catch ex As Exception
                MsgBox("Impossible d'ouvrir le fichier demandé.", vbCritical)
            End Try

        End If
    End Sub

    Private Sub MenuItem_Click_2(sender As Object, e As RoutedEventArgs)
        Dim tvi As TreeViewItem = TVSelected
        If tvi Is Nothing Then
            Exit Sub
        Else
            Try
                Dim it As ItemTV = TryCast(tvi.DataContext, ItemTV)
                FctionCATIA.Createfrom(it.ItemCATIA)
            Catch ex As Exception
                Dim MsgErr As New MessageErreur("Une erreur s'est produite. Vérifier que le fichier soit sauvegardé", Notifications.Wpf.NotificationType.Error)
            End Try

        End If
    End Sub

    Private Sub MenuItem_Click_3(sender As Object, e As RoutedEventArgs)
        Dim tvi As TreeViewItem = TVSelected
        If tvi Is Nothing Then
            Exit Sub
        Else
            Try
                Dim it As ItemTV = TryCast(tvi.DataContext, ItemTV)
                FctionCATIA.CreerPlan(it.ItemCATIA)
            Catch ex As Exception
                Dim MsgErr As New MessageErreur("Une erreur s'est produite lors de la création du plan", Notifications.Wpf.NotificationType.Error)
            End Try

        End If
    End Sub

    Private Sub ButtonGrille_Click(sender As Object, e As RoutedEventArgs)
        Drill.CATMain()

    End Sub

    Private Sub ButtonLink_Click(sender As Object, e As RoutedEventArgs)
        GoRelinkDraw.RelinkDoc()
    End Sub

    Private Sub ButtonPDF2D_Click(sender As Object, e As RoutedEventArgs)
        WindowPDF.ShowDialog()
    End Sub

    Private Sub ButtonUnits_Click(sender As Object, e As RoutedEventArgs)
        WindowUnitDraw.ShowDialog()

    End Sub


    Private Sub ButtonCatiaTest_Click(sender As Object, e As RoutedEventArgs)

        Go()



    End Sub 'GO

    Sub Go()

        ButtonCatiaTest.Visibility = Visibility.Collapsed

        PercentNbElements = 1
        NbElements = 0

        ListDocuments.Clear()
        ListFabriques.Clear()
        ListAchetes.Clear()
        ListInconnus.Clear()
        ListItemCatia.Clear()
        ListItemTV.Clear()
        ListPartNumber.Clear()
        MonTV.Items.Clear()


        LoadProgressBar.Visibility = Visibility.Visible
        LabelLoad.Visibility = Visibility.Visible
        FctionLOAD.Main()


        MajVisuColonnesTableau()

    End Sub
    Sub MajVisuColonnesTableau()

        If GetEnv() = "SPIRIT AEROSYSTEMS" Then
            CRev.Visibility = Visibility.Visible
            CDetailNumber.Visibility = Visibility.Visible
            CStockSize.Visibility = Visibility.Visible
            cMaterial.Visibility = Visibility.Visible
            CBOM.Visibility = Visibility.Visible
            CSym.Visibility = Visibility.Hidden
            CTTS.Visibility = Visibility.Hidden
            CFour.Visibility = Visibility.Hidden
            CRef.Visibility = Visibility.Hidden
            CObservation.Visibility = Visibility.Hidden
            CAnglais.Visibility = Visibility.Hidden
            Perso1.Visibility = Visibility.Hidden
            Perso2.Visibility = Visibility.Hidden
            Perso3.Visibility = Visibility.Hidden
            Perso4.Visibility = Visibility.Hidden
            Perso5.Visibility = Visibility.Hidden
            Perso6.Visibility = Visibility.Hidden
            Perso7.Visibility = Visibility.Hidden
            Perso8.Visibility = Visibility.Hidden
            Perso9.Visibility = Visibility.Hidden
            Perso10.Visibility = Visibility.Hidden
            ButtonSYM.Visibility = Visibility.Collapsed
            ButtonOBS.Visibility = Visibility.Collapsed
            ButtonBOM.Visibility = Visibility.Collapsed
            SeparatorToHide.Visibility = Visibility.Visible
            cMaterial.Header = INIProperties.GetString(GetEnv, "ProprieteMATERIAL", "Matière")
        ElseIf GetEnv() = "AIRBUS" Then
            CRev.Visibility = Visibility.Hidden
            CBOM.Visibility = Visibility.Visible
            CDetailNumber.Visibility = Visibility.Hidden
            CStockSize.Visibility = Visibility.Hidden
            cMaterial.Visibility = Visibility.Visible
            cMaterial.Header = INIProperties.GetString(GetEnv, "ProprieteMATERIAL", "Matière")
            CTTS.Header = INIProperties.GetString(GetEnv, "ProprieteTTS", "TTS")
            CSym.Visibility = Visibility.Visible
            CTTS.Visibility = Visibility.Visible
            CFour.Visibility = Visibility.Visible
            CRef.Visibility = Visibility.Visible
            CObservation.Visibility = Visibility.Visible
            CAnglais.Visibility = Visibility.Hidden
            Perso1.Visibility = Visibility.Hidden
            Perso2.Visibility = Visibility.Hidden
            Perso3.Visibility = Visibility.Hidden
            Perso4.Visibility = Visibility.Hidden
            Perso5.Visibility = Visibility.Hidden
            Perso6.Visibility = Visibility.Hidden
            Perso7.Visibility = Visibility.Hidden
            Perso8.Visibility = Visibility.Hidden
            Perso9.Visibility = Visibility.Hidden
            Perso10.Visibility = Visibility.Hidden
            ButtonSYM.Visibility = Visibility.Visible
            ButtonOBS.Visibility = Visibility.Visible
            ButtonBOM.Visibility = Visibility.Visible
            SeparatorToHide.Visibility = Visibility.Visible
        ElseIf GetEnv() = "DASSAULT AVIATION" Then
            CBOM.Visibility = Visibility.Hidden
            Perso1.Visibility = Visibility.Hidden
            Perso2.Visibility = Visibility.Hidden
            Perso3.Visibility = Visibility.Hidden
            Perso4.Visibility = Visibility.Hidden
            Perso5.Visibility = Visibility.Hidden
            Perso6.Visibility = Visibility.Hidden
            Perso7.Visibility = Visibility.Hidden 'DESIGNATION
            Perso8.Visibility = Visibility.Hidden
            Perso9.Visibility = Visibility.Hidden
            Perso10.Visibility = Visibility.Hidden
            Dim i As Integer = 1
            For Each item As ItemProperties In DataGridProperties.Items
                Select Case i
                    Case 1 'MARQUAGE
                        Perso1.Visibility = Visibility.Hidden
                        Perso1.Header = item.Properties
                        nPerso1 = item.Properties
                        GetBindMatiere(Perso1, item.Properties)
                        GetBindTTS(Perso1, item.Properties)
                    Case 2 'TRAITEMENT
                        Perso2.Visibility = Visibility.Visible
                        Perso2.Header = item.Properties
                        nPerso2 = item.Properties
                        GetBindMatiere(Perso2, item.Properties)
                        GetBindTTS(Perso2, item.Properties)
                    Case 3 'PROTECTION
                        Perso3.Visibility = Visibility.Hidden
                        Perso3.Header = item.Properties
                        nPerso3 = item.Properties
                        GetBindMatiere(Perso3, item.Properties)
                        GetBindTTS(Perso3, item.Properties)
                    Case 4 'DIM BRUT
                        Perso4.Visibility = Visibility.Hidden
                        Perso4.Header = item.Properties
                        GetBindMatiere(Perso4, item.Properties)
                        GetBindTTS(Perso4, item.Properties)
                        nPerso4 = item.Properties
                    Case 5 'MASSE
                        Perso5.Visibility = Visibility.Hidden
                        Perso5.Header = item.Properties
                        GetBindMatiere(Perso5, item.Properties)
                        GetBindTTS(Perso5, item.Properties)
                        nPerso5 = item.Properties
                    Case 6 'MATIERE
                        Perso6.Visibility = Visibility.Visible
                        Perso6.Header = item.Properties
                        nPerso6 = item.Properties
                        GetBindMatiere(Perso6, item.Properties)
                        GetBindTTS(Perso6, item.Properties)
                    Case 7 'DESIGNATION
                        Perso7.Visibility = Visibility.Hidden
                        Perso7.Header = item.Properties
                        nPerso7 = item.Properties
                        GetBindMatiere(Perso7, item.Properties)
                        GetBindTTS(Perso7, item.Properties)
                    Case 8 'INDICE
                        Perso8.Visibility = Visibility.Hidden
                        Perso8.Header = item.Properties
                        nPerso8 = item.Properties
                        GetBindMatiere(Perso8, item.Properties)
                        GetBindTTS(Perso8, item.Properties)
                    Case 9 'PLANCHE
                        Perso9.Visibility = Visibility.Visible
                        Perso9.Header = item.Properties
                        nPerso9 = item.Properties
                        GetBindMatiere(Perso9, item.Properties)
                        GetBindTTS(Perso9, item.Properties)
                    Case 10
                        Perso10.Visibility = Visibility.Visible
                        Perso10.Header = item.Properties
                        nPerso10 = item.Properties
                        GetBindMatiere(Perso10, item.Properties)
                        GetBindTTS(Perso10, item.Properties)
                End Select
                i = i + 1
            Next
            CSym.Visibility = Visibility.Hidden
            cMaterial.Visibility = Visibility.Hidden
            CTTS.Visibility = Visibility.Hidden
            CFour.Visibility = Visibility.Hidden
            CRef.Visibility = Visibility.Hidden
            CObservation.Visibility = Visibility.Hidden
            CAnglais.Visibility = Visibility.Hidden
            CRev.Visibility = Visibility.Visible
            CDetailNumber.Visibility = Visibility.Hidden
            CStockSize.Visibility = Visibility.Hidden
            ButtonSYM.Visibility = Visibility.Collapsed
            ButtonOBS.Visibility = Visibility.Collapsed
            ButtonBOM.Visibility = Visibility.Collapsed
            SeparatorToHide.Visibility = Visibility.Collapsed
        Else
            CBOM.Visibility = Visibility.Visible
            Perso1.Visibility = Visibility.Hidden
            Perso2.Visibility = Visibility.Hidden
            Perso3.Visibility = Visibility.Hidden
            Perso4.Visibility = Visibility.Hidden
            Perso5.Visibility = Visibility.Hidden
            Perso6.Visibility = Visibility.Hidden
            Perso7.Visibility = Visibility.Hidden
            Perso8.Visibility = Visibility.Hidden
            Perso9.Visibility = Visibility.Hidden
            Perso10.Visibility = Visibility.Hidden
            Dim i As Integer = 1
            For Each item As ItemProperties In DataGridProperties.Items
                Select Case i
                    Case 1
                        Perso1.Visibility = Visibility.Visible
                        Perso1.Header = item.Properties
                        nPerso1 = item.Properties
                        GetBindMatiere(Perso1, item.Properties)
                        GetBindTTS(Perso1, item.Properties)
                    Case 2
                        Perso2.Visibility = Visibility.Visible
                        Perso2.Header = item.Properties
                        nPerso2 = item.Properties
                        GetBindMatiere(Perso2, item.Properties)
                        GetBindTTS(Perso2, item.Properties)
                    Case 3
                        Perso3.Visibility = Visibility.Visible
                        Perso3.Header = item.Properties
                        nPerso3 = item.Properties
                        GetBindMatiere(Perso3, item.Properties)
                        GetBindTTS(Perso3, item.Properties)
                    Case 4
                        Perso4.Visibility = Visibility.Visible
                        Perso4.Header = item.Properties
                        GetBindMatiere(Perso4, item.Properties)
                        GetBindTTS(Perso4, item.Properties)
                        nPerso4 = item.Properties
                    Case 5
                        Perso5.Visibility = Visibility.Visible
                        Perso5.Header = item.Properties
                        GetBindMatiere(Perso5, item.Properties)
                        GetBindTTS(Perso5, item.Properties)
                        nPerso5 = item.Properties
                    Case 6
                        Perso6.Visibility = Visibility.Visible
                        Perso6.Header = item.Properties
                        nPerso6 = item.Properties
                        GetBindMatiere(Perso6, item.Properties)
                        GetBindTTS(Perso6, item.Properties)
                    Case 7
                        Perso7.Visibility = Visibility.Visible
                        Perso7.Header = item.Properties
                        nPerso7 = item.Properties
                        GetBindMatiere(Perso7, item.Properties)
                        GetBindTTS(Perso7, item.Properties)
                    Case 8
                        Perso8.Visibility = Visibility.Visible
                        Perso8.Header = item.Properties
                        nPerso8 = item.Properties
                        GetBindMatiere(Perso8, item.Properties)
                        GetBindTTS(Perso8, item.Properties)
                    Case 9
                        Perso9.Visibility = Visibility.Visible
                        Perso9.Header = item.Properties
                        nPerso9 = item.Properties
                        GetBindMatiere(Perso9, item.Properties)
                        GetBindTTS(Perso9, item.Properties)
                    Case 10
                        Perso10.Visibility = Visibility.Visible
                        Perso10.Header = item.Properties
                        nPerso10 = item.Properties
                        GetBindMatiere(Perso10, item.Properties)
                        GetBindTTS(Perso10, item.Properties)
                End Select
                i = i + 1
            Next
            CSym.Visibility = Visibility.Hidden
            cMaterial.Visibility = Visibility.Hidden
            CTTS.Visibility = Visibility.Hidden
            CFour.Visibility = Visibility.Hidden
            CRef.Visibility = Visibility.Hidden
            CObservation.Visibility = Visibility.Hidden
            CAnglais.Visibility = Visibility.Hidden
            CRev.Visibility = Visibility.Hidden
            CDetailNumber.Visibility = Visibility.Hidden
            CStockSize.Visibility = Visibility.Hidden
            ButtonSYM.Visibility = Visibility.Collapsed
            ButtonOBS.Visibility = Visibility.Collapsed
            ButtonBOM.Visibility = Visibility.Collapsed
            SeparatorToHide.Visibility = Visibility.Collapsed
        End If

    End Sub 'MAJ SPIRIT

    Sub GetBindMatiere(c As DataGridTextColumn, s As String)
        If s = INIProperties.GetString(GetEnv, "ProprieteMATERIAL", "") Then
            c.Binding = New Binding("Matiere")
        End If
    End Sub
    Sub GetBindTTS(c As DataGridTextColumn, s As String)
        If s = INIProperties.GetString(GetEnv, "ProprieteTTS", "") Then
            c.Binding = New Binding("Traitement")
        End If
    End Sub
    Private Sub MaDataGrid_CellEditEnding(sender As Object, e As DataGridCellEditEndingEventArgs)
        Dim ic As ItemCatia = e.EditingElement.DataContext
        '   FctionCATIA.LinkICtoTV()
        If ic.Type = "COMPOSANT" Then
        Else

            Dim macolonne As String = e.Column.Header

            FctionCATIA.AddParamatres(ic.Owner, ic)

            Select Case macolonne
                Case Perso1.Header
                    Dim t As TextBox = e.EditingElement
                    ic.Perso1 = t.Text
                    Dim p As Parameters = ic.ProductCATIA.UserRefProperties
                    Try
                        p.Item(Perso1.Header).value = ic.Perso1
                    Catch ex As Exception
                    End Try
                Case Perso2.Header
                    Dim t As TextBox = e.EditingElement
                    ic.Perso2 = t.Text
                    Dim p As Parameters = ic.ProductCATIA.UserRefProperties
                    Try
                        p.Item(Perso2.Header).value = ic.Perso2
                    Catch ex As Exception
                    End Try
                Case Perso3.Header
                    Dim t As TextBox = e.EditingElement
                    ic.Perso3 = t.Text
                    Dim p As Parameters = ic.ProductCATIA.UserRefProperties
                    Try
                        p.Item(Perso3.Header).value = ic.Perso3
                    Catch ex As Exception
                    End Try
                Case Perso4.Header
                    Dim t As TextBox = e.EditingElement
                    ic.Perso4 = t.Text
                    Dim p As Parameters = ic.ProductCATIA.UserRefProperties
                    Try
                        p.Item(Perso4.Header).value = ic.Perso4
                    Catch ex As Exception
                    End Try
                Case Perso5.Header
                    Dim t As TextBox = e.EditingElement
                    ic.Perso5 = t.Text
                    Dim p As Parameters = ic.ProductCATIA.UserRefProperties
                    Try
                        p.Item(Perso5.Header).value = ic.Perso5
                    Catch ex As Exception
                    End Try
                Case Perso6.Header
                    Dim t As TextBox = e.EditingElement
                    ic.Perso6 = t.Text
                    Dim p As Parameters = ic.ProductCATIA.UserRefProperties
                    Try
                        p.Item(Perso6.Header).value = ic.Perso6
                    Catch ex As Exception
                    End Try
                Case Perso7.Header
                    Dim t As TextBox = e.EditingElement
                    ic.Perso7 = t.Text
                    Dim p As Parameters = ic.ProductCATIA.UserRefProperties
                    Try
                        p.Item(Perso7.Header).value = ic.Perso7
                    Catch ex As Exception
                    End Try
                Case Perso8.Header
                    Dim t As TextBox = e.EditingElement
                    ic.Perso8 = t.Text
                    Dim p As Parameters = ic.ProductCATIA.UserRefProperties
                    Try
                        p.Item(Perso8.Header).value = ic.Perso8
                    Catch ex As Exception
                    End Try
                Case Perso9.Header
                    Dim t As TextBox = e.EditingElement
                    ic.Perso9 = t.Text
                    Dim p As Parameters = ic.ProductCATIA.UserRefProperties
                    Try
                        p.Item(Perso9.Header).value = ic.Perso9
                    Catch ex As Exception
                    End Try

                Case Perso10.Header
                    Dim t As TextBox = e.EditingElement
                    ic.Perso10 = t.Text
                    Dim p As Parameters = ic.ProductCATIA.UserRefProperties
                    Try
                        p.Item(Perso10.Header).value = ic.Perso10
                    Catch ex As Exception
                    End Try

                Case "Detail Number"
                    Dim t As TextBox = e.EditingElement
                    ic.DetailNumber = t.Text
                    Dim p As Parameters = ic.ProductCATIA.UserRefProperties
                    Try
                        p.Item("DETAIL NUMBER").value = ic.DetailNumber
                    Catch ex As Exception
                    End Try
                Case "Stock Size"
                    Dim t As TextBox = e.EditingElement
                    ic.StockSize = t.Text
                    Dim p As Parameters = ic.ProductCATIA.UserRefProperties
                    Try
                        p.Item("STOCK SIZE").value = ic.StockSize
                    Catch ex As Exception
                    End Try
                Case "Nomenclature"
                    Dim t As TextBox = e.EditingElement
                    ic.Nomenclature = t.Text
                    ic.ProductCATIA.Nomenclature = ic.Nomenclature
                Case "Désignation"
                    Dim t As TextBox = e.EditingElement
                    ic.DescriptionRef = t.Text
                    ic.Anglais = ic.Anglais
                    If ic.Anglais <> "" Then
                        ic.ProductCATIA.DescriptionRef = ic.DescriptionRef & vbNewLine & ic.Anglais
                    Else
                        ic.ProductCATIA.DescriptionRef = ic.DescriptionRef
                    End If
                    If Env = "[DASSAULT AVIATION]" Then
                        ic.Perso7 = ic.DescriptionRef
                        Dim p As Parameters = ic.ProductCATIA.UserRefProperties
                        Try
                            p.Item(Perso7.Header).value = ic.Perso7
                        Catch ex As Exception
                        End Try
                        ' FctionCATIA.SaveIC(ic)
                    End If
                Case "PartNumber"
                    'MA55800Z00-PIECE_500006
                    Dim t As TextBox = e.EditingElement
                    If ListPartNumber.Contains(t.Text) And ic.PartNumber.Length > 0 Then
                        t.Text = ic.PartNumber
                    Else
                        ic.PartNumber = t.Text
                        ic.ProductCATIA.PartNumber = ic.PartNumber
                        FctionCATIA.SaveIC(ic)
                        ListPartNumber.Add(ic.PartNumber)
                        If Env = "[DASSAULT AVIATION]" Then
                            If ic.PartNumber Like "??#####?##-*_######" Then
                                Dim ss() As String = Strings.Split(ic.PartNumber, "_")
                                Dim p As Parameters = ic.ProductCATIA.UserRefProperties
                                Try
                                    p.Item(Perso9.Header).value = ss(1)
                                Catch ex As Exception
                                End Try
                                ic.Perso9 = ss(1)
                                BoolHaveTORefresh = True
                            End If
                        End If
                    End If

                Case "Rev"
                    Dim t As TextBox = e.EditingElement
                    ic.Indice = t.Text
                    ic.Revision = ic.Indice
                    ic.ProductCATIA.Revision = ic.Indice
                    Dim p As Parameters = ic.ProductCATIA.UserRefProperties
                    Try
                        p.Item("ISSUE").Value = ic.Indice
                    Catch ex As Exception
                    End Try
                    If Env = "[DASSAULT AVIATION]" Then
                        ic.Perso8 = ic.Indice
                        p.Item("NomPuls_Indice").Value = ic.Indice
                    End If

                Case "Anglais"
                    Dim t As TextBox = e.EditingElement
                    ic.Anglais = t.Text
                    ic.DescriptionRef = ic.DescriptionRef
                    Dim str As String = ic.DescriptionRef & vbNewLine & ic.Anglais
                    If ic.Anglais <> "" Then
                        ic.ProductCATIA.DescriptionRef = str
                    Else
                        ic.ProductCATIA.DescriptionRef = ic.DescriptionRef
                    End If
                Case "Part Symétrique"
                    Dim t As TextBox = e.EditingElement
                    ic.SymPart = t.Text
                    Dim p As Parameters = ic.ProductCATIA.UserRefProperties
                    Try
                        p.Item("SYM").value = ic.SymPart
                    Catch ex As Exception
                    End Try
                Case "Traitement/Aspect"
                    Dim t As TextBox = e.EditingElement
                    ic.Traitement = t.Text
                    Dim p As Parameters = ic.ProductCATIA.UserRefProperties
                    Try
                        p.Item(INIProperties.GetString(GetEnv, "ProprieteTTS", "TTS")).value = ic.Traitement
                    Catch ex As Exception
                    End Try
                Case "Fournisseur"
                    Dim t As TextBox = e.EditingElement
                    ic.Fournisseur = t.Text
                    Dim p As Parameters = ic.ProductCATIA.UserRefProperties
                    Try
                        p.Item("SUPPLIER").value = ic.Fournisseur
                    Catch ex As Exception
                    End Try
                Case "Référence"
                    Dim t As TextBox = e.EditingElement
                    ic.Reference = t.Text
                    Dim p As Parameters = ic.ProductCATIA.UserRefProperties
                    Try
                        p.Item("REF").value = ic.Reference
                    Catch ex As Exception
                    End Try

                Case "Observations"
                    Dim t As TextBox = e.EditingElement
                    ic.Observation = t.Text
                    Dim p As Parameters = ic.ProductCATIA.UserRefProperties
                    Try
                        p.Item("OBSERVATIONS").value = ic.Observation
                    Catch ex As Exception
                    End Try

            End Select

        End If

        For Each tv As TreeViewItem In MonTV.Items
            majHederTV(tv)
        Next

        MaDataGrid.IsReadOnly = True
    End Sub



    Sub majHederTV(tv As TreeViewItem)


        Dim sp As StackPanel = tv.Header
        Dim im As Image = sp.Children.Item(0)
        Dim lb As Label = sp.Children.Item(1)
        Dim itemtv__ As ItemTV = tv.DataContext
        Dim ic As ItemCatia = itemtv__.ItemCATIA

        If Not ic Is Nothing Then
            If ic.DescriptionRef Is Nothing Then
                lb.Content = ic.PartNumber.ToString & " | "
            Else
                lb.Content = ic.PartNumber.ToString & " | " & ic.DescriptionRef.ToString
            End If
        End If

        For Each _tv As TreeViewItem In tv.Items
            majHederTV(_tv)
        Next

    End Sub


    Private Sub SearchBarDataTable_TextChanged(sender As Object, e As TextChangedEventArgs)

        Try
            MaDataGrid.CancelEdit()
            ColDoc.Filter = New Predicate(Of Object)(AddressOf FilterList)
            ColDoc.Refresh()
        Catch ex As Exception
        End Try

    End Sub

    Private Sub ExportDrawing_Click(sender As Object, e As RoutedEventArgs)

        Dim l As New List(Of ItemCatia)
        For Each ic As ItemCatia In ColDoc
            l.Add(ic)
        Next

        If ComboBOM.SelectedIndex = 0 Then
            FctionCATIA.CreerNomenclature2D(l, True)
        Else
            FctionCATIA.CreerNomenclature2D(l, False)
        End If

    End Sub
    Function FilterList(item As ItemCatia) As Boolean

        Try
            Dim strIC() = Strings.Split(ICRacine.Doc.Name, ".CATProduct")
            Dim strIC_ As String = strIC(0)

            If ComboBOM.SelectedValue.ToString = "Ensemble des éléments" Then
                For Each ic In ListDocuments
                    If ic.Type = "PRODUCT" Then
                        ic.Qte = ""
                    End If
                    ic.Visible = True
                Next
            End If

            Try
                Dim value As ItemCatia = item
                If (Not value Is Nothing) And (Not value Is DBNull.Value) Then
                    Dim MonString As String = UCase(SearchBarDataTable.Text)
                    If (UCase(value.PartNumber).Contains(MonString) Or UCase(value.DescriptionRef).Contains(MonString) _
                      Or UCase(value.Nomenclature).Contains(MonString)) And value.Visible = True Then
                        Return True
                    Else
                        Return False
                    End If
                Else
                    Return False
                End If
            Catch ex As Exception
            End Try

            Return False

        Catch ex As Exception
            Return False
        End Try

    End Function


    Dim TVSelected As TreeViewItem = Nothing

    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)

        If GridContentBDD.Visibility = Visibility.Visible Then
            If MsgBox("Vous êtes sur le point de supprimer l'ensemble des propriétés de chacun des éléments de l'assemblage. Continuer ?", vbInformation + vbYesNo) = MsgBoxResult.No Then
                Exit Sub

            Else
                ProgressTableau.Visibility = Visibility.Visible
                GoProgress("Properties")
                BadgeReload.Badge = "!"
                NeedReload = True
            End If
        ElseIf GridContentCATPART.Visibility = Visibility.Visible Then
            If MsgBox("Vous êtes sur le point de supprimer l'ensemble des propriétés de la Part. Continuer ?", vbInformation + vbYesNo) = MsgBoxResult.No Then
                Exit Sub

            Else
                ProgressTableau.Visibility = Visibility.Visible
                GoProgress("Properties")
                ListPropertiesPart.Clear()
                FctionCATIA.GoPart()
                DataGridPropertiesPart.Items.Refresh()

            End If
        End If

    End Sub

    Private Sub Button_Click_1(sender As Object, e As RoutedEventArgs)

        WindowRenameDassault_.ShowDialog()

    End Sub

    Private Sub Button_Click_2(sender As Object, e As RoutedEventArgs)


        If MsgBox("Vous êtes sur le de renommer l'ensemble des instances. Cela risque de prendre quelques minutes. Continuer ?", vbInformation + vbYesNo) = MsgBoxResult.No Then
            Exit Sub

        Else
            ProgressTableau.Visibility = Visibility.Visible
            GoProgress("Instance")

        End If

    End Sub

    Private Sub Button_Click_3(sender As Object, e As RoutedEventArgs)
        Dim c As String = INIProperties.GetString(GetEnv, "BibliothequeMaterial", "")
        c = GetAs(c)
        Process.Start(c)
    End Sub

    Private Sub Button_Click_4(sender As Object, e As RoutedEventArgs)
        Try
            Process.Start(TreeSauv)
        Catch ex As Exception
        End Try

    End Sub

    Private Sub Button_Click_5(sender As Object, e As RoutedEventArgs)
        Try
            MaDataGrid.SelectAllCells()
            MaDataGrid.ClipboardCopyMode = DataGridClipboardCopyMode.IncludeHeader
            ApplicationCommands.Copy.Execute(Nothing, MaDataGrid)

            Dim Result As String = Clipboard.GetData(DataFormats.CommaSeparatedValue)
            Dim Res As String = Clipboard.GetData(DataFormats.Text)

            Dim fs As FileStream = New FileStream(CATIA.ActiveDocument.Path & "\data.xls", FileMode.Create)
            Using sw As StreamWriter = New StreamWriter(fs, Text.Encoding.GetEncoding("iso-8859-1"))
                sw.WriteLine(Res.Replace(",", " "))
                sw.Close()
            End Using


            MaDataGrid.UnselectAllCells()
            Process.Start(CATIA.ActiveDocument.Path & "\data.xls")

        Catch ex As Exception
            Dim nms As New MessageErreur("Impossible de générer de fichier EXCEL à partir d'un tableau vide", Notifications.Wpf.NotificationType.Error)
        End Try
    End Sub



    Dim ListICOKSym As New List(Of ItemCatia)
    Public Function IsPair(i As Integer) As Boolean
        IsPair = (i Mod 2) = 0
    End Function
    Private Sub Button_Click_6(sender As Object, e As RoutedEventArgs)

        ListICOKSym.Clear()
        If Env = "[AIRBUS]" Then
            For Each ic As ItemCatia In ListDocuments
                If Not ListICOKSym.Contains(ic) Then
                    If ic.PartNumber.Length = 18 Then
                        If ic.Type = "PART" Then
                            Dim str1 As String = Right(ic.PartNumber, 4)
                            Dim int1 As Integer = 0
                            Try
                                int1 = str1
                            Catch ex As Exception
                                int1 = 0
                            End Try
                            For Each ic_ As ItemCatia In ListDocuments
                                If Not ListICOKSym.Contains(ic_) And ic.Type = ic_.Type Then
                                    Dim str2 As String = Right(ic_.PartNumber, 4)
                                    Dim int2 As Integer = 0
                                    Try
                                        int2 = str2
                                    Catch ex As Exception
                                        int2 = 0
                                    End Try
                                    If IsPair(int1) = True Then
                                        If int1 = int2 - 1 Then
                                            ic.SymPart = ic_.PartNumber
                                            ic_.SymPart = ic.PartNumber
                                            ListICOKSym.Add(ic)
                                            ListICOKSym.Add(ic_)
                                            Try
                                                ic.ProductCATIA.UserRefProperties.Item("SYM").ValuateFromString(ic.SymPart)
                                            Catch ex As Exception
                                                FctionCATIA.AddParamatres(ic.Owner, ic)
                                                ic.ProductCATIA.UserRefProperties.Item("SYM").ValuateFromString(ic.SymPart)
                                            End Try
                                            Try
                                                ic_.ProductCATIA.UserRefProperties.Item("SYM").ValuateFromString(ic_.SymPart)
                                            Catch ex As Exception
                                                FctionCATIA.AddParamatres(ic_.Owner, ic_)
                                                ic_.ProductCATIA.UserRefProperties.Item("SYM").ValuateFromString(ic_.SymPart)
                                            End Try
                                        End If
                                    End If
                                End If
                            Next
                        End If
                        If ic.Type = "PRODUCT" Then
                            Dim MaLettre1 As String = ""
                            Dim str1 As String = Right(ic.PartNumber, 4)
                            MaLettre1 = Strings.Left(str1, 1)
                            str1 = Right(str1, 3)
                            Dim int1 As Integer = 0
                            Try
                                int1 = str1
                            Catch ex As Exception
                                int1 = 0
                            End Try
                            For Each ic_ As ItemCatia In ListDocuments
                                If Not ListICOKSym.Contains(ic_) And ic.Type = ic_.Type Then
                                    Dim MaLettre2 As String = ""
                                    Dim str2 As String = Right(ic_.PartNumber, 4)
                                    MaLettre2 = Strings.Left(str2, 1)
                                    str2 = Right(str2, 3)
                                    Dim int2 As Integer = 0
                                    Try
                                        int2 = str2
                                    Catch ex As Exception
                                        int2 = 0
                                    End Try
                                    If IsPair(int1) = True Then
                                        If int1 = int2 - 1 And MaLettre1 = MaLettre2 Then
                                            ic.SymPart = ic_.PartNumber
                                            ic_.SymPart = ic.PartNumber
                                            ListICOKSym.Add(ic)
                                            ListICOKSym.Add(ic_)
                                            Try
                                                ic.ProductCATIA.UserRefProperties.Item("SYM").ValuateFromString(ic.SymPart)
                                            Catch ex As Exception
                                                FctionCATIA.AddParamatres(ic.Owner, ic)
                                                ic.ProductCATIA.UserRefProperties.Item("SYM").ValuateFromString(ic.SymPart)
                                            End Try
                                            Try
                                                ic_.ProductCATIA.UserRefProperties.Item("SYM").ValuateFromString(ic_.SymPart)
                                            Catch ex As Exception
                                                FctionCATIA.AddParamatres(ic_.Owner, ic_)
                                                ic_.ProductCATIA.UserRefProperties.Item("SYM").ValuateFromString(ic_.SymPart)
                                            End Try
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
            Next
            Me.MaDataGrid.Items.Refresh()
        Else
            Dim m As New MessageErreur("Fonction inutile dans un environnement autre que AIRBUS", Notifications.Wpf.NotificationType.Warning)
        End If
    End Sub

    Private Sub Button_Click_7(sender As Object, e As RoutedEventArgs)
        If MsgBox("Vous êtes sur le point de fixer l'assemblage ouvert. Quelques minutes peuvent être nécéssaires. Continuer ?", vbInformation + vbYesNo) = MsgBoxResult.No Then
            Exit Sub

        Else

            ProgressTableau.Visibility = Visibility.Visible
            GoProgress("Fix")

        End If


    End Sub

    Private Sub Button_Click_8(sender As Object, e As RoutedEventArgs)


        If Env = "[AIRBUS]" Then
            For Each ic As ItemCatia In ListDocuments
                If ic.Reference <> "" Then
                    ic.Observation = ic.Fournisseur & " '" & ic.Reference
                ElseIf ic.Traitement <> "" Then
                    ic.Observation = ic.Traitement
                End If

                Try
                    ic.ProductCATIA.UserRefProperties.Item("OBSERVATIONS").ValuateFromString(ic.Observation)
                Catch ex As Exception
                    FctionCATIA.AddParamatres(ic.Owner, ic)
                    ic.ProductCATIA.UserRefProperties.Item("OBSERVATIONS").ValuateFromString(ic.Observation)
                End Try

            Next
        Else
            Dim m As New MessageErreur("Fonction inutile pour un sous-ensemble en dehors de l'envrionnement AIRBUS", Notifications.Wpf.NotificationType.Warning)
        End If


        MaDataGrid.Items.Refresh()
    End Sub

    Private Sub Button_Click_9(sender As Object, e As RoutedEventArgs)


        If GridContentBDD.Visibility = Visibility.Visible Then
            If MsgBox("Vous êtes sur le point de calculer la masse de l'ensemble des pièces. Quelques minutes peuvent être nécéssaires. Continuer ?", vbInformation + vbYesNo) = MsgBoxResult.No Then
                Exit Sub
            Else
                ProgressTableau.Visibility = Visibility.Visible
                GoProgress("Masse")
                BadgeReloadMass.Badge = ""

            End If
        ElseIf GridContentCATPART.Visibility = Visibility.Visible Then

            ProgressTableau.Visibility = Visibility.Visible
            Try
                GoProgress("Masse")
            Catch ex As Exception
            End Try

            ListPropertiesPart.Clear()
            FctionCATIA.GoPart()
            DataGridPropertiesPart.Items.Refresh()
        End If





    End Sub

    Private Sub Button_Click_10(sender As Object, e As RoutedEventArgs)

        Dim Nber As String = Nothing
        If Env = "[AIRBUS]" Then
            For Each ic As ItemCatia In ListDocuments
                If ic.PartNumber Like "??????????-????????" Then '"T000561017-98982P01" Then
                    Nber = Right(ic.PartNumber, 8)
                    Nber = Strings.Left(Nber, 6)
                Else
                    If ic.PartNumber.Length = 18 Then
                        Nber = Right(ic.PartNumber, 4)
                    End If
                End If
                If Nber <> Nothing Then
                    ic.Nomenclature = Nber
                    ic.ProductCATIA.Nomenclature = Nber
                End If
                Nber = Nothing
            Next
        ElseIf Env = "[SPIRIT AEROSYSTEMS]" Then

            For Each ic As ItemCatia In ListDocuments
                If ic.PartNumber Like "*_*_*" Then '8FME15V09003_411_-
                    Dim a() As String = Strings.Split(ic.PartNumber, "_")
                    Nber = a(1)
                    Nber = Replace(Nber, "_", "")
                    ic.DetailNumber = Nber
                    Try
                        ic.ProductCATIA.UserRefProperties.Item("DETAIL NUMBER").ValuateFromString(ic.DetailNumber)
                    Catch ex As Exception
                        FctionCATIA.AddParamatres(ic.Owner, ic)
                        ic.ProductCATIA.UserRefProperties.Item("DETAIL NUMBER").ValuateFromString(ic.DetailNumber)
                    End Try

                End If
            Next

        Else
            Dim m As New MessageErreur("Inutile en dehors des environnements AIRBUS et SPIRIT", Notifications.Wpf.NotificationType.Error)
        End If

        MaDataGrid.Items.Refresh()

    End Sub

    Private Sub ComboBOM_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        Try
            If ComboBOM.Items.Count > 0 Then MajBOm(Nothing)
        Catch ex As Exception
        End Try
    End Sub

    Sub MajBOm(s As String)
        If s Is Nothing Then s = ComboBOM.SelectedValue.ToString
        FctionCATIA.GoNomenclature(s)
        Try
            ColDoc.Filter = New Predicate(Of Object)(AddressOf FilterList)
        Catch ex As Exception
        End Try

        ColDoc.Refresh()
    End Sub

    Sub ButtonListMenu(sender As Object, e As RoutedEventArgs)

        Dim t As String = sender.content

        For Each ic As ItemCatia In Me.MaDataGrid.SelectedItems
            If ic.Type = "PART" Then
                If t = "Aucun" Then t = ""
                ic.Traitement = t
                Dim p As Parameters = ic.ProductCATIA.UserRefProperties
                Try
                    p.Item(INIProperties.GetString(GetEnv, "ProprieteTTS", "TTS")).Value = ic.Traitement
                Catch ex As Exception
                    FctionCATIA.AddParamatres(ic.Owner, ic)
                    p.Item(INIProperties.GetString(GetEnv, "ProprieteTTS", "TTS")).Value = ic.Traitement
                End Try
            End If
        Next

        Me.MaDataGrid.Items.Refresh()
        Me.MonTV.Items.Refresh()
        PopupTraitements.IsOpen = False

    End Sub
    Sub ButtonListMaterial_click(sender As Object, e As RoutedEventArgs)

        Dim t As String = sender.content

        For Each ic As ItemCatia In Me.MaDataGrid.SelectedItems
            If ic.Type = "PART" Then
                ic.Matiere = t
                FctionCATIA.AddParamatres(ic.Owner, ic)
                FctionCATIA.AppliqueMaterial(t, ic.Owner, ic)
            End If
        Next

        Me.MaDataGrid.Items.Refresh()
        Me.MonTV.Items.Refresh()
        PopupMaterials.IsOpen = False

        If MaDataGrid.Items.Count > 0 Then BadgeReloadMass.Badge = "!"


    End Sub

    Private Sub AcheteButton_Click(sender As Object, e As RoutedEventArgs)

        For Each ic As ItemCatia In MaDataGrid.SelectedItems
            ic.Source = "Acheté"
            ic.ProductCATIA.Source = CatProductSource.catProductBought
        Next

        Me.MaDataGrid.Items.Refresh()
        Me.MonTV.Items.Refresh()

        PopupSource.IsOpen = False
    End Sub
    Private Sub FabriqueButton_Click(sender As Object, e As RoutedEventArgs)

        For Each ic As ItemCatia In MaDataGrid.SelectedItems

            ic.Source = "Fabriqué"
            ic.ProductCATIA.Source = CatProductSource.catProductMade
        Next
        Me.MaDataGrid.Items.Refresh()
        Me.MonTV.Items.Refresh()
        PopupSource.IsOpen = False
    End Sub
    Private Sub InconnuButton_Click(sender As Object, e As RoutedEventArgs)
        For Each ic As ItemCatia In MaDataGrid.SelectedItems

            ic.Source = "Inconnu"
            ic.ProductCATIA.Source = CatProductSource.catProductSourceUnknown
        Next

        Me.MaDataGrid.Items.Refresh()
        Me.MonTV.Items.Refresh()
        PopupSource.IsOpen = False
    End Sub

    Sub EventCellClick(sender As Object, e As MouseButtonEventArgs)

        PopupSource.IsOpen = False
        PopupMaterials.IsOpen = False
        PopupTraitements.IsOpen = False

        Dim ic As ItemCatia = Me.MaDataGrid.SelectedItem
        Dim c As DataGridCell = sender

        If Not ic Is Nothing Then
            If ic.Type = "PRODUCT" Or ic.Type = "RACINE" Or ic.Type = "PART" Then
                If c.Column.Header = "Source" Then
                    PopupSource.PlacementTarget = c
                    PopupSource.IsOpen = True
                End If
            End If

            If ic.Type = "PART" Then
                If c.Column.Header = INIProperties.GetString(GetEnv, "ProprieteMATERIAL", "Matière") Then
                    PopupMaterials.PlacementTarget = c
                    PopupMaterials.PopupAnimation = Primitives.PopupAnimation.Scroll
                    PopupMaterials.IsOpen = True
                End If
                If c.Column.Header = INIProperties.GetString(GetEnv, "ProprieteTTS", "TTS") Then
                    PopupTraitements.PlacementTarget = c
                    PopupTraitements.PopupAnimation = Primitives.PopupAnimation.Scroll
                    PopupTraitements.IsOpen = True
                End If
            End If
        End If




    End Sub

    Private Sub ContextMenu_Opened(sender As Object, e As RoutedEventArgs)

        Dim ic As ItemCatia = Me.MaDataGrid.SelectedItem
        If ic Is Nothing Then
        Else
            If ic.Type = "PRODUCT" Or ic.Type = "RACINE" Then
                ButtonNomenclature.IsEnabled = True
            Else
                ButtonNomenclature.IsEnabled = False
            End If
        End If
    End Sub

    Private Sub MenuItem_Click_4(sender As Object, e As RoutedEventArgs)
        Dim ic As ItemCatia = Me.MaDataGrid.SelectedItem
        If ic Is Nothing Then
            Exit Sub
        Else
            Try
                FctionCATIA.SelectCATIA(ic)
            Catch ex As Exception
                Dim MsgErr As New MessageErreur("Erreur lors de la recherche du fichier dans CATIA", Notifications.Wpf.NotificationType.Error)
            End Try

        End If
    End Sub

    Private Sub MenuItem_Click_5(sender As Object, e As RoutedEventArgs)
        Dim ic As ItemCatia = Me.MaDataGrid.SelectedItem
        If ic Is Nothing Then
            Exit Sub
        Else
            Try

                FctionCATIA.OpenFile(ic.FileName)
            Catch ex As Exception
                Dim MsgErr As New MessageErreur("Impossible d'ouvrir le fichier demandé.", Notifications.Wpf.NotificationType.Error)
            End Try

        End If
    End Sub

    Private Sub MenuItem_Click_6(sender As Object, e As RoutedEventArgs)
        Dim ic As ItemCatia = Me.MaDataGrid.SelectedItem
        If ic Is Nothing Then
            Exit Sub
        Else
            Try
                FctionCATIA.Createfrom(ic)
            Catch ex As Exception
                Dim MsgErr As New MessageErreur("Erreur lors de la recherche du fichier dans CATIA", Notifications.Wpf.NotificationType.Error)
            End Try

        End If
    End Sub

    Private Sub MenuItem_Click_7(sender As Object, e As RoutedEventArgs)
        Dim ic As ItemCatia = Me.MaDataGrid.SelectedItem

        If ic Is Nothing Then
            Exit Sub
        Else
            FctionCATIA.CreerPlan(ic)
        End If
    End Sub

    Private Sub MenuItem_Click_8(sender As Object, e As RoutedEventArgs)
        Dim ic As ItemCatia = Me.MaDataGrid.SelectedItem
        If ic Is Nothing Then
            Exit Sub
        Else
            Dim str() As String = Strings.Split(ic.Doc.Name, ".CATProduct")
            Dim str_ As String = str(0)
            ComboBOM.SelectedValue = str_
            MajBOm(Nothing)
        End If
    End Sub

    Private Sub MaDataGrid_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)

        Dim ic As ItemCatia = MaDataGrid.SelectedItem
        For Each tv As TreeViewItem In MonTV.Items
            majSelectionfromDTtoTV(tv, ic)
        Next

        Exit Sub

        Dim l As New List(Of ItemCatia)
        For Each item As ItemCatia In MaDataGrid.SelectedItems
            l.Add(item)
        Next

        For i = 0 To MaDataGrid.Items.Count - 1
            Dim t As TextBlock = MaDataGrid.Columns(4).GetCellContent(MaDataGrid.Items(i))
            If Not t Is Nothing Then
                If l.Contains(MaDataGrid.Items(i)) Then
                    t.Foreground = New SolidColorBrush(Colors.White)
                Else
                    If t.Text = "Acheté" Then t.Foreground = New SolidColorBrush(Colors.IndianRed)
                    If t.Text = "Fabriqué" Then t.Foreground = New SolidColorBrush(Colors.DarkGreen)
                    If t.Text = "Inconnu" Then t.Foreground = New SolidColorBrush(Colors.DarkBlue)
                End If

            End If
            i = i + 1
        Next


    End Sub

    Sub majSelectionfromDTtoTV(tv As TreeViewItem, ic As ItemCatia)


        Dim ictv As ItemTV = tv.DataContext
        If ictv.ItemCATIA Is ic Then
            tv.IsSelected = True
        End If

        For Each _tv As TreeViewItem In tv.Items
            majSelectionfromDTtoTV(_tv, ic)
        Next

    End Sub

    Private Sub Window_SizeChanged(sender As Object, e As SizeChangedEventArgs)
        VerifFenetre()

    End Sub
    Sub VerifFenetre()
        If Me.Width < 1400 Then
            columnTV.Width = New GridLength(0)
        Else
            columnTV.Width = New GridLength(350)
        End If
        If Me.WindowState = WindowState.Maximized Then columnTV.Width = New GridLength(350)
    End Sub


    '------------------- CLOWN PDF ---------------------------
    Dim MonTWM As String = ""

    Private Sub But_Click(sender As Object, e As RoutedEventArgs)

        If ListViewPDF.SelectedIndex = 0 Then

            Dim OpenFileDialog1 As New Microsoft.Win32.OpenFileDialog With {
                .Title = "Sélétion du fichier PDF",
                .AddExtension = True,
                .Multiselect = False,
                .Filter = "Fichier PDF|*.pdf"
            }

            OpenFileDialog1.ShowDialog()

            If OpenFileDialog1.FileName <> "" Then
                Dim Name() As String = Strings.Split(OpenFileDialog1.FileName, "\")
                Dim NomFichier As String = Name(UBound(Name))

                FctionPDF.split(OpenFileDialog1.FileName, Strings.Left(NomFichier, Len(NomFichier) - 4))
            End If
        End If


        If ListViewPDF.SelectedIndex = 1 Then

            Dim OpenFileDialog1 As New Microsoft.Win32.OpenFileDialog
            OpenFileDialog1.Title = "Sélétion des fichiers PDF"
            OpenFileDialog1.AddExtension = True
            OpenFileDialog1.Multiselect = True
            OpenFileDialog1.Filter = "Fichier PDF|*.pdf"
            OpenFileDialog1.ShowDialog()

            Dim ListStrFil As New List(Of String)
            ListStrFil.Clear()

            For Each StrFil As String In OpenFileDialog1.FileNames
                ListStrFil.Add(StrFil)
            Next

            If ListStrFil.Count > 0 Then
                Dim Name() As String = Strings.Split(ListStrFil(0), "\")
                Dim NomFichier As String = Name(UBound(Name))

                FctionPDF.merge(ListStrFil, Strings.Left(NomFichier, Len(NomFichier) - 4))
            End If
        End If

        If ListViewPDF.SelectedIndex = 2 Then

            Dim OpenFileDialog1 As New Microsoft.Win32.OpenFileDialog
            OpenFileDialog1.Title = "Sélétion du fichier PDF"
            OpenFileDialog1.AddExtension = True
            OpenFileDialog1.Multiselect = False
            OpenFileDialog1.Filter = "Fichier PDF|*.pdf"

            OpenFileDialog1.ShowDialog()

            If OpenFileDialog1.FileName <> "" Then
                Dim Name() As String = Strings.Split(OpenFileDialog1.FileName, "\")
                Dim NomFichier As String = Name(UBound(Name))

                FctionPDF.Marker(OpenFileDialog1.FileName, Strings.Left(NomFichier, Len(NomFichier) - 4), MonTWM)
            End If
        End If
    End Sub

    Private Sub TextWM_TextChanged(sender As Object, e As TextChangedEventArgs)
        MonTWM = TextWM.Text
    End Sub

    Private Sub DragPanel_Drop(sender As Object, e As DragEventArgs)

        If e.Data.GetDataPresent(DataFormats.FileDrop) Then

            If ListViewPDF.SelectedIndex = 0 Then

                Dim f As String() = e.Data.GetData(DataFormats.FileDrop)

                Dim Name() As String = Strings.Split(f(0), "\")
                Dim NomFichier As String = Name(UBound(Name))

                If UCase(Right(NomFichier, 3)) = "PDF" Then

                    Split(f(0).ToString, Strings.Left(NomFichier, Len(NomFichier) - 4))

                End If

            End If

            If ListViewPDF.SelectedIndex = 1 Then

                Dim f As String() = e.Data.GetData(DataFormats.FileDrop)

                Dim ListStrFil As New List(Of String)
                ListStrFil.Clear()

                Dim MonNomFichierOK As String = ""
                For Each StrFil As String In f
                    Dim Name() As String = Strings.Split(f(0), "\")
                    Dim NomFichier As String = Name(UBound(Name))
                    If UCase(Right(NomFichier, 3)) = "PDF" Then ListStrFil.Add(StrFil)

                    If StrFil = f(0) Then MonNomFichierOK = NomFichier
                Next

                FctionPDF.merge(ListStrFil, Strings.Left(MonNomFichierOK, Len(MonNomFichierOK) - 4))

            End If

            If ListViewPDF.SelectedIndex = 2 Then

                Dim f As String() = e.Data.GetData(DataFormats.FileDrop)

                Dim Name() As String = Strings.Split(f(0), "\")
                Dim NomFichier As String = Name(UBound(Name))

                If UCase(Right(NomFichier, 3)) = "PDF" Then

                    FctionPDF.Marker(f(0).ToString, Strings.Left(NomFichier, Len(NomFichier) - 4), MonTWM)

                End If

            End If

        End If
    End Sub

    Private Sub ListViewPDF_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ListViewPDF.SelectionChanged

        If ListViewPDF.SelectedItems.Count = 0 Then ListViewPDF.SelectedIndex = 0

        Select Case ListViewPDF.SelectedIndex
            Case 0
                LabelPDF.Text = "A partir d'un fichier PDF de plusieurs pages,"
                Label2PDF.Text = "fractionner les pages en plusieurs fichiers"
                TextWM.Visibility = Visibility.Hidden
                Tex.Visibility = Visibility.Visible
                But.Visibility = Visibility.Visible
            Case 1
                LabelPDF.Text = "A partir de plusieurs fichiers PDF,"
                Label2PDF.Text = "fusionner les fichiers en un seul"
                TextWM.Visibility = Visibility.Hidden
                Tex.Visibility = Visibility.Visible
                But.Visibility = Visibility.Visible
            Case 2
                LabelPDF.Text = "A partir d'un fichier PDF,"
                Label2PDF.Text = "ajouter une notation filigrane sur l'ensemble des pages"
                TextWM.Visibility = Visibility.Visible
                Tex.Visibility = Visibility.Visible
                But.Visibility = Visibility.Visible
        End Select
    End Sub



    Private Sub ButtonSettingsCATIA_Click(sender As Object, e As RoutedEventArgs)

        My.Settings.PLANT = TextPlant.Text
        My.Settings.DRN = TextDessinateur.Text
        My.Settings.TITLE1 = TextTitre1.Text
        My.Settings.TITLE2 = TextTitre2.Text
        My.Settings.PROGRAM = TextProgram.Text
        My.Settings.CAGECODE = TextCageCode.Text
        My.Settings.Save()


        Dim m As New MessageErreur("Les options ont bien été sauvegardées", Notifications.Wpf.NotificationType.Information)
    End Sub

    Private Sub reportBUGbutton_Click(sender As Object, e As RoutedEventArgs)
        Process.Start("https://docs.google.com/spreadsheets/d/1IKvGSCUcCIcX1kBaWi9LIOPTso6A1NBqbHrjCPkLREU/edit#gid=0")
    End Sub

    Private Sub BanButton_Click(sender As Object, e As RoutedEventArgs)

        Dim u As ClassUsers = DataGridAdmin.SelectedItem
        Dim t As Integer = u.Ban
        If t = 0 Then
            t = 1
        Else
            t = 0
        End If
        u.Ban = t
        ColDocUsers.Refresh()

        cSQLUsers.BanUser(u.IPLocale, u.IPPublic, t)

    End Sub

    Private Sub ButtonSettingsG_Click(sender As Object, e As RoutedEventArgs)

        cSQLUsers.UpdateTableSettingsG(TextUpdateàJour.Text, TextBanAll.Text)


    End Sub
    Private Sub DragPanelDropStep_Drop(sender As Object, e As DragEventArgs)

        If e.Data.GetDataPresent(DataFormats.FileDrop) Then

            Dim f As String() = e.Data.GetData(DataFormats.FileDrop)

            Dim ListStrFil As New List(Of String)
            ListStrFil.Clear()

            For Each StrFil As String In f
                Dim Name() As String = Strings.Split(f(0), "\")
                Dim NomFichier As String = Name(UBound(Name))
                If UCase(Right(NomFichier, 7)) = "CATPART" Or UCase(Right(NomFichier, 10)) = "CATPRODUCT" Or UCase(Right(NomFichier, 3)) = "STP" Then ListStrFil.Add(StrFil)
            Next

            '  FctionPDF.merge(ListStrFil, Strings.Left(MonNomFichierOK, Len(MonNomFichierOK) - 4))
            CreateXMLFile(ListStrFil)

        End If


    End Sub

    Sub CreateXMLFile(l As List(Of String))

        Dim k As Integer = 1
        Dim FoldReception As String = Me.URLFolderSTEPReception.Text
        Dim FullPathCatia As String = Me.URLCatiaSTEP.Text

        If Directory.Exists(FoldReception) And Directory.Exists(FullPathCatia) Then
            Dim sw As New StreamWriter(FoldReception & "\genSTEP.xml")
            sw.WriteLine("<?xml version=""1.0"" encoding=""UTF-8""?>")
            sw.WriteLine("<!DOCTYPE root SYSTEM ""Parameters.dtd"">")
            sw.WriteLine("<root batch_name=""BatchDataExchange"">")
            sw.WriteLine("<inputParameters>")
            For Each i As String In l
                sw.WriteLine("<file id=""FileToProcess"" destination="""" filePath=""" & i & """ type=""bin"" upLoadable=""RightNow"" automatic=""1""/>")
            Next
            sw.WriteLine("</inputParameters>")
            sw.WriteLine("<outputParameters>")
            For Each i As String In l
                sw.WriteLine("<folder id=""OutputFolder"" destination=""" & FoldReception & """ folderPath=""" & FoldReception & """ type=""bin"" upLoadable=""RightNow"" extension=""*"" automatic=""1""/>")
                If UCase(Right(i, 3)) = "STP" Then
                    sw.WriteLine("<simple_arg id=""OutputExtension" & k & """ value=""CATPart/CATProduct""/>")
                Else
                    sw.WriteLine("<simple_arg id=""OutputExtension" & k & """ value=""stp""/>")
                End If
                k += 1
            Next
            sw.WriteLine("</outputParameters>")
            sw.WriteLine("<PCList>")
            sw.WriteLine("<PC name = ""HD2.slt"" />")
            sw.WriteLine("<PC name = ""ST1.prd"" />")
            sw.WriteLine("</PCList>")
            sw.WriteLine("</root>")
            sw.Close()


            Dim sw_ As New StreamWriter(FoldReception & "\genSTEP.bat")
            Dim guill As String = """"
            Dim str As String = guill & FullPathCatia & "\CATBatchStarter" & guill & " -input " & guill & FoldReception & "\genSTEP.xml" & guill
            If NameMachineDistanteSTEP.Text.Length > 0 Then
                str = str & " -driver " & guill & "BB" & guill & " -host " & guill & NameMachineDistanteSTEP.Text & guill
            End If
            sw_.WriteLine(str)
            sw_.WriteLine("del " & guill & FoldReception & "\genSTEP.xml""")
            sw_.WriteLine("del " & guill & FoldReception & "\genSTEP.bat""")
            sw_.Close()

            '   Process.Start(FoldReception & "\genSTEP.bat")

            Dim pi As New ProcessStartInfo
            pi.FileName = FoldReception & "\genSTEP.bat"
            pi.CreateNoWindow = True
            pi.WindowStyle = ProcessWindowStyle.Hidden
            Process.Start(pi)
            Process.Start(FoldReception)
            Dim m As New MessageErreur("La conversion des fichiers en arrière plan est en cours...", Notifications.Wpf.NotificationType.Information)
        Else
            Dim m As New MessageErreur("Vérifier l'existence des dossiers choisis", Notifications.Wpf.NotificationType.Error)
        End If

    End Sub

    Sub RunCommandCom(command As String, arguments As String, permanent As Boolean)
        Dim p As Process = New Process()
        Dim pi As ProcessStartInfo = New ProcessStartInfo()
        pi.Arguments = " " + If(permanent = True, "/K", "/C") + " " + command + " " + arguments
        pi.FileName = "cmd.exe"
        p.StartInfo = pi
        p.Start()
    End Sub

    Private Sub ButtonStep_Click(sender As Object, e As RoutedEventArgs)

        Dim OpenFileDialog1 As New Microsoft.Win32.OpenFileDialog
        OpenFileDialog1.Title = "Sélétion des fichiers à convertir en STEP"
        OpenFileDialog1.AddExtension = True
        OpenFileDialog1.Multiselect = True
        OpenFileDialog1.Filter = "Fichiers CATIA|*.CATPart;*.CATProduct;*.stp"
        OpenFileDialog1.ShowDialog()

        Dim ListStrFil As New List(Of String)
        ListStrFil.Clear()

        For Each StrFil As String In OpenFileDialog1.FileNames
            ListStrFil.Add(StrFil)
        Next

        If ListStrFil.Count > 0 Then CreateXMLFile(ListStrFil)

    End Sub

    Private Sub Hyperlink_RequestNavigate(sender As Object, e As RequestNavigateEventArgs)
        System.Diagnostics.Process.Start(e.Uri.AbsoluteUri)
    End Sub

    Private Sub ButtonDeketeUser_Click(sender As Object, e As RoutedEventArgs)
        Dim u As ClassUsers = DataGridAdmin.SelectedItem

        cSQLUsers.RemoveUser(u.IPLocale, u.IPPublic)

        ColDocUsers.Remove(u)
        ColDocUsers.Refresh()

    End Sub

    Private Sub Button_Click_11(sender As Object, e As RoutedEventArgs)

        Process.Start(DossierBase & "\Données\Environnements.ini")
    End Sub

    Private Sub Button_Click_12(sender As Object, e As RoutedEventArgs)

        ButtonCatiaTest.Visibility = Visibility.Visible
        LoadProgressBar.Visibility = Visibility.Collapsed
        LabelLoad.Visibility = Visibility.Collapsed

        Go()

        GridContentBDD.Visibility = Visibility.Collapsed
        GridContentCATPART.Visibility = Visibility.Collapsed
        GridContentCATPART.Visibility = Visibility.Collapsed

        GridLoaded.Visibility = Visibility.Visible

        BadgeReload.Badge = ""
        NeedReload = False
        VerifFenetre()

    End Sub

    Private Sub MaDataGrid_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs)

        MaDataGrid.IsReadOnly = False
        MaDataGrid.BeginEdit()



    End Sub

    Private Sub MaDataGrid_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Down Or e.Key = Key.Up Or e.Key = Key.Left Or e.Key = Key.Right Then
        Else

            MaDataGrid.IsReadOnly = False
            MaDataGrid.BeginEdit()
        End If


    End Sub

    Private Sub Chip_Click(sender As Object, e As RoutedEventArgs)



        Select Case sender.content
            Case "Environnement Clients"
                chip0.IconBackground = New SolidColorBrush(Color.FromRgb(118, 194, 175))
                chip1.IconBackground = New SolidColorBrush(Colors.Gray)
                chip2.IconBackground = New SolidColorBrush(Colors.Gray)
                chip3.IconBackground = New SolidColorBrush(Colors.Gray)
                TabOptions.SelectedIndex = 0
            Case "Bibliothèque TTS"
                chip0.IconBackground = New SolidColorBrush(Colors.Gray)
                chip1.IconBackground = New SolidColorBrush(Color.FromRgb(118, 194, 175))
                chip2.IconBackground = New SolidColorBrush(Colors.Gray)
                chip3.IconBackground = New SolidColorBrush(Colors.Gray)
                TabOptions.SelectedIndex = 1
            Case "Drawings et Cartouches"
                chip0.IconBackground = New SolidColorBrush(Colors.Gray)
                chip1.IconBackground = New SolidColorBrush(Colors.Gray)
                chip2.IconBackground = New SolidColorBrush(Color.FromRgb(118, 194, 175))
                chip3.IconBackground = New SolidColorBrush(Colors.Gray)
                TabOptions.SelectedIndex = 2
            Case "Général"
                chip0.IconBackground = New SolidColorBrush(Colors.Gray)
                chip1.IconBackground = New SolidColorBrush(Colors.Gray)
                chip2.IconBackground = New SolidColorBrush(Colors.Gray)
                chip3.IconBackground = New SolidColorBrush(Color.FromRgb(118, 194, 175))
                TabOptions.SelectedIndex = 3
        End Select
    End Sub

    Private Sub textNameUser_TextChanged(sender As Object, e As TextChangedEventArgs)
        labelNameUser.Text = textNameUser.Text
    End Sub
    Sub Getname()

        Dim name = System.Security.Principal.WindowsIdentity.GetCurrent().Name
        name = Strings.Replace(name, "EXCENT", "")
        name = Strings.Replace(name, "\", "")
        name = Strings.Replace(name, ".", " ")
        textNameUser.Text = name

    End Sub

    Function GetAs(s As String) As String
        If Strings.Left(s, 1) = "*" Then
            s = DossierBase & "\Données\" & Right(s, Len(s) - 1)
        Else
            s = s
        End If
        Return s
    End Function

    Private Sub TabOptions_Drop(sender As Object, e As DragEventArgs)

        If e.Data.GetDataPresent(DataFormats.FileDrop) Then

            Dim f As String() = e.Data.GetData(DataFormats.FileDrop)

            Dim Name() As String = Strings.Split(f(0), "\")
            Dim NomFichier As String = Name(UBound(Name))

            If UCase(Right(NomFichier, 3)) = "TXT" Then

                RemplirTableauEnvPerso(f(0))
                Dim sr As New StreamWriter(GetAs(INIProperties.GetString(GetEnv, "FichierProperties", "PersoProperties.txt")))

                Using reader As StreamReader = New StreamReader(f(0))
                    While Not reader.EndOfStream
                        Dim ligne As String = reader.ReadLine
                        sr.WriteLine(ligne)
                    End While
                    sr.Close()
                End Using

            End If


        End If

    End Sub

    Sub RemplirTableauEnvPerso(f As String)


        DataGridProperties.Items.Clear()

        If File.Exists(f) Then

            Using reader As StreamReader = New StreamReader(f)
                While Not reader.EndOfStream
                    Dim ligne As String = reader.ReadLine
                    Dim Splitligne() As String = Split(ligne, vbTab)

                    Dim i As New ItemProperties
                    i.Properties = Splitligne(0)
                    Try
                        i.Valeur = Splitligne(1)
                    Catch ex As Exception
                        i.Valeur = ""
                    End Try
                    Try
                        i.Type = Splitligne(2)
                    Catch ex As Exception
                        i.Type = "String"
                    End Try

                    DataGridProperties.Items.Add(i)

                End While
            End Using

        End If

        DataGridProperties.Items.Refresh()


    End Sub

    Private Sub Button_Click_13(sender As Object, e As RoutedEventArgs)
        Dim dlg As New OpenFileDialog()
        dlg.DefaultExt = ".txt"
        dlg.Filter = "Text documents (.txt)|*.txt"
        Dim result As System.Nullable(Of Boolean) = dlg.ShowDialog()
        If result = True Then
            RemplirTableauEnvPerso(dlg.FileName)
            Dim sr As New StreamWriter(GetAs(INIProperties.GetString(GetEnv, "FichierProperties", "PersoProperties.txt")))

            Using reader As StreamReader = New StreamReader(dlg.FileName)
                While Not reader.EndOfStream
                    Dim ligne As String = reader.ReadLine
                    sr.WriteLine(ligne)
                End While
                sr.Close()
            End Using
        End If
    End Sub

    Private Sub DataGridPropertiesPart_CellEditEnding(sender As Object, e As DataGridCellEditEndingEventArgs)
        Dim ic As PropertiesPart = e.EditingElement.DataContext
        Dim t As TextBox = e.EditingElement
        ic.Value = t.Text
        ic.MajProperties()


    End Sub

    Private Sub Button_Click_14(sender As Object, e As RoutedEventArgs)
        Process.Start(GetAs(INIProperties.GetString(GetEnv, "FichierProperties", "PersoProperties.txt")))
        Dim m As New MessageErreur("Il est conseillé de redemarrer l'application après modification du fichier", Notifications.Wpf.NotificationType.Information)
    End Sub

    Private Sub Button_Click_15(sender As Object, e As RoutedEventArgs)

        If GetEnv() = "DASSAULT AVIATION" Then
            FctionCATIA.CheckActiveDoc()
            If TypeActiveDoc = "DRAWING" Then
                FctionCATIA.MajCartoucheDassault()
            Else
                Dim m As New MessageErreur("Ouvrir un plan pour pouvoir utiliser cette fonction", Notifications.Wpf.NotificationType.Error)

            End If
        Else
            Dim m As New MessageErreur("Fonction non disponible pour les clients autre que DASSAULT", Notifications.Wpf.NotificationType.Error)
        End If
    End Sub

    Dim BoolHaveTORefresh As Boolean = False
    Private Sub MaDataGrid_CurrentCellChanged(sender As Object, e As EventArgs)
        If BoolHaveTORefresh = True Then ColDoc.Refresh()
        BoolHaveTORefresh = False
    End Sub
End Class

Public Class ItemProperties

    Public Property Properties As String
    Public Property Type As String
    Public Property Valeur As String



End Class





















