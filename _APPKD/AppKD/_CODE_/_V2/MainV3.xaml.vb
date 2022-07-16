
'couleur blanche : #F9F9F9
'couleur bleue : #0E1D31
'couleur rose : #EE4865
'couleur vert : #76C2AF

Imports System.ComponentModel
Imports System.IO
Imports System.Windows.Controls.Primitives
Imports KnowledgewareTypeLib
Imports Microsoft.Win32
Imports Microsoft.WindowsAPICodePack.Shell
Imports ProductStructureTypeLib

Public Class MainV3

    Dim bgwProgress As BackgroundWorker
    Dim NeedReload As Boolean = False
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)



        cSQL.ConnectionToDB()
        Getname()

        URLFolderSTEPReception.Text = Path.GetTempPath
        MonMainV3 = Me
        ResetonLoad()
        ListViewPDF.SelectedIndex = 0

        ListeEnvironements()
        CreerColonnes()
        CreerCheckBoxEnvironnementSettings()
        INIFiles = New GestionINIFiles(DossierBase & "\Request.ini")
        FichierTreeTxt = My.Computer.FileSystem.SpecialDirectories.Temp & "\CatiaTreeTxt.txt"
        TreeSauv = My.Computer.FileSystem.SpecialDirectories.Temp & "\CatiaTreeSauv.txt"


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

        'Exit sub
        '  FctionCATIA.test()


    End Sub 'LOAD

    Sub CreerColonnes()

        Dim i As Integer = 0
        For Each item In listPropertiesall
            Dim c As New DataGridTextColumn
            c.Header = item
            c.Binding = New Binding("l[" & i & "].Value")
            MaDataGrid.Columns.Add(c)
            i += 1
        Next
    End Sub
    Sub ResetonLoad()
        GridContentBDD.Visibility = Visibility.Collapsed
        GridChart.Visibility = Visibility.Collapsed
        GridBibliotheque.Visibility = Visibility.Collapsed
        GridAbout.Visibility = Visibility.Collapsed
        GridLoaded.Visibility = Visibility.Visible
        GridStep.Visibility = Visibility.Collapsed
        GridContentCATPART.Visibility = Visibility.Collapsed
    End Sub

    Sub CheckLeToogleClient(s As String)
        For Each g As ToggleButton In ListToogleSettingsEnv
            If g.Tag = s Then
                g.IsChecked = True
                Env = g.Tag
                ToolbarPerso.IsEnabled = False
                TabOptions.AllowDrop = False
                Dim f As String = GetAs(DossierBase & "\" & INIProperties.GetString(GetEnv, "FichierProperties", ""))
                RemplirTableauEnvPerso(f)
                labelClient.Text = GetEnv()
            End If
        Next
    End Sub
    Sub ListeEnvironements()
        DossierBase = DossierBase & "\Données"
        INIProperties = New GestionINIFiles(DossierBase & "\Environnements.ini")
        Using sr As New StreamReader(INIProperties.FileName)
            While Not sr.EndOfStream
                Dim l As String = sr.ReadLine
                If Strings.Left(l, 1) = "[" And Strings.Right(l, 1) = "]" Then
                    ListEnvironnements.Add(l)
                End If
                If Strings.Left(l, 17) = "FichierProperties" Then
                    Dim s() As String = Strings.Split(l, "=")
                    Using sr_ As New StreamReader(DossierBase & "\" & s(1))
                        While Not sr_.EndOfStream
                            Dim l_() As String = Strings.Split(sr_.ReadLine, vbTab)
                            Dim l__ As String = l_(0)
                            If listPropertiesall.Contains(l__) Then
                            Else
                                listPropertiesall.Add(l__)
                            End If
                        End While
                    End Using
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
                Dim f As String = DossierBase & "\" & INIProperties.GetString(n, "FichierProperties", "")
                If f = "" Then
                    DataGridProperties.Items.Clear()
                    DataGridProperties.Items.Refresh()
                Else
                    f = GetAs(f)
                    RemplirTableauEnvPerso(f)
                End If

                ToolbarPerso.IsEnabled = False
                TabOptions.AllowDrop = False


            Else
                g.IsChecked = False
            End If
            If GetEnv() <> s Then
                NeedReload = True
                MajVisuColonnesTableau()

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
            AddHandler s.MouseDown, AddressOf ToogleSettingsEnv_Click

        Next


        ' StackPanelSettingsEnv
    End Sub

    Private Sub Window_Closed(sender As Object, e As EventArgs)
        cSQL.CLoseConnexion()
        End
    End Sub 'CLOSE
    Private Sub Window_MouseMove(sender As Object, e As MouseEventArgs)

        Try
            FctionCATIA.GetCATIA()
        Catch ex As Exception
        End Try

        If CATIA Is Nothing Then
            LabelNameCatiaProduct.Content = "[Impossible de lier l'application avec CATIA]"
        Else
            Try
                LabelNameCatiaProduct.Content = "[" & CATIA.ActiveDocument.Name & "]"
                URLCatiaSTEP.Text = FctionCATIA.GetPathCATIA()
            Catch ex As Exception
                LabelNameCatiaProduct.Content = "[Aucun élément ne semble ouvert dans CATIA]"
                URLCatiaSTEP.Text = FctionCATIA.GetPathCATIA()
            End Try


            If ICRacine IsNot Nothing Then

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
            End If


        End If


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
        Dim m As New MessageErreur("Properties deleted successfully. It is preferable to restart a calculation of the application to update the table.", Notifications.Wpf.NotificationType.Information)
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
                        ic.l(getItemListProperties("STOCK SIZE")).Value = ic.ProductCATIA.UserRefProperties.Item(sm).ValueAsString
                    End If
                Else
                    Try
                        ic.ProductCATIA.UserRefProperties.Item(sm).ValuateFromString(MaMasse)
                    Catch ex As Exception
                        FctionCATIA.AddParamatres(ic.Owner, ic)
                        ic.ProductCATIA.UserRefProperties.Item(sm).ValuateFromString(MaMasse)
                    End Try
                    If Env = "[DASSAULT AVIATION]" Then
                        'ERREUR KEVIN ic.pers5 =>>>> chercher la variable ic correspondante à l'id colonne
                        ic.l(getItemListProperties("NomPuls_Masse")).Value = MaMasse
                    End If
                End If


            Next
        Catch ex As Exception
        End Try


        Dim m As New MessageErreur("Mass calculation completed", Notifications.Wpf.NotificationType.Information)

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

        s = "Loading ... | " & i & " %"
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
                ReloadSimple()

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
                Dim MsgErr As New MessageErreur("An error has occurred. Check that the file is saved", Notifications.Wpf.NotificationType.Error)
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
                Dim MsgErr As New MessageErreur("An error occurred while creating the plan", Notifications.Wpf.NotificationType.Error)
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
            ButtonSYM.Visibility = Visibility.Collapsed
            ButtonOBS.Visibility = Visibility.Collapsed
            ButtonBOM.Visibility = Visibility.Collapsed
            SeparatorToHide.Visibility = Visibility.Collapsed
            CBOM.Width = 220
            cPN.Width = 270
            cDesignation.Width = 330
        ElseIf GetEnv() = "AIRBUS" Then
            ButtonSYM.Visibility = Visibility.Visible
            ButtonOBS.Visibility = Visibility.Visible
            ButtonBOM.Visibility = Visibility.Visible
            SeparatorToHide.Visibility = Visibility.Visible
            CBOM.Width = 120
            cPN.Width = 150
            cDesignation.Width = 172
        ElseIf GetEnv() = "DASSAULT AVIATION" Then
            ButtonSYM.Visibility = Visibility.Collapsed
            ButtonOBS.Visibility = Visibility.Collapsed
            ButtonBOM.Visibility = Visibility.Collapsed
            SeparatorToHide.Visibility = Visibility.Collapsed
            CBOM.Width = 120
            cPN.Width = 150
            cDesignation.Width = 172
        End If
        Dim l As New List(Of String)
        l.Add("Qté")
        l.Add("Type")
        l.Add("Nomenclature")
        l.Add("PartNumber")
        l.Add("Source")
        l.Add("Désignation")
        l.Add("Rev")

        For Each c As DataGridColumn In MaDataGrid.Columns
            If l.Contains(c.Header) Then
                c.Visibility = Visibility.Visible
            Else
                c.Visibility = Visibility.Collapsed
            End If

        Next

        For Each item As ItemProperties In DataGridProperties.Items
            For Each c As DataGridColumn In MaDataGrid.Columns
                If c.Header = item.Properties Then
                    c.Visibility = GetVisu(item)
                End If
            Next
        Next

        DataGridProperties.Items.Refresh()
        DataGridProperties.Items.Refresh()

    End Sub 'MAJ SPIRIT
    Function GetVisu(item As ItemProperties) As Visibility

        If item.Visible = True Then
            Return Visibility.Visible
        Else
            Return Visibility.Hidden
        End If
    End Function

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
            'UserRefProperties

            Dim l_ As New List(Of itemCATIAProperties)
            For Each item As Parameter In ic.ProductCATIA.UserRefProperties
                Dim myP As New itemCATIAProperties(item.Name, item.ValueAsString)
                l_.Add(myP)
            Next
            For Each item In l_
                For j = 0 To listPropertiesall.Count - 1
                    If item.Name = listPropertiesall(j) Then
                        ic.l(j) = item
                        Exit For
                    End If
                Next
            Next
            Dim t As TextBox = e.EditingElement
            Dim p As Parameters = ic.ProductCATIA.UserRefProperties
            For Each item In ic.l
                If item Is Nothing Then
                    p.Item(e.Column.Header).value = t.Text
                Else
                    If item.Name = e.Column.Header Then
                        item.Value = t.Text
                        p.Item(e.Column.Header).value = item.Value
                    End If
                End If
            Next

            Select Case e.Column.Header
                Case "Nomenclature"
                    Dim te As TextBox = e.EditingElement
                    ic.Nomenclature = te.Text
                    ic.ProductCATIA.Nomenclature = ic.Nomenclature
                Case "Désignation"
                    Dim te As TextBox = e.EditingElement
                    ic.DescriptionRef = te.Text
                    ic.ProductCATIA.DescriptionRef = ic.DescriptionRef
                    If Env = "[DASSAULT AVIATION]" Then
                        ic.l(getItemListProperties("NomPuls_Designation")).Value = ic.DescriptionRef
                        p.Item("NomPuls_Designation").value = ic.DescriptionRef
                    End If
                Case "PartNumber"
                    'MA55800Z00-PIECE_500006
                    Dim te As TextBox = e.EditingElement
                    If ListPartNumber.Contains(te.Text) And ic.PartNumber.Length > 0 Then
                        te.Text = ic.PartNumber
                    Else
                        ic.PartNumber = te.Text
                        ic.ProductCATIA.PartNumber = ic.PartNumber
                        FctionCATIA.SaveIC(ic)
                        ListPartNumber.Add(ic.PartNumber)
                        If Env = "[DASSAULT AVIATION]" Then
                            If ic.PartNumber Like "??#####?##-*_######" Then
                                Dim ss() As String = Strings.Split(ic.PartNumber, "_")
                                Try
                                    p.Item(e.Column.Header).value = ss(1)
                                Catch ex As Exception
                                End Try
                                ic.l(getItemListProperties("NomPuls_Dlanche")).Value = ss(1)
                                BoolHaveTORefresh = True
                            End If
                        End If
                    End If

                Case "Rev"
                    ic.Revision = t.Text
                    ic.ProductCATIA.Revision = ic.Revision
                    If Env = "[DASSAULT AVIATION]" Then
                        ic.l(getItemListProperties("NomPuls_Indice")).Value = ic.Revision
                        p.Item("NomPuls_Indice").Value = ic.Revision
                    End If


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
        Dim c As String = DossierBase & "\" & INIProperties.GetString(GetEnv, "BibliothequeMaterial", "")
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
            Dim nms As New MessageErreur("Unable to generate EXCEL file from an empty table", Notifications.Wpf.NotificationType.Error)
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
                                            ic.l(getItemListProperties("SYM")).Value = ic_.PartNumber
                                            ic_.l(getItemListProperties("SYM")).Value = ic.PartNumber
                                            ListICOKSym.Add(ic)
                                            ListICOKSym.Add(ic_)
                                            Try
                                                ic.ProductCATIA.UserRefProperties.Item("SYM").ValuateFromString(ic.l(getItemListProperties("SYM")).Value)
                                            Catch ex As Exception
                                                FctionCATIA.AddParamatres(ic.Owner, ic)
                                                ic.ProductCATIA.UserRefProperties.Item("SYM").ValuateFromString(ic.l(getItemListProperties("SYM")).Value)
                                            End Try
                                            Try
                                                ic_.ProductCATIA.UserRefProperties.Item("SYM").ValuateFromString(ic_.l(getItemListProperties("SYM")).Value)
                                            Catch ex As Exception
                                                FctionCATIA.AddParamatres(ic_.Owner, ic_)
                                                ic_.ProductCATIA.UserRefProperties.Item("SYM").ValuateFromString(ic_.l(getItemListProperties("SYM")).Value)
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
                                            ic.l(getItemListProperties("SYM")).Value = ic_.PartNumber
                                            ic_.l(getItemListProperties("SYM")).Value = ic.PartNumber
                                            ListICOKSym.Add(ic)
                                            ListICOKSym.Add(ic_)
                                            Try
                                                ic.ProductCATIA.UserRefProperties.Item("SYM").ValuateFromString(ic.l(getItemListProperties("SYM")).Value)
                                            Catch ex As Exception
                                                FctionCATIA.AddParamatres(ic.Owner, ic)
                                                ic.ProductCATIA.UserRefProperties.Item("SYM").ValuateFromString(ic.l(getItemListProperties("SYM")).Value)
                                            End Try
                                            Try
                                                ic_.ProductCATIA.UserRefProperties.Item("SYM").ValuateFromString(ic_.l(getItemListProperties("SYM")).Value)
                                            Catch ex As Exception
                                                FctionCATIA.AddParamatres(ic_.Owner, ic_)
                                                ic_.ProductCATIA.UserRefProperties.Item("SYM").ValuateFromString(ic_.l(getItemListProperties("SYM")).Value)
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
            Dim m As New MessageErreur("Useless function in an environment other than AIRBUS", Notifications.Wpf.NotificationType.Warning)
        End If
    End Sub


    Private Sub Button_Click_7(sender As Object, e As RoutedEventArgs)
        If MsgBox("You are about to fix the open assembly. A few minutes may be necessary. Continue?", vbInformation + vbYesNo) = MsgBoxResult.No Then
            Exit Sub

        Else

            ProgressTableau.Visibility = Visibility.Visible
            GoProgress("Fix")

        End If


    End Sub

    Private Sub Button_Click_8(sender As Object, e As RoutedEventArgs)


        If Env = "[AIRBUS]" Then
            For Each ic As ItemCatia In ListDocuments
                If ic.l(getItemListProperties("REF")).Value <> "" Then
                    ic.l(getItemListProperties("OBSERVATIONS")).Value = ic.l(getItemListProperties("SUPPLIER")).Value & " '" & ic.l(getItemListProperties("REF")).Value
                ElseIf ic.l(getItemListProperties(INIProperties.GetString(GetEnv, "ProprieteTTS", ""))).Value <> "" Then
                    ic.l(getItemListProperties("OBSERVATIONS")).Value = ic.l(getItemListProperties(INIProperties.GetString(GetEnv, "ProprieteTTS", ""))).Value
                End If

                Try
                    ic.ProductCATIA.UserRefProperties.Item("OBSERVATIONS").ValuateFromString(ic.l(getItemListProperties("OBSERVATIONS")).Value)
                Catch ex As Exception
                    FctionCATIA.AddParamatres(ic.Owner, ic)
                    ic.ProductCATIA.UserRefProperties.Item("OBSERVATIONS").ValuateFromString(ic.l(getItemListProperties("OBSERVATIONS")).Value)
                End Try

            Next
        Else
            Dim m As New MessageErreur("Useless function for a sub-assembly outside the AIRBUS environment", Notifications.Wpf.NotificationType.Warning)
        End If

        RefreshTable()


    End Sub

    Private Sub Button_Click_9(sender As Object, e As RoutedEventArgs)


        If GridContentBDD.Visibility = Visibility.Visible Then
            If MsgBox("You are about to calculate the mass of all the parts. A few minutes may be necessary. Continue?", vbInformation + vbYesNo) = MsgBoxResult.No Then
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
                    ic.l(getItemListProperties("DETAIL NUMBER")).Value = Nber
                    Try
                        ic.ProductCATIA.UserRefProperties.Item("DETAIL NUMBER").ValuateFromString(ic.l(getItemListProperties("DETAIL NUMBER")).Value)
                    Catch ex As Exception
                        FctionCATIA.AddParamatres(ic.Owner, ic)
                        ic.ProductCATIA.UserRefProperties.Item("DETAIL NUMBER").ValuateFromString(ic.l(getItemListProperties("DETAIL NUMBER")).Value)
                    End Try

                End If
            Next

        Else
            Dim m As New MessageErreur("Inutile en dehors des environnements AIRBUS et SPIRIT", Notifications.Wpf.NotificationType.Error)
        End If

        RefreshTable()




    End Sub

    Sub RefreshTable()
        MaDataGrid.IsReadOnly = True
        MaDataGrid.Items.Refresh()
    End Sub

    Private Sub ComboBOM_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)

        If BoolStartBOM = True Then
            If ComboBOM.Items.Count > 0 Then FctionGetBOM.GoBOM(ComboBOM.SelectedValue.ToString)
            Try
                ColDoc.Filter = New Predicate(Of Object)(AddressOf FilterList)
            Catch ex As Exception
            End Try

            Try
                ColDoc.Refresh()
            Catch ex As Exception
            End Try
        End If
        BoolStartBOM = True
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
                ic.l(getItemListProperties(INIProperties.GetString(GetEnv, "ProprieteTTS", ""))).Value = t
                Dim p As Parameters = ic.ProductCATIA.UserRefProperties
                Try
                    p.Item(INIProperties.GetString(GetEnv, "ProprieteTTS", "TTS")).Value = t
                Catch ex As Exception
                    FctionCATIA.AddParamatres(ic.Owner, ic)
                    Try
                        p.Item(INIProperties.GetString(GetEnv, "ProprieteTTS", "TTS")).Value = t
                    Catch ex_ As Exception
                    End Try
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
                ic.l(getItemListProperties(INIProperties.GetString(GetEnv, "ProprieteMATERIAL", ""))).Value = t
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
                SelCATIA_Grid.IsEnabled = False
            Else
                ButtonNomenclature.IsEnabled = False
                SelCATIA_Grid.IsEnabled = True
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
                ReloadSimple()
            Catch ex As Exception
                Dim MsgErr As New MessageErreur("Unable to open requested file", Notifications.Wpf.NotificationType.Error)
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
            Dim str() As String = Strings.Split(ic.Doc.Name, ".CATProduct") 'kevin à corriger
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

        Try
            Dim pathFile As String = ic.FileName
            Dim s As ShellFile = ShellFile.FromFilePath(pathFile)
            Dim b As System.Drawing.Bitmap = s.Thumbnail.ExtraLargeBitmap
            '    b = modifCouleur(b)
            Dim bimg As New BitmapImage
            bimg = bitmapToImgSource(b)
            Me.imgSource.Source = bimg
        Catch ex As Exception
        End Try


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

    Function modifCouleur(img As System.Drawing.Bitmap) As System.Drawing.Bitmap



        With img
            For i As Integer = 0 To .Width - 1
                For j As Integer = 0 To .Height - 1
                    If img.GetPixel(i, j) = System.Drawing.Color.FromArgb(53, 51, 101) Then
                        .SetPixel(i, j, System.Drawing.Color.FromArgb(255, 255, 255))
                    End If
                Next
            Next
        End With

        Return img


    End Function
    Function bitmapToImgSource(img As System.Drawing.Bitmap) As BitmapImage


        Dim m As New MemoryStream
        img.Save(m, System.Drawing.Imaging.ImageFormat.Bmp)
        m.Position = 0
        Dim bimg As BitmapImage = New BitmapImage
        bimg.BeginInit()
        bimg.StreamSource = m
        bimg.CacheOption = BitmapCacheOption.OnLoad
        bimg.EndInit()

        Return bimg

    End Function

    Sub majSelectionfromDTtoTV(tv As TreeViewItem, ic As ItemCatia)


        Dim ictv As ItemTV = tv.DataContext
        If ictv.ItemCATIA Is ic Then
            tv.IsSelected = True
        End If

        For Each _tv As TreeViewItem In tv.Items
            majSelectionfromDTtoTV(_tv, ic)
        Next

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
                LabelPDF.Text = "From a PDF file of several pages,"
                Label2PDF.Text = "split pages into multiple files"
                TextWM.Visibility = Visibility.Hidden
                Tex.Visibility = Visibility.Visible
                But.Visibility = Visibility.Visible
            Case 1
                LabelPDF.Text = "From multiple PDF files,"
                Label2PDF.Text = "merge files into one"
                TextWM.Visibility = Visibility.Hidden
                Tex.Visibility = Visibility.Visible
                But.Visibility = Visibility.Visible
            Case 2
                LabelPDF.Text = "From a PDF file,"
                Label2PDF.Text = "add watermark notation on all pages"
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


        Dim m As New MessageErreur("The settings have been saved", Notifications.Wpf.NotificationType.Information)
    End Sub

    Private Sub reportBUGbutton_Click(sender As Object, e As RoutedEventArgs)
        Process.Start("https://www.catiavb.net/")
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
            sw.WriteLine("<PC name = ""MD2.slt"" />") 'HD2 ?
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
            Dim m As New MessageErreur("Background file conversion is in progress...", Notifications.Wpf.NotificationType.Information)
        Else
            Dim m As New MessageErreur("Check the existence of the selected folders", Notifications.Wpf.NotificationType.Error)
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


    Private Sub Button_Click_12(sender As Object, e As RoutedEventArgs)

        Reload()

    End Sub

    Sub ReloadSimple()

        Exit Sub

        ButtonCatiaTest.Visibility = Visibility.Collapsed

        PercentNbElements = 1
        NbElements = 0

        LoadProgressBar.Visibility = Visibility.Visible
        LabelLoad.Visibility = Visibility.Visible
        FctionLOAD.Main()

        GridContentBDD.Visibility = Visibility.Collapsed
        GridContentCATPART.Visibility = Visibility.Collapsed
        GridLoaded.Visibility = Visibility.Visible

        BadgeReload.Badge = ""
        NeedReload = False

    End Sub
    Sub Reload()
        ButtonCatiaTest.Visibility = Visibility.Visible
        LoadProgressBar.Visibility = Visibility.Collapsed
        LabelLoad.Visibility = Visibility.Collapsed

        Go()
        GridContentBDD.Visibility = Visibility.Collapsed
        GridContentCATPART.Visibility = Visibility.Collapsed
        GridLoaded.Visibility = Visibility.Visible

        BadgeReload.Badge = ""
        NeedReload = False
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
            Case "Environnements"
                chip0.IconBackground = New SolidColorBrush(Color.FromRgb(118, 194, 175))
                chip1.IconBackground = New SolidColorBrush(Colors.Gray)
                chip2.IconBackground = New SolidColorBrush(Colors.Gray)
                chip3.IconBackground = New SolidColorBrush(Colors.Gray)
                TabOptions.SelectedIndex = 0
            Case "TTS library"
                chip0.IconBackground = New SolidColorBrush(Colors.Gray)
                chip1.IconBackground = New SolidColorBrush(Color.FromRgb(118, 194, 175))
                chip2.IconBackground = New SolidColorBrush(Colors.Gray)
                chip3.IconBackground = New SolidColorBrush(Colors.Gray)
                TabOptions.SelectedIndex = 1
            Case "Drawings and title blocks"
                chip0.IconBackground = New SolidColorBrush(Colors.Gray)
                chip1.IconBackground = New SolidColorBrush(Colors.Gray)
                chip2.IconBackground = New SolidColorBrush(Color.FromRgb(118, 194, 175))
                chip3.IconBackground = New SolidColorBrush(Colors.Gray)
                TabOptions.SelectedIndex = 2
            Case "Settings"
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

        textNameUser.Text = "Kévin DESVOIS"

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
                Dim sr As New StreamWriter(GetAs(DossierBase & "\" & INIProperties.GetString(GetEnv, "FichierProperties", "PersoProperties.txt")))

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
                    i.Visible = True
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

        If GetEnv() = "AIRBUS" Then
            Dim i As Integer = 0
            For Each item As ItemProperties In DataGridProperties.Items
                Select Case i
                    Case 0
                        item.Visible = True 'Materila
                    Case 1
                        item.Visible = True 'Observation
                    Case 2
                        item.Visible = False 'Length
                    Case 3
                        item.Visible = False 'width
                    Case 4
                        item.Visible = False 'Diameter
                    Case 5
                        item.Visible = False 'Mass
                    Case 6
                        item.Visible = True 'Supplier
                    Case 7
                        item.Visible = True 'Ref
                    Case 8
                        item.Visible = True 'tts
                    Case 9
                        item.Visible = True 'sym
                End Select
                i += 1
            Next
        ElseIf GetEnv() = "SPIRIT AEROSYSTEMS" Then
            Dim i As Integer = 0
            For Each item As ItemProperties In DataGridProperties.Items
                Select Case i
                    Case 0
                        item.Visible = True 'Material
                    Case 1
                        item.Visible = True 'DETAIL NUMBER
                    Case 2
                        item.Visible = True 'STOCK SIZE
                End Select
                i += 1
            Next
        ElseIf GetEnv() = "DASSAULT AVIATION" Then
            Dim i As Integer = 0
            For Each item As ItemProperties In DataGridProperties.Items
                Select Case i
                    Case 0
                        item.Visible = False 'Marquage
                    Case 1
                        item.Visible = True 'traitement
                    Case 2
                        item.Visible = False 'protection
                    Case 3
                        item.Visible = True 'dim
                    Case 4
                        item.Visible = False 'mass
                    Case 5
                        item.Visible = True 'material
                    Case 6
                        item.Visible = False 'designation
                    Case 7
                        item.Visible = False 'indice
                    Case 8
                        item.Visible = True 'planche
                    Case 9
                        item.Visible = True 'num out
                End Select
                i += 1
            Next
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
            Dim sr As New StreamWriter(GetAs(DossierBase & "\" & INIProperties.GetString(GetEnv, "FichierProperties", "PersoProperties.txt")))

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
        Process.Start(GetAs(DossierBase & "\" & INIProperties.GetString(GetEnv, "FichierProperties", "PersoProperties.txt")))
        Dim m As New MessageErreur("It is recommended to restart the application after modifying the file", Notifications.Wpf.NotificationType.Information)
    End Sub

    Private Sub Button_Click_15(sender As Object, e As RoutedEventArgs)

        If GetEnv() = "DASSAULT AVIATION" Then
            FctionCATIA.CheckActiveDoc()
            If TypeActiveDoc = "DRAWING" Then
                FctionCATIA.MajCartoucheDassault()
            Else
                Dim m As New MessageErreur("Open a draw to be able to use this function", Notifications.Wpf.NotificationType.Error)

            End If
        Else
            Dim m As New MessageErreur("Function not available for customers other than DASSAULT", Notifications.Wpf.NotificationType.Error)
        End If
    End Sub

    Dim BoolHaveTORefresh As Boolean = False
    Private Sub MaDataGrid_CurrentCellChanged(sender As Object, e As EventArgs)
        If BoolHaveTORefresh = True Then ColDoc.Refresh()
        BoolHaveTORefresh = False
    End Sub

    Private Sub Button_Click_16(sender As Object, e As RoutedEventArgs)
        FctionRenameDassault.Rename("test")

    End Sub


    Private Sub DataGridProperties_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs)
        Dim i As ItemProperties = DataGridProperties.SelectedItem
        If i.Visible = True Then
            i.Visible = False
        Else
            i.Visible = True
        End If
        DataGridProperties.Items.Refresh()
        MajVisuColonnesTableau()
    End Sub

    Private Sub ButtonOpenDocClients_Click(sender As Object, e As RoutedEventArgs)
        Process.Start(ICRacine.Doc.Path)

    End Sub

    Private Sub ButtonHome_Click(sender As Object, e As RoutedEventArgs)


        GridContentBDD.Visibility = Visibility.Visible
        GridContentCATPART.Visibility = Visibility.Collapsed
        GridLoaded.Visibility = Visibility.Collapsed
        ColDoc = New ListCollectionView(ListDocuments)
        RefreshTable()



    End Sub

    Private Sub SettingsButtonGen_Click(sender As Object, e As RoutedEventArgs)
        Process.Start(DossierBase & "\Environnements.ini")
    End Sub
End Class

Public Class ItemProperties

    Public Property Properties As String
    Public Property Type As String
    Public Property Valeur As String
    Public Property Visible As Boolean



End Class





















