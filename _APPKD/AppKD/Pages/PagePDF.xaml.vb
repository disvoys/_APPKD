Imports org.pdfclown
Imports org.pdfclown.documents
Imports org.pdfclown.files
Imports org.pdfclown.tools
Imports org.pdfclown.objects
Imports System.IO
Imports File = org.pdfclown.files.File
Imports org.pdfclown.files.SerializationModeEnum
Imports org.pdfclown.documents.contents.composition
Imports org.pdfclown.documents.contents.fonts
Imports bytes = org.pdfclown.bytes
Imports org.pdfclown.documents.interaction
Imports org.pdfclown.documents.interchange.metadata
Imports org.pdfclown.documents.interaction.viewer
Imports files = org.pdfclown.files
Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports org.pdfclown.documents.contents.xObjects
Imports org.pdfclown.documents.contents
Imports org.pdfclown.documents.contents.colorSpaces




Class PagePDF

    Dim MonTWM As String = ""

    Private Sub Page_Loaded(sender As Object, e As RoutedEventArgs)

        ListViewPDF.SelectedIndex = 0







    End Sub


    Sub split(Chemin As String, NameFile As String)

        Using f As New File(Chemin)
            Dim oldD As Document = f.Document
            If Not Directory.Exists(Path.GetTempPath & "\PDF_Maniement_AppKD") Then MkDir(Path.GetTempPath & "\PDF_Maniement_AppKD")

            For i As Integer = 0 To oldD.Pages.Count - 1
                Dim D As Document = New PageManager(oldD).Extract(i, i + 1)
                D.File.Save(Path.GetTempPath & "\PDF_Maniement_AppKD\" & NameFile & "_" & i & ".pdf", SerializationModeEnum.Standard)
            Next i

        End Using


        Dim k As New MessageErreur("Split du fichier réussi avec succès", Notifications.Wpf.NotificationType.Information)
        Process.Start(Path.GetTempPath & "\PDF_Maniement_AppKD")

    End Sub
    Sub merge(Chemin As List(Of String), NameFile As String)



        Using f As New File(Chemin(0))
            Dim D As Document = f.Document
            If Not Directory.Exists(Path.GetTempPath & "\PDF_Maniement_AppKD") Then MkDir(Path.GetTempPath & "\PDF_Maniement_AppKD")
            Dim Pm As New PageManager(D)
            For i = 1 To Chemin.Count - 1
                Dim fi As New File(Chemin(i))
                Pm.Add(fi.Document)
            Next

            D.File.Save(Path.GetTempPath & "\PDF_Maniement_AppKD\" & NameFile & "_" & ".pdf", SerializationModeEnum.Standard)
        End Using


        Dim k As New MessageErreur("Fusion des fichiers réussie avec succès", Notifications.Wpf.NotificationType.Information)
        Process.Start(Path.GetTempPath & "\PDF_Maniement_AppKD")

    End Sub

    Sub Marker(Chemin As String, NameFile As String, TextAffiche As String)

        Using f As New File(Chemin)
            Dim d As Document = f.Document
            If Not Directory.Exists(Path.GetTempPath & "\PDF_Maniement_AppKD") Then MkDir(Path.GetTempPath & "\PDF_Maniement_AppKD")

            Dim Stamp As New PageStamper
            Dim Size As SizeF = d.GetSize
            Dim WaterMark As New FormXObject(d, Size)
            Dim compo As New PrimitiveComposer(WaterMark)
            compo.SetFont(New StandardType1Font(d, StandardType1Font.FamilyEnum.Helvetica, False, False), 120)
            compo.SetFillColor(New DeviceRGBColor(157 / 255D, 27 / 255D, 48 / 255D))
            Dim state As New ExtGState(d)
            state.FillAlpha = 0.3
            compo.ApplyState(state)

            compo.ShowText(TextAffiche, New PointF(Size.Width / 2.0F, Size.Height / 2.0F), XAlignmentEnum.Center, YAlignmentEnum.Middle, 50)
            compo.Flush()



            For i As Integer = 0 To d.Pages.Count - 1
                Stamp.Page = d.Pages(i)
                Dim Composer As PrimitiveComposer = Stamp.Foreground
                Composer.ShowXObject(WaterMark)
                Stamp.Flush()

            Next i


            d.File.Save(Path.GetTempPath & "\PDF_Maniement_AppKD\" & NameFile & "_" & ".pdf", SerializationModeEnum.Standard)

        End Using

        Dim k As New MessageErreur("Ajout du marker avec succès", Notifications.Wpf.NotificationType.Information)
        Process.Start(Path.GetTempPath & "\PDF_Maniement_AppKD")

    End Sub


    Private Sub DragPanel_Drop(sender As Object, e As DragEventArgs)

        If e.Data.GetDataPresent(DataFormats.FileDrop) Then

            If ListViewPDF.SelectedIndex = 0 Then

                Dim f As String() = e.Data.GetData(DataFormats.FileDrop)

                Dim Name() As String = Strings.Split(f(0), "\")
                Dim NomFichier As String = Name(UBound(Name))

                If UCase(Right(NomFichier, 3)) = "PDF" Then

                    split(f(0).ToString, Left(NomFichier, Len(NomFichier) - 4))

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

                merge(ListStrFil, Left(MonNomFichierOK, Len(MonNomFichierOK) - 4))

            End If

            If ListViewPDF.SelectedIndex = 2 Then

                Dim f As String() = e.Data.GetData(DataFormats.FileDrop)

                Dim Name() As String = Strings.Split(f(0), "\")
                Dim NomFichier As String = Name(UBound(Name))

                If UCase(Right(NomFichier, 3)) = "PDF" Then

                    Marker(f(0).ToString, Left(NomFichier, Len(NomFichier) - 4), MonTWM)

                End If

            End If

        End If
    End Sub

    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)

        If ListViewPDF.SelectedIndex = 0 Then

            Dim OpenFileDialog1 As New Microsoft.Win32.OpenFileDialog
            OpenFileDialog1.Title = "Sélétion du fichier PDF"
            OpenFileDialog1.AddExtension = True
            OpenFileDialog1.Multiselect = False
            OpenFileDialog1.Filter = "Fichier PDF|*.pdf"

            OpenFileDialog1.ShowDialog()

            If OpenFileDialog1.FileName <> "" Then
                Dim Name() As String = Strings.Split(OpenFileDialog1.FileName, "\")
                Dim NomFichier As String = Name(UBound(Name))

                split(OpenFileDialog1.FileName, Left(NomFichier, Len(NomFichier) - 4))
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

            If ListStrFil(0) <> "" Then
                Dim Name() As String = Strings.Split(ListStrFil(0), "\")
                Dim NomFichier As String = Name(UBound(Name))

                merge(ListStrFil, Left(NomFichier, Len(NomFichier) - 4))
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

                Marker(OpenFileDialog1.FileName, Left(NomFichier, Len(NomFichier) - 4), MonTWM)
            End If
        End If
    End Sub

    Private Sub ListViewPDF_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ListViewPDF.SelectionChanged

        Select Case ListViewPDF.SelectedIndex
            Case 0
                LabelPDF.Text = "A partir d'un fichier PDF de plusieurs pages,"
                Label2PDF.Text = "fractionner les pages en plusieurs fichiers"
                TitlePage.Text = "SPLIT"
                TitlePage.Width = 20
                TextWM.Visibility = Visibility.Hidden
            Case 1
                LabelPDF.Text = "A partir de plusieurs fichiers PDF,"
                Label2PDF.Text = "fusionner les fichiers en un seul"
                TitlePage.Text = "MERGE"
                TitlePage.Width = 23
                TextWM.Visibility = Visibility.Hidden
            Case 2
                LabelPDF.Text = "A partir d'un fichier PDF,"
                Label2PDF.Text = "ajouter une notation filigrane sur l'ensemble des pages"
                TitlePage.Text = "MARKER"
                TitlePage.Width = 23
                TextWM.Visibility = Visibility.Visible
            Case 4
                LabelPDF.Text = "A partir d'un fichier PDF,"
                Label2PDF.Text = "obtenir un fichier similaire mais à taille réduite"
                TitlePage.Text = "COMPRESSER"
                TitlePage.Width = 23
                TextWM.Visibility = Visibility.Hidden
        End Select
    End Sub

    Private Sub TextWM_TextChanged(sender As Object, e As TextChangedEventArgs)
        MonTWM = TextWM.Text
    End Sub
End Class

